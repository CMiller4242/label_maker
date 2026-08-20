import { useCallback, useEffect, useState } from "react";
import type { Product } from "@label-maker/shared";
import { ApiError, NetworkError } from "../api/client";
import { createProduct, listProducts, updateProduct } from "../api/endpoints";

export type ProductLoadStatus = "idle" | "loading" | "loaded" | "failed";

/** A row not yet persisted to the server (added client-side via "Add row"). */
export interface DraftRow {
  tempId: string;
  isNew: true;
  sku: string;
  description: string;
  priceCents: number | null;
  include: true;
}

/** A persisted Product, plus per-row save bookkeeping the review table needs. */
export interface PersistedRow extends Product {
  isNew: false;
}

export type ProductRow = PersistedRow | DraftRow;

export interface UseProductsResult {
  status: ProductLoadStatus;
  errorMessage: string | null;
  rows: ProductRow[];
  rowErrors: Record<string, string>;
  savingRowIds: Set<string>;
  refetch: () => void;
  addRow: () => void;
  removeRow: (rowId: string) => void;
  /** Commits an edited field for one row: updates local state immediately, persists via PATCH (existing rows) or POST (once a draft row has every required field). */
  editRow: (rowId: string, patch: Partial<Pick<Product, "sku" | "description" | "priceCents">>) => void;
  approveRow: (rowId: string) => void;
  setInclude: (rowId: string, include: boolean) => void;
}

function describeError(error: unknown): string {
  if (error instanceof ApiError) return error.message;
  if (error instanceof NetworkError) return error.message;
  if (error instanceof Error) return error.message;
  return "Something went wrong.";
}

function rowId(row: ProductRow): string {
  return row.isNew ? row.tempId : row.id;
}

function makeDraftRow(): DraftRow {
  return {
    tempId: `draft-${crypto.randomUUID()}`,
    isNew: true,
    sku: "",
    description: "",
    priceCents: null,
    include: true,
  };
}

export function useProducts(documentId: string | null): UseProductsResult {
  const [status, setStatus] = useState<ProductLoadStatus>("idle");
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [rows, setRows] = useState<ProductRow[]>([]);
  const [rowErrors, setRowErrors] = useState<Record<string, string>>({});
  const [savingRowIds, setSavingRowIds] = useState<Set<string>>(new Set());

  const load = useCallback(() => {
    if (!documentId) return;
    setStatus("loading");
    setErrorMessage(null);
    void (async () => {
      try {
        const result = await listProducts(documentId);
        setRows(result.products.map((p): PersistedRow => ({ ...p, isNew: false })));
        setStatus("loaded");
      } catch (error) {
        setStatus("failed");
        setErrorMessage(describeError(error));
      }
    })();
  }, [documentId]);

  useEffect(load, [load]);

  const withSaving = useCallback(<T,>(id: string, fn: () => Promise<T>): Promise<T> => {
    setSavingRowIds((prev) => new Set(prev).add(id));
    setRowErrors((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    return fn().finally(() => {
      setSavingRowIds((prev) => {
        const next = new Set(prev);
        next.delete(id);
        return next;
      });
    });
  }, []);

  const addRow = useCallback(() => {
    setRows((prev) => [...prev, makeDraftRow()]);
  }, []);

  const removeRow = useCallback((id: string) => {
    setRows((prev) => {
      const target = prev.find((r) => rowId(r) === id);
      if (!target) return prev;
      if (target.isNew) {
        // Never persisted - just drop it, nothing to undo server-side.
        return prev.filter((r) => rowId(r) !== id);
      }
      return prev; // persisted rows are "removed" via setInclude(false) below, not deletion.
    });
    const target = rows.find((r) => rowId(r) === id);
    if (target && !target.isNew) {
      void withSaving(id, async () => {
        try {
          const updated = await updateProduct(target.id, { include: false });
          setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? { ...updated, isNew: false } : r)));
        } catch (error) {
          setRowErrors((prev) => ({ ...prev, [id]: describeError(error) }));
        }
      });
    }
  }, [rows, withSaving]);

  const setInclude = useCallback(
    (id: string, include: boolean) => {
      const target = rows.find((r) => rowId(r) === id);
      if (!target) return;
      if (target.isNew) {
        setRows((prev) =>
          prev.map((r) => (r.isNew && r.tempId === id ? { ...r, include: true } : r)),
        );
        return; // drafts are always include:true until persisted
      }
      setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? { ...r, include } : r)));
      void withSaving(id, async () => {
        try {
          const updated = await updateProduct(target.id, { include });
          setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? { ...updated, isNew: false } : r)));
        } catch (error) {
          setRowErrors((prev) => ({ ...prev, [id]: describeError(error) }));
          setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? target : r))); // revert
        }
      });
    },
    [rows, withSaving],
  );

  const approveRow = useCallback(
    (id: string) => {
      const target = rows.find((r) => rowId(r) === id);
      if (!target || target.isNew) return;
      void withSaving(id, async () => {
        try {
          const updated = await updateProduct(target.id, { status: "APPROVED" });
          setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? { ...updated, isNew: false } : r)));
        } catch (error) {
          setRowErrors((prev) => ({ ...prev, [id]: describeError(error) }));
        }
      });
    },
    [rows, withSaving],
  );

  const editRow = useCallback(
    (id: string, patch: Partial<Pick<Product, "sku" | "description" | "priceCents">>) => {
      setRows((prev) =>
        prev.map((r) => (rowId(r) === id ? ({ ...r, ...patch } as ProductRow) : r)),
      );

      const current = rows.find((r) => rowId(r) === id);
      if (!current) return;

      if (current.isNew) {
        const nextSku = patch.sku ?? current.sku;
        const nextDescription = patch.description ?? current.description;
        const nextPriceCents = patch.priceCents ?? current.priceCents;
        const ready = nextSku.trim().length > 0 && nextDescription.trim().length > 0 && nextPriceCents !== null;
        if (!ready) return; // keep editing locally until every required field is present
        void withSaving(id, async () => {
          try {
            const created = await createProduct(documentId as string, {
              sku: nextSku,
              description: nextDescription,
              priceCents: nextPriceCents,
              include: true,
            });
            setRows((prev) => prev.map((r) => (r.isNew && r.tempId === id ? { ...created, isNew: false } : r)));
          } catch (error) {
            setRowErrors((prev) => ({ ...prev, [id]: describeError(error) }));
          }
        });
        return;
      }

      const next = { ...current, ...patch };
      // Editing content on a NEEDS_REVIEW/EXCLUDED row is treated as the
      // human reviewing it - approve it in the same request, so it doesn't
      // silently stay ineligible for generation after being corrected.
      const shouldApprove = next.status !== "APPROVED" && next.status !== "AUTO_ACCEPTED";
      void withSaving(id, async () => {
        try {
          const updated = await updateProduct(next.id, {
            ...patch,
            ...(shouldApprove ? { status: "APPROVED" as const } : {}),
          });
          setRows((prev) => prev.map((r) => (!r.isNew && r.id === id ? { ...updated, isNew: false } : r)));
        } catch (error) {
          setRowErrors((prev) => ({ ...prev, [id]: describeError(error) }));
        }
      });
    },
    [rows, documentId, withSaving],
  );

  return {
    status,
    errorMessage,
    rows,
    rowErrors,
    savingRowIds,
    refetch: load,
    addRow,
    removeRow,
    editRow,
    approveRow,
    setInclude,
  };
}
