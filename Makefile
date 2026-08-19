.PHONY: up down logs install dev build lint typecheck test db-generate db-migrate db-seed clean

up:
	docker compose up -d

down:
	docker compose down

logs:
	docker compose logs -f

install:
	pnpm install

dev:
	pnpm dev

build:
	pnpm build

lint:
	pnpm lint

typecheck:
	pnpm typecheck

test:
	pnpm test

db-generate:
	pnpm db:generate

db-migrate:
	pnpm db:migrate

db-seed:
	pnpm db:seed

clean:
	docker compose down -v
