# Label template fixtures

No physical Word (.docx) Avery template files are committed yet. The
`avery-5155` `LabelTemplate` preset (seeded by
`packages/database/prisma/seed.ts`) currently ships with placeholder
geometry only - see the comments in that seed file.

TODO: once a verified physical Avery 5155 template is measured and
confirmed to print correctly from Microsoft Word, add the source `.dotx`/
`.docx` template here, wire `templateStorageKey` on the seeded
`LabelTemplate` to reference it, and replace every placeholder value in
`configJson.geometry` with the verified measurements.
