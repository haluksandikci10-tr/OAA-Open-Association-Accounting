# OAA v2 — Open Association Accounting

OAA is an association management and accounting application developed with Microsoft Excel VBA.

## Platform
- Microsoft Excel 2010 and later are the compatibility target.
- Macro-enabled Excel workbook (`.xlsm`).
- VBA.
- Excel 365 is not a requirement; compatibility must be verified in supported Excel versions.

## Architecture
`UserForms -> Service -> Repository -> Excel ListObject tables`

The existing layered architecture is preserved. UserForms do not implement business rules or access worksheet data directly. Services own business rules; repositories own Excel table access.

## Current scope
- Association and settings
- Members
- Member charges and dues rates
- Income and expenses
- Receipts and collections
- Cash/bank ledger

Roadmap items are not necessarily implemented features.

## Source layout
- `src/classes` — VBA classes
- `src/forms` — `.frm` and matching `.frx` files
- `src/modules` — standard modules
- `tests` — test organization
- `docs` — approved project documents
- `application` — development workbooks and backups

The current test source is `src/modules/modTests.bas`.

## Test procedure
Compile the VBA project before running tests. Run tests in order; fix a failing test before proceeding. Clean up temporary test records. The latest user-confirmed passing test is TEST-144; TEST-145 is next.

## Status
OAA v2 is under active development and is not yet a stable production release. Development workbooks may contain test data and must be cleaned and reviewed before public distribution.
