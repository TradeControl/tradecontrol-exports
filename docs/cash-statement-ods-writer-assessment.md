# Cash Statement ODS writer assessment

**Date:** 16 October 2026  
**Decision:** accepted — retain the Phase 0A internal C# writer and proceed to the full Phase 6 renderer  
**Owner acceptance:** Accepted 16 October 2026

## Scope

Phase 5 turns the Phase 0A spike into a production-shaped capability without implementing the Cash Statement layout. The package writer consumes physical worksheets whose cells carry the existing renderer-neutral `StatementValue`, `StatementNumericKind`, `StatementPresentationRole`, and `StatementFormula` types. It does not query a database, calculate financial values, choose rows, or interpret request options.

The full row and column allocation remains a Phase 6 responsibility. This separation lets the hostile capability fixture exercise the same model primitives that the production renderer will use without hiding report-specific logic inside the low-level ODS package mechanism.

## Dependency and licence decision

The writer adds no NuGet package, Python module, native library, office automation API, or server-side LibreOffice dependency. Package creation uses the .NET runtime's `System.IO.Compression`, `System.Xml.Linq`, streams, immutable collections, and invariant formatting support already present in the application runtime.

There is therefore no additional third-party ODS licence, package publisher, release cadence, transitive dependency, or software-supply-chain source to approve. LibreOffice remains an acceptance-test application and supported client, not a deployed generator dependency.

## Maintenance decision

The maintained surface is deliberately narrow:

- immutable physical ODS worksheet, row, column, and cell records;
- semantic-formula to OpenFormula translation;
- six deterministic package parts;
- native styles for the capability set;
- view settings, including hidden columns and frozen panes; and
- admission limits and caller-owned stream handling.

XML is produced through `XDocument`; financial values are formatted with the invariant culture. ZIP entry order and timestamps are deterministic. The implementation avoids a general office-document object model and exposes only features needed by the Cash Statement.

## Security and privacy review

- No database connection, credential, tenant identifier, temporary filename, or report request is accepted by the package writer.
- XML text and attribute values are escaped by the XML library.
- The generated package contains no macro, script, embedded object, external link, or remote resource.
- Formula references can target only semantic cells allocated by the caller. Missing allocations fail generation.
- Management expressions remain intentionally spreadsheet-engine expressions. Their category tokens are bound before serialization; the writer does not evaluate their language or treat them as trusted financial calculations.
- Output is first created in a bounded private memory buffer and copied to the destination only after the package closes successfully. A failed or cancelled write therefore does not expose a partial archive through the caller-owned destination.

## Limits, streaming, cancellation, and disposal

Default writer limits are three worksheets, 1,024 physical columns per worksheet, 10,000 rows, 2,000,000 cells, and 64 MiB of generated package bytes. These sit behind the shared-model admission checks and provide a second renderer boundary.

The implementation deliberately buffers one bounded package before copying it to the response stream. This is not constant-memory streaming, but it gives deterministic failure semantics and places a hard upper bound on writer-owned memory. Direct streaming can be reconsidered only if measured Azure workloads require it; it must not remove the byte limit or expose partial ZIP files.

Cancellation is checked before package construction, while worksheets and rows are serialized, and during the final asynchronous copy. ZIP entries, writers, buffers, and archives are deterministically disposed. The destination stream remains owned and usable by the caller.

## Verification result

The permanent hostile fixture uses real shared-model values and formula nodes for literal formulas, direct and forward references, aggregates, arithmetic, guarded division, a terminal management expression, a three-level dependency chain, negative values, explicit zero, missing values, a non-January period layout, totals, percentage formatting, a hidden marker column, and frozen panes.

Automated tests verify:

- `mimetype` is the first, stored ZIP entry with no extra field;
- exact generated entry order and required manifest relationships;
- ODF and OpenFormula namespaces;
- authoritative cached values on every formula;
- deterministic byte-for-byte output;
- rejection of missing parts, bad manifest relationships, wrong entry order, wrong namespaces, and missing cached values;
- row, column, cell, and output-byte limits;
- pre-cancelled requests;
- destination stream ownership; and
- failure on an unallocated semantic dependency.

LibreOffice 26.2.6.3 recalculated and saved the fixture, and a second open/save preserved all 20 formula-cell cached values with no error tokens. The PDF export is a readable three-page A4 representation with the expected negative-number, total-fill, header, annual-border, and percentage presentation.

During this native test, LibreOffice exposed a missing `xmlns:of` declaration in the first integrated package. The permanent package-level test now protects the corrected OpenFormula namespace. This is useful evidence that Phase 5 is testing the integration rather than simply copying the successful spike.

## Remaining boundaries

- The fixture proves the mechanism and formula adapter, not the complete Cash Statement layout.
- Full cross-section row allocation, complete native styling, protection, print ranges, and all four business comparisons belong to Phase 6.
- Multi-tenant concurrency, cancellation during database acquisition/model construction, telemetry, Azure resource behaviour, and feature flags belong to Phase 7.
- The existing XLSX handler remains unaffected.

The selected internal C# writer is accepted for Phase 6. No evidenced capability gap justifies Python or an external C# ODS package.
