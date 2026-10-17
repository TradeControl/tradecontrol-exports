# Requirements

## Scope
- Provide ERP-driven export services callable from ASP.NET endpoints.
- Supported Cash Statement formats: Office Excel (XLSX) and OpenDocument Spreadsheet (ODS).
- Other formats such as CSV and PDF are optional future work.
- Internal consumption only: this library is vendored into the ASP.NET repo (no NuGet distribution).

## Functional Requirements
- Accept a database connection string, an export subject (e.g., “Cash Statement”), and a parameter object.
- Generate a document and return it as a downloadable file stream to the web client.
- Log and surface errors to the caller with meaningful messages.

## Non-Functional Requirements
- Works with the ASP.NET application on .NET 9.
- Clear separation between the document-handler contract (`IDocumentHandler`) and concrete implementations.
- Testability: unit-level for data shaping; integration-level for document generation.
- Packaging: .NET 9 class library consumed via ProjectReference.
- Cash Statement generation is in-process C# and must not require Python, ODBC, office automation, or server-side LibreOffice.

## Out of Scope (for now)
- CSV and PDF Cash Statement renderers.
- Public NuGet packaging and distribution.


## Configuration
- Connection strings via ASP.NET configuration.
- A feature flag can disable ODS independently without changing the XLSX path.

## Acceptance
- XLSX and ODS exports function end-to-end through the web app using a ProjectReference to the vendored library.
- Both renderers consume the same accepted immutable statement model.
- ODS can be disabled without impacting XLSX.
- Published application output contains no Python source or Python/ODBC/`odfdo`/`lxml` runtime dependency.
- Documentation reflects internal-only usage and deployment guidance.
