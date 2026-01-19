# Requirements

## Scope
- Provide ERP-driven export services callable from ASP.NET endpoints.
- Primary target format: Office Excel (XLSX).
- Secondary formats (ODS, CSV, PDF) are optional and may be introduced later.
- Internal consumption only: this library is vendored into the ASP.NET repo (no NuGet distribution).

## Functional Requirements
- Accept a database connection string, an export subject (e.g., “Cash Statement”), and a parameter object.
- Generate a document and return it as a downloadable file stream to the web client.
- Log and surface errors to the caller with meaningful messages.

## Non-Functional Requirements
- Works with ASP.NET (.NET 8).
- Clear separation between exporter contract (`IExporter`) and concrete implementations.
- Testability: unit-level for data shaping; integration-level for document generation.
- Packaging: .NET 8 class library consumed via ProjectReference.

## Out of Scope (for now)
- Libre ODS Cash Statement formatting and load-time value enforcement.
- Row-level styling and complex post-processing of ODS content.
- Public NuGet packaging and distribution.


## Configuration
- Connection strings via ASP.NET configuration.
- Feature flags to enable/disable experimental exporters (disable ODS in production).

## Acceptance
- XLSX exports function end-to-end through the web app using a ProjectReference to the vendored library.
- ODS path is disabled without impacting the rest of the system.
- Documentation reflects internal-only usage and deployment guidance.