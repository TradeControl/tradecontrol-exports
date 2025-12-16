# Requirements for tradecontrol-exports

This document captures the functional and technical requirements for the `tradecontrol-exports` project.  
It will evolve as design decisions are clarified.

---

## 1. Inputs
- Test API ASP.NET endpoint: single generic `ExportData()` function.
- Payload: JSON body including Identity (injected by ASP.NET), connection string, source document, and params.
- Identity: must be authenticated with permitted roles before payload is constructed.
- JSON schema: enforced at the controller boundary to validate payload structure.

---

## 2. Data Retrieval
- Provider abstraction: `IDataProvider` interface.
- Current provider: SQL Server implementation.
- Future cut-over: PostgreSQL implementation will replace SQL Server (no dual support required).
- Connection strings: passed securely, never logged. Identity authentication required to proceed.

---

## 3. Output Formats
- Supported: Office (XLSX), Libre (ODS).
- Fallback: PDF if ODS/XLSX are not available.
- Format parity: documents must look the same to the user, but internal formatting may differ per target.

---

## 4. File Delivery
- Delivery method: stream file back to browser (download prompt).
- No server-side “Downloads” folder.
- Filename pattern: `companyName_documentName_timestamp.ext`.
- Filenames must be regex-sanitized to ensure validity and safety.

---

## 5. Extensibility
- Pattern: pluggable exporters behind `IExporter` (e.g., `OfficeExporter`, `LibreExporter`, `PdfExporter`).
- Dispatcher: single entry point routes `documentType` to the correct exporter.

---

## 6. Security
- Authentication: handled by ASP.NET Identity middleware before controller executes.
- Authorization: enforced at controller via roles/policies.
- Payload identity: injected by server from authenticated `User` claims, never accepted directly from client.
- Connection strings: handled securely, masked from logs.

---

## 7. Testing
- Unit tests for each exporter implementation.
- Integration tests for ASP.NET endpoint.
- Proof-of-concept: initial “Success!” text file export to validate pipeline end-to-end.

---

## 8. Documentation
- README.md for overview.
- requirements.md for evolving spec.
- Inline code comments for maintainability.

---

## 9. Next Steps
- Implement JSON schema validation in controller.
- Scaffold `IDataProvider` and `IExporter` interfaces.
- Build proof-of-concept pipeline: ASP.NET → wrapper → Python → “Success!” file.
