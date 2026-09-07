# tradecontrol-exports

ERP-driven data export library for Trade Control Ltd.

## Overview
`tradecontrol-exports` provides a service layer that connects to the ERP database and generates downloadable documents for end users. It is consumed by the ASP.NET web application (`tradecontrol.web`) via a direct project reference, brought into that repo as a Git submodule (or subtree).

## Features
- Export ERP data to **Office XLSX**
- Pluggable document handlers for additional formats
- Experimental native **OpenDocument Spreadsheet (ODS)** Cash Statement generator
- Integration with ASP.NET endpoints
- Clean separation of requirements and implementation

## Current Status
- XLSX Cash Statement: operational and the current supported export path.
- ODS (Libre) Cash Statement: feasibility confirmed across the four synthetic company and sole-trader template scenarios. Native ODS files can be generated and opened by LibreOffice, but numerical parity and deployment issues remain; the route is experimental and must not yet be used for production financial reporting.
- ODS development is deferred while tax/accounts integration is the active priority. Future work will use one format-neutral Cash Statement model with thin XLSX and ODS renderers rather than maintaining two independent calculation engines.
- Packaging: internal only; no NuGet publishing. The ASP.NET repo vendors this repo and references the project directly.

The investigation findings, confirmed defects, recommended architecture, implementation phases, and release acceptance matrix are recorded in [ODS Cash Statement Generation](docs/ods-statement-generation.md).

## Integrating with tradecontrol.web

### Git submodule
1) In the tradecontrol.web repo root:
   - `git submodule add https://github.com/TradeControl/tradecontrol-exports.git src/TCExports`
   - `git submodule update --init --recursive`
2) Open the web solution and add a ProjectReference to:
   - `src\TCExports\src\TCExports.Generator\TCExports.Generator.csproj`
   - Visual Studio: right-click web project > __Add > Project Reference...__ > __Browse...__ to the `.csproj`.
3) Commit the updated `.sln`, `.csproj`, and `.gitmodules`.

### Configuration
- Provide the ERP connection string via ASP.NET configuration.
- Disable ODS features via a feature flag until reinstated.

### Developer notes
- Build: __Build Solution__
- Local testing: `TCExports.WebHarness` remains a harness in this repo and is not deployed.

## Roadmap (paused items)
- Extract a shared, format-neutral Cash Statement model while preserving XLSX behaviour.
- Reduce the existing Python ODS implementation to a renderer of that shared model.
- Qualify ODS for Accounts Mode using the four-fixture numerical parity and LibreOffice acceptance suite.
- Add CSV/PDF exporters through the document-handler contract.

## License
Licensed under the **GNU General Public License v3.0**.
See [LICENSE](LICENSE) for details.
