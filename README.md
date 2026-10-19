# tradecontrol-exports

ERP-driven data export library for Trade Control Ltd.

## Overview
`tradecontrol-exports` provides a service layer that connects to the ERP database and generates downloadable documents for end users. It is consumed by the ASP.NET web application (`tradecontrol.web`) via a direct project reference, brought into that repo as a Git submodule (or subtree).

## Features
- Export ERP data to **Office XLSX**
- Export the Cash Statement to native **OpenDocument Spreadsheet (ODS)**
- One immutable, format-neutral Cash Statement model rendered by both formats
- Integration with ASP.NET endpoints
- In-process .NET generation with no Python, ODBC, office automation, or server-side LibreOffice dependency

## Current Status
- XLSX and ODS Cash Statements are supported through the same C# builder and accepted native renderers.
- The four-fixture equivalence suite, exhaustive tri-format comparison, authenticated UI checks, and Azure service qualification are complete.
- ODS can be disabled independently while the accepted XLSX path remains available.
- The former Python ODS prototype and child-process route have been removed from the source and production publish graph. Their history remains available in Git.
- All active generator, harness and test projects target .NET 10 LTS.
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
- Use the Cash Statement service settings to control ODS availability and bounded generation.

### Developer notes
- Build: __Build Solution__
- Local testing: `TCExports.WebHarness` remains a harness in this repo and is not deployed.

## Roadmap
- Add further formats only through the shared document-handler contract and with their own acceptance evidence.

## License
Licensed under the **GNU General Public License v3.0**.
See [LICENSE](LICENSE) for details.
