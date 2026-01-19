# tradecontrol-exports

ERP-driven data export library for Trade Control Ltd.

## Overview
`tradecontrol-exports` provides a service layer that connects to the ERP database and generates downloadable documents for end users. It is consumed by the ASP.NET web application (`tradecontrol.web`) via a direct project reference, brought into that repo as a Git submodule (or subtree).

## Features
- Export ERP data to **Office XLSX**
- Pluggable exporter interface (`IExporter`) for future formats (CSV, PDF, ODS)
- Integration with ASP.NET endpoints
- Clean separation of requirements and implementation

## Current Status
- XLSX exports: operational.
- ODS (Libre) Cash Statement: paused due to unresolved formatting and load-time evaluation issues. Code remains in the repo but is not enabled.
- Packaging: internal only; no NuGet publishing. The ASP.NET repo vendors this repo and references the project directly.

## Integrating with tradecontrol.web

### Git submodule
1) In the tradecontrol.web repo root:
   - `git submodule add https://github.com/TradeControl/tradecontrol-exports.git src/TCExports`
   - `git submodule update --init --recursive`
2) Open the web solution and add a ProjectReference to:
   - `src\externals\tradecontrol-exports\src\TCExports.Generator\TCExports.Generator.csproj`
   - Visual Studio: right-click web project > __Add > Project Reference...__ > __Browse...__ to the `.csproj`.
3) Commit the updated `.sln`, `.csproj`, and `.gitmodules`.

### Configuration
- Provide the ERP connection string via ASP.NET configuration.
- Disable ODS features via a feature flag until reinstated.

### Developer notes
- Build: __Build Solution__
- Local testing: `TCExports.WebHarness` remains a harness in this repo and is not deployed.

## Roadmap (paused items)
- Resume Libre ODS Cash Statement once a stable approach is agreed.
- Add CSV/PDF exporters behind `IExporter`.

## License
Licensed under the **GNU General Public License v3.0**.
See [LICENSE](LICENSE) for details.