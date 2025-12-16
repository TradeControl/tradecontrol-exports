# tradecontrol-exports

ERP-driven data export library for Trade Control Ltd.  
This project rewrites legacy VSTO/Excel logic into modern Python and .NET components, producing spreadsheet outputs in **Office (XLSX)** and **Libre (ODS)** formats.

## Overview
`tradecontrol-exports` provides a service layer that connects to the ERP database and generates downloadable documents for end users. It integrates with the ASP.NET website (`tradecontrol.web`) via a class library that passes:
- A connection string
- The object to be exported
- An associated parameter object

The library then produces the required document and streams it back to the user’s download folder.

## Features
- Export ERP data to **Office XLSX** or **Libre ODS**
- Pluggable exporter interface (`IExporter`) for future formats (CSV, PDF, etc.)
- Integration with ASP.NET endpoints
- Clean separation of requirements and implementation

## Project Status
This is a **provisional scaffold**. Requirements are documented in [`docs/requirements.md`](docs/requirements.md) and will evolve as the project develops.

## License
Licensed under the **GNU General Public License v3.0**.  
See [LICENSE](LICENSE) for details.
