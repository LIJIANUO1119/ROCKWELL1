## Capstone - SMT Line Allocation

### Prerequisites
- **.NET SDK**: install .NET 8 SDK (required to build/run the WPF app).
- **SQL Server Express**: install SQL Server Express (SQLEXPRESS instance).
  - Ensure Windows Authentication is enabled.
  - Optional but recommended: install SSMS to view data.

### Configure database connection
- Edit `SmtLineAllocationUI/appsettings.json`:
  - Default is:
    - `Server=localhost\\SQLEXPRESS;Database=SmtLineAllocation;Trusted_Connection=True;TrustServerCertificate=True`

### Run
- Open the project folder in Visual Studio (or build via `dotnet` after SDK install).
- Start the WPF app from `SmtLineAllocationUI`.

### Data files
- Input files are in `data/` (including:
  - `SGP_COMPONENT_DATA_GSM_GE_HEAD GC13.csv`
  - `SGP_FAMILY_GROUPINGS_DETAIL.csv`
  - `SGP_ASSEMBLY_BUILD_TO_MODULE.csv`
  - `SGP_GEM_MACHINE_CYCLETIME 30days 30oct25.csv`
  - example `.etf` files)

### Product Cycletime – `Upload build-to-module` (SGP_ASSEMBLY_BUILD_TO_MODULE.csv)
- **Columns persisted** (four headers): **`Build #`**, **`Build Key`**, **`Module #`**, **`PCB`** → stored in `dbo.BuildToModuleMapping` as `BuildNumber`, `BuildKey`, `ModuleNumber`, `PCB` (other source columns are not imported).
- **Cleaning rules** (applied in order):
  1. Skip rows where **`PCB`** is empty.
  2. Skip rows where **`Build Key`** is not numeric.
  3. Skip rows where **`Build #`** is empty.
  4. For each distinct **`Build #`**, keep **only one** row: the row whose **`Build Key`** is **numerically largest** (ties: the later row in the file wins).

