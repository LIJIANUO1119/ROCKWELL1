-- SMT Line Allocation System - Detailed Database Schema
-- Target DBMS: SQL Server (easily adaptable to SQLite with minor changes)
-- NOTE:
-- - ID fields use INT IDENTITY for simplicity.
-- - Use NVARCHAR for texts to support Chinese/English.
-- - Adjust lengths / indexes based on real data volume when implementing.

------------------------------------------------------------
-- 0. Utility: Common schema & default
------------------------------------------------------------

IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = N'dbo')
BEGIN
    EXEC('CREATE SCHEMA dbo');
END;
GO

------------------------------------------------------------
-- 1. User & Permission Tables
------------------------------------------------------------

CREATE TABLE dbo.[User] (
    UserId          INT IDENTITY(1,1) PRIMARY KEY,
    UserName        NVARCHAR(50)  NOT NULL UNIQUE,   -- Login account
    DisplayName     NVARCHAR(100) NULL,              -- Display name
    PasswordHash    VARBINARY(256) NOT NULL,         -- Password hash (handled by the application)
    IsActive        BIT           NOT NULL DEFAULT 1,
    CreatedAt       DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt       DATETIME2(0)  NULL
);

CREATE TABLE dbo.[Role] (
    RoleId          INT IDENTITY(1,1) PRIMARY KEY,
    RoleName        NVARCHAR(50)  NOT NULL UNIQUE,   -- Admin / Engineer / Planner
    Description     NVARCHAR(200) NULL
);

CREATE TABLE dbo.UserRole (
    UserId          INT NOT NULL,
    RoleId          INT NOT NULL,
    PRIMARY KEY (UserId, RoleId),
    CONSTRAINT FK_UserRole_User FOREIGN KEY (UserId) REFERENCES dbo.[User](UserId),
    CONSTRAINT FK_UserRole_Role FOREIGN KEY (RoleId) REFERENCES dbo.[Role](RoleId)
);

------------------------------------------------------------
-- 2. Line / Machine / Nozzle Basic Tables
------------------------------------------------------------

CREATE TABLE dbo.Line (
    LineId          INT IDENTITY(1,1) PRIMARY KEY,
    LineCode        NVARCHAR(50)  NOT NULL UNIQUE,   -- Line name/code from Excel (e.g. "SMT1")
    LineName        NVARCHAR(100) NULL,
    Description     NVARCHAR(200) NULL,
    IsActive        BIT           NOT NULL DEFAULT 1,
    CreatedAt       DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt       DATETIME2(0)  NULL
);

CREATE TABLE dbo.Machine (
    MachineId       INT IDENTITY(1,1) PRIMARY KEY,
    LineId          INT          NOT NULL,
    MachineCode     NVARCHAR(50) NOT NULL,           -- Machine code (e.g. "GSM1-1")
    MachineName     NVARCHAR(100) NULL,              -- Machine name
    MachineType     NVARCHAR(50)  NULL,              -- e.g. GSM / GC13 / SPI / AOI
    PositionInLine  INT           NULL,              -- Order within the line (1,2,3...)
    IsActive        BIT           NOT NULL DEFAULT 1,
    CreatedAt       DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt       DATETIME2(0)  NULL,
    CONSTRAINT FK_Machine_Line FOREIGN KEY (LineId) REFERENCES dbo.Line(LineId)
);

CREATE UNIQUE INDEX UX_Machine_Line_MachineCode
ON dbo.Machine(LineId, MachineCode);

-- Optional: machine software/program information (from line configuration Excel, e.g. "Program Name", "Software Version", etc.)
CREATE TABLE dbo.MachineSoftware (
    MachineSoftwareId  INT IDENTITY(1,1) PRIMARY KEY,
    MachineId          INT           NOT NULL,
    SoftwareName       NVARCHAR(100) NOT NULL,       -- e.g. placement program name
    SoftwareVersion    NVARCHAR(50)  NULL,
    FilePathOrNumber   NVARCHAR(260) NULL,           -- Program number/path, etc.
    IsActive           BIT           NOT NULL DEFAULT 1,
    CreatedAt          DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt          DATETIME2(0)  NULL,
    CONSTRAINT FK_MachineSoftware_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId)
);

------------------------------------------------------------
-- Nozzle Definition & Machine Nozzle Config
-- From nozzle type data in SGP_COMPONENT_DATA_GSM_GE_HEAD_GC13.csv
------------------------------------------------------------

CREATE TABLE dbo.NozzleType (
    NozzleTypeId    INT IDENTITY(1,1) PRIMARY KEY,
    NozzleCode      NVARCHAR(50)  NOT NULL UNIQUE,   -- Nozzle code/model (e.g. "GC13-NZL-01")
    NozzleModel     NVARCHAR(100) NULL,
    HeightMinMm     DECIMAL(8,3)  NULL,
    HeightMaxMm     DECIMAL(8,3)  NULL,
    PrecisionLevel  NVARCHAR(50)  NULL,              -- Precision class / supported precision
    Vendor          NVARCHAR(100) NULL,
    Description     NVARCHAR(200) NULL,
    IsActive        BIT           NOT NULL DEFAULT 1,
    CreatedAt       DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt       DATETIME2(0)  NULL
);

-- Machine nozzle installation configuration (from line config files or manual maintenance)
CREATE TABLE dbo.MachineNozzleConfig (
    MachineNozzleConfigId INT IDENTITY(1,1) PRIMARY KEY,
    MachineId             INT          NOT NULL,
    NozzleTypeId          INT          NOT NULL,
    HeadIndex             INT          NULL,         -- Head index (e.g. HEAD 1 / HEAD 2)
    HeadLabel             NVARCHAR(30) NULL,         -- Text head label (e.g. Changer01 / InLine07)
    SlotIndex             INT          NOT NULL,     -- Slot index
    Quantity              INT          NOT NULL DEFAULT 1,
    IsActive              BIT          NOT NULL DEFAULT 1,
    CreatedAt             DATETIME2(0) NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt             DATETIME2(0) NULL,
    CONSTRAINT FK_MachineNozzleConfig_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId),
    CONSTRAINT FK_MachineNozzleConfig_NozzleType FOREIGN KEY (NozzleTypeId) REFERENCES dbo.NozzleType(NozzleTypeId)
);

-- Uniqueness must include HeadLabel; otherwise "ChangerXX" and other heads
-- (e.g. "Spindle Nozzel", "IN Line 7") sharing the same SlotIndex would overwrite.
CREATE UNIQUE INDEX UX_MachineNozzleConfig_Machine_Head_Slot
ON dbo.MachineNozzleConfig(MachineId, HeadLabel, SlotIndex);
GO

-- Idempotent migration: drop old uniqueness index if it exists.
IF EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_MachineNozzleConfig_Machine_Slot'
      AND object_id = OBJECT_ID(N'dbo.MachineNozzleConfig')
)
BEGIN
    DROP INDEX UX_MachineNozzleConfig_Machine_Slot ON dbo.MachineNozzleConfig;
END;
GO

-- Schema upgrade (idempotent): add HeadLabel if the table already existed before this column was introduced.
IF COL_LENGTH('dbo.MachineNozzleConfig', 'HeadLabel') IS NULL
BEGIN
    ALTER TABLE dbo.MachineNozzleConfig
    ADD HeadLabel NVARCHAR(30) NULL;
END;
GO

-- Machine model from Excel Equip Type (e.g. "Momentum II" from "MPM Momentum II")
IF COL_LENGTH('dbo.Machine', 'MachineModel') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD MachineModel NVARCHAR(100) NULL;
END;
GO

-- Extra machine configuration columns for "SMT LINE CONFIGURATION" matrix view
IF COL_LENGTH('dbo.Machine', 'EquipType') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD EquipType NVARCHAR(100) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'SerialNumberOrComputerName') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD SerialNumberOrComputerName NVARCHAR(100) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'SoftwareVersion') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD SoftwareVersion NVARCHAR(50) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'IpAddress') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD IpAddress NVARCHAR(50) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'Dns') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD Dns NVARCHAR(100) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'Gateway') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD Gateway NVARCHAR(50) NULL;
END;
GO
IF COL_LENGTH('dbo.Machine', 'Os') IS NULL
BEGIN
    ALTER TABLE dbo.Machine ADD Os NVARCHAR(50) NULL;
END;
GO

------------------------------------------------------------
-- 3. Product / Component / BOM (ProductComponent)
--   - .etf: process/BOM files including Board Size, Target CycleTime, BOM, etc.
--   - SGP_ASSEMBLY_BUILD_TO_MODULE.csv: assembly structure / module mapping (can be mapped to BOM or submodules).
------------------------------------------------------------

CREATE TABLE dbo.Product (
    ProductId        INT IDENTITY(1,1) PRIMARY KEY,
    ProductCode      NVARCHAR(50)  NOT NULL UNIQUE,  -- e.g. "PN123456"
    ProductName      NVARCHAR(200) NULL,
    FamilyCode       NVARCHAR(50)  NULL,             -- Product family code (for linking with SGP_FAMILY_GROUPINGS_DETAIL)
    BoardLengthMm    DECIMAL(10,3) NULL,             -- Board length
    BoardWidthMm     DECIMAL(10,3) NULL,             -- Board width
    BoardThicknessMm DECIMAL(8,3)  NULL,
    HasHighComponents BIT          NOT NULL DEFAULT 0,
    TargetCycleTimeSec DECIMAL(10,3) NULL,           -- Target cycle time (sec/board)
    IsActive         BIT           NOT NULL DEFAULT 1,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL
);

CREATE TABLE dbo.Component (
    ComponentId      INT IDENTITY(1,1) PRIMARY KEY,
    PartNumber       NVARCHAR(100) NOT NULL,         -- PartNumber from SGP_COMPONENT_DATA
    Description      NVARCHAR(200) NULL,
    HeightMm         DECIMAL(8,3)  NULL,
    PrecisionClass   NVARCHAR(50)  NULL,             -- Precision class (e.g. FinePitch / BGA)
    ComponentType    NVARCHAR(50)  NULL,             -- e.g. "R", "C", "IC"
    IsActive         BIT           NOT NULL DEFAULT 1,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL
);

CREATE UNIQUE INDEX UX_Component_PartNumber
ON dbo.Component(PartNumber);

-- Component -> Nozzle requirements
-- Derived from SGP_COMPONENT_DATA_GSM_GE_HEAD GC13.csv (columns Nozzle1/Nozzle2)
CREATE TABLE dbo.ComponentNozzleRequirement (
    ComponentNozzleRequirementId INT IDENTITY(1,1) PRIMARY KEY,
    ComponentId      INT NOT NULL,
    NozzleTypeId     INT NOT NULL,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    CONSTRAINT FK_ComponentNozzleRequirement_Component FOREIGN KEY (ComponentId) REFERENCES dbo.Component(ComponentId),
    CONSTRAINT FK_ComponentNozzleRequirement_NozzleType FOREIGN KEY (NozzleTypeId) REFERENCES dbo.NozzleType(NozzleTypeId),
    CONSTRAINT UX_ComponentNozzleRequirement UNIQUE (ComponentId, NozzleTypeId)
);

-- Product BOM (based on .etf + SGP_ASSEMBLY_BUILD_TO_MODULE.csv)
CREATE TABLE dbo.ProductComponent (
    ProductComponentId INT IDENTITY(1,1) PRIMARY KEY,
    ProductId          INT           NOT NULL,
    ComponentId        INT           NOT NULL,
    ReferenceDesignator NVARCHAR(100) NULL,          -- Reference designator (e.g. R1, C1), optional
    QuantityPerBoard   DECIMAL(10,3) NOT NULL DEFAULT 1,
    ModuleCode         NVARCHAR(50)  NULL,           -- Module code (if sourced from BUILD_TO_MODULE)
    StageName          NVARCHAR(50)  NULL,           -- Stage/process (e.g. Top/Bottom/Module)
    NozzleTypeId       INT           NULL,           -- Nozzle type mapped from SGP_COMPONENT_DATA_GSM_GE_HEAD_GC13
    CreatedAt          DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt          DATETIME2(0)  NULL,
    CONSTRAINT FK_ProductComponent_Product FOREIGN KEY (ProductId) REFERENCES dbo.Product(ProductId),
    CONSTRAINT FK_ProductComponent_Component FOREIGN KEY (ComponentId) REFERENCES dbo.Component(ComponentId),
    CONSTRAINT FK_ProductComponent_NozzleType FOREIGN KEY (NozzleTypeId) REFERENCES dbo.NozzleType(NozzleTypeId)
);

CREATE INDEX IX_ProductComponent_ProductId ON dbo.ProductComponent(ProductId);
CREATE INDEX IX_ProductComponent_ComponentId ON dbo.ProductComponent(ComponentId);

------------------------------------------------------------
-- 4. Family / Line Allocation
--   - From SGP_FAMILY_GROUPINGS_DETAIL.csv
------------------------------------------------------------

CREATE TABLE dbo.ProductFamily (
    ProductFamilyId  INT IDENTITY(1,1) PRIMARY KEY,
    FamilyCode       NVARCHAR(50)  NOT NULL UNIQUE,  -- Family group code
    FamilyName       NVARCHAR(200) NULL,
    Description      NVARCHAR(200) NULL,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL
);

-- Product-to-family mapping (if a product belongs to only one family, Product.FamilyCode alone may be sufficient)
CREATE TABLE dbo.ProductFamilyMapping (
    ProductId       INT NOT NULL,
    ProductFamilyId INT NOT NULL,
    PRIMARY KEY (ProductId, ProductFamilyId),
    CONSTRAINT FK_ProductFamilyMapping_Product FOREIGN KEY (ProductId) REFERENCES dbo.Product(ProductId),
    CONSTRAINT FK_ProductFamilyMapping_ProductFamily FOREIGN KEY (ProductFamilyId) REFERENCES dbo.ProductFamily(ProductFamilyId)
);

-- Current and historical line allocation
CREATE TABLE dbo.ProductLineAllocation (
    AllocationId     INT IDENTITY(1,1) PRIMARY KEY,
    ProductId        INT           NOT NULL,
    LineId           INT           NOT NULL,
    ProductFamilyId  INT           NULL,             -- Family from SGP_FAMILY_GROUPINGS_DETAIL
    PriorityInFamily INT           NULL,             -- Line priority within family (mapped from Excel)
    EffectiveFrom    DATE          NOT NULL,
    EffectiveTo      DATE          NULL,
    IsCurrent        BIT           NOT NULL DEFAULT 1,
    Remark           NVARCHAR(200) NULL,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL,
    CONSTRAINT FK_ProductLineAllocation_Product FOREIGN KEY (ProductId) REFERENCES dbo.Product(ProductId),
    CONSTRAINT FK_ProductLineAllocation_Line FOREIGN KEY (LineId) REFERENCES dbo.Line(LineId),
    CONSTRAINT FK_ProductLineAllocation_ProductFamily FOREIGN KEY (ProductFamilyId) REFERENCES dbo.ProductFamily(ProductFamilyId)
);

CREATE INDEX IX_ProductLineAllocation_ProductId ON dbo.ProductLineAllocation(ProductId);
CREATE INDEX IX_ProductLineAllocation_LineId ON dbo.ProductLineAllocation(LineId);

------------------------------------------------------------
-- 4b. Manual Line Allocation (per Family) - user maintained
-- Allows multiple rows per Family; no dedupe (append-only unless user deletes).
------------------------------------------------------------
IF OBJECT_ID(N'dbo.FamilyLineAllocation', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FamilyLineAllocation (
        FamilyLineAllocationId INT IDENTITY(1,1) PRIMARY KEY,
        FamilyCode     NVARCHAR(50)  NOT NULL,
        PrimaryLine    NVARCHAR(50)  NULL,
        SecondaryLine  NVARCHAR(50)  NULL,
        TertiaryLine   NVARCHAR(50)  NULL,
        ModuleNumber   NVARCHAR(100) NULL,
        Side           NVARCHAR(20)  NULL,
        Constraints    NVARCHAR(MAX) NULL,
        CreatedAt      DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );

    CREATE INDEX IX_FamilyLineAllocation_FamilyCode
    ON dbo.FamilyLineAllocation(FamilyCode);
END;

-- Idempotent upgrade: add Constraints column if table already exists.
IF OBJECT_ID(N'dbo.FamilyLineAllocation', N'U') IS NOT NULL
   AND COL_LENGTH('dbo.FamilyLineAllocation', 'Constraints') IS NULL
BEGIN
    ALTER TABLE dbo.FamilyLineAllocation ADD Constraints NVARCHAR(MAX) NULL;
END;
GO

------------------------------------------------------------
-- 4c. Family groupings detail (SGP_FAMILY_GROUPINGS_DETAIL upload)
-- Separate batch so upgrades run even when section 3 tables already exist.
------------------------------------------------------------
IF OBJECT_ID(N'dbo.FamilyGroupingDetail', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FamilyGroupingDetail (
        FamilyGroupingDetailId INT IDENTITY(1,1) PRIMARY KEY,
        AssemblyNo         NVARCHAR(100) NOT NULL,   -- ASY
        Pcb                NVARCHAR(50)  NULL,
        FamilyName         NVARCHAR(200) NULL,       -- Family
        FamilyNumber       NVARCHAR(50)  NULL,       -- Family #
        TopPriority1       NVARCHAR(50)  NULL,
        TopPriority2       NVARCHAR(50)  NULL,
        TopPriority3       NVARCHAR(50)  NULL,
        TopCycleTimeSec    DECIMAL(10,3) NULL,
        BotPriority1       NVARCHAR(50)  NULL,
        BotPriority2       NVARCHAR(50)  NULL,
        BotPriority3       NVARCHAR(50)  NULL,
        BotCycleTimeSec    DECIMAL(10,3) NULL,
        CircuitCount       INT           NULL,
        SourceFileName     NVARCHAR(260) NULL,
        CreatedAt          DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );

    CREATE INDEX IX_FamilyGroupingDetail_AssemblyNo
    ON dbo.FamilyGroupingDetail(AssemblyNo);
END;
GO

------------------------------------------------------------
-- 5. Constraints (Line / Machine)
------------------------------------------------------------

CREATE TABLE dbo.LineConstraint (
    LineConstraintId INT IDENTITY(1,1) PRIMARY KEY,
    LineId           INT          NOT NULL,
    MaxBoardLengthMm DECIMAL(10,3) NULL,
    MaxBoardWidthMm  DECIMAL(10,3) NULL,
    MaxComponentHeightMm DECIMAL(8,3) NULL,
    HighPrecisionSupport  BIT      NULL,             -- Supports high-precision components
    HighComponentSupport  BIT      NULL,             -- Supports tall/high components
    GlueDispenserRequired BIT      NULL,             -- Hard constraint: product requires glue dispenser support
    LargeBoardOver10InchSupported BIT NULL,          -- Hard constraint: supports large board width > 10"
    ReflowNonRetractableCenterSupport BIT NULL,      -- Hard constraint: reflow non-retractable center support
    Remark           NVARCHAR(200) NULL,
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL,
    CONSTRAINT FK_LineConstraint_Line FOREIGN KEY (LineId) REFERENCES dbo.Line(LineId)
);

CREATE UNIQUE INDEX UX_LineConstraint_LineId ON dbo.LineConstraint(LineId);
GO

-- Schema upgrade (idempotent): add constraint columns if missing.
IF COL_LENGTH('dbo.LineConstraint', 'GlueDispenserRequired') IS NULL
BEGIN
    ALTER TABLE dbo.LineConstraint ADD GlueDispenserRequired BIT NULL;
END;
GO
IF COL_LENGTH('dbo.LineConstraint', 'LargeBoardOver10InchSupported') IS NULL
BEGIN
    ALTER TABLE dbo.LineConstraint ADD LargeBoardOver10InchSupported BIT NULL;
END;
GO
IF COL_LENGTH('dbo.LineConstraint', 'ReflowNonRetractableCenterSupport') IS NULL
BEGIN
    ALTER TABLE dbo.LineConstraint ADD ReflowNonRetractableCenterSupport BIT NULL;
END;
GO

-- Optional: machine-level constraints (e.g. board width limits, nozzle head limits, etc.)
CREATE TABLE dbo.MachineConstraint (
    MachineConstraintId INT IDENTITY(1,1) PRIMARY KEY,
    MachineId           INT           NOT NULL,
    MaxBoardLengthMm    DECIMAL(10,3) NULL,
    MaxBoardWidthMm     DECIMAL(10,3) NULL,
    MaxComponentHeightMm DECIMAL(8,3) NULL,
    HighPrecisionSupport BIT          NULL,
    HighComponentSupport BIT          NULL,
    Remark              NVARCHAR(200) NULL,
    CreatedAt           DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt           DATETIME2(0)  NULL,
    CONSTRAINT FK_MachineConstraint_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId)
);

------------------------------------------------------------
-- 6. Production History & Cycle Time
--   - Historical cycle times, machine cycle times, etc. from line logs and uploaded CSV files
------------------------------------------------------------

CREATE TABLE dbo.ProductionHistory (
    HistoryId        INT IDENTITY(1,1) PRIMARY KEY,
    ProductId        INT           NOT NULL,
    LineId           INT           NOT NULL,
    MachineId        INT           NULL,             -- Optional: record at machine level
    WorkOrder        NVARCHAR(50)  NULL,
    BoardsProduced   INT           NULL,
    StartTime        DATETIME2(0)  NOT NULL,
    EndTime          DATETIME2(0)  NULL,
    AvgCycleTimeSec  DECIMAL(10,3) NULL,            -- Average cycle time (sec/board)
    IsSimulated      BIT           NOT NULL DEFAULT 0,
    SourceFileName   NVARCHAR(260) NULL,            -- Source file name (for traceability)
    CreatedAt        DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt        DATETIME2(0)  NULL,
    CONSTRAINT FK_ProductionHistory_Product FOREIGN KEY (ProductId) REFERENCES dbo.Product(ProductId),
    CONSTRAINT FK_ProductionHistory_Line FOREIGN KEY (LineId) REFERENCES dbo.Line(LineId),
    CONSTRAINT FK_ProductionHistory_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId)
);

CREATE INDEX IX_ProductionHistory_Product_Line ON dbo.ProductionHistory(ProductId, LineId);
CREATE INDEX IX_ProductionHistory_StartTime ON dbo.ProductionHistory(StartTime);

-- Machine-level cycle time data (Upload machine cycletime)
CREATE TABLE dbo.MachineCycleTime (
    MachineCycleTimeId INT IDENTITY(1,1) PRIMARY KEY,
    MachineId          INT           NOT NULL,
    ProductId          INT           NULL,          -- For a specific product, or NULL for generic
    StageName          NVARCHAR(50)  NULL,         -- Stage/process (e.g. "Placement", "Reflow")
    CycleTimeSec       DECIMAL(10,3) NOT NULL,
    SampleCount        INT           NULL,         -- Sample count
    MeasuredAt         DATETIME2(0)  NULL,         -- Measured/aggregated time
    SourceFileName     NVARCHAR(260) NULL,
    CreatedAt          DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt          DATETIME2(0)  NULL,
    CONSTRAINT FK_MachineCycleTime_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId),
    CONSTRAINT FK_MachineCycleTime_Product FOREIGN KEY (ProductId) REFERENCES dbo.Product(ProductId)
);

CREATE INDEX IX_MachineCycleTime_Machine_Product ON dbo.MachineCycleTime(MachineId, ProductId);

------------------------------------------------------------
-- 6b. Import-support: cleaned machine cycletime snapshot
--  - Persist only the columns needed for bottleneck analysis.
--  - Each upload appends rows; previous snapshot rows are kept (handled by importer).
------------------------------------------------------------
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MachineCycleTimeSnapshot (
        MachineCycleTimeSnapshotId INT IDENTITY(1,1) PRIMARY KEY,
        LineCode          NVARCHAR(50)  NOT NULL,
        AssemblyNo        NVARCHAR(100) NOT NULL,
        Side              NVARCHAR(20)  NULL,
        MachineName       NVARCHAR(100) NOT NULL,
        PanelEndTime      DATETIME2(0)  NULL,
        BoardsProcessed   INT           NULL,
        MachineMinCycleTimeSec     DECIMAL(10,3) NULL,
        MachineMediumCycleTimeSec  DECIMAL(10,3) NULL,
        AdmCurrentCycleTimeSec     DECIMAL(10,3) NULL,
        SourceFileName    NVARCHAR(260) NULL,
        CreatedAt         DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );

    CREATE INDEX IX_MachineCycleTimeSnapshot_Key
    ON dbo.MachineCycleTimeSnapshot(LineCode, AssemblyNo, Side, MachineName);
END;

-- Idempotent upgrade: add PanelEndTime column if missing.
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NOT NULL
   AND COL_LENGTH('dbo.MachineCycleTimeSnapshot', 'PanelEndTime') IS NULL
BEGIN
    ALTER TABLE dbo.MachineCycleTimeSnapshot ADD PanelEndTime DATETIME2(0) NULL;
END;

-- Idempotent upgrade: remove WorkOrderNo column (no longer used).
IF OBJECT_ID(N'dbo.MachineCycleTimeSnapshot', N'U') IS NOT NULL
   AND COL_LENGTH('dbo.MachineCycleTimeSnapshot', 'WorkOrderNo') IS NOT NULL
BEGIN
    ALTER TABLE dbo.MachineCycleTimeSnapshot DROP COLUMN WorkOrderNo;
END;

------------------------------------------------------------
-- 7. Import File Log
--   - Logs metadata and results for Excel/CSV imports
------------------------------------------------------------

CREATE TABLE dbo.ImportFileLog (
    ImportFileLogId INT IDENTITY(1,1) PRIMARY KEY,
    FileName        NVARCHAR(260) NOT NULL,
    FileType        NVARCHAR(50)  NOT NULL,         -- e.g. "LINE_CONFIG", "PRODUCT_ETF", "SGP_COMPONENT_DATA", "SGP_ASSEMBLY", "SGP_FAMILY_GROUPINGS"
    ImportedByUserId INT          NULL,
    ImportedAt      DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    RecordCount     INT           NULL,
    SuccessCount    INT           NULL,
    FailureCount    INT           NULL,
    Status          NVARCHAR(20)  NOT NULL DEFAULT 'SUCCESS', -- SUCCESS / PARTIAL / FAILED
    ErrorMessage    NVARCHAR(MAX) NULL,
    CONSTRAINT FK_ImportFileLog_User FOREIGN KEY (ImportedByUserId) REFERENCES dbo.[User](UserId)
);

------------------------------------------------------------
-- 7b. Import-support tables for CSV sources
------------------------------------------------------------

-- Raw mapping from SGP_ASSEMBLY_BUILD_TO_MODULE.csv (kept for traceability)
IF OBJECT_ID(N'dbo.BuildToModuleMapping', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BuildToModuleMapping (
        BuildToModuleMappingId INT IDENTITY(1,1) PRIMARY KEY,
        BuildNumber       NVARCHAR(50)  NULL,
        BuildKey          NVARCHAR(50)  NULL,
        ModuleNumber      NVARCHAR(50)  NULL,
        ModuleKey         NVARCHAR(50)  NULL,
        PCB               NVARCHAR(50)  NULL,
        ModulesPerPanel   INT           NULL,
        OptelCreateDate   DATETIME2(0)  NULL,
        LastChanged       DATETIME2(0)  NULL,
        SapBuildDescription  NVARCHAR(200) NULL,
        SapModuleDescription NVARCHAR(200) NULL,
        OptelDescription     NVARCHAR(500) NULL,
        SourceFileName    NVARCHAR(260) NULL,
        CreatedAt         DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME()
    );
END;

------------------------------------------------------------
-- 8. Supporting / Lookup Tables (Optional)
------------------------------------------------------------

-- Simple lookup table for enums (e.g. product types, stage names, etc.). Extend as needed.
CREATE TABLE dbo.Lookup (
    LookupId        INT IDENTITY(1,1) PRIMARY KEY,
    Category        NVARCHAR(50)  NOT NULL,         -- e.g. "ProductType", "StageName"
    Code            NVARCHAR(50)  NOT NULL,
    Name            NVARCHAR(100) NULL,
    Description     NVARCHAR(200) NULL,
    IsActive        BIT           NOT NULL DEFAULT 1,
    SortOrder       INT           NULL,
    CreatedAt       DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
    UpdatedAt       DATETIME2(0)  NULL
);

CREATE UNIQUE INDEX UX_Lookup_Category_Code ON dbo.Lookup(Category, Code);

------------------------------------------------------------
-- 9. Recommended Indexes for Main Queries
--    (adjust/extend based on real data volume)
------------------------------------------------------------

-- Query product cycle time by line/time range
-- ProductionHistory(ProductId, LineId, StartTime) is already indexed

-- Query products that do not meet the target cycle time
-- Compare Product.TargetCycleTimeSec and ProductionHistory.AvgCycleTimeSec
-- If performance becomes an issue, consider a materialized view or a summary table.

------------------------------------------------------------
-- 9b. User-maintained constraint options (Line + Machine boolean toggles)
--     Uses dbo.SmtConstraintOption (prefix avoids name collisions / failed partial deploys).
--     OptionKey is reserved for built-ins that sync to dbo.LineConstraint legacy BIT columns.
------------------------------------------------------------

-- Repair failed partial deploy: child table existed without a valid parent user table.
IF OBJECT_ID(N'dbo.LineConstraintOptionValue', N'U') IS NOT NULL
   AND OBJECT_ID(N'dbo.ConstraintOption', N'U') IS NULL
BEGIN
    DROP TABLE dbo.LineConstraintOptionValue;
END;
GO

IF OBJECT_ID(N'dbo.MachineConstraintOptionValue', N'U') IS NOT NULL
   AND OBJECT_ID(N'dbo.ConstraintOption', N'U') IS NULL
BEGIN
    DROP TABLE dbo.MachineConstraintOptionValue;
END;
GO

IF OBJECT_ID(N'dbo.SmtConstraintOption', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SmtConstraintOption (
        SmtConstraintOptionId INT IDENTITY(1,1) PRIMARY KEY,
        OptionKey             NVARCHAR(80)  NULL,
        DisplayName           NVARCHAR(200) NOT NULL,
        SortOrder             INT           NOT NULL DEFAULT 0,
        IsActive              BIT           NOT NULL DEFAULT 1,
        CreatedAt             DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
        UpdatedAt             DATETIME2(0)  NULL
    );
END;
GO

IF OBJECT_ID(N'dbo.SmtConstraintOption', N'U') IS NOT NULL
   AND NOT EXISTS (
       SELECT 1
       FROM sys.indexes i
       WHERE i.object_id = OBJECT_ID(N'dbo.SmtConstraintOption')
         AND i.name = N'UX_SmtConstraintOption_OptionKey'
   )
BEGIN
    CREATE UNIQUE INDEX UX_SmtConstraintOption_OptionKey
    ON dbo.SmtConstraintOption(OptionKey)
    WHERE OptionKey IS NOT NULL;
END;
GO

IF OBJECT_ID(N'dbo.SmtLineConstraintOptionValue', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SmtLineConstraintOptionValue (
        SmtLineConstraintOptionValueId INT IDENTITY(1,1) PRIMARY KEY,
        LineId                         INT NOT NULL,
        SmtConstraintOptionId          INT NOT NULL,
        BitValue                       BIT NOT NULL DEFAULT 0,
        UpdatedAt                      DATETIME2(0) NULL,
        CONSTRAINT FK_SmtLineConstraintOptionValue_Line FOREIGN KEY (LineId) REFERENCES dbo.Line(LineId) ON DELETE CASCADE,
        CONSTRAINT FK_SmtLineConstraintOptionValue_Option FOREIGN KEY (SmtConstraintOptionId) REFERENCES dbo.SmtConstraintOption(SmtConstraintOptionId) ON DELETE CASCADE,
        CONSTRAINT UX_SmtLineConstraintOptionValue_Line_Option UNIQUE (LineId, SmtConstraintOptionId)
    );
END;
GO

IF OBJECT_ID(N'dbo.SmtMachineConstraintOptionValue', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SmtMachineConstraintOptionValue (
        SmtMachineConstraintOptionValueId INT IDENTITY(1,1) PRIMARY KEY,
        MachineId                         INT NOT NULL,
        SmtConstraintOptionId             INT NOT NULL,
        BitValue                          BIT NOT NULL DEFAULT 0,
        UpdatedAt                         DATETIME2(0) NULL,
        CONSTRAINT FK_SmtMachineConstraintOptionValue_Machine FOREIGN KEY (MachineId) REFERENCES dbo.Machine(MachineId) ON DELETE CASCADE,
        CONSTRAINT FK_SmtMachineConstraintOptionValue_Option FOREIGN KEY (SmtConstraintOptionId) REFERENCES dbo.SmtConstraintOption(SmtConstraintOptionId) ON DELETE CASCADE,
        CONSTRAINT UX_SmtMachineConstraintOptionValue_Machine_Option UNIQUE (MachineId, SmtConstraintOptionId)
    );
END;
GO

-- Seed built-in options (idempotent)
IF OBJECT_ID(N'dbo.SmtConstraintOption', N'U') IS NOT NULL
BEGIN
    IF NOT EXISTS (SELECT 1 FROM dbo.SmtConstraintOption WHERE OptionKey = N'GLUE_DISPENSER')
        INSERT INTO dbo.SmtConstraintOption (OptionKey, DisplayName, SortOrder, IsActive)
        VALUES (N'GLUE_DISPENSER', N'Glue dispenser required', 10, 1);

    IF NOT EXISTS (SELECT 1 FROM dbo.SmtConstraintOption WHERE OptionKey = N'LARGE_BOARD_10IN')
        INSERT INTO dbo.SmtConstraintOption (OptionKey, DisplayName, SortOrder, IsActive)
        VALUES (N'LARGE_BOARD_10IN', N'Large board width > 10 inch supported', 20, 1);

    IF NOT EXISTS (SELECT 1 FROM dbo.SmtConstraintOption WHERE OptionKey = N'REFLOW_CENTER_SUPPORT')
        INSERT INTO dbo.SmtConstraintOption (OptionKey, DisplayName, SortOrder, IsActive)
        VALUES (N'REFLOW_CENTER_SUPPORT', N'Reflow non-retractable center support', 30, 1);
END;
GO

-- If an older build created dbo.ConstraintOption / *OptionValue, move data into dbo.Smt* then drop old tables.
IF OBJECT_ID(N'dbo.ConstraintOption', N'U') IS NOT NULL
   AND OBJECT_ID(N'dbo.SmtConstraintOption', N'U') IS NOT NULL
BEGIN
    INSERT INTO dbo.SmtConstraintOption (OptionKey, DisplayName, SortOrder, IsActive, CreatedAt, UpdatedAt)
    SELECT o.OptionKey, o.DisplayName, o.SortOrder, o.IsActive, o.CreatedAt, o.UpdatedAt
    FROM dbo.ConstraintOption o
    WHERE NOT EXISTS (
        SELECT 1
        FROM dbo.SmtConstraintOption s
        WHERE (o.OptionKey IS NOT NULL AND s.OptionKey = o.OptionKey)
           OR (
                o.OptionKey IS NULL
                AND s.OptionKey IS NULL
                AND s.DisplayName = o.DisplayName
                AND s.SortOrder = o.SortOrder
              )
    );

    IF OBJECT_ID(N'dbo.LineConstraintOptionValue', N'U') IS NOT NULL
    BEGIN
        INSERT INTO dbo.SmtLineConstraintOptionValue (LineId, SmtConstraintOptionId, BitValue, UpdatedAt)
        SELECT v.LineId, s.SmtConstraintOptionId, v.BitValue, SYSDATETIME()
        FROM dbo.LineConstraintOptionValue v
        INNER JOIN dbo.ConstraintOption o ON o.ConstraintOptionId = v.ConstraintOptionId
        INNER JOIN dbo.SmtConstraintOption s
            ON (o.OptionKey IS NOT NULL AND s.OptionKey = o.OptionKey)
            OR (
                 o.OptionKey IS NULL
                 AND s.OptionKey IS NULL
                 AND s.DisplayName = o.DisplayName
                 AND s.SortOrder = o.SortOrder
               )
        WHERE NOT EXISTS (
            SELECT 1 FROM dbo.SmtLineConstraintOptionValue x
            WHERE x.LineId = v.LineId AND x.SmtConstraintOptionId = s.SmtConstraintOptionId
        );

        DROP TABLE dbo.LineConstraintOptionValue;
    END;

    IF OBJECT_ID(N'dbo.MachineConstraintOptionValue', N'U') IS NOT NULL
    BEGIN
        INSERT INTO dbo.SmtMachineConstraintOptionValue (MachineId, SmtConstraintOptionId, BitValue, UpdatedAt)
        SELECT v.MachineId, s.SmtConstraintOptionId, v.BitValue, SYSDATETIME()
        FROM dbo.MachineConstraintOptionValue v
        INNER JOIN dbo.ConstraintOption o ON o.ConstraintOptionId = v.ConstraintOptionId
        INNER JOIN dbo.SmtConstraintOption s
            ON (o.OptionKey IS NOT NULL AND s.OptionKey = o.OptionKey)
            OR (
                 o.OptionKey IS NULL
                 AND s.OptionKey IS NULL
                 AND s.DisplayName = o.DisplayName
                 AND s.SortOrder = o.SortOrder
               )
        WHERE NOT EXISTS (
            SELECT 1 FROM dbo.SmtMachineConstraintOptionValue x
            WHERE x.MachineId = v.MachineId AND x.SmtConstraintOptionId = s.SmtConstraintOptionId
        );

        DROP TABLE dbo.MachineConstraintOptionValue;
    END;

    DROP TABLE dbo.ConstraintOption;
END;
GO

-- One-time copy from legacy LineConstraint BIT columns into SmtLineConstraintOptionValue
IF OBJECT_ID(N'dbo.SmtLineConstraintOptionValue', N'U') IS NOT NULL
   AND OBJECT_ID(N'dbo.LineConstraint', N'U') IS NOT NULL
   AND OBJECT_ID(N'dbo.SmtConstraintOption', N'U') IS NOT NULL
BEGIN
    INSERT INTO dbo.SmtLineConstraintOptionValue (LineId, SmtConstraintOptionId, BitValue, UpdatedAt)
    SELECT lc.LineId, o.SmtConstraintOptionId, COALESCE(lc.GlueDispenserRequired, CAST(0 AS BIT)), SYSDATETIME()
    FROM dbo.LineConstraint lc
    INNER JOIN dbo.SmtConstraintOption o ON o.OptionKey = N'GLUE_DISPENSER'
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.SmtLineConstraintOptionValue v
        WHERE v.LineId = lc.LineId AND v.SmtConstraintOptionId = o.SmtConstraintOptionId
    );

    INSERT INTO dbo.SmtLineConstraintOptionValue (LineId, SmtConstraintOptionId, BitValue, UpdatedAt)
    SELECT lc.LineId, o.SmtConstraintOptionId, COALESCE(lc.LargeBoardOver10InchSupported, CAST(0 AS BIT)), SYSDATETIME()
    FROM dbo.LineConstraint lc
    INNER JOIN dbo.SmtConstraintOption o ON o.OptionKey = N'LARGE_BOARD_10IN'
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.SmtLineConstraintOptionValue v
        WHERE v.LineId = lc.LineId AND v.SmtConstraintOptionId = o.SmtConstraintOptionId
    );

    INSERT INTO dbo.SmtLineConstraintOptionValue (LineId, SmtConstraintOptionId, BitValue, UpdatedAt)
    SELECT lc.LineId, o.SmtConstraintOptionId, COALESCE(lc.ReflowNonRetractableCenterSupport, CAST(0 AS BIT)), SYSDATETIME()
    FROM dbo.LineConstraint lc
    INNER JOIN dbo.SmtConstraintOption o ON o.OptionKey = N'REFLOW_CENTER_SUPPORT'
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.SmtLineConstraintOptionValue v
        WHERE v.LineId = lc.LineId AND v.SmtConstraintOptionId = o.SmtConstraintOptionId
    );
END;
GO

------------------------------------------------------------
-- 9c. Machine nozzle matrix (wide Excel: Machine ID + nozzle type quantity columns)
------------------------------------------------------------

IF OBJECT_ID(N'dbo.SmtMachineNozzleColumn', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SmtMachineNozzleColumn (
        NozzleCode  NVARCHAR(50) NOT NULL PRIMARY KEY,
        SortOrder   INT          NOT NULL DEFAULT 0
    );
END;
GO

IF OBJECT_ID(N'dbo.SmtMachineNozzleMatrixRow', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SmtMachineNozzleMatrixRow (
        SmtMachineNozzleMatrixRowId INT IDENTITY(1,1) PRIMARY KEY,
        MachineCode                 NVARCHAR(50)  NOT NULL,
        QuantitiesJson              NVARCHAR(MAX) NOT NULL CONSTRAINT DF_SmtMachineNozzleMatrixRow_QuantitiesJson DEFAULT (N'{}'),
        CreatedAt                   DATETIME2(0)  NOT NULL DEFAULT SYSDATETIME(),
        UpdatedAt                   DATETIME2(0)  NULL
    );

    CREATE UNIQUE INDEX UX_SmtMachineNozzleMatrixRow_MachineCode
    ON dbo.SmtMachineNozzleMatrixRow(MachineCode);
END;
GO

------------------------------------------------------------
-- END OF SCHEMA
------------------------------------------------------------

