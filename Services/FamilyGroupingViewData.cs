using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services;

/// <summary>Family grouping upload rows (priority columns in display order).</summary>
public static class FamilyGroupingViewData
{
    public const string EnsureTableSql = """
IF OBJECT_ID(N'dbo.FamilyGroupingDetail', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.FamilyGroupingDetail (
        FamilyGroupingDetailId INT IDENTITY(1,1) PRIMARY KEY,
        AssemblyNo         NVARCHAR(100) NOT NULL,
        Pcb                NVARCHAR(50)  NULL,
        FamilyName         NVARCHAR(200) NULL,
        FamilyNumber       NVARCHAR(50)  NULL,
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
    CREATE INDEX IX_FamilyGroupingDetail_AssemblyNo ON dbo.FamilyGroupingDetail(AssemblyNo);
END;
""";

    public const string GridSql = """
SELECT
    FamilyGroupingDetailId,
    AssemblyNo AS [ASY],
    Pcb AS [PCB],
    FamilyName AS [Family],
    FamilyNumber AS [Family#],
    TopPriority1 AS [Top Priority 1],
    TopPriority2 AS [Top Priority 2],
    TopPriority3 AS [Top Priority 3],
    TopCycleTimeSec AS [Top CycleTime],
    BotPriority1 AS [Bot Priority 1],
    BotPriority2 AS [Bot Priority 2],
    BotPriority3 AS [Bot Priority 3],
    BotCycleTimeSec AS [Bot CycleTime],
    CircuitCount AS [# Circuits]
FROM dbo.FamilyGroupingDetail
ORDER BY AssemblyNo, FamilyGroupingDetailId;
""";

    public const string ExportSql = """
IF OBJECT_ID(N'dbo.FamilyGroupingDetail', N'U') IS NULL
BEGIN
    SELECT
        CAST(NULL AS nvarchar(100)) AS [ASY],
        CAST(NULL AS nvarchar(50))  AS [PCB],
        CAST(NULL AS nvarchar(200)) AS [Family],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 1],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 2],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 3],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 1],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 2],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 3],
        CAST(NULL AS int)           AS [# Circuits]
    WHERE 1 = 0;
END
ELSE
BEGIN
    SELECT
        AssemblyNo AS [ASY],
        Pcb AS [PCB],
        FamilyName AS [Family],
        TopPriority1 AS [Top Priority 1],
        TopPriority2 AS [Top Priority 2],
        TopPriority3 AS [Top Priority 3],
        BotPriority1 AS [Bot Priority 1],
        BotPriority2 AS [Bot Priority 2],
        BotPriority3 AS [Bot Priority 3],
        CircuitCount AS [# Circuits]
    FROM dbo.FamilyGroupingDetail
    ORDER BY AssemblyNo, FamilyGroupingDetailId;
END
""";

    public static string BuildFilterSql(string filterField)
    {
        var where = filterField switch
        {
            "ASY" => "WHERE UPPER(LTRIM(RTRIM(AssemblyNo))) = UPPER(LTRIM(RTRIM(@v)))",
            "Family" => "WHERE UPPER(LTRIM(RTRIM(FamilyName))) = UPPER(LTRIM(RTRIM(@v)))",
            _ => throw new InvalidOperationException("Unsupported filter field.")
        };

        return $"""
IF OBJECT_ID(N'dbo.FamilyGroupingDetail', N'U') IS NULL
BEGIN
    SELECT
        CAST(NULL AS nvarchar(100)) AS [ASY],
        CAST(NULL AS nvarchar(50))  AS [PCB],
        CAST(NULL AS nvarchar(200)) AS [Family],
        CAST(NULL AS nvarchar(50))  AS [Family#],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 1],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 2],
        CAST(NULL AS nvarchar(50))  AS [Top Priority 3],
        CAST(NULL AS decimal(10,3)) AS [Top CycleTime],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 1],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 2],
        CAST(NULL AS nvarchar(50))  AS [Bot Priority 3],
        CAST(NULL AS decimal(10,3)) AS [Bot CycleTime],
        CAST(NULL AS int)           AS [# Circuits]
    WHERE 1 = 0;
END
ELSE
BEGIN
    SELECT
        AssemblyNo AS [ASY],
        Pcb AS [PCB],
        FamilyName AS [Family],
        FamilyNumber AS [Family#],
        TopPriority1 AS [Top Priority 1],
        TopPriority2 AS [Top Priority 2],
        TopPriority3 AS [Top Priority 3],
        TopCycleTimeSec AS [Top CycleTime],
        BotPriority1 AS [Bot Priority 1],
        BotPriority2 AS [Bot Priority 2],
        BotPriority3 AS [Bot Priority 3],
        BotCycleTimeSec AS [Bot CycleTime],
        CircuitCount AS [# Circuits]
    FROM dbo.FamilyGroupingDetail
    {where}
    ORDER BY AssemblyNo, FamilyGroupingDetailId;
END
""";
    }

    public static async Task EnsureTableAsync(SqlConnection conn, SqlTransaction? tx, CancellationToken ct = default)
    {
        await using var cmd = new SqlCommand(EnsureTableSql, conn, tx);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    public static async Task EnsureTableAsync(SqlConnectionFactory connectionFactory, CancellationToken ct = default)
    {
        await using var conn = connectionFactory.Open();
        await EnsureTableAsync(conn, null, ct);
    }

    public static async Task<DataTable> LoadGridAsync(SqlConnectionFactory connectionFactory, CancellationToken ct = default)
    {
        var initializer = new DbInitializer(connectionFactory);
        await initializer.EnsureDatabaseAndSchemaAsync(@"database\schema.sql", ct);
        await EnsureTableAsync(connectionFactory, ct);

        await using var conn = connectionFactory.Open();
        await using var cmd = new SqlCommand(GridSql, conn);
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        var table = new DataTable();
        table.Load(reader);
        return table;
    }
}
