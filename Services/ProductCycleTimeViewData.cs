namespace SmtLineAllocationUI.Services;

/// <summary>SQL for product cycletime views (build-to-module + machine cycletime snapshot).</summary>
public static class ProductCycleTimeViewData
{
    /// <summary>
    /// Full joined dataset: BuildToModuleMapping.ModuleNumber = MachineCycleTimeSnapshot.AssemblyNo.
  /// </summary>
    public const string JoinedUploadedDataSql = """
WITH joined AS (
    SELECT
        b.BuildNumber AS [Build],
        b.BuildKey AS [Build Key],
        b.ModuleNumber AS [Module#],
        b.PCB AS [PCB],
        s.LineCode AS [LINE],
        s.AssemblyNo AS [AssemblyNo],
        s.Side AS [Side],
        s.MachineName AS [MACHINENAME],
        s.BoardsProcessed AS [BOARDSPROCESSED],
        s.MachineMinCycleTimeSec AS [MACHINE_MIN_CYCLETIME],
        s.MachineMediumCycleTimeSec AS [MACHINE_MEDIUM_CYCLETIME],
        s.AdmCurrentCycleTimeSec AS [ADM_CURRENT_CYCLETIME],
        s.PanelEndTime AS [PANELENDTIME],
        s.SourceFileName AS [MachineCycletimeSource],
        s.CreatedAt AS [MachineSnapshotAt],
        b.SourceFileName AS [BuildToModuleSource],
        b.CreatedAt AS [BuildToModuleAt]
    FROM dbo.BuildToModuleMapping b
    INNER JOIN dbo.MachineCycleTimeSnapshot s
        ON UPPER(LTRIM(RTRIM(b.ModuleNumber))) = UPPER(LTRIM(RTRIM(s.AssemblyNo)))
),
withBn AS (
    SELECT
        *,
        MAX([MACHINE_MEDIUM_CYCLETIME]) OVER (PARTITION BY [Build], [LINE], [Module#], [Side]) AS [Bottleneck]
    FROM joined
)
SELECT
    [Build],
    [Build Key],
    [Module#],
    [PCB],
    [LINE],
    [Side],
    [MACHINENAME],
    [BOARDSPROCESSED],
    [MACHINE_MIN_CYCLETIME],
    [MACHINE_MEDIUM_CYCLETIME],
    [ADM_CURRENT_CYCLETIME],
    [Bottleneck],
    [PANELENDTIME],
    [MachineCycletimeSource],
    [MachineSnapshotAt],
    [BuildToModuleSource],
    [BuildToModuleAt]
FROM withBn
ORDER BY [Build], [LINE], [Module#], [Side], [MACHINENAME];
""";
}
