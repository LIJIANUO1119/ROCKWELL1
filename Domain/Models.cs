using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;

namespace SmtLineAllocationUI.Domain;

public sealed record Line(
    string LineId,
    string LineName,
    bool IsActive,
    LineConstraint? Constraint,
    ImmutableArray<Machine> Machines
);

public sealed record Machine(
    string MachineId,
    string LineId,
    string MachineName,
    int PositionInLine,
    bool IsActive,
    ImmutableArray<MachineNozzleConfig> Nozzles
);

public sealed record NozzleType(
    string NozzleTypeId,
    string NozzleModel,
    decimal? MinComponentHeightMm,
    decimal? MaxComponentHeightMm,
    string? Vendor
);

public sealed record MachineNozzleConfig(
    string MachineId,
    string NozzleTypeId,
    int SlotIndex,
    int Quantity
);

public sealed record LineConstraint(
    decimal? MaxBoardLengthMm,
    decimal? MaxBoardWidthMm,
    decimal? MaxComponentHeightMm,
    bool HighPrecisionSupport
);

public sealed record Product(
    string ProductId,
    string ProductCode,
    string? ProductName,
    string? FamilyGroup,
    BoardSpec Board,
    ProductRequirement Requirement,
    TimeSpan? TargetCycleTime
);

public sealed record BoardSpec(
    decimal? LengthMm,
    decimal? WidthMm
);

public sealed record ProductRequirement(
    bool HasHighComponents,
    decimal? MaxComponentHeightMm,
    bool RequiresHighPrecision
);

public sealed record Component(
    string ComponentId,
    string PartNumber,
    decimal? HeightMm,
    string? PrecisionClass
);

public sealed record ProductComponent(
    string ProductId,
    string ComponentId,
    int QuantityPerBoard
);

public sealed record ComponentNozzleRequirement(
    string ComponentId,
    string NozzleTypeId
);

public sealed record ProductionHistory(
    string ProductId,
    string LineId,
    DateTimeOffset StartTime,
    DateTimeOffset EndTime,
    int BoardsProduced,
    TimeSpan AvgCycleTime,
    bool IsSimulated
);

public static class DomainValidation
{
    public static void EnsureNotNullOrWhiteSpace(string value, string paramName)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Value cannot be null/empty.", paramName);
    }

    public static ImmutableArray<T> ToImmutableArrayOrEmpty<T>(IEnumerable<T>? values)
        => values is null ? ImmutableArray<T>.Empty : values.ToImmutableArray();
}

