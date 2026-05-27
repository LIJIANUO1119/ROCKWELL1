using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using SmtLineAllocationUI.Domain;

namespace SmtLineAllocationUI.Services;

public sealed record CycleTimeSummary(
    string ProductId,
    string LineId,
    int SampleCount,
    TimeSpan Mean,
    TimeSpan P50,
    TimeSpan P90,
    TimeSpan Min,
    TimeSpan Max
);

public sealed record CycleTimeTargetEvaluation(
    string ProductId,
    string LineId,
    TimeSpan Target,
    TimeSpan Actual,
    bool ExceedsTarget
);

public sealed class CycleTimeStatisticsService
{
    public ImmutableArray<CycleTimeSummary> SummarizeByProductAndLine(IEnumerable<ProductionHistory> histories)
    {
        if (histories is null) throw new ArgumentNullException(nameof(histories));

        return histories
            .GroupBy(h => (h.ProductId, h.LineId), StringTupleComparer.Ordinal)
            .Select(g =>
            {
                var samples = g.Select(x => x.AvgCycleTime).ToArray();
                if (samples.Length == 0)
                {
                    return new CycleTimeSummary(
                        g.Key.ProductId, g.Key.LineId, 0,
                        TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero, TimeSpan.Zero
                    );
                }

                Array.Sort(samples);
                var meanTicks = (long)samples.Select(s => s.Ticks).Average();
                return new CycleTimeSummary(
                    g.Key.ProductId,
                    g.Key.LineId,
                    samples.Length,
                    TimeSpan.FromTicks(meanTicks),
                    PercentileSorted(samples, 0.50),
                    PercentileSorted(samples, 0.90),
                    samples[0],
                    samples[^1]
                );
            })
            .ToImmutableArray();
    }

    public ImmutableArray<CycleTimeTargetEvaluation> EvaluateTargets(
        IEnumerable<Product> products,
        IEnumerable<CycleTimeSummary> summaries
    )
    {
        if (products is null) throw new ArgumentNullException(nameof(products));
        if (summaries is null) throw new ArgumentNullException(nameof(summaries));

        var targetByProductId = products
            .Where(p => p.TargetCycleTime is not null)
            .ToDictionary(p => p.ProductId, p => p.TargetCycleTime!.Value, StringComparer.Ordinal);

        return summaries
            .Where(s => targetByProductId.ContainsKey(s.ProductId))
            .Select(s =>
            {
                var target = targetByProductId[s.ProductId];
                var actual = s.Mean;
                return new CycleTimeTargetEvaluation(
                    s.ProductId,
                    s.LineId,
                    target,
                    actual,
                    actual > target
                );
            })
            .ToImmutableArray();
    }

    public ImmutableArray<string> ProductsExceedingTarget(
        IEnumerable<CycleTimeTargetEvaluation> evals
    )
    {
        if (evals is null) throw new ArgumentNullException(nameof(evals));

        return evals
            .Where(e => e.ExceedsTarget)
            .Select(e => e.ProductId)
            .Distinct(StringComparer.Ordinal)
            .ToImmutableArray();
    }

    private static TimeSpan PercentileSorted(TimeSpan[] sorted, double p)
    {
        if (sorted.Length == 0) return TimeSpan.Zero;
        if (p <= 0) return sorted[0];
        if (p >= 1) return sorted[^1];

        var idx = p * (sorted.Length - 1);
        var lo = (int)Math.Floor(idx);
        var hi = (int)Math.Ceiling(idx);
        if (lo == hi) return sorted[lo];

        var t = idx - lo;
        var loTicks = sorted[lo].Ticks;
        var hiTicks = sorted[hi].Ticks;
        var lerp = loTicks + (long)Math.Round((hiTicks - loTicks) * t);
        return TimeSpan.FromTicks(lerp);
    }

    private sealed class StringTupleComparer : IEqualityComparer<(string ProductId, string LineId)>
    {
        public static readonly StringTupleComparer Ordinal = new(StringComparer.Ordinal);

        private readonly StringComparer _cmp;

        private StringTupleComparer(StringComparer cmp) => _cmp = cmp;

        public bool Equals((string ProductId, string LineId) x, (string ProductId, string LineId) y)
            => _cmp.Equals(x.ProductId, y.ProductId) && _cmp.Equals(x.LineId, y.LineId);

        public int GetHashCode((string ProductId, string LineId) obj)
            => HashCode.Combine(_cmp.GetHashCode(obj.ProductId), _cmp.GetHashCode(obj.LineId));
    }
}

