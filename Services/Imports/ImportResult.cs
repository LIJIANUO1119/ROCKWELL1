using System.Collections.Generic;

namespace SmtLineAllocationUI.Services.Imports;

public sealed record ImportResult(
    string FileName,
    int ReadCount,
    int SuccessCount,
    int FailureCount,
    IReadOnlyList<string> Errors
);

