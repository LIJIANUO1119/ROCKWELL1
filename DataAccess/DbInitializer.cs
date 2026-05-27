using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace SmtLineAllocationUI.DataAccess;

public sealed class DbInitializer
{
    private readonly SqlConnectionFactory _connectionFactory;

    public DbInitializer(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task EnsureDatabaseAndSchemaAsync(string schemaSqlRelativePath, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(schemaSqlRelativePath))
            throw new ArgumentException("Schema SQL path cannot be empty.", nameof(schemaSqlRelativePath));

        var dbName = _connectionFactory.DatabaseName;
        if (string.IsNullOrWhiteSpace(dbName))
            throw new InvalidOperationException("Connection string must specify a Database/Initial Catalog.");

        await EnsureDatabaseExistsAsync(dbName, ct);

        var schemaPath = Path.Combine(AppContext.BaseDirectory, schemaSqlRelativePath);
        if (!File.Exists(schemaPath))
            throw new FileNotFoundException($"Schema file not found at '{schemaPath}'.", schemaPath);

        var rawSql = await File.ReadAllTextAsync(schemaPath, Encoding.UTF8, ct);
        var batches = SplitOnGo(rawSql).Where(b => !string.IsNullOrWhiteSpace(b)).ToArray();

        await using var conn = _connectionFactory.Open();
        foreach (var batch in batches)
        {
            ct.ThrowIfCancellationRequested();
            await using var cmd = new SqlCommand(batch, conn)
            {
                CommandTimeout = 120
            };

            try
            {
                await cmd.ExecuteNonQueryAsync(ct);
            }
            catch (SqlException ex) when (IsIgnorableCreateExistsError(ex))
            {
                // Schema already applied (e.g. "There is already an object named 'User'").
                // Safe to ignore so that EnsureDatabaseAndSchemaAsync can be called many times.
            }
        }
    }

    private async Task EnsureDatabaseExistsAsync(string dbName, CancellationToken ct)
    {
        // Connect to master to create database if missing.
        await using var master = _connectionFactory.OpenMaster();

        await using var cmd = new SqlCommand("""
IF DB_ID(@dbName) IS NULL
BEGIN
    DECLARE @sql nvarchar(max) = N'CREATE DATABASE [' + REPLACE(@dbName, ']', ']]') + N']';
    EXEC(@sql);
END
""", master)
        {
            CommandTimeout = 120
        };

        cmd.Parameters.AddWithValue("@dbName", dbName);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static IEnumerable<string> SplitOnGo(string sql)
    {
        using var reader = new StringReader(sql);
        var sb = new StringBuilder();
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (string.Equals(line.Trim(), "GO", StringComparison.OrdinalIgnoreCase))
            {
                yield return sb.ToString();
                sb.Clear();
                continue;
            }

            sb.AppendLine(line);
        }

        if (sb.Length > 0)
            yield return sb.ToString();
    }

    private static bool IsIgnorableCreateExistsError(SqlException ex)
    {
        // SQL Server error 2714: "There is already an object named '%.*ls' in the database."
        if (ex.Number == 2714)
            return true;

        // If it's a batch with multiple errors, check all.
        foreach (SqlError err in ex.Errors)
        {
            if (err.Number != 2714)
                return false;
        }

        return ex.Errors.Count > 0;
    }
}

