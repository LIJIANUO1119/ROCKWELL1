using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using SmtLineAllocationUI.DataAccess;

namespace SmtLineAllocationUI.Services.Imports;

public sealed class MachineSoftwareVerExcelImporter
{
    private readonly SqlConnectionFactory _connectionFactory;

    public MachineSoftwareVerExcelImporter(SqlConnectionFactory connectionFactory)
    {
        _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
    }

    public async Task<ImportResult> ImportAsync(string xlsxPath, string sheetName, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath)) throw new ArgumentException("Path cannot be empty.", nameof(xlsxPath));
        if (!File.Exists(xlsxPath)) throw new FileNotFoundException("Excel file not found.", xlsxPath);
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));

        var errors = new List<string>();
        var read = 0;
        var ok = 0;
        var failed = 0;

        // Import strategy for SMT LINE CONFIGURATION matrix view (Excel columns):
        // - Process Line  -> dbo.Line.LineCode
        // - Machine ID    -> dbo.Machine.MachineCode + MachineName
        // - Machine model -> dbo.Machine.MachineModel
        // - Equip Type    -> dbo.Machine.EquipType (also mirrors into MachineType for compatibility)
        // - Serial Number/Computer Name, Software Version, IP Address, DNS, Gateway, OS
        //   -> stored on dbo.Machine for direct matrix display
        // - Optional legacy columns (e.g. Software Level) are still stored into dbo.MachineSoftware when present.

        await using var conn = _connectionFactory.Open();
        await using var tx = (SqlTransaction)await conn.BeginTransactionAsync(ct);

        string? tempCopyPath = null;
        try
        {
            // If the user has the Excel open, Windows may lock the file.
            // Workaround: copy to a temp file with shared read and open the copy.
            tempCopyPath = await CopyToTempAsync(xlsxPath, ct);

            using var wb = new XLWorkbook(tempCopyPath);
            var ws = wb.Worksheet(sheetName);

            // Header row is row 1.
            var headerRow = ws.Row(1);
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var cell in headerRow.CellsUsed())
            {
                var name = cell.GetString()?.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    var s = cell.Value.ToString()?.Trim();
                    name = s;
                }
                if (string.IsNullOrWhiteSpace(name)) continue;
                headers[name] = cell.Address.ColumnNumber;
            }

            int ColAny(params string[] names)
            {
                foreach (var n in names)
                {
                    if (headers.TryGetValue(n, out var c)) return c;
                }
                throw new InvalidDataException($"Missing required column '{names[0]}'.");
            }

            int ColOptional(params string[] names)
            {
                foreach (var n in names)
                {
                    if (headers.TryGetValue(n, out var c)) return c;
                }
                return -1;
            }

            int ColOptionalFuzzy(string token1, string? token2)
            {
                token1 = token1?.Trim() ?? "";
                token2 = token2?.Trim();
                if (string.IsNullOrWhiteSpace(token1)) return -1;

                foreach (var (k, v) in headers)
                {
                    var kk = k ?? "";
                    // header text is already trimmed, but we still allow flexible whitespace.
                    if (!kk.Contains(token1, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!string.IsNullOrWhiteSpace(token2) && !kk.Contains(token2, StringComparison.OrdinalIgnoreCase)) continue;
                    return v;
                }

                return -1;
            }

            var colLine = ColAny("Process Line");
            var colMachineId = ColAny("Machine ID");
            var colMachineModel = ColAny("Machine model", "Machine Model");
            var colEquipType = ColAny("Equip Type");
            static string NormalizeHeaderName(string s)
            {
                // Make "Serial Number/Computer Name", "Serial Number / Computer", etc comparable.
                // Keep digits + letters only.
                var chars = (s ?? string.Empty).Trim().ToUpperInvariant().Where(ch => char.IsLetterOrDigit(ch)).ToArray();
                return new string(chars);
            }

            int ColByTokens(params string[] tokens)
            {
                if (tokens is null || tokens.Length == 0) return -1;
                foreach (var (k, v) in headers)
                {
                    var kk = NormalizeHeaderName(k ?? "");
                    var all = true;
                    foreach (var t in tokens)
                    {
                        var nt = NormalizeHeaderName(t);
                        if (string.IsNullOrWhiteSpace(nt) || !kk.Contains(nt, StringComparison.OrdinalIgnoreCase))
                        {
                            all = false;
                            break;
                        }
                    }
                    if (all) return v;
                }
                return -1;
            }

            // Prefer exact meaning: Serial + Computer both present in the header.
            var colSerial = ColByTokens("SERIAL", "COMPUTER");
            if (colSerial < 0) colSerial = ColByTokens("SERIAL");
            if (colSerial < 0) colSerial = ColByTokens("COMPUTER");
            var colSoftwareVersion = ColOptional("Software Version");
            var colSoftwareLevel = ColOptional("Software Level");
            var colIp = ColOptional("IP Address");
            var colDns = ColOptional("DNS");
            var colGateway = ColOptional("Gateway");
            var colOs = ColOptional("OS");

            static string CellToString(IXLCell cell)
            {
                var s = cell.GetString()?.Trim();
                if (!string.IsNullOrWhiteSpace(s)) return s;
                return cell.Value.ToString()?.Trim() ?? "";
            }

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            string? lastLineCode = null;
            string? lastMachineIdCol = null;
            string? lastSerialOrComputer = null;
            for (var r = 2; r <= lastRow; r++)
            {
                ct.ThrowIfCancellationRequested();
                read++;

                try
                {
                    var row = ws.Row(r);
                    var lineCodeRaw = CellToString(row.Cell(colLine));
                    if (!string.IsNullOrWhiteSpace(lineCodeRaw))
                        lastLineCode = lineCodeRaw;
                    var lineCode = string.IsNullOrWhiteSpace(lineCodeRaw) ? (lastLineCode ?? "") : lineCodeRaw;

                    var machineIdColRaw = CellToString(row.Cell(colMachineId));  // Machine ID -> display name
                    if (!string.IsNullOrWhiteSpace(machineIdColRaw))
                        lastMachineIdCol = machineIdColRaw;
                    var machineIdCol = string.IsNullOrWhiteSpace(machineIdColRaw) ? (lastMachineIdCol ?? "") : machineIdColRaw;

                    var machineModel = colMachineModel > 0 ? CellToString(row.Cell(colMachineModel)) : "";
                    var equipType = colEquipType > 0 ? CellToString(row.Cell(colEquipType)) : "";
                    var serialOrComputerRaw = colSerial > 0 ? CellToString(row.Cell(colSerial)) : "";
                    var serialOrComputer = string.IsNullOrWhiteSpace(serialOrComputerRaw) ? lastSerialOrComputer ?? "" : serialOrComputerRaw;
                    if (!string.IsNullOrWhiteSpace(serialOrComputerRaw))
                        lastSerialOrComputer = serialOrComputerRaw;
                    var softwareVersion = colSoftwareVersion > 0 ? CellToString(row.Cell(colSoftwareVersion)) : "";
                    var softwareLevel = colSoftwareLevel > 0 ? CellToString(row.Cell(colSoftwareLevel)) : "";
                    var ip = colIp > 0 ? CellToString(row.Cell(colIp)) : "";
                    var dns = colDns > 0 ? CellToString(row.Cell(colDns)) : "";
                    var gateway = colGateway > 0 ? CellToString(row.Cell(colGateway)) : "";
                    var os = colOs > 0 ? CellToString(row.Cell(colOs)) : "";

                    // Many excel exports use merged cells; blank rows/sections should be skipped rather than treated as failures.
                    if (string.IsNullOrWhiteSpace(lineCode) && string.IsNullOrWhiteSpace(machineIdCol))
                        continue;
                    if (string.IsNullOrWhiteSpace(lineCode) || string.IsNullOrWhiteSpace(machineIdCol))
                        continue;

                    var lineId = await EnsureLineAsync(conn, tx, lineCode, ct);
                    var machineId = await EnsureMachineAsync(
                        conn,
                        tx,
                        lineId,
                        machineCode: machineIdCol,
                        machineName: machineIdCol,
                        machineModel: machineModel,
                        equipType: equipType,
                        serialNumberOrComputerName: serialOrComputer,
                        softwareVersion: softwareVersion,
                        ipAddress: ip,
                        dns: dns,
                        gateway: gateway,
                        os: os,
                        ct
                    );

                    // Optional: save software info when Serial/Software Level present
                    var softwareName = string.IsNullOrWhiteSpace(equipType) ? "MachineSoftware" : equipType;
                    var fileOrNumber = !string.IsNullOrWhiteSpace(ip) ? ip : serialOrComputer;
                    if (!string.IsNullOrWhiteSpace(softwareLevel) || !string.IsNullOrWhiteSpace(fileOrNumber))
                    {
                        await UpsertMachineSoftwareAsync(conn, tx, machineId, softwareName, softwareLevel, fileOrNumber, ct);
                    }

                    ok++;
                }
                catch (Exception ex)
                {
                    failed++;
                    errors.Add($"Row {r}: {ex.Message}");
                }
            }

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }
        finally
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(tempCopyPath) && File.Exists(tempCopyPath))
                    File.Delete(tempCopyPath);
            }
            catch
            {
                // ignore temp cleanup failure
            }
        }

        return new ImportResult(Path.GetFileName(xlsxPath), read, ok, failed, errors);
    }

    private static async Task<string> CopyToTempAsync(string sourcePath, CancellationToken ct)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "SmtLineAllocationUI");
        Directory.CreateDirectory(tempDir);

        var tempPath = Path.Combine(
            tempDir,
            $"{Path.GetFileNameWithoutExtension(sourcePath)}_{DateTime.Now:yyyyMMdd_HHmmss_fff}{Path.GetExtension(sourcePath)}"
        );

        await using var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var dst = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        await src.CopyToAsync(dst, 1024 * 1024, ct);
        return tempPath;
    }

    private static async Task<int> EnsureLineAsync(SqlConnection conn, SqlTransaction tx, string lineCode, CancellationToken ct)
    {
        await using (var cmd = new SqlCommand("SELECT LineId FROM dbo.Line WHERE LineCode=@code;", conn, tx))
        {
            cmd.Parameters.AddWithValue("@code", lineCode);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id) return id;
        }

        await using (var cmd = new SqlCommand("""
INSERT INTO dbo.Line (LineCode, LineName) VALUES (@code, @name);
SELECT CAST(SCOPE_IDENTITY() AS int);
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@code", lineCode);
            cmd.Parameters.AddWithValue("@name", lineCode);
            var id = await cmd.ExecuteScalarAsync(ct);
            return (int)id!;
        }
    }

    private static async Task<int> EnsureMachineAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int lineId,
        string machineCode,
        string? machineName,
        string? machineModel,
        string? equipType,
        string? serialNumberOrComputerName,
        string? softwareVersion,
        string? ipAddress,
        string? dns,
        string? gateway,
        string? os,
        CancellationToken ct
    )
    {
        await using (var cmd = new SqlCommand("""
SELECT MachineId FROM dbo.Machine WHERE LineId=@lineId AND MachineCode=@code;
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@lineId", lineId);
            cmd.Parameters.AddWithValue("@code", machineCode);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id)
            {
                await using var up = new SqlCommand("""
UPDATE dbo.Machine
SET MachineName = COALESCE(NULLIF(@name,''), MachineName),
    -- Store "Machine model" into MachineType (temporary) to avoid MachineModel write issues.
    MachineType = COALESCE(NULLIF(@model,''), MachineType),
    EquipType = COALESCE(NULLIF(@equipType,''), EquipType),
    IsActive = 1, -- mark as active because machine appears in the imported machine list
    SerialNumberOrComputerName = COALESCE(NULLIF(@serial,''), SerialNumberOrComputerName),
    SoftwareVersion = COALESCE(NULLIF(@swVer,''), SoftwareVersion),
    IpAddress = COALESCE(NULLIF(@ip,''), IpAddress),
    Dns = COALESCE(NULLIF(@dns,''), Dns),
    Gateway = COALESCE(NULLIF(@gateway,''), Gateway),
    Os = COALESCE(NULLIF(@os,''), Os),
    UpdatedAt = SYSDATETIME()
WHERE MachineId=@id;
""", conn, tx);
                up.Parameters.AddWithValue("@id", id);
                up.Parameters.AddWithValue("@name", machineName ?? "");
                up.Parameters.AddWithValue("@equipType", equipType ?? "");
                up.Parameters.AddWithValue("@model", machineModel ?? "");
                up.Parameters.AddWithValue("@serial", serialNumberOrComputerName ?? "");
                up.Parameters.AddWithValue("@swVer", softwareVersion ?? "");
                up.Parameters.AddWithValue("@ip", ipAddress ?? "");
                up.Parameters.AddWithValue("@dns", dns ?? "");
                up.Parameters.AddWithValue("@gateway", gateway ?? "");
                up.Parameters.AddWithValue("@os", os ?? "");
                await up.ExecuteNonQueryAsync(ct);
                return id;
            }
        }

        await using (var cmd = new SqlCommand("""
INSERT INTO dbo.Machine
    (LineId, MachineCode, MachineName, MachineType, EquipType,
     SerialNumberOrComputerName, SoftwareVersion, IpAddress, Dns, Gateway, Os,
     PositionInLine, IsActive)
VALUES
    (@lineId, @code, @name, @model, @equipType,
     @serial, @swVer, @ip, @dns, @gateway, @os,
     NULL, 1);
SELECT CAST(SCOPE_IDENTITY() AS int);
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@lineId", lineId);
            cmd.Parameters.AddWithValue("@code", machineCode);
            cmd.Parameters.AddWithValue("@name", (object?)(string.IsNullOrWhiteSpace(machineName) ? machineCode : machineName) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@equipType", (object?)(string.IsNullOrWhiteSpace(equipType) ? null : equipType) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@model", (object?)machineModel ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@serial", (object?)(string.IsNullOrWhiteSpace(serialNumberOrComputerName) ? null : serialNumberOrComputerName) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@swVer", (object?)(string.IsNullOrWhiteSpace(softwareVersion) ? null : softwareVersion) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ip", (object?)(string.IsNullOrWhiteSpace(ipAddress) ? null : ipAddress) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@dns", (object?)(string.IsNullOrWhiteSpace(dns) ? null : dns) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@gateway", (object?)(string.IsNullOrWhiteSpace(gateway) ? null : gateway) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@os", (object?)(string.IsNullOrWhiteSpace(os) ? null : os) ?? DBNull.Value);
            var id = await cmd.ExecuteScalarAsync(ct);
            return (int)id!;
        }
    }

    private static async Task UpsertMachineSoftwareAsync(
        SqlConnection conn,
        SqlTransaction tx,
        int machineId,
        string softwareName,
        string? softwareVersion,
        string? filePathOrNumber,
        CancellationToken ct
    )
    {
        await using (var cmd = new SqlCommand("""
SELECT MachineSoftwareId FROM dbo.MachineSoftware
WHERE MachineId=@mid AND SoftwareName=@name;
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@mid", machineId);
            cmd.Parameters.AddWithValue("@name", softwareName);
            var existing = await cmd.ExecuteScalarAsync(ct);
            if (existing is int id)
            {
                await using var up = new SqlCommand("""
UPDATE dbo.MachineSoftware
SET SoftwareVersion = COALESCE(NULLIF(@ver,''), SoftwareVersion),
    FilePathOrNumber = COALESCE(NULLIF(@fp,''), FilePathOrNumber),
    UpdatedAt = SYSDATETIME()
WHERE MachineSoftwareId=@id;
""", conn, tx);
                up.Parameters.AddWithValue("@id", id);
                up.Parameters.AddWithValue("@ver", softwareVersion ?? "");
                up.Parameters.AddWithValue("@fp", filePathOrNumber ?? "");
                await up.ExecuteNonQueryAsync(ct);
                return;
            }
        }

        await using (var cmd = new SqlCommand("""
INSERT INTO dbo.MachineSoftware (MachineId, SoftwareName, SoftwareVersion, FilePathOrNumber)
VALUES (@mid, @name, @ver, @fp);
""", conn, tx))
        {
            cmd.Parameters.AddWithValue("@mid", machineId);
            cmd.Parameters.AddWithValue("@name", softwareName);
            cmd.Parameters.AddWithValue("@ver", (object?)(string.IsNullOrWhiteSpace(softwareVersion) ? null : softwareVersion) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@fp", (object?)(string.IsNullOrWhiteSpace(filePathOrNumber) ? null : filePathOrNumber) ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
        }
    }
}

