using System;
using System.IO;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace SmtLineAllocationUI.DataAccess;

public sealed class SqlConnectionFactory
{
    private readonly string _connectionString;
    private readonly SqlConnectionStringBuilder _builder;

    public SqlConnectionFactory(string connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
            throw new ArgumentException("Connection string cannot be empty.", nameof(connectionString));

        _connectionString = connectionString;
        _builder = new SqlConnectionStringBuilder(connectionString);
    }

    public static SqlConnectionFactory FromAppSettings()
    {
        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .Build();

        var cs = config.GetConnectionString("SmtLineAllocation");
        if (string.IsNullOrWhiteSpace(cs))
            throw new InvalidDataException("Missing connection string 'ConnectionStrings:SmtLineAllocation' in appsettings.json.");

        return new SqlConnectionFactory(cs);
    }

    public SqlConnection Open()
    {
        var conn = new SqlConnection(_connectionString);
        conn.Open();
        return conn;
    }

    public string? DatabaseName
        => string.IsNullOrWhiteSpace(_builder.InitialCatalog) ? null : _builder.InitialCatalog;

    public SqlConnection OpenMaster()
    {
        var master = new SqlConnectionStringBuilder(_builder.ConnectionString)
        {
            InitialCatalog = "master"
        };

        var conn = new SqlConnection(master.ConnectionString);
        conn.Open();
        return conn;
    }
}

