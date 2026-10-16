
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using MySqlConnector;
using System.Globalization;

string filePath = @"C:\QMSData\MilestoneDeliverables.xlsx";

if (!File.Exists(filePath))
{
    Console.WriteLine($"Excel file not found: {filePath}");
    return;
}

var configuration = new ConfigurationBuilder()
    .AddUserSecrets<Program>()
    .Build();

string? connectionString =
    configuration.GetConnectionString("DefaultConnection");

if (string.IsNullOrWhiteSpace(connectionString))
{
    Console.WriteLine("Database connection string not found.");
    return;
}

using var workbook = new XLWorkbook(filePath);
var worksheet = workbook.Worksheet("MilestoneDeliverables");
int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

static object DbText(IXLCell cell)
{
    string value = cell.GetString().Trim();
    return string.IsNullOrEmpty(value) ? DBNull.Value : value;
}

static object DbInt(IXLCell cell)
{
    if (cell.IsEmpty()) return DBNull.Value;
    return cell.TryGetValue<int>(out int value)
        ? value : DBNull.Value;
}



static object DbDate(IXLCell cell)
{
    if (cell.IsEmpty())
        return DBNull.Value;

    if (cell.TryGetValue<DateTime>(out DateTime date))
        return date.Date;

    string value = cell.GetString().Trim();

    if (string.IsNullOrWhiteSpace(value) ||
        value.Equals("NA", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("N.A.", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("NULL", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("-", StringComparison.OrdinalIgnoreCase) ||
        !DateTime.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out date))
    {
        return DBNull.Value;
    }

    return date.Date;
}


const string sql = """
INSERT INTO `QMSPortalDb`.`MilestoneDeliverables`
(
    `S.No`,
    `Milestone Name`,
    `Deliverable Name`,
    `Deliverable Type`,
    `Milestone Target Date`,
    `Deliverable Target Date`,
    `Assigned To`,
    `Any Schedule Deviation`,
    `Deliverable at Which Activity`,
    `Reason for Schedule Deviation`,
    `Deliverable Actual Completed Date`,
    `Milestone Achieved Date`,
    `Status`,
    `Remarks`
)
VALUES
(
    @serial, @milestone, @deliverable, @type,
    @milestoneDate, @deliverableDate, @assigned,
    @deviation, @activity, @reason, @completedDate,
    @achievedDate, @status, @remarks
);
""";

await using var connection = new MySqlConnection(connectionString);
await connection.OpenAsync();

await using var transaction = await connection.BeginTransactionAsync();

try
{
    int imported = 0;
    int skipped = 0;
    int serialNumber = 0;

    for (int row = 2; row <= lastRow; row++)
    {
        var milestone = worksheet.Cell(row, 3).GetString().Trim();
        var deliverable = worksheet.Cell(row, 4).GetString().Trim();

        // Skip completely empty rows.
        if (worksheet.Row(row).CellsUsed().Count() == 0)
        {
            skipped++;
            continue;
        }

        // Assign a serial number whenever Deliverable Name is filled.
        object serialValue = DBNull.Value;

        if (!string.IsNullOrWhiteSpace(deliverable))
        {
            serialNumber++;
            serialValue = serialNumber;
        }
        

        await using var command = new MySqlCommand(sql, connection, transaction);

        command.Parameters.AddWithValue("@serial", serialValue);
        command.Parameters.AddWithValue("@milestone", milestone);
        command.Parameters.AddWithValue("@deliverable", deliverable);
        command.Parameters.AddWithValue("@type", DbText(worksheet.Cell(row, 5)));
        command.Parameters.AddWithValue("@milestoneDate", DbDate(worksheet.Cell(row, 6)));
        command.Parameters.AddWithValue("@deliverableDate", DbDate(worksheet.Cell(row, 7)));
        command.Parameters.AddWithValue("@assigned", DbText(worksheet.Cell(row, 8)));
        command.Parameters.AddWithValue("@deviation", DbText(worksheet.Cell(row, 9)));
        command.Parameters.AddWithValue("@activity", DbText(worksheet.Cell(row, 10)));
        command.Parameters.AddWithValue("@reason", DbText(worksheet.Cell(row, 11)));
        command.Parameters.AddWithValue("@completedDate", DbDate(worksheet.Cell(row, 12)));
        command.Parameters.AddWithValue("@achievedDate", DbDate(worksheet.Cell(row, 13)));
        command.Parameters.AddWithValue("@status", DbText(worksheet.Cell(row, 14)));
        command.Parameters.AddWithValue("@remarks", DbText(worksheet.Cell(row, 15)));

        await command.ExecuteNonQueryAsync();
        imported++;
    }

    await transaction.CommitAsync();

    Console.WriteLine($"Import successful. Records imported: {imported}");
    Console.WriteLine($"Empty rows skipped: {skipped}");
}
catch (Exception ex)
{
    await transaction.RollbackAsync();
    Console.WriteLine("Import failed. Changes were rolled back.");
    Console.WriteLine(ex.Message);
}
