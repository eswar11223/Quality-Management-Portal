using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.MandD
{
    public class MandDModel : PageModel
    {
        public string FirstName { get; set; } = "";

        public List<MandDRecord> Records { get; set; } = new();

        private readonly string SourceExcelFilePath =
            @"C:\QMSData\QR-014-MDTR Labelling (Milestone & Deliverable's Tracking Record).xlsx";

        private readonly string ExcelFilePath =
            @"C:\QMSData\MilestoneDeliverables.xlsx";

        private static readonly object ExcelLock = new object();

        public IActionResult OnGet()
        {
            string? username = HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";

            ImportSourceWorkbookIfNeeded();

            LoadRecordsFromExcel();

            return Page();
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }

        private void ImportSourceWorkbookIfNeeded()
        {
            lock (ExcelLock)
            {
                // Do not import again if the application workbook already exists.
                if (System.IO.File.Exists(ExcelFilePath))
                {
                    return;
                }

                if (!System.IO.File.Exists(SourceExcelFilePath))
                {
                    throw new Exception(
                        "Source M&D workbook was not found at: " +
                        SourceExcelFilePath);
                }

                Directory.CreateDirectory(@"C:\QMSData");

                using var sourceWorkbook = new XLWorkbook(SourceExcelFilePath);

                var sourceWorksheet =
                    sourceWorkbook.Worksheet("Milestone & Deliverables");

                var sourceUsedRange = sourceWorksheet.RangeUsed();

                if (sourceUsedRange == null)
                {
                    throw new Exception(
                        "The 'Milestone & Deliverables' sheet does not contain any data.");
                }

                // Find the actual header row.
                // The source workbook has logo/title rows above the table.
                var headerRow = FindHeaderRow(sourceUsedRange);

                // Find all required source columns.
                int sNoColumn =
                    FindColumn(headerRow, "S.No", "S.No.");

                int milestoneNameColumn =
                    FindColumn(headerRow, "Milestone Name");

                int deliverableNameColumn =
                    FindColumn(headerRow, "Deliverable Name");

                int deliverableTypeColumn =
                    FindColumn(headerRow, "Deliverable Type");

                int milestoneTargetDateColumn =
                    FindColumn(headerRow, "Milestone Target Date");

                int deliverableTargetDateColumn =
                    FindColumn(
                        headerRow,
                        "Deliverable Target Date",
                        "Deliverable Target date");

                int assignedToColumn =
                    FindColumn(
                        headerRow,
                        "Assigned To",
                        "Assigned to");

                int scheduleDeviationColumn =
                    FindColumn(
                        headerRow,
                        "Any Schedule Deviation");

                int activityColumn =
                    FindColumn(
                        headerRow,
                        "Deliverable at Which Activity",
                        "Deliverable at which Activity");

                int reasonColumn =
                    FindColumn(
                        headerRow,
                        "Reason for Schedule Deviation",
                        "Reson for Schedule Deviation");

                int actualCompletedDateColumn =
                    FindColumn(
                        headerRow,
                        "Deliverable Actual Completed Date",
                        "Deliverable Actual completed Date");

                int milestoneAchievedDateColumn =
                    FindColumn(
                        headerRow,
                        "Milestone Achieved Date",
                        "Milestone achieved Date");

                int statusColumn =
                    FindColumn(headerRow, "Status");

                int remarksColumn =
                    FindColumn(headerRow, "Remarks");

                // Create the application workbook.
                using var targetWorkbook = new XLWorkbook();

                var targetWorksheet =
                    targetWorkbook.Worksheets.Add("MilestoneDeliverables");

                // Required 15 columns.
                targetWorksheet.Cell(1, 1).Value = "Record ID";
                targetWorksheet.Cell(1, 2).Value = "S.No";
                targetWorksheet.Cell(1, 3).Value = "Milestone Name";
                targetWorksheet.Cell(1, 4).Value = "Deliverable Name";
                targetWorksheet.Cell(1, 5).Value = "Deliverable Type";
                targetWorksheet.Cell(1, 6).Value = "Milestone Target Date";
                targetWorksheet.Cell(1, 7).Value = "Deliverable Target Date";
                targetWorksheet.Cell(1, 8).Value = "Assigned To";
                targetWorksheet.Cell(1, 9).Value = "Any Schedule Deviation";
                targetWorksheet.Cell(1, 10).Value = "Deliverable at Which Activity";
                targetWorksheet.Cell(1, 11).Value = "Reason for Schedule Deviation";
                targetWorksheet.Cell(1, 12).Value = "Deliverable Actual Completed Date";
                targetWorksheet.Cell(1, 13).Value = "Milestone Achieved Date";
                targetWorksheet.Cell(1, 14).Value = "Status";
                targetWorksheet.Cell(1, 15).Value = "Remarks";

                targetWorksheet.Row(1).Style.Font.Bold = true;

                int targetRow = 2;
                int recordNumber = 1;

                // Read rows below the actual header row.
                foreach (var sourceRow in sourceWorksheet
                    .Rows(headerRow.RowNumber() + 1, sourceUsedRange.LastRow().RowNumber()))
                {
                    bool rowHasData =
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(sourceRow, milestoneNameColumn)) ||
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(sourceRow, deliverableNameColumn)) ||
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(sourceRow, deliverableTypeColumn));

                    if (!rowHasData)
                    {
                        continue;
                    }

                    // Generate permanent Record ID.
                    string recordId =
                        $"MD-{recordNumber:0000}";

                    targetWorksheet.Cell(targetRow, 1).Value = recordId;

                    // Original S.No remains unchanged.
                    targetWorksheet.Cell(targetRow, 2).Value =
                        GetCellValue(sourceRow, sNoColumn);

                    targetWorksheet.Cell(targetRow, 3).Value =
                        GetCellValue(sourceRow, milestoneNameColumn);

                    targetWorksheet.Cell(targetRow, 4).Value =
                        GetCellValue(sourceRow, deliverableNameColumn);

                    targetWorksheet.Cell(targetRow, 5).Value =
                        GetCellValue(sourceRow, deliverableTypeColumn);

                    targetWorksheet.Cell(targetRow, 6).Value =
                        GetCellValue(sourceRow, milestoneTargetDateColumn);

                    targetWorksheet.Cell(targetRow, 7).Value =
                        GetCellValue(sourceRow, deliverableTargetDateColumn);

                    targetWorksheet.Cell(targetRow, 8).Value =
                        GetCellValue(sourceRow, assignedToColumn);

                    targetWorksheet.Cell(targetRow, 9).Value =
                        GetCellValue(sourceRow, scheduleDeviationColumn);

                    targetWorksheet.Cell(targetRow, 10).Value =
                        GetCellValue(sourceRow, activityColumn);

                    targetWorksheet.Cell(targetRow, 11).Value =
                        GetCellValue(sourceRow, reasonColumn);

                    targetWorksheet.Cell(targetRow, 12).Value =
                        GetCellValue(sourceRow, actualCompletedDateColumn);

                    targetWorksheet.Cell(targetRow, 13).Value =
                        GetCellValue(sourceRow, milestoneAchievedDateColumn);

                    string status =
                        GetCellValue(sourceRow, statusColumn);

                    // Convert the source spelling to the required display value.
                    if (status.Equals(
                        "Deffered",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        status = "Deferred";
                    }

                    targetWorksheet.Cell(targetRow, 14).Value = status;

                    targetWorksheet.Cell(targetRow, 15).Value =
                        GetCellValue(sourceRow, remarksColumn);

                    targetRow++;
                    recordNumber++;
                }

                targetWorksheet.Columns().AdjustToContents();

                targetWorkbook.SaveAs(ExcelFilePath);
            }
        }

        private IXLRangeRow FindHeaderRow(IXLRange usedRange)
        {
            foreach (var row in usedRange.RowsUsed())
            {
                bool hasSNo = false;
                bool hasMilestoneName = false;
                bool hasDeliverableName = false;

                foreach (var cell in row.CellsUsed())
                {
                    string header =
                        NormalizeHeader(cell.GetString());

                    if (header == NormalizeHeader("S.No") ||
                        header == NormalizeHeader("S.No."))
                    {
                        hasSNo = true;
                    }

                    if (header == NormalizeHeader("Milestone Name"))
                    {
                        hasMilestoneName = true;
                    }

                    if (header == NormalizeHeader("Deliverable Name"))
                    {
                        hasDeliverableName = true;
                    }
                }

                if (hasSNo && hasMilestoneName && hasDeliverableName)
                {
                    return row;
                }
            }

            throw new Exception(
                "The M&D table header row could not be found in the source workbook.");
        }

        private string GetCellValue(
            IXLRow row,
            int columnNumber)
        {
            if (columnNumber <= 0)
            {
                return "";
            }

            return row.Cell(columnNumber)
                .GetFormattedString()
                .Trim();
        }

        private int FindColumn(
            IXLRangeRow headerRow,
            params string[] possibleNames)
        {
            foreach (var cell in headerRow.CellsUsed())
            {
                string actualHeader =
                    NormalizeHeader(cell.GetString());

                foreach (string possibleName in possibleNames)
                {
                    if (actualHeader ==
                        NormalizeHeader(possibleName))
                    {
                        return cell.Address.ColumnNumber;
                    }
                }
            }

            throw new Exception(
                "Required column was not found in the source workbook: " +
                string.Join(" / ", possibleNames));
        }

        private string NormalizeHeader(string value)
        {
            return new string(
                value
                    .Trim()
                    .ToLowerInvariant()
                    .Where(char.IsLetterOrDigit)
                    .ToArray());
        }

        private void LoadRecordsFromExcel()
        {
            Records.Clear();

            if (!System.IO.File.Exists(ExcelFilePath))
            {
                return;
            }

            lock (ExcelLock)
            {
                using var workbook = new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange = worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return;
                }

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    Records.Add(new MandDRecord
                    {
                        RecordId =
                            row.Cell(1).GetFormattedString(),

                        SNo =
                            row.Cell(2).GetFormattedString(),

                        MilestoneName =
                            row.Cell(3).GetFormattedString(),

                        DeliverableName =
                            row.Cell(4).GetFormattedString(),

                        DeliverableType =
                            row.Cell(5).GetFormattedString(),

                        MilestoneTargetDate =
                            row.Cell(6).GetFormattedString(),

                        DeliverableTargetDate =
                            row.Cell(7).GetFormattedString(),

                        AssignedTo =
                            row.Cell(8).GetFormattedString(),

                        ScheduleDeviation =
                            row.Cell(9).GetFormattedString(),

                        DeliverableActivity =
                            row.Cell(10).GetFormattedString(),

                        ScheduleDeviationReason =
                            row.Cell(11).GetFormattedString(),

                        ActualCompletedDate =
                            row.Cell(12).GetFormattedString(),

                        MilestoneAchievedDate =
                            row.Cell(13).GetFormattedString(),

                        Status =
                            row.Cell(14).GetFormattedString(),

                        Remarks =
                            row.Cell(15).GetFormattedString()
                    });
                }
            }
        }

        public class MandDRecord
        {
            public string RecordId { get; set; } = "";
            public string SNo { get; set; } = "";
            public string MilestoneName { get; set; } = "";
            public string DeliverableName { get; set; } = "";
            public string DeliverableType { get; set; } = "";
            public string MilestoneTargetDate { get; set; } = "";
            public string DeliverableTargetDate { get; set; } = "";
            public string AssignedTo { get; set; } = "";
            public string ScheduleDeviation { get; set; } = "";
            public string DeliverableActivity { get; set; } = "";
            public string ScheduleDeviationReason { get; set; } = "";
            public string ActualCompletedDate { get; set; } = "";
            public string MilestoneAchievedDate { get; set; } = "";
            public string Status { get; set; } = "";
            public string Remarks { get; set; } = "";
        }
    }
}