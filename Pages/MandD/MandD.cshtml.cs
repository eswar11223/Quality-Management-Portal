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

        // =========================================================
        // SEARCH
        // =========================================================

        [BindProperty(SupportsGet = true)]
        public string SearchText { get; set; } = "";

        // =========================================================
        // FILTERS
        // =========================================================

        [BindProperty(SupportsGet = true)]
        public string FilterMilestoneName { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public string FilterAssignedTo { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public string FilterStatus { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public string FilterDeliverableType { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public string FilterScheduleDeviation { get; set; } = "";

        // =========================================================
        // DROPDOWN OPTIONS
        // =========================================================

        public List<string> MilestoneNames { get; set; } = new();

        public List<string> AssignedUsers { get; set; } = new();

        public List<string> Statuses { get; set; } = new();

        public List<string> DeliverableTypes { get; set; } = new();

        // =========================================================
        // GET M&D PAGE
        // =========================================================

        public IActionResult OnGet()
        {
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            ImportSourceWorkbookIfNeeded();

            LoadRecordsFromExcel();

            LoadFilterOptions();

            ApplySearchAndFilters();

            return Page();
        }

        // =========================================================
        // LOGOUT
        // =========================================================

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }

        // =========================================================
        // EXPORT TO EXCEL
        // =========================================================

        public IActionResult OnGetExportToExcel()
        {
            // Check whether the user is logged in.
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            try
            {
                lock (ExcelLock)
                {
                    // Make sure the application workbook exists.
                    if (!System.IO.File.Exists(ExcelFilePath))
                    {
                        return BadRequest(
                            "MilestoneDeliverables.xlsx was not found at: " +
                            ExcelFilePath);
                    }

                    // Open the current M&D workbook.
                    using var workbook =
                        new XLWorkbook(ExcelFilePath);

                    var worksheet =
                        workbook.Worksheet(
                            "MilestoneDeliverables");

                    var usedRange =
                        worksheet.RangeUsed();

                    if (usedRange == null)
                    {
                        return BadRequest(
                            "There are no M&D records available to export.");
                    }

                    // =================================================
                    // READ ALL RECORDS
                    // =================================================

                    List<MandDRecord> exportRecords =
                        new List<MandDRecord>();

                    foreach (var row in
                        usedRange.RowsUsed().Skip(1))
                    {
                        // Skip completely empty rows.
                        bool rowHasData =
                            !string.IsNullOrWhiteSpace(
                                row.Cell(3).GetFormattedString())
                            ||
                            !string.IsNullOrWhiteSpace(
                                row.Cell(4).GetFormattedString())
                            ||
                            !string.IsNullOrWhiteSpace(
                                row.Cell(5).GetFormattedString());

                        if (!rowHasData)
                        {
                            continue;
                        }

                        string status =
                            NormalizeStatus(
                                row.Cell(14)
                                    .GetFormattedString()
                                    .Trim());

                        exportRecords.Add(
                            new MandDRecord
                            {
                                RecordId =
                                    row.Cell(1)
                                        .GetFormattedString(),

                                SNo =
                                    row.Cell(2)
                                        .GetFormattedString(),

                                MilestoneName =
                                    row.Cell(3)
                                        .GetFormattedString(),

                                DeliverableName =
                                    row.Cell(4)
                                        .GetFormattedString(),

                                DeliverableType =
                                    row.Cell(5)
                                        .GetFormattedString(),

                                MilestoneTargetDate =
                                    row.Cell(6)
                                        .GetFormattedString(),

                                DeliverableTargetDate =
                                    row.Cell(7)
                                        .GetFormattedString(),

                                AssignedTo =
                                    row.Cell(8)
                                        .GetFormattedString(),

                                ScheduleDeviation =
                                    row.Cell(9)
                                        .GetFormattedString(),

                                DeliverableActivity =
                                    row.Cell(10)
                                        .GetFormattedString(),

                                ScheduleDeviationReason =
                                    row.Cell(11)
                                        .GetFormattedString(),

                                ActualCompletedDate =
                                    row.Cell(12)
                                        .GetFormattedString(),

                                MilestoneAchievedDate =
                                    row.Cell(13)
                                        .GetFormattedString(),

                                Status =
                                    status,

                                Remarks =
                                    row.Cell(15)
                                        .GetFormattedString(),

                                IsOverdue =
                                    IsRecordOverdue(
                                        row.Cell(7),
                                        status)
                            });
                    }

                    // =================================================
                    // APPLY SAME SEARCH
                    // =================================================

                    IEnumerable<MandDRecord> filteredRecords =
                        exportRecords;

                    if (!string.IsNullOrWhiteSpace(SearchText))
                    {
                        string search =
                            SearchText.Trim();

                        filteredRecords =
                            filteredRecords.Where(record =>
                                record.MilestoneName.Contains(
                                    search,
                                    StringComparison.OrdinalIgnoreCase)
                                ||
                                record.DeliverableName.Contains(
                                    search,
                                    StringComparison.OrdinalIgnoreCase)
                                ||
                                record.AssignedTo.Contains(
                                    search,
                                    StringComparison.OrdinalIgnoreCase)
                                ||
                                record.Status.Contains(
                                    search,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    // =================================================
                    // APPLY MILESTONE FILTER
                    // =================================================

                    if (!string.IsNullOrWhiteSpace(
                        FilterMilestoneName))
                    {
                        filteredRecords =
                            filteredRecords.Where(record =>
                                record.MilestoneName.Equals(
                                    FilterMilestoneName,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    // =================================================
                    // APPLY ASSIGNED TO FILTER
                    // =================================================

                    if (!string.IsNullOrWhiteSpace(
                        FilterAssignedTo))
                    {
                        filteredRecords =
                            filteredRecords.Where(record =>
                                record.AssignedTo.Equals(
                                    FilterAssignedTo,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    // =================================================
                    // APPLY STATUS FILTER
                    // =================================================

                    if (!string.IsNullOrWhiteSpace(
                        FilterStatus))
                    {
                        filteredRecords =
                            filteredRecords.Where(record =>
                                NormalizeStatus(
                                    record.Status)
                                .Equals(
                                    FilterStatus,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    // =================================================
                    // APPLY DELIVERABLE TYPE FILTER
                    // =================================================

                    if (!string.IsNullOrWhiteSpace(
                        FilterDeliverableType))
                    {
                        filteredRecords =
                            filteredRecords.Where(record =>
                                record.DeliverableType.Equals(
                                    FilterDeliverableType,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    // =================================================
                    // APPLY SCHEDULE DEVIATION FILTER
                    // =================================================

                    if (!string.IsNullOrWhiteSpace(
                        FilterScheduleDeviation))
                    {
                        filteredRecords =
                            filteredRecords.Where(record =>
                                record.ScheduleDeviation.Equals(
                                    FilterScheduleDeviation,
                                    StringComparison.OrdinalIgnoreCase));
                    }

                    List<MandDRecord> finalRecords =
                        filteredRecords.ToList();

                    // =================================================
                    // NO RECORDS
                    // =================================================

                    if (finalRecords.Count == 0)
                    {
                        return BadRequest(
                            "No M&D records match the selected search or filters.");
                    }

                    // =================================================
                    // CREATE NEW EXCEL WORKBOOK
                    // =================================================

                    using var exportWorkbook =
                        new XLWorkbook();

                    var exportWorksheet =
                        exportWorkbook.Worksheets.Add(
                            "M&D Export");

                    // =================================================
                    // HEADERS
                    // =================================================

                    string[] headers =
                    {
                        "Record ID",
                        "S.No",
                        "Milestone Name",
                        "Deliverable Name",
                        "Deliverable Type",
                        "Milestone Target Date",
                        "Deliverable Target Date",
                        "Assigned To",
                        "Any Schedule Deviation",
                        "Deliverable at Which Activity",
                        "Reason for Schedule Deviation",
                        "Deliverable Actual Completed Date",
                        "Milestone Achieved Date",
                        "Status",
                        "Remarks"
                    };

                    for (int column = 0;
                         column < headers.Length;
                         column++)
                    {
                        exportWorksheet.Cell(
                            1,
                            column + 1)
                            .Value = headers[column];
                    }

                    // =================================================
                    // HEADER STYLE
                    // =================================================

                    var headerRange =
                        exportWorksheet.Range(
                            1,
                            1,
                            1,
                            headers.Length);

                    headerRange.Style.Font.Bold = true;

                    headerRange.Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                    headerRange.Style.Alignment.Vertical =
                        XLAlignmentVerticalValues.Center;

                    // =================================================
                    // WRITE DATA
                    // =================================================

                    int exportRow = 2;

                    foreach (var record in finalRecords)
                    {
                        exportWorksheet.Cell(
                            exportRow,
                            1).Value =
                            record.RecordId;

                        exportWorksheet.Cell(
                            exportRow,
                            2).Value =
                            record.SNo;

                        exportWorksheet.Cell(
                            exportRow,
                            3).Value =
                            record.MilestoneName;

                        exportWorksheet.Cell(
                            exportRow,
                            4).Value =
                            record.DeliverableName;

                        exportWorksheet.Cell(
                            exportRow,
                            5).Value =
                            record.DeliverableType;

                        exportWorksheet.Cell(
                            exportRow,
                            6).Value =
                            record.MilestoneTargetDate;

                        exportWorksheet.Cell(
                            exportRow,
                            7).Value =
                            record.DeliverableTargetDate;

                        exportWorksheet.Cell(
                            exportRow,
                            8).Value =
                            record.AssignedTo;

                        exportWorksheet.Cell(
                            exportRow,
                            9).Value =
                            record.ScheduleDeviation;

                        exportWorksheet.Cell(
                            exportRow,
                            10).Value =
                            record.DeliverableActivity;

                        exportWorksheet.Cell(
                            exportRow,
                            11).Value =
                            record.ScheduleDeviationReason;

                        exportWorksheet.Cell(
                            exportRow,
                            12).Value =
                            record.ActualCompletedDate;

                        exportWorksheet.Cell(
                            exportRow,
                            13).Value =
                            record.MilestoneAchievedDate;

                        exportWorksheet.Cell(
                            exportRow,
                            14).Value =
                            record.Status;

                        exportWorksheet.Cell(
                            exportRow,
                            15).Value =
                            record.Remarks;

                        exportRow++;
                    }

                    // =================================================
                    // EXCEL FORMATTING
                    // =================================================

                    exportWorksheet
                        .Columns()
                        .AdjustToContents();

                    // Prevent extremely wide columns.
                    for (int column = 1;
                         column <= headers.Length;
                         column++)
                    {
                        if (exportWorksheet
                            .Column(column)
                            .Width > 50)
                        {
                            exportWorksheet
                                .Column(column)
                                .Width = 50;
                        }
                    }

                    // =================================================
                    // CREATE FILE IN MEMORY
                    // =================================================

                    using var memoryStream =
                        new MemoryStream();

                    exportWorkbook.SaveAs(
                        memoryStream);

                    byte[] fileBytes =
                        memoryStream.ToArray();

                    string fileName =
                        "MandD_Export_" +
                        DateTime.Now.ToString(
                            "yyyyMMdd_HHmmss") +
                        ".xlsx";

                    // =================================================
                    // DOWNLOAD FILE
                    // =================================================

                    return File(
                        fileBytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(
                    "Export failed: " +
                    ex.Message);
            }
        }

        // =========================================================
        // IMPORT SOURCE WORKBOOK
        // =========================================================

        private void ImportSourceWorkbookIfNeeded()
        {
            lock (ExcelLock)
            {
                // Do not import again if the application workbook
                // already exists.
                if (System.IO.File.Exists(ExcelFilePath))
                {
                    return;
                }

                if (!System.IO.File.Exists(
                    SourceExcelFilePath))
                {
                    throw new Exception(
                        "Source M&D workbook was not found at: " +
                        SourceExcelFilePath);
                }

                Directory.CreateDirectory(
                    @"C:\QMSData");

                using var sourceWorkbook =
                    new XLWorkbook(
                        SourceExcelFilePath);

                var sourceWorksheet =
                    sourceWorkbook.Worksheet(
                        "Milestone & Deliverables");

                var sourceUsedRange =
                    sourceWorksheet.RangeUsed();

                if (sourceUsedRange == null)
                {
                    throw new Exception(
                        "The 'Milestone & Deliverables' sheet does not contain any data.");
                }

                // Find actual header row.
                var headerRow =
                    FindHeaderRow(
                        sourceUsedRange);

                // Find required columns.
                int sNoColumn =
                    FindColumn(
                        headerRow,
                        "S.No",
                        "S.No.");

                int milestoneNameColumn =
                    FindColumn(
                        headerRow,
                        "Milestone Name");

                int deliverableNameColumn =
                    FindColumn(
                        headerRow,
                        "Deliverable Name");

                int deliverableTypeColumn =
                    FindColumn(
                        headerRow,
                        "Deliverable Type");

                int milestoneTargetDateColumn =
                    FindColumn(
                        headerRow,
                        "Milestone Target Date");

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
                    FindColumn(
                        headerRow,
                        "Status");

                int remarksColumn =
                    FindColumn(
                        headerRow,
                        "Remarks");

                // =================================================
                // CREATE APPLICATION WORKBOOK
                // =================================================

                using var targetWorkbook =
                    new XLWorkbook();

                var targetWorksheet =
                    targetWorkbook.Worksheets.Add(
                        "MilestoneDeliverables");

                // =================================================
                // 15 REQUIRED COLUMNS
                // =================================================

                targetWorksheet.Cell(1, 1).Value =
                    "Record ID";

                targetWorksheet.Cell(1, 2).Value =
                    "S.No";

                targetWorksheet.Cell(1, 3).Value =
                    "Milestone Name";

                targetWorksheet.Cell(1, 4).Value =
                    "Deliverable Name";

                targetWorksheet.Cell(1, 5).Value =
                    "Deliverable Type";

                targetWorksheet.Cell(1, 6).Value =
                    "Milestone Target Date";

                targetWorksheet.Cell(1, 7).Value =
                    "Deliverable Target Date";

                targetWorksheet.Cell(1, 8).Value =
                    "Assigned To";

                targetWorksheet.Cell(1, 9).Value =
                    "Any Schedule Deviation";

                targetWorksheet.Cell(1, 10).Value =
                    "Deliverable at Which Activity";

                targetWorksheet.Cell(1, 11).Value =
                    "Reason for Schedule Deviation";

                targetWorksheet.Cell(1, 12).Value =
                    "Deliverable Actual Completed Date";

                targetWorksheet.Cell(1, 13).Value =
                    "Milestone Achieved Date";

                targetWorksheet.Cell(1, 14).Value =
                    "Status";

                targetWorksheet.Cell(1, 15).Value =
                    "Remarks";

                targetWorksheet.Row(1)
                    .Style
                    .Font
                    .Bold = true;

                int targetRow = 2;

                int recordNumber = 1;

                foreach (var sourceRow in
                    sourceWorksheet.Rows(
                        headerRow.RowNumber() + 1,
                        sourceUsedRange.LastRow().RowNumber()))
                {
                    bool rowHasData =
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(
                                sourceRow,
                                milestoneNameColumn))
                        ||
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(
                                sourceRow,
                                deliverableNameColumn))
                        ||
                        !string.IsNullOrWhiteSpace(
                            GetCellValue(
                                sourceRow,
                                deliverableTypeColumn));

                    if (!rowHasData)
                    {
                        continue;
                    }

                    string recordId =
                        $"MD-{recordNumber:0000}";

                    targetWorksheet.Cell(
                        targetRow,
                        1).Value =
                        recordId;

                    targetWorksheet.Cell(
                        targetRow,
                        2).Value =
                        GetCellValue(
                            sourceRow,
                            sNoColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        3).Value =
                        GetCellValue(
                            sourceRow,
                            milestoneNameColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        4).Value =
                        GetCellValue(
                            sourceRow,
                            deliverableNameColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        5).Value =
                        GetCellValue(
                            sourceRow,
                            deliverableTypeColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        6).Value =
                        GetCellValue(
                            sourceRow,
                            milestoneTargetDateColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        7).Value =
                        GetCellValue(
                            sourceRow,
                            deliverableTargetDateColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        8).Value =
                        GetCellValue(
                            sourceRow,
                            assignedToColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        9).Value =
                        GetCellValue(
                            sourceRow,
                            scheduleDeviationColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        10).Value =
                        GetCellValue(
                            sourceRow,
                            activityColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        11).Value =
                        GetCellValue(
                            sourceRow,
                            reasonColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        12).Value =
                        GetCellValue(
                            sourceRow,
                            actualCompletedDateColumn);

                    targetWorksheet.Cell(
                        targetRow,
                        13).Value =
                        GetCellValue(
                            sourceRow,
                            milestoneAchievedDateColumn);

                    string status =
                        GetCellValue(
                            sourceRow,
                            statusColumn);

                    status =
                        NormalizeStatus(status);

                    targetWorksheet.Cell(
                        targetRow,
                        14).Value =
                        status;

                    targetWorksheet.Cell(
                        targetRow,
                        15).Value =
                        GetCellValue(
                            sourceRow,
                            remarksColumn);

                    targetRow++;

                    recordNumber++;
                }

                targetWorksheet
                    .Columns()
                    .AdjustToContents();

                targetWorkbook.SaveAs(
                    ExcelFilePath);
            }
        }

        // =========================================================
        // FIND HEADER ROW
        // =========================================================

        private IXLRangeRow FindHeaderRow(
            IXLRange usedRange)
        {
            foreach (var row in usedRange.RowsUsed())
            {
                bool hasSNo = false;

                bool hasMilestoneName = false;

                bool hasDeliverableName = false;

                foreach (var cell in row.CellsUsed())
                {
                    string header =
                        NormalizeHeader(
                            cell.GetString());

                    if (header ==
                        NormalizeHeader("S.No")
                        ||
                        header ==
                        NormalizeHeader("S.No."))
                    {
                        hasSNo = true;
                    }

                    if (header ==
                        NormalizeHeader("Milestone Name"))
                    {
                        hasMilestoneName = true;
                    }

                    if (header ==
                        NormalizeHeader("Deliverable Name"))
                    {
                        hasDeliverableName = true;
                    }
                }

                if (hasSNo &&
                    hasMilestoneName &&
                    hasDeliverableName)
                {
                    return row;
                }
            }

            throw new Exception(
                "The M&D table header row could not be found in the source workbook.");
        }

        // =========================================================
        // GET CELL VALUE
        // =========================================================

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

        // =========================================================
        // FIND COLUMN
        // =========================================================

        private int FindColumn(
            IXLRangeRow headerRow,
            params string[] possibleNames)
        {
            foreach (var cell in headerRow.CellsUsed())
            {
                string actualHeader =
                    NormalizeHeader(
                        cell.GetString());

                foreach (string possibleName
                    in possibleNames)
                {
                    if (actualHeader ==
                        NormalizeHeader(
                            possibleName))
                    {
                        return cell.Address.ColumnNumber;
                    }
                }
            }

            throw new Exception(
                "Required column was not found in the source workbook: " +
                string.Join(
                    " / ",
                    possibleNames));
        }

        // =========================================================
        // NORMALIZE HEADER
        // =========================================================

        private string NormalizeHeader(
            string value)
        {
            return new string(
                value
                    .Trim()
                    .ToLowerInvariant()
                    .Where(
                        char.IsLetterOrDigit)
                    .ToArray());
        }

        // =========================================================
        // LOAD RECORDS
        // =========================================================

        private void LoadRecordsFromExcel()
        {
            Records.Clear();

            if (!System.IO.File.Exists(
                ExcelFilePath))
            {
                return;
            }

            lock (ExcelLock)
            {
                using var workbook =
                    new XLWorkbook(
                        ExcelFilePath);

                var worksheet =
                    workbook.Worksheet(
                        "MilestoneDeliverables");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return;
                }

                foreach (var row in
                    usedRange.RowsUsed().Skip(1))
                {
                    string status =
                        row.Cell(14)
                            .GetFormattedString();

                    status =
                        NormalizeStatus(status);

                    string deliverableTargetDate =
                        row.Cell(7)
                            .GetFormattedString();

                    Records.Add(
                        new MandDRecord
                        {
                            RecordId =
                                row.Cell(1)
                                    .GetFormattedString(),

                            SNo =
                                row.Cell(2)
                                    .GetFormattedString(),

                            MilestoneName =
                                row.Cell(3)
                                    .GetFormattedString(),

                            DeliverableName =
                                row.Cell(4)
                                    .GetFormattedString(),

                            DeliverableType =
                                row.Cell(5)
                                    .GetFormattedString(),

                            MilestoneTargetDate =
                                row.Cell(6)
                                    .GetFormattedString(),

                            DeliverableTargetDate =
                                deliverableTargetDate,

                            AssignedTo =
                                row.Cell(8)
                                    .GetFormattedString(),

                            ScheduleDeviation =
                                row.Cell(9)
                                    .GetFormattedString(),

                            DeliverableActivity =
                                row.Cell(10)
                                    .GetFormattedString(),

                            ScheduleDeviationReason =
                                row.Cell(11)
                                    .GetFormattedString(),

                            ActualCompletedDate =
                                row.Cell(12)
                                    .GetFormattedString(),

                            MilestoneAchievedDate =
                                row.Cell(13)
                                    .GetFormattedString(),

                            Status =
                                status,

                            Remarks =
                                row.Cell(15)
                                    .GetFormattedString(),

                            IsOverdue =
                                IsRecordOverdue(
                                    row.Cell(7),
                                    status)
                        });
                }
            }
        }

        // =========================================================
        // NORMALIZE STATUS
        // =========================================================

        private string NormalizeStatus(
            string status)
        {
            if (status.Equals(
                "Deffered",
                StringComparison.OrdinalIgnoreCase))
            {
                return "Deferred";
            }

            return status;
        }

        // =========================================================
        // LOAD FILTER OPTIONS
        // =========================================================

        private void LoadFilterOptions()
        {
            MilestoneNames =
                Records
                    .Select(
                        x => x.MilestoneName)
                    .Where(
                        x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(
                        StringComparer.OrdinalIgnoreCase)
                    .OrderBy(
                        x => x)
                    .ToList();

            AssignedUsers =
                Records
                    .Select(
                        x => x.AssignedTo)
                    .Where(
                        x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(
                        StringComparer.OrdinalIgnoreCase)
                    .OrderBy(
                        x => x)
                    .ToList();

            Statuses =
                Records
                    .Select(
                        x => NormalizeStatus(x.Status))
                    .Where(
                        x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(
                        StringComparer.OrdinalIgnoreCase)
                    .OrderBy(
                        x => x)
                    .ToList();

            DeliverableTypes =
                Records
                    .Select(
                        x => x.DeliverableType)
                    .Where(
                        x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(
                        StringComparer.OrdinalIgnoreCase)
                    .OrderBy(
                        x => x)
                    .ToList();
        }

        // =========================================================
        // SEARCH AND FILTER
        // =========================================================

        private void ApplySearchAndFilters()
        {
            IEnumerable<MandDRecord> filtered =
                Records;

            // Combined search.
            if (!string.IsNullOrWhiteSpace(
                SearchText))
            {
                string search =
                    SearchText.Trim();

                filtered =
                    filtered.Where(record =>
                        record.MilestoneName.Contains(
                            search,
                            StringComparison.OrdinalIgnoreCase)
                        ||
                        record.DeliverableName.Contains(
                            search,
                            StringComparison.OrdinalIgnoreCase)
                        ||
                        record.AssignedTo.Contains(
                            search,
                            StringComparison.OrdinalIgnoreCase)
                        ||
                        record.Status.Contains(
                            search,
                            StringComparison.OrdinalIgnoreCase));
            }

            // Milestone filter.
            if (!string.IsNullOrWhiteSpace(
                FilterMilestoneName))
            {
                filtered =
                    filtered.Where(record =>
                        record.MilestoneName.Equals(
                            FilterMilestoneName,
                            StringComparison.OrdinalIgnoreCase));
            }

            // Assigned To filter.
            if (!string.IsNullOrWhiteSpace(
                FilterAssignedTo))
            {
                filtered =
                    filtered.Where(record =>
                        record.AssignedTo.Equals(
                            FilterAssignedTo,
                            StringComparison.OrdinalIgnoreCase));
            }

            // Status filter.
            if (!string.IsNullOrWhiteSpace(
                FilterStatus))
            {
                filtered =
                    filtered.Where(record =>
                        NormalizeStatus(
                            record.Status)
                            .Equals(
                                FilterStatus,
                                StringComparison.OrdinalIgnoreCase));
            }

            // Deliverable Type filter.
            if (!string.IsNullOrWhiteSpace(
                FilterDeliverableType))
            {
                filtered =
                    filtered.Where(record =>
                        record.DeliverableType.Equals(
                            FilterDeliverableType,
                            StringComparison.OrdinalIgnoreCase));
            }

            // Schedule Deviation filter.
            if (!string.IsNullOrWhiteSpace(
                FilterScheduleDeviation))
            {
                filtered =
                    filtered.Where(record =>
                        record.ScheduleDeviation.Equals(
                            FilterScheduleDeviation,
                            StringComparison.OrdinalIgnoreCase));
            }

            Records =
                filtered.ToList();
        }

        // =========================================================
        // OVERDUE CHECK
        // =========================================================

        private bool IsRecordOverdue(
            IXLCell dateCell,
            string status)
        {
            string normalizedStatus =
                NormalizeStatus(status);

            // Completed, Deferred and Cancelled
            // can never be overdue.
            if (normalizedStatus.Equals(
                    "Completed",
                    StringComparison.OrdinalIgnoreCase)
                ||
                normalizedStatus.Equals(
                    "Deferred",
                    StringComparison.OrdinalIgnoreCase)
                ||
                normalizedStatus.Equals(
                    "Cancelled",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (dateCell.IsEmpty())
            {
                return false;
            }

            // First try the actual Excel date value.
            if (dateCell.TryGetValue<DateTime>(
                out DateTime targetDate))
            {
                return targetDate.Date <
                       DateTime.Today;
            }

            // If the date is stored as text,
            // try to read the text value.
            string dateText =
                dateCell
                    .GetFormattedString()
                    .Trim();

            if (string.IsNullOrWhiteSpace(
                dateText))
            {
                return false;
            }

            if (DateTime.TryParse(
                dateText,
                out DateTime parsedDate))
            {
                return parsedDate.Date <
                       DateTime.Today;
            }

            return false;
        }

        // =========================================================
        // M&D RECORD MODEL
        // =========================================================

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

            public bool IsOverdue { get; set; }
        }
    }
}