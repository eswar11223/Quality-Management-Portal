using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.MandD.Edit
{
    public class EditModel : PageModel
    {
        public string FirstName { get; set; } = "";

        public List<string> RegisteredUsers { get; set; } = new();

        [BindProperty]
        public string RecordId { get; set; } = "";

        [BindProperty]
        public string SNo { get; set; } = "";

        [BindProperty]
        public string MilestoneName { get; set; } = "";

        [BindProperty]
        public string DeliverableName { get; set; } = "";

        [BindProperty]
        public string DeliverableType { get; set; } = "";

        [BindProperty]
        public string CustomDeliverableType { get; set; } = "";

        [BindProperty]
        public string MilestoneTargetDate { get; set; } = "";

        [BindProperty]
        public string DeliverableTargetDate { get; set; } = "";

        [BindProperty]
        public string AssignedTo { get; set; } = "";

        [BindProperty]
        public string ScheduleDeviation { get; set; } = "";

        [BindProperty]
        public List<string> DeliverableActivity { get; set; } = new();

        [BindProperty]
        public string ScheduleDeviationReason { get; set; } = "";

        [BindProperty]
        public string ActualCompletedDate { get; set; } = "";

        [BindProperty]
        public string MilestoneAchievedDate { get; set; } = "";

        [BindProperty]
        public string Status { get; set; } = "";

        [BindProperty]
        public string Remarks { get; set; } = "";

        private readonly string ExcelFilePath =
            @"C:\QMSData\MilestoneDeliverables.xlsx";

        private readonly string AuditFilePath =
            @"C:\QMSData\MDAuditLog.xlsx";

        private readonly string UsersExcelFilePath =
            @"C:\QMSData\Users.xlsx";

        private static readonly object ExcelLock = new object();

        public IActionResult OnGet(string? recordId)
        {
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
                return Redirect("/Login");

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            RegisteredUsers = GetRegisteredUsers();

            if (string.IsNullOrWhiteSpace(recordId))
                return Redirect("/MandD");

            if (!LoadRecord(recordId))
                return Redirect("/MandD");

            if (DeliverableType != "UGs" &&
                DeliverableType != "Clarification Doc" &&
                DeliverableType != "Data Sheets" &&
                DeliverableType != "Other" &&
                !string.IsNullOrWhiteSpace(DeliverableType))
            {
                CustomDeliverableType = DeliverableType;
                DeliverableType = "Other";
            }

            return Page();
        }

        public IActionResult OnPost()
        {
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
                return Redirect("/Login");

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            RegisteredUsers = GetRegisteredUsers();

            try
            {
                if (string.IsNullOrWhiteSpace(RecordId))
                    ModelState.AddModelError("", "Record ID is missing.");

                if (string.IsNullOrWhiteSpace(MilestoneName))
                    ModelState.AddModelError(
                        "MilestoneName",
                        "Milestone Name is required.");

                if (string.IsNullOrWhiteSpace(DeliverableName))
                    ModelState.AddModelError(
                        "DeliverableName",
                        "Deliverable Name is required.");

                if (string.IsNullOrWhiteSpace(DeliverableType))
                    ModelState.AddModelError(
                        "DeliverableType",
                        "Deliverable Type is required.");

                if (DeliverableType == "Other" &&
                    string.IsNullOrWhiteSpace(CustomDeliverableType))
                {
                    ModelState.AddModelError(
                        "CustomDeliverableType",
                        "Please enter the custom Deliverable Type.");
                }

                if (string.IsNullOrWhiteSpace(MilestoneTargetDate))
                    ModelState.AddModelError(
                        "MilestoneTargetDate",
                        "Milestone Target Date is required.");

                if (string.IsNullOrWhiteSpace(DeliverableTargetDate))
                    ModelState.AddModelError(
                        "DeliverableTargetDate",
                        "Deliverable Target Date is required.");

                if (string.IsNullOrWhiteSpace(AssignedTo))
                    ModelState.AddModelError(
                        "AssignedTo",
                        "Assigned To is required.");

                if (string.IsNullOrWhiteSpace(Status))
                    ModelState.AddModelError(
                        "Status",
                        "Status is required.");

                if (string.IsNullOrWhiteSpace(ScheduleDeviation))
                {
                    ModelState.AddModelError(
                        "ScheduleDeviation",
                        "Please select Any Schedule Deviation.");
                }
                else if (ScheduleDeviation == "Yes")
                {
                    if (DeliverableActivity == null ||
                        DeliverableActivity.Count == 0)
                    {
                        ModelState.AddModelError(
                            "DeliverableActivity",
                            "Please select at least one activity.");
                    }
                    else
                    {
                        string[] allowedActivities =
                        {
                            "GN",
                            "AdHoc",
                            "Agile",
                            "Other"
                        };

                        if (DeliverableActivity.Any(
                            activity => !allowedActivities.Contains(activity)))
                        {
                            ModelState.AddModelError(
                                "DeliverableActivity",
                                "Invalid activity selected.");
                        }
                    }

                    if (string.IsNullOrWhiteSpace(ScheduleDeviationReason))
                    {
                        ModelState.AddModelError(
                            "ScheduleDeviationReason",
                            "Reason for Schedule Deviation is required.");
                    }
                }
                else if (ScheduleDeviation == "No")
                {
                    DeliverableActivity = new List<string>
                    {
                        "N/A"
                    };

                    ScheduleDeviationReason = "N/A";
                }
                else
                {
                    ModelState.AddModelError(
                        "ScheduleDeviation",
                        "Please select Yes or No.");
                }

                if (!ModelState.IsValid)
                    return Page();

                if (DeliverableType == "Other")
                    DeliverableType = CustomDeliverableType.Trim();

                /*
                 * Deliverable Actual Completed Date
                 *
                 * When Status is Completed:
                 * - If the date is empty, today's date is added.
                 * - User can still manually edit the date.
                 *
                 * When Status is not Completed:
                 * - The date is cleared.
                 */
                if (Status == "Completed")
                {
                    if (string.IsNullOrWhiteSpace(ActualCompletedDate))
                    {
                        ActualCompletedDate =
                            DateTime.Today.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    ActualCompletedDate = "";
                }

                bool updated =
                    UpdateRecord(username);

                if (!updated)
                {
                    ModelState.AddModelError(
                        "",
                        "Record could not be found in MilestoneDeliverables.xlsx.");

                    return Page();
                }

                return Redirect("/MandD");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(
                    "",
                    "Save failed: " + ex.Message);

                return Page();
            }
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }

        private bool LoadRecord(string recordId)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                    return false;

                using var workbook =
                    new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                    return false;

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string currentRecordId =
                        row.Cell(1)
                           .GetFormattedString()
                           .Trim();

                    if (!currentRecordId.Equals(
                        recordId.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    RecordId = currentRecordId;

                    SNo =
                        row.Cell(2)
                           .GetFormattedString();

                    MilestoneName =
                        row.Cell(3)
                           .GetFormattedString();

                    DeliverableName =
                        row.Cell(4)
                           .GetFormattedString();

                    DeliverableType =
                        row.Cell(5)
                           .GetFormattedString();

                    MilestoneTargetDate =
                        row.Cell(6)
                           .GetFormattedString();

                    DeliverableTargetDate =
                        row.Cell(7)
                           .GetFormattedString();

                    AssignedTo =
                        row.Cell(8)
                           .GetFormattedString();

                    ScheduleDeviation =
                        row.Cell(9)
                           .GetFormattedString();

                    string activityValue =
                        row.Cell(10)
                           .GetFormattedString();

                    DeliverableActivity =
                        activityValue
                            .Split(
                                ',',
                                StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .Where(x =>
                                !string.IsNullOrWhiteSpace(x))
                            .ToList();

                    ScheduleDeviationReason =
                        row.Cell(11)
                           .GetFormattedString();

                    ActualCompletedDate =
                        row.Cell(12)
                           .GetFormattedString();

                    MilestoneAchievedDate =
                        row.Cell(13)
                           .GetFormattedString();

                    Status =
                        row.Cell(14)
                           .GetFormattedString();

                    Remarks =
                        row.Cell(15)
                           .GetFormattedString();

                    return true;
                }

                return false;
            }
        }

        private bool UpdateRecord(string username)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                    return false;

                using var workbook =
                    new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                    return false;

                IXLRangeRow? currentRow = null;

                /*
                 * First find the record being edited.
                 */
                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string currentRecordId =
                        row.Cell(1)
                           .GetFormattedString()
                           .Trim();

                    if (currentRecordId.Equals(
                        RecordId.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        currentRow = row;
                        break;
                    }
                }

                if (currentRow == null)
                    return false;

                /*
                 * Save the old values before changing anything.
                 */
                var changes =
                    new List<(string Field, string OldValue, string NewValue)>
                    {
                        (
                            "Milestone Name",
                            currentRow.Cell(3).GetFormattedString(),
                            MilestoneName
                        ),

                        (
                            "Deliverable Name",
                            currentRow.Cell(4).GetFormattedString(),
                            DeliverableName
                        ),

                        (
                            "Deliverable Type",
                            currentRow.Cell(5).GetFormattedString(),
                            DeliverableType
                        ),

                        (
                            "Milestone Target Date",
                            currentRow.Cell(6).GetFormattedString(),
                            MilestoneTargetDate
                        ),

                        (
                            "Deliverable Target Date",
                            currentRow.Cell(7).GetFormattedString(),
                            DeliverableTargetDate
                        ),

                        (
                            "Assigned To",
                            currentRow.Cell(8).GetFormattedString(),
                            AssignedTo
                        ),

                        (
                            "Any Schedule Deviation",
                            currentRow.Cell(9).GetFormattedString(),
                            ScheduleDeviation
                        ),

                        (
                            "Deliverable at Which Activity",
                            currentRow.Cell(10).GetFormattedString(),
                            string.Join(", ", DeliverableActivity)
                        ),

                        (
                            "Reason for Schedule Deviation",
                            currentRow.Cell(11).GetFormattedString(),
                            ScheduleDeviationReason
                        ),

                        (
                            "Deliverable Actual Completed Date",
                            currentRow.Cell(12).GetFormattedString(),
                            ActualCompletedDate
                        ),

                        (
                            "Milestone Achieved Date",
                            currentRow.Cell(13).GetFormattedString(),
                            MilestoneAchievedDate
                        ),

                        (
                            "Status",
                            NormalizeStatus(
                                currentRow.Cell(14).GetFormattedString()),
                            NormalizeStatus(Status)
                        ),

                        (
                            "Remarks",
                            currentRow.Cell(15).GetFormattedString(),
                            Remarks
                        )
                    };

                /*
                 * Save the current record.
                 */
                currentRow.Cell(3).Value =
                    MilestoneName;

                currentRow.Cell(4).Value =
                    DeliverableName;

                currentRow.Cell(5).Value =
                    DeliverableType;

                currentRow.Cell(6).Value =
                    MilestoneTargetDate;

                currentRow.Cell(7).Value =
                    DeliverableTargetDate;

                currentRow.Cell(8).Value =
                    AssignedTo;

                currentRow.Cell(9).Value =
                    ScheduleDeviation;

                string activityValue =
                    string.Join(", ", DeliverableActivity);

                currentRow.Cell(10).Value =
                    activityValue;

                currentRow.Cell(11).Value =
                    ScheduleDeviationReason;

                currentRow.Cell(12).Value =
                    ActualCompletedDate;

                currentRow.Cell(13).Value =
                    MilestoneAchievedDate;

                currentRow.Cell(14).Value =
                    NormalizeStatus(Status);

                currentRow.Cell(15).Value =
                    Remarks;

                /*
                 * Milestone Achieved Date automation.
                 *
                 * Check all deliverables belonging to the
                 * same milestone.
                 */
                bool milestoneCompleted =
                    AreAllDeliverablesCompleted(
                        worksheet,
                        MilestoneName);

                if (milestoneCompleted)
                {
                    string automaticMilestoneDate =
                        DateTime.Today.ToString("yyyy-MM-dd");

                    /*
                     * Set the date only when it is empty.
                     *
                     * This allows the user to manually edit
                     * the date without the system overwriting
                     * it every time the record is saved.
                     */
                    if (string.IsNullOrWhiteSpace(
                        currentRow.Cell(13)
                                  .GetFormattedString()))
                    {
                        currentRow.Cell(13).Value =
                            automaticMilestoneDate;

                        MilestoneAchievedDate =
                            automaticMilestoneDate;

                        changes.Add(
                            (
                                "Milestone Achieved Date",
                                "",
                                automaticMilestoneDate
                            ));
                    }
                }

                /*
                 * Save the workbook.
                 */
                workbook.Save();

                /*
                 * Write audit entries.
                 */
                foreach (var change in changes)
                {
                    AddAuditIfChanged(
                        username,
                        RecordId,
                        MilestoneName,
                        DeliverableName,
                        change.Field,
                        change.OldValue,
                        change.NewValue);
                }

                return true;
            }
        }

        private bool AreAllDeliverablesCompleted(
            IXLWorksheet worksheet,
            string milestoneName)
        {
            bool foundMilestoneRecord = false;

            foreach (var row in worksheet.RangeUsed()!
                                      .RowsUsed()
                                      .Skip(1))
            {
                string rowMilestoneName =
                    row.Cell(3)
                       .GetFormattedString()
                       .Trim();

                if (!rowMilestoneName.Equals(
                    milestoneName.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foundMilestoneRecord = true;

                string rowStatus =
                    row.Cell(14)
                       .GetFormattedString()
                       .Trim();

                rowStatus =
                    NormalizeStatus(rowStatus);

                if (!rowStatus.Equals(
                    "Completed",
                    StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return foundMilestoneRecord;
        }

        private string NormalizeStatus(string status)
        {
            status = status.Trim();

            /*
             * Requirement says imported "Deffered"
             * should display/store as "Deferred".
             */
            if (status.Equals(
                "Deffered",
                StringComparison.OrdinalIgnoreCase))
            {
                return "Deferred";
            }

            return status;
        }

        private void AddAuditIfChanged(
            string username,
            string recordId,
            string milestoneName,
            string deliverableName,
            string fieldChanged,
            string oldValue,
            string newValue)
        {
            if (oldValue == newValue)
                return;

            CreateAuditFileIfNeeded();

            using var auditWorkbook =
                new XLWorkbook(AuditFilePath);

            var auditWorksheet =
                auditWorkbook.Worksheet("MDAuditLog");

            int nextRow =
                (auditWorksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

            string auditId =
                GenerateAuditId(auditWorksheet);

            auditWorksheet.Cell(nextRow, 1).Value =
                auditId;

            auditWorksheet.Cell(nextRow, 2).Value =
                DateTime.Now;

            auditWorksheet.Cell(nextRow, 2)
                          .Style
                          .DateFormat
                          .Format =
                "dd-MM-yyyy HH:mm:ss";

            auditWorksheet.Cell(nextRow, 3).Value =
                username;

            auditWorksheet.Cell(nextRow, 4).Value =
                recordId;

            auditWorksheet.Cell(nextRow, 5).Value =
                milestoneName;

            auditWorksheet.Cell(nextRow, 6).Value =
                deliverableName;

            auditWorksheet.Cell(nextRow, 7).Value =
                fieldChanged;

            auditWorksheet.Cell(nextRow, 8).Value =
                oldValue;

            auditWorksheet.Cell(nextRow, 9).Value =
                newValue;

            auditWorksheet.Columns()
                          .AdjustToContents();

            auditWorkbook.Save();
        }

        private void CreateAuditFileIfNeeded()
        {
            if (System.IO.File.Exists(AuditFilePath))
                return;

            using var workbook =
                new XLWorkbook();

            var worksheet =
                workbook.Worksheets.Add("MDAuditLog");

            worksheet.Cell(1, 1).Value =
                "Audit ID";

            worksheet.Cell(1, 2).Value =
                "Date and Time";

            worksheet.Cell(1, 3).Value =
                "Logged-in Username";

            worksheet.Cell(1, 4).Value =
                "Record ID";

            worksheet.Cell(1, 5).Value =
                "Milestone Name";

            worksheet.Cell(1, 6).Value =
                "Deliverable Name";

            worksheet.Cell(1, 7).Value =
                "Field Changed";

            worksheet.Cell(1, 8).Value =
                "Old Value";

            worksheet.Cell(1, 9).Value =
                "New Value";

            worksheet.Range("A1:I1")
                     .Style
                     .Font
                     .Bold = true;

            worksheet.Range("A1:I1")
                     .Style
                     .Fill
                     .SetBackgroundColor(
                         XLColor.LightGray);

            worksheet.Columns()
                     .AdjustToContents();

            workbook.SaveAs(AuditFilePath);
        }

        private string GenerateAuditId(
            IXLWorksheet worksheet)
        {
            int lastRow =
                worksheet.LastRowUsed()?.RowNumber() ?? 1;

            int highestNumber = 0;

            for (int rowNumber = 2;
                 rowNumber <= lastRow;
                 rowNumber++)
            {
                string auditId =
                    worksheet.Cell(rowNumber, 1)
                             .GetString()
                             .Trim();

                if (!auditId.StartsWith(
                    "AUD-",
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string numberPart =
                    auditId.Substring(4);

                if (int.TryParse(
                    numberPart,
                    out int number))
                {
                    if (number > highestNumber)
                        highestNumber = number;
                }
            }

            return $"AUD-{(highestNumber + 1):D4}";
        }

        private List<string> GetRegisteredUsers()
        {
            var users =
                new List<string>();

            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(
                    UsersExcelFilePath))
                {
                    return users;
                }

                using var workbook =
                    new XLWorkbook(
                        UsersExcelFilePath);

                var worksheet =
                    workbook.Worksheet("Users");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                    return users;

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string registeredUsername =
                        row.Cell(3)
                           .GetFormattedString()
                           .Trim();

                    if (!string.IsNullOrWhiteSpace(
                        registeredUsername))
                    {
                        users.Add(
                            registeredUsername);
                    }
                }
            }

            return users
                .Distinct(
                    StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }
    }
}