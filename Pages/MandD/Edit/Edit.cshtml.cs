
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

        // Load the selected record for editing
        public IActionResult OnGet(string? recordId)
        {
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            RegisteredUsers = GetRegisteredUsers();

            if (string.IsNullOrWhiteSpace(recordId))
            {
                return Redirect("/MandD");
            }

            if (!LoadRecord(recordId))
            {
                return Redirect("/MandD");
            }

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

        // Save edited record
        public IActionResult OnPost()
        {
            string? username =
                HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            RegisteredUsers = GetRegisteredUsers();

            try
            {
                // Required-field validation
                if (string.IsNullOrWhiteSpace(RecordId))
                {
                    ModelState.AddModelError(
                        "",
                        "Record ID is missing.");
                }

                if (string.IsNullOrWhiteSpace(MilestoneName))
                {
                    ModelState.AddModelError(
                        "MilestoneName",
                        "Milestone Name is required.");
                }

                if (string.IsNullOrWhiteSpace(DeliverableName))
                {
                    ModelState.AddModelError(
                        "DeliverableName",
                        "Deliverable Name is required.");
                }

                if (string.IsNullOrWhiteSpace(DeliverableType))
                {
                    ModelState.AddModelError(
                        "DeliverableType",
                        "Deliverable Type is required.");
                }

                if (DeliverableType == "Other" &&
                    string.IsNullOrWhiteSpace(CustomDeliverableType))
                {
                    ModelState.AddModelError(
                        "CustomDeliverableType",
                        "Please enter the custom Deliverable Type.");
                }

                if (string.IsNullOrWhiteSpace(MilestoneTargetDate))
                {
                    ModelState.AddModelError(
                        "MilestoneTargetDate",
                        "Milestone Target Date is required.");
                }

                if (string.IsNullOrWhiteSpace(DeliverableTargetDate))
                {
                    ModelState.AddModelError(
                        "DeliverableTargetDate",
                        "Deliverable Target Date is required.");
                }

                if (string.IsNullOrWhiteSpace(AssignedTo))
                {
                    ModelState.AddModelError(
                        "AssignedTo",
                        "Assigned To is required.");
                }

                if (string.IsNullOrWhiteSpace(Status))
                {
                    ModelState.AddModelError(
                        "Status",
                        "Status is required.");
                }

                // Schedule Deviation validation
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
                        // Allow only the specified activities
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
                    DeliverableActivity = new List<string> { "N/A" };
                    ScheduleDeviationReason = "N/A";
                }
                else
                {
                    ModelState.AddModelError(
                        "ScheduleDeviation",
                        "Please select Yes or No.");
                }

                if (!ModelState.IsValid)
                {
                    return Page();
                }

                
                // Save custom Deliverable Type
                if (DeliverableType == "Other")
                {
                    DeliverableType = CustomDeliverableType.Trim();
                }

                
                // Automatically manage actual completion date
                if (Status == "Completed")
                {
                    if (string.IsNullOrWhiteSpace(ActualCompletedDate))
                    {
                        ActualCompletedDate = DateTime.Today.ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    ActualCompletedDate = "";
            }

                bool updated = UpdateRecord(username);

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

        // Logout
        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }

        // Load a record from the M&D Excel workbook
        private bool LoadRecord(string recordId)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    return false;
                }

                using var workbook =
                    new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return false;
                }

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string currentRecordId =
                        row.Cell(1).GetFormattedString().Trim();

                    if (!currentRecordId.Equals(
                        recordId.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    RecordId = currentRecordId;
                    SNo = row.Cell(2).GetFormattedString();
                    MilestoneName = row.Cell(3).GetFormattedString();
                    DeliverableName = row.Cell(4).GetFormattedString();
                    DeliverableType = row.Cell(5).GetFormattedString();
                    MilestoneTargetDate = row.Cell(6).GetFormattedString();
                    DeliverableTargetDate = row.Cell(7).GetFormattedString();
                    AssignedTo = row.Cell(8).GetFormattedString();
                    ScheduleDeviation = row.Cell(9).GetFormattedString();

                    // Read multiple activities from Excel
                    string activityValue =
                        row.Cell(10).GetFormattedString();

                    DeliverableActivity = activityValue
                        .Split(
                            ',',
                            StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .ToList();

                    ScheduleDeviationReason =
                        row.Cell(11).GetFormattedString();

                    ActualCompletedDate =
                        row.Cell(12).GetFormattedString();

                    MilestoneAchievedDate =
                        row.Cell(13).GetFormattedString();

                    Status = row.Cell(14).GetFormattedString();
                    Remarks = row.Cell(15).GetFormattedString();

                    return true;
                }

                return false;
            }
        }

        // Update the record and create audit entries
        private bool UpdateRecord(string username)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    return false;
                }

                using var workbook =
                    new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return false;
                }

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string currentRecordId =
                        row.Cell(1).GetFormattedString().Trim();

                    if (!currentRecordId.Equals(
                        RecordId.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Convert selected activities into Excel text
                    string activityValue =
                        string.Join(", ", DeliverableActivity);

                    // Capture the old values before updating
                    var changes = new List<(string Field, string OldValue, string NewValue)>
                    {
                        ("Milestone Name",
                            row.Cell(3).GetFormattedString(),
                            MilestoneName),

                        ("Deliverable Name",
                            row.Cell(4).GetFormattedString(),
                            DeliverableName),

                        ("Deliverable Type",
                            row.Cell(5).GetFormattedString(),
                            DeliverableType),

                        ("Milestone Target Date",
                            row.Cell(6).GetFormattedString(),
                            MilestoneTargetDate),

                        ("Deliverable Target Date",
                            row.Cell(7).GetFormattedString(),
                            DeliverableTargetDate),

                        ("Assigned To",
                            row.Cell(8).GetFormattedString(),
                            AssignedTo),

                        ("Any Schedule Deviation",
                            row.Cell(9).GetFormattedString(),
                            ScheduleDeviation),

                        ("Deliverable at Which Activity",
                            row.Cell(10).GetFormattedString(),
                            activityValue),

                        ("Reason for Schedule Deviation",
                            row.Cell(11).GetFormattedString(),
                            ScheduleDeviationReason),

                        ("Deliverable Actual Completed Date",
                            row.Cell(12).GetFormattedString(),
                            ActualCompletedDate),

                        ("Milestone Achieved Date",
                            row.Cell(13).GetFormattedString(),
                            MilestoneAchievedDate),

                        ("Status",
                            row.Cell(14).GetFormattedString(),
                            Status),

                        ("Remarks",
                            row.Cell(15).GetFormattedString(),
                            Remarks)
                    };

                    // Update the M&D Excel record
                    row.Cell(3).Value = MilestoneName;
                    row.Cell(4).Value = DeliverableName;
                    row.Cell(5).Value = DeliverableType;
                    row.Cell(6).Value = MilestoneTargetDate;
                    row.Cell(7).Value = DeliverableTargetDate;
                    row.Cell(8).Value = AssignedTo;
                    row.Cell(9).Value = ScheduleDeviation;
                    row.Cell(10).Value = activityValue;
                    row.Cell(11).Value = ScheduleDeviationReason;
                    row.Cell(12).Value = ActualCompletedDate;
                    row.Cell(13).Value = MilestoneAchievedDate;
                    row.Cell(14).Value = Status;
                    row.Cell(15).Value = Remarks;

                    // Save the M&D workbook
                    workbook.Save();

                    // Write one audit row for each changed field
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

                return false;
            }
        }

        // Add an audit entry only when a value changes
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
            {
                return;
            }

            CreateAuditFileIfNeeded();

            using var auditWorkbook =
                new XLWorkbook(AuditFilePath);

            var auditWorksheet =
                auditWorkbook.Worksheet("MDAuditLog");

            int nextRow =
                (auditWorksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

            string auditId =
                GenerateAuditId(auditWorksheet);

            auditWorksheet.Cell(nextRow, 1).Value = auditId;

            auditWorksheet.Cell(nextRow, 2).Value = DateTime.Now;
            auditWorksheet.Cell(nextRow, 2).Style.DateFormat.Format =
                "dd-MM-yyyy HH:mm:ss";

            auditWorksheet.Cell(nextRow, 3).Value = username;
            auditWorksheet.Cell(nextRow, 4).Value = recordId;
            auditWorksheet.Cell(nextRow, 5).Value = milestoneName;
            auditWorksheet.Cell(nextRow, 6).Value = deliverableName;
            auditWorksheet.Cell(nextRow, 7).Value = fieldChanged;
            auditWorksheet.Cell(nextRow, 8).Value = oldValue;
            auditWorksheet.Cell(nextRow, 9).Value = newValue;

            auditWorksheet.Columns().AdjustToContents();

            auditWorkbook.Save();
        }

        // Create the audit workbook if it does not exist
        private void CreateAuditFileIfNeeded()
        {
            if (System.IO.File.Exists(AuditFilePath))
            {
                return;
            }

            using var workbook = new XLWorkbook();

            var worksheet =
                workbook.Worksheets.Add("MDAuditLog");

            worksheet.Cell(1, 1).Value = "Audit ID";
            worksheet.Cell(1, 2).Value = "Date and Time";
            worksheet.Cell(1, 3).Value = "Logged-in Username";
            worksheet.Cell(1, 4).Value = "Record ID";
            worksheet.Cell(1, 5).Value = "Milestone Name";
            worksheet.Cell(1, 6).Value = "Deliverable Name";
            worksheet.Cell(1, 7).Value = "Field Changed";
            worksheet.Cell(1, 8).Value = "Old Value";
            worksheet.Cell(1, 9).Value = "New Value";

            worksheet.Range("A1:I1").Style.Font.Bold = true;

            worksheet.Range("A1:I1").Style.Fill
                .SetBackgroundColor(XLColor.LightGray);

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(AuditFilePath);
        }

        // Generate the next Audit ID
        private string GenerateAuditId(IXLWorksheet worksheet)
        {
            int lastRow =
                worksheet.LastRowUsed()?.RowNumber() ?? 1;

            int highestNumber = 0;

            for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
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

                string numberPart = auditId.Substring(4);

                if (int.TryParse(numberPart, out int number))
                {
                    if (number > highestNumber)
                    {
                        highestNumber = number;
                    }
                }
            }

            return $"AUD-{(highestNumber + 1):D4}";
        }

        // Read registered usernames from Users.xlsx
        private List<string> GetRegisteredUsers()
        {
            var users = new List<string>();

            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(UsersExcelFilePath))
                {
                    return users;
                }

                using var workbook =
                    new XLWorkbook(UsersExcelFilePath);

                var worksheet =
                    workbook.Worksheet("Users");

                var usedRange =
                    worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return users;
                }

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    // Username is in column 3 of Users.xlsx
                    string registeredUsername =
                        row.Cell(3).GetFormattedString().Trim();

                    if (!string.IsNullOrWhiteSpace(registeredUsername))
                    {
                        users.Add(registeredUsername);
                    }
                }
            }

            return users
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }
    }
}