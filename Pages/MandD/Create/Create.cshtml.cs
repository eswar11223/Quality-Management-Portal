
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.MandD.Create
{
    public class CreateModel : PageModel
    {
        public string FirstName { get; set; } = "";
        public List<string> RegisteredUsers { get; set; } = new();

        [BindProperty] public string MilestoneName { get; set; } = "";
        [BindProperty] public string DeliverableName { get; set; } = "";
        [BindProperty] public string DeliverableType { get; set; } = "";
        [BindProperty] public string CustomDeliverableType { get; set; } = "";
        [BindProperty] public string MilestoneTargetDate { get; set; } = "";
        [BindProperty] public string DeliverableTargetDate { get; set; } = "";
        [BindProperty] public string AssignedTo { get; set; } = "";
        [BindProperty] public string ScheduleDeviation { get; set; } = "";
        [BindProperty] public List<string> DeliverableActivity { get; set; } = new();
        [BindProperty] public string ScheduleDeviationReason { get; set; } = "";
        [BindProperty] public string ActualCompletedDate { get; set; } = "";
        [BindProperty] public string MilestoneAchievedDate { get; set; } = "";
        [BindProperty] public string Status { get; set; } = "Not Started";
        [BindProperty] public string Remarks { get; set; } = "";

        private readonly string ExcelFilePath =
            @"C:\QMSData\MilestoneDeliverables.xlsx";

        private readonly string AuditFilePath =
            @"C:\QMSData\MDAuditLog.xlsx";

        private readonly string UsersExcelFilePath =
            @"C:\QMSData\Users.xlsx";

        private static readonly object ExcelLock = new();

        public IActionResult OnGet()
        {
            if (!IsLoggedIn())
                return Redirect("/Login");

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";
            RegisteredUsers = GetRegisteredUsers();

            return Page();
        }

        public IActionResult OnPost()
        {
            if (!IsLoggedIn())
                return Redirect("/Login");

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";
            RegisteredUsers = GetRegisteredUsers();

            ValidateForm();

            if (!ModelState.IsValid)
                return Page();

            if (DeliverableType == "Other")
                DeliverableType = CustomDeliverableType.Trim();

            if (ScheduleDeviation == "No")
            {
                DeliverableActivity = new List<string> { "N/A" };
                ScheduleDeviationReason = "N/A";
            }

            if (Status == "Completed")
            {
                if (string.IsNullOrWhiteSpace(ActualCompletedDate))
                    ActualCompletedDate = DateTime.Today.ToString("yyyy-MM-dd");
            }
            else
            {
                ActualCompletedDate = "";
            }

            string username = HttpContext.Session.GetString("Username") ?? "";

            try
            {
                lock (ExcelLock)
                {
                    if (!System.IO.File.Exists(ExcelFilePath))
                    {
                        ModelState.AddModelError("",
                            "MilestoneDeliverables.xlsx was not found. Please ensure the M&D workbook has been initialized.");
                        return Page();
                    }

                    using var workbook = new XLWorkbook(ExcelFilePath);
                    var worksheet = workbook.Worksheet("MilestoneDeliverables");

                    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    int nextRow = Math.Max(2, lastRow + 1);

                    string recordId = GenerateRecordId(worksheet);
                    int nextSNo = GenerateNextSNo(worksheet);

                    worksheet.Cell(nextRow, 1).Value = recordId;
                    worksheet.Cell(nextRow, 2).Value = nextSNo;
                    worksheet.Cell(nextRow, 3).Value = MilestoneName.Trim();
                    worksheet.Cell(nextRow, 4).Value = DeliverableName.Trim();
                    worksheet.Cell(nextRow, 5).Value = DeliverableType.Trim();
                    worksheet.Cell(nextRow, 6).Value = MilestoneTargetDate.Trim();
                    worksheet.Cell(nextRow, 7).Value = DeliverableTargetDate.Trim();
                    worksheet.Cell(nextRow, 8).Value = AssignedTo.Trim();
                    worksheet.Cell(nextRow, 9).Value = ScheduleDeviation;
                    worksheet.Cell(nextRow, 10).Value =
                        string.Join(", ", DeliverableActivity);
                    worksheet.Cell(nextRow, 11).Value = ScheduleDeviationReason.Trim();
                    worksheet.Cell(nextRow, 12).Value = ActualCompletedDate.Trim();
                    worksheet.Cell(nextRow, 13).Value = MilestoneAchievedDate.Trim();
                    worksheet.Cell(nextRow, 14).Value = NormalizeStatus(Status);
                    worksheet.Cell(nextRow, 15).Value = Remarks.Trim();

                    workbook.Save();

                    WriteCreationAudit(
                        username, recordId, MilestoneName.Trim(),
                        DeliverableName.Trim());
                }

                return Redirect("/MandD");
            }
            catch (Exception)
            {
                ModelState.AddModelError("",
                    "The deliverable could not be saved. Check whether the Excel workbook is open or locked, then try again.");
                return Page();
            }
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();
            return Redirect("/Login");
        }

        private bool IsLoggedIn() =>
            !string.IsNullOrWhiteSpace(
                HttpContext.Session.GetString("Username"));

        private void ValidateForm()
        {
            if (string.IsNullOrWhiteSpace(MilestoneName))
                ModelState.AddModelError(nameof(MilestoneName), "Milestone Name is required.");

            if (string.IsNullOrWhiteSpace(DeliverableName))
                ModelState.AddModelError(nameof(DeliverableName), "Deliverable Name is required.");

            if (string.IsNullOrWhiteSpace(DeliverableType))
                ModelState.AddModelError(nameof(DeliverableType), "Deliverable Type is required.");

            if (DeliverableType == "Other" &&
                string.IsNullOrWhiteSpace(CustomDeliverableType))
                ModelState.AddModelError(nameof(CustomDeliverableType),
                    "Enter the custom deliverable type.");

            if (string.IsNullOrWhiteSpace(MilestoneTargetDate))
                ModelState.AddModelError(nameof(MilestoneTargetDate), "Milestone Target Date is required.");

            if (string.IsNullOrWhiteSpace(DeliverableTargetDate))
                ModelState.AddModelError(nameof(DeliverableTargetDate), "Deliverable Target Date is required.");

            if (string.IsNullOrWhiteSpace(AssignedTo))
                ModelState.AddModelError(nameof(AssignedTo), "Assigned To is required.");

            if (string.IsNullOrWhiteSpace(Status))
                ModelState.AddModelError(nameof(Status), "Status is required.");

            if (ScheduleDeviation != "Yes" && ScheduleDeviation != "No")
                ModelState.AddModelError(nameof(ScheduleDeviation), "Select Yes or No for Schedule Deviation.");

            if (ScheduleDeviation == "Yes")
            {
                string[] allowed = { "GN", "AdHoc", "Agile", "Other" };

                if (DeliverableActivity == null ||
                    DeliverableActivity.Count == 0 ||
                    DeliverableActivity.Any(x => !allowed.Contains(x)))
                    ModelState.AddModelError(nameof(DeliverableActivity),
                        "Select at least one valid activity.");

                if (string.IsNullOrWhiteSpace(ScheduleDeviationReason))
                    ModelState.AddModelError(nameof(ScheduleDeviationReason),
                        "Reason for Schedule Deviation is required.");
            }
        }

        private string GenerateRecordId(IXLWorksheet worksheet)
        {
            int highest = 0;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int r = 2; r <= lastRow; r++)
            {
                string id = worksheet.Cell(r, 1).GetString().Trim();

                if (id.StartsWith("MD-", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(id.Substring(3), out int number))
                    highest = Math.Max(highest, number);
            }

            return $"MD-{highest + 1:D4}";
        }

        private int GenerateNextSNo(IXLWorksheet worksheet)
        {
            int highest = 0;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int r = 2; r <= lastRow; r++)
            {
                if (int.TryParse(worksheet.Cell(r, 2).GetString(), out int number))
                    highest = Math.Max(highest, number);
            }

            return highest + 1;
        }

        private string NormalizeStatus(string status) =>
            status.Trim().Equals("Deffered", StringComparison.OrdinalIgnoreCase)
                ? "Deferred"
                : status.Trim();

        private List<string> GetRegisteredUsers()
        {
            var users = new List<string>();

            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(UsersExcelFilePath))
                    return users;

                using var workbook = new XLWorkbook(UsersExcelFilePath);
                var worksheet = workbook.Worksheet("Users");

                foreach (var row in worksheet.RangeUsed()?.RowsUsed().Skip(1)
                             ?? Enumerable.Empty<IXLRangeRow>())
                {
                    string user = row.Cell(3).GetString().Trim();

                    if (!string.IsNullOrWhiteSpace(user))
                        users.Add(user);
                }
            }

            return users.Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(x => x).ToList();
        }

        private void WriteCreationAudit(
            string username, string recordId,
            string milestoneName, string deliverableName)
        {
            if (!System.IO.File.Exists(AuditFilePath))
            {
                using var newWorkbook = new XLWorkbook();
                var sheet = newWorkbook.Worksheets.Add("MDAuditLog");

                string[] headers =
                {
                    "Audit ID", "Date and Time", "Logged-in Username",
                    "Record ID", "Milestone Name", "Deliverable Name",
                    "Field Changed", "Old Value", "New Value"
                };

                for (int i = 0; i < headers.Length; i++)
                    sheet.Cell(1, i + 1).Value = headers[i];

                newWorkbook.SaveAs(AuditFilePath);
            }

            using var auditWorkbook = new XLWorkbook(AuditFilePath);
            var audit = auditWorkbook.Worksheet("MDAuditLog");
            int row = (audit.LastRowUsed()?.RowNumber() ?? 1) + 1;

            int highest = 0;
            for (int r = 2; r < row; r++)
            {
                string id = audit.Cell(r, 1).GetString().Trim();
                if (id.StartsWith("AUD-", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(id.Substring(4), out int number))
                    highest = Math.Max(highest, number);
            }

            audit.Cell(row, 1).Value = $"AUD-{highest + 1:D4}";
            audit.Cell(row, 2).Value = DateTime.Now;
            audit.Cell(row, 2).Style.DateFormat.Format = "dd-MM-yyyy HH:mm:ss";
            audit.Cell(row, 3).Value = username;
            audit.Cell(row, 4).Value = recordId;
            audit.Cell(row, 5).Value = milestoneName;
            audit.Cell(row, 6).Value = deliverableName;
            audit.Cell(row, 7).Value = "Record Created";
            audit.Cell(row, 8).Value = "";
            audit.Cell(row, 9).Value = "New deliverable created";

            auditWorkbook.Save();
        }
    }
}
