using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.MandD.View
{
    public class ViewModel : PageModel
    {
        public string FirstName { get; set; } = "";

        public MandDRecord? Record { get; set; }

        private readonly string ExcelFilePath =
            @"C:\QMSData\MilestoneDeliverables.xlsx";

        private static readonly object ExcelLock = new object();

        public IActionResult OnGet(string? recordId)
        {
            string? username = HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";

            if (string.IsNullOrWhiteSpace(recordId))
            {
                return Redirect("/MandD");
            }

            Record = GetRecordFromExcel(recordId);

            if (Record == null)
            {
                return Redirect("/MandD");
            }

            return Page();
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }

        private MandDRecord? GetRecordFromExcel(string recordId)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    return null;
                }

                using var workbook = new XLWorkbook(ExcelFilePath);

                var worksheet =
                    workbook.Worksheet("MilestoneDeliverables");

                var usedRange = worksheet.RangeUsed();

                if (usedRange == null)
                {
                    return null;
                }

                foreach (var row in usedRange.RowsUsed().Skip(1))
                {
                    string currentRecordId =
                        row.Cell(1).GetFormattedString().Trim();

                    if (currentRecordId.Equals(
                        recordId.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return new MandDRecord
                        {
                            RecordId = currentRecordId,
                            SNo = row.Cell(2).GetFormattedString(),
                            MilestoneName = row.Cell(3).GetFormattedString(),
                            DeliverableName = row.Cell(4).GetFormattedString(),
                            DeliverableType = row.Cell(5).GetFormattedString(),
                            MilestoneTargetDate = row.Cell(6).GetFormattedString(),
                            DeliverableTargetDate = row.Cell(7).GetFormattedString(),
                            AssignedTo = row.Cell(8).GetFormattedString(),
                            ScheduleDeviation = row.Cell(9).GetFormattedString(),
                            DeliverableActivity = row.Cell(10).GetFormattedString(),
                            ScheduleDeviationReason = row.Cell(11).GetFormattedString(),
                            ActualCompletedDate = row.Cell(12).GetFormattedString(),
                            MilestoneAchievedDate = row.Cell(13).GetFormattedString(),
                            Status = row.Cell(14).GetFormattedString(),
                            Remarks = row.Cell(15).GetFormattedString()
                        };
                    }
                }

                return null;
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