
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.GlobalSearch
{
    public class GlobalSearchModel : PageModel
    {
        private readonly string ExcelFilePath =
            @"C:\QMSData\MilestoneDeliverables.xlsx";

        public string FirstName { get; set; } = "";
        public string Query { get; set; } = "";
        public List<SearchResult> Results { get; set; } = new();

        public class SearchResult
        {
            public string RecordId { get; set; } = "";
            public string MilestoneName { get; set; } = "";
            public string DeliverableName { get; set; } = "";
            public string DeliverableType { get; set; } = "";
            public string AssignedTo { get; set; } = "";
            public string Status { get; set; } = "";
            public string TargetDate { get; set; } = "";
        }

        public IActionResult OnGet(string? query)
        {
            if (string.IsNullOrWhiteSpace(
                HttpContext.Session.GetString("Username")))
            {
                return Redirect("/Login");
            }

            FirstName =
                HttpContext.Session.GetString("FirstName") ?? "";

            Query = query?.Trim() ?? "";

            if (!string.IsNullOrWhiteSpace(Query))
            {
                Results = FindRecords(Query);
            }

            return Page();
        }

        public IActionResult OnGetSuggestions(string? query)
        {
            if (string.IsNullOrWhiteSpace(
                HttpContext.Session.GetString("Username")))
            {
                return Unauthorized();
            }

            string search = query?.Trim() ?? "";

            if (search.Length < 2)
            {
                return new JsonResult(
                    Array.Empty<SearchResult>());
            }

            var results = FindRecords(search)
                .Take(8)
                .ToList();

            return new JsonResult(results);
        }

        private List<SearchResult> FindRecords(string search)
        {
            var results = new List<SearchResult>();

            if (!System.IO.File.Exists(ExcelFilePath))
            {
                return results;
            }

            using var workbook = new XLWorkbook(ExcelFilePath);

            var worksheet =
                workbook.Worksheet("MilestoneDeliverables");

            var usedRange = worksheet.RangeUsed();

            if (usedRange == null)
            {
                return results;
            }

            foreach (var row in usedRange.RowsUsed().Skip(1))
            {
                string recordId =
                    row.Cell(1).GetFormattedString().Trim();

                string milestone =
                    row.Cell(3).GetFormattedString().Trim();

                string deliverable =
                    row.Cell(4).GetFormattedString().Trim();

                string type =
                    row.Cell(5).GetFormattedString().Trim();

                string targetDate =
                    row.Cell(7).GetFormattedString().Trim();

                string assignedTo =
                    row.Cell(8).GetFormattedString().Trim();

                string deviation =
                    row.Cell(9).GetFormattedString().Trim();

                string deviationReason =
                    row.Cell(11).GetFormattedString().Trim();

                string actualCompletedDate =
                    row.Cell(12).GetFormattedString().Trim();

                string milestoneAchievedDate =
                    row.Cell(13).GetFormattedString().Trim();

                string status =
                    row.Cell(14).GetFormattedString().Trim();

                string remarks =
                    row.Cell(15).GetFormattedString().Trim();

                string activity =
                    row.Cell(10).GetFormattedString().Trim();

                string milestoneTargetDate =
                    row.Cell(6).GetFormattedString().Trim();

                bool matches =
                    Contains(recordId, search) ||
                    Contains(milestone, search) ||
                    Contains(deliverable, search) ||
                    Contains(type, search) ||
                    Contains(assignedTo, search) ||
                    Contains(status, search) ||
                    Contains(deviation, search) ||
                    Contains(deviationReason, search) ||
                    Contains(activity, search) ||
                    Contains(remarks, search) ||
                    Contains(targetDate, search) ||
                    Contains(milestoneTargetDate, search) ||
                    Contains(actualCompletedDate, search) ||
                    Contains(milestoneAchievedDate, search);

                if (!matches)
                {
                    continue;
                }

                results.Add(new SearchResult
                {
                    RecordId = recordId,
                    MilestoneName = milestone,
                    DeliverableName = deliverable,
                    DeliverableType = type,
                    AssignedTo = assignedTo,
                    Status = status,
                    TargetDate = targetDate
                });
            }

            return results;
        }

        private static bool Contains(
            string? value,
            string search)
        {
            return !string.IsNullOrEmpty(value) &&
                   value.Contains(
                       search,
                       StringComparison.OrdinalIgnoreCase);
        }
    }
}
