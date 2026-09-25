using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.Register
{
    public class RegisterModel : PageModel
    {
        [BindProperty]
        public string FirstName { get; set; } = "";

        [BindProperty]
        public string LastName { get; set; } = "";

        [BindProperty]
        public string Username { get; set; } = "";

        [BindProperty]
        public string Password { get; set; } = "";

        [BindProperty]
        public string Email { get; set; } = "";

        [BindProperty]
        public string MobileNumber { get; set; } = "";

        [BindProperty]
        public string Otp { get; set; } = "";

        public bool OtpSent { get; set; }

        public string ErrorMessage { get; set; } = "";

        public string SuccessMessage { get; set; } = "";

        private readonly string ExcelFilePath = @"C:\QMSData\Users.xlsx";

        public void OnGet()
        {
        }

        public IActionResult OnPostSendOtp()
        {
            if (string.IsNullOrWhiteSpace(FirstName) ||
                string.IsNullOrWhiteSpace(LastName) ||
                string.IsNullOrWhiteSpace(Username) ||
                string.IsNullOrWhiteSpace(Password) ||
                string.IsNullOrWhiteSpace(Email) ||
                string.IsNullOrWhiteSpace(MobileNumber))
            {
                ErrorMessage = "Please fill in all fields.";
                return Page();
            }

            if (UserExists(Username, Email))
            {
                ErrorMessage = "Username or email already exists.";
                return Page();
            }

            OtpSent = true;

            return Page();
        }

        public IActionResult OnPostVerifyOtp()
        {
            OtpSent = true;

            if (Otp != "123456")
            {
                ErrorMessage = "Invalid OTP. Please enter 123456 for the demo.";
                return Page();
            }

            if (UserExists(Username, Email))
            {
                ErrorMessage = "Username or email already exists.";
                return Page();
            }

            SaveUser();

            return Redirect("/Login?registered=true");
        }

        private bool UserExists(string username, string email)
        {
            if (!System.IO.File.Exists(ExcelFilePath))
            {
                return false;
            }

            using var workbook = new XLWorkbook(ExcelFilePath);

            var worksheet = workbook.Worksheet("Users");

            var rows = worksheet.RangeUsed()?.RowsUsed();

            if (rows == null)
            {
                return false;
            }

            foreach (var row in rows.Skip(1))
            {
                string existingUsername = row.Cell(3).GetString();
                string existingEmail = row.Cell(5).GetString();

                if (existingUsername.Equals(username, StringComparison.OrdinalIgnoreCase) ||
                    existingEmail.Equals(email, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void SaveUser()
        {
            XLWorkbook workbook;

            if (System.IO.File.Exists(ExcelFilePath))
            {
                workbook = new XLWorkbook(ExcelFilePath);
            }
            else
            {
                workbook = new XLWorkbook();

                var worksheet = workbook.Worksheets.Add("Users");

                worksheet.Cell(1, 1).Value = "FirstName";
                worksheet.Cell(1, 2).Value = "LastName";
                worksheet.Cell(1, 3).Value = "Username";
                worksheet.Cell(1, 4).Value = "Password";
                worksheet.Cell(1, 5).Value = "Email";
                worksheet.Cell(1, 6).Value = "MobileNumber";
                worksheet.Cell(1, 7).Value = "Role";
            }

            var sheet = workbook.Worksheet("Users");

            int nextRow = (sheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

            sheet.Cell(nextRow, 1).Value = FirstName;
            sheet.Cell(nextRow, 2).Value = LastName;
            sheet.Cell(nextRow, 3).Value = Username;
            sheet.Cell(nextRow, 4).Value = Password;
            sheet.Cell(nextRow, 5).Value = Email;
            sheet.Cell(nextRow, 6).Value = MobileNumber;
            sheet.Cell(nextRow, 7).Value = "User";

            sheet.Columns().AdjustToContents();

            workbook.SaveAs(ExcelFilePath);

            workbook.Dispose();
        }
    }
}