using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.Login
{
    public class LoginModel : PageModel
    {
        private readonly string ExcelFilePath = @"C:\QMSData\Users.xlsx";

        private static readonly object ExcelLock = new object();

        [BindProperty]
        public string UsernameOrEmail { get; set; } = "";

        [BindProperty]
        public string Password { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public bool Registered { get; set; }

        [BindProperty(SupportsGet = true)]
        public bool PasswordReset { get; set; }

        public string ErrorMessage { get; set; } = "";

        public void OnGet()
        {
        }

        public IActionResult OnPost()
        {
            // Check Username or Email
            if (string.IsNullOrWhiteSpace(UsernameOrEmail))
            {
                ModelState.AddModelError(
                    "UsernameOrEmail",
                    "Username or Email is required."
                );
            }

            // Check Password
            if (string.IsNullOrWhiteSpace(Password))
            {
                ModelState.AddModelError(
                    "Password",
                    "Password is required."
                );
            }

            if (!ModelState.IsValid)
            {
                return Page();
            }

            try
            {
                // Check login details from Excel
                // First check whether Username or Email exists in Excel
            var user = GetUserFromExcel(UsernameOrEmail);

            if (user == null)
            {
                ErrorMessage = "No user data found";
                return Page();
            }

            // User exists, now check password
            if (user.Password != Password.Trim())
            {
                ErrorMessage = "Username or password is wrong.";
                return Page();
            }

            // Store login information in Session
            HttpContext.Session.SetString(
                "Username",
                user.Username
            );

            HttpContext.Session.SetString(
                "FirstName",
                user.FirstName
            );

            HttpContext.Session.SetString(
                "Role",
                user.Role
            );

            // Login successful
            return Redirect("/Welcome");            }
            catch (Exception ex)
            {
                ErrorMessage =
                    "Error reading Users.xlsx: " + ex.Message;

                return Page();
            }
        }

        // =====================================================
        // GET USER FROM EXCEL
        // =====================================================

        private UserDetails? GetUserFromExcel(
            string usernameOrEmail)
        {
            lock (ExcelLock)
            {
                // Check whether Excel file exists
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    throw new Exception(
                        "Users.xlsx was not found at C:\\QMSData\\Users.xlsx"
                    );
                }

                // Open Excel file
                using var workbook =
                    new XLWorkbook(ExcelFilePath);

                // Open Users sheet
                var worksheet =
                    workbook.Worksheet("Users");

                // Get all used rows
                var rows =
                    worksheet.RangeUsed()?.RowsUsed();

                if (rows == null)
                {
                    return null;
                }

                // Skip first row because it contains column names
                foreach (var row in rows.Skip(1))
                {
                    string firstName =
                        row.Cell(1).GetString().Trim();

                    string excelUsername =
                        row.Cell(3).GetString().Trim();

                    string excelPassword =
                        row.Cell(4).GetString().Trim();

                    string email =
                        row.Cell(5).GetString().Trim();

                    string role =
                        row.Cell(7).GetString().Trim();

                    // Check username
                    // Check username
                    bool usernameMatches =
                        excelUsername.Equals(
                            usernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        );

                    // Check email
                    bool emailMatches =
                        email.Equals(
                            usernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        );

                    // Username OR Email found
                    if (usernameMatches || emailMatches)
                    {
                        return new UserDetails
                    {
                        FirstName = firstName,
                        Username = excelUsername,
                        Password = excelPassword,
                        Role = role
                    };
                }
                }

                // No matching user
                return null;
            }
        }

        // =====================================================
        // USER DETAILS
        // =====================================================

        private class UserDetails
        {
            public string FirstName { get; set; } = "";

            public string Username { get; set; } = "";
            public string Password { get; set; } = "";

            public string Role { get; set; } = "";
            
        }
    }
}