using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.ForgotPassword
{
    public class ForgotPasswordModel : PageModel
    {
        private readonly string ExcelFilePath = @"C:\QMSData\Users.xlsx";

        private static readonly object ExcelLock = new object();

        [BindProperty]
        public string UsernameOrEmail { get; set; } = "";

        [BindProperty]
        public string MobileNumber { get; set; } = "";

        [BindProperty]
        public string Otp { get; set; } = "";

        [BindProperty]
        public string NewPassword { get; set; } = "";

        [BindProperty]
        public string ConfirmPassword { get; set; } = "";

        public bool OtpSent { get; set; }

        public bool PasswordReset { get; set; }

        public string ErrorMessage { get; set; } = "";

        public string SuccessMessage { get; set; } = "";

        public void OnGet()
        {
        }

        // SEND OTP
        public IActionResult OnPostSendOtp()
        {
            if (string.IsNullOrWhiteSpace(UsernameOrEmail))
            {
                ErrorMessage = "Username or Email is required.";
                return Page();
            }

            if (string.IsNullOrWhiteSpace(MobileNumber))
            {
                ErrorMessage = "Mobile Number is required.";
                return Page();
            }

            try
            {
                if (!UserExists(UsernameOrEmail, MobileNumber))
                {
                    ErrorMessage = "Username/Email or Mobile Number is incorrect.";
                    return Page();
                }

                // Demo OTP
                OtpSent = true;

                SuccessMessage = "OTP sent successfully. Demo OTP is 123456.";

                return Page();
            }
            catch (Exception ex)
            {
                ErrorMessage = "Error reading Users.xlsx: " + ex.Message;
                return Page();
            }
        }

        // VERIFY OTP
        public IActionResult OnPostVerifyOtp()
        {
            OtpSent = true;

            if (Otp != "123456")
            {
                ErrorMessage = "Invalid OTP. Please enter 123456.";
                return Page();
            }

            PasswordReset = true;

            SuccessMessage = "OTP verified successfully.";

            return Page();
        }

        // RESET PASSWORD
        public IActionResult OnPostResetPassword()
        {
            PasswordReset = true;

            if (string.IsNullOrWhiteSpace(NewPassword))
            {
                ErrorMessage = "New Password is required.";
                return Page();
            }

            if (string.IsNullOrWhiteSpace(ConfirmPassword))
            {
                ErrorMessage = "Please confirm your new password.";
                return Page();
            }

            if (NewPassword != ConfirmPassword)
            {
                ErrorMessage = "Passwords do not match.";
                return Page();
            }

            if (string.IsNullOrWhiteSpace(UsernameOrEmail))
            {
                ErrorMessage = "Username or Email information is missing.";
                return Page();
            }

            if (string.IsNullOrWhiteSpace(MobileNumber))
            {
                ErrorMessage = "Mobile Number information is missing.";
                return Page();
            }

            try
            {
                UpdatePassword();

                return Redirect("/Login?passwordreset=true");
            }
            catch (Exception ex)
            {
                ErrorMessage = "Error updating password: " + ex.Message;
                return Page();
            }
        }

        // CHECK USER
        private bool UserExists(string usernameOrEmail, string mobileNumber)
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    throw new Exception(
                        "Users.xlsx was not found at C:\\QMSData\\Users.xlsx"
                    );
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
                    string username = row.Cell(3).GetString().Trim();

                    string email = row.Cell(5).GetString().Trim();

                    string mobile = row.Cell(6).GetString().Trim();

                    bool usernameOrEmailMatches =
                        username.Equals(
                            usernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        )
                        ||
                        email.Equals(
                            usernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        );

                    bool mobileMatches =
                        mobile.Equals(mobileNumber.Trim());

                    if (usernameOrEmailMatches && mobileMatches)
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        // UPDATE PASSWORD
        private void UpdatePassword()
        {
            lock (ExcelLock)
            {
                if (!System.IO.File.Exists(ExcelFilePath))
                {
                    throw new Exception(
                        "Users.xlsx was not found at C:\\QMSData\\Users.xlsx"
                    );
                }

                using var workbook = new XLWorkbook(ExcelFilePath);

                var worksheet = workbook.Worksheet("Users");

                var rows = worksheet.RangeUsed()?.RowsUsed();

                if (rows == null)
                {
                    throw new Exception(
                        "Users.xlsx does not contain any users."
                    );
                }

                foreach (var row in rows.Skip(1))
                {
                    string username = row.Cell(3).GetString().Trim();

                    string email = row.Cell(5).GetString().Trim();

                    string mobile = row.Cell(6).GetString().Trim();

                    bool usernameOrEmailMatches =
                        username.Equals(
                            UsernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        )
                        ||
                        email.Equals(
                            UsernameOrEmail.Trim(),
                            StringComparison.OrdinalIgnoreCase
                        );

                    bool mobileMatches =
                        mobile.Equals(MobileNumber.Trim());

                    if (usernameOrEmailMatches && mobileMatches)
                    {
                        // Column 4 = Password
                        row.Cell(4).Value = NewPassword;

                        workbook.Save();

                        return;
                    }
                }

                throw new Exception(
                    "User was not found in Users.xlsx. Please check Username/Email and Mobile Number."
                );
            }
        }
    }
}