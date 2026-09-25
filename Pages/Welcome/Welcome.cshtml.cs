using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.Welcome
{
    public class WelcomeModel : PageModel
    {
        public string FirstName { get; set; } = "";

        public IActionResult OnGet()
        {
            string? username = HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";

            return Page();
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }
    }
}