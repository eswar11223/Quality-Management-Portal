using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Quality_management_portal.Pages.ComingSoon
{
    public class ComingSoonModel : PageModel
    {
        public string FirstName { get; set; } = "";

        [BindProperty(SupportsGet = true)]
        public string Module { get; set; } = "";

        public string ModuleName { get; set; } = "";

        public IActionResult OnGet()
        {
            string? username = HttpContext.Session.GetString("Username");

            if (string.IsNullOrEmpty(username))
            {
                return Redirect("/Login");
            }

            FirstName = HttpContext.Session.GetString("FirstName") ?? "";

            ModuleName = string.IsNullOrWhiteSpace(Module)
                ? "Module"
                : Module;

            return Page();
        }

        public IActionResult OnPostLogout()
        {
            HttpContext.Session.Clear();

            return Redirect("/Login");
        }
    }
}