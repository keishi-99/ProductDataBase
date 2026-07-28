using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Auth;

namespace ProductWebViewer.Pages;

public class LoginModel : PageModel {
    private readonly IConfiguration _configuration;

    public LoginModel(IConfiguration configuration) {
        _configuration = configuration;
    }

    [BindProperty] public string? Password { get; set; }
    public string? ErrorMessage { get; private set; }

    public void OnGet() { }

    public async Task<IActionResult> OnPostAsync(string? returnUrl) {
        var storedHash = _configuration["Auth:AdminPasswordHash"];
        if (string.IsNullOrEmpty(storedHash) || string.IsNullOrEmpty(Password) || !PasswordHasher.Verify(Password, storedHash)) {
            ErrorMessage = "パスワードが正しくありません。";
            return Page();
        }

        var claims = new List<Claim> { new(ClaimTypes.Name, "管理者") };
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));

        return LocalRedirect(Url.IsLocalUrl(returnUrl) ? returnUrl! : "/");
    }
}
