using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

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
        var adminPassword = _configuration["Auth:AdminPassword"];
        if (string.IsNullOrEmpty(adminPassword) || string.IsNullOrEmpty(Password) || !FixedTimeEquals(Password, adminPassword)) {
            ErrorMessage = "パスワードが正しくありません。";
            return Page();
        }

        var claims = new List<Claim> { new(ClaimTypes.Name, "管理者") };
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));

        return LocalRedirect(Url.IsLocalUrl(returnUrl) ? returnUrl! : "/");
    }

    // 平文比較でもタイミング攻撃を避けるため定数時間比較を使う
    private static bool FixedTimeEquals(string a, string b) =>
        CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(a), Encoding.UTF8.GetBytes(b));
}
