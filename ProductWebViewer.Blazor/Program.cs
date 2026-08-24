using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using ProductWebViewer.Blazor.Components;
using ProductWebViewer.Blazor.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// リポジトリはクエリごとに接続を開閉するステートレス設計のため Singleton で問題ない（本家 ProductWebViewer と同じ方針）
builder.Services.AddSingleton<ProductRecordRepository>();
builder.Services.AddSingleton<SubstrateRecordRepository>();
builder.Services.AddSingleton<ProductWriteRepository>();
builder.Services.AddSingleton<SubstrateWriteRepository>();
builder.Services.AddHostedService<DbInitializer>();

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options => {
        options.LoginPath = "/login";
        options.AccessDeniedPath = "/login";
        options.ExpireTimeSpan = TimeSpan.FromHours(8);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAuthentication();
app.UseAuthorization();

app.UseAntiforgery();

// Blazorのインタラクティブコンポーネントは SignalR 経由で動作するため HttpContext.SignInAsync/SignOutAsync を直接呼べない。
// そのため認証操作は通常の（非Blazorな）エンドポイントとして用意し、フォームPOSTで叩く方式にしている。
app.MapPost("/login", async (HttpContext context, IConfiguration configuration, string? returnUrl) => {
    var form = await context.Request.ReadFormAsync();
    var password = form["Password"].ToString();
    var adminPassword = configuration["Auth:AdminPassword"];

    if (string.IsNullOrEmpty(adminPassword) || string.IsNullOrEmpty(password) || !FixedTimeEquals(password, adminPassword)) {
        return Results.Redirect($"/login?error=1&returnUrl={Uri.EscapeDataString(returnUrl ?? "/")}");
    }

    var claims = new List<Claim> { new(ClaimTypes.Name, "管理者") };
    var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    await context.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));

    return Results.LocalRedirect(!string.IsNullOrEmpty(returnUrl) && returnUrl.StartsWith('/') ? returnUrl : "/");
});

app.MapPost("/logout", async (HttpContext context) => {
    await context.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.LocalRedirect("/");
});

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();

// 平文比較でもタイミング攻撃を避けるため定数時間比較を使う
static bool FixedTimeEquals(string a, string b) =>
    CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(a), Encoding.UTF8.GetBytes(b));
