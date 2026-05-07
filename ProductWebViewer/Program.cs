using ProductWebViewer.Data;

var builder = WebApplication.CreateBuilder(args);

// コンソールウィンドウなしで Windows サービスとして起動できるようにする
builder.Host.UseWindowsService();

builder.Services.AddRazorPages();
// リポジトリはクエリごとに接続を開閉するステートレス設計のため Singleton で問題ない
builder.Services.AddSingleton<ProductRecordRepository>();
builder.Services.AddSingleton<SubstrateRecordRepository>();

var app = builder.Build();

app.UseHttpsRedirection();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseRouting();

app.UseAuthorization();

app.UseStaticFiles();
app.MapRazorPages();

app.Run();
