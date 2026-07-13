using ProductWebViewer.Data;

var builder = WebApplication.CreateBuilder(args);

// コンソールウィンドウなしで Windows サービスとして起動できるようにする
builder.Host.UseWindowsService();

builder.Services.AddRazorPages();
// リポジトリはクエリごとに接続を開閉するステートレス設計のため Singleton で問題ない
builder.Services.AddSingleton<ProductRecordRepository>();
builder.Services.AddSingleton<SubstrateRecordRepository>();
// リポジトリのDB疎通確認・ビュー定義検証をリクエスト処理開始前に実行する
builder.Services.AddHostedService<ViewDefinitionStartupCheck>();

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
