using ProductWebViewer.Data;

var builder = WebApplication.CreateBuilder(args);

// コンソールウィンドウなしで Windows サービスとして起動できるようにする
builder.Host.UseWindowsService();

builder.Services.AddRazorPages();
// リポジトリはクエリごとに接続を開閉するステートレス設計のため Singleton で問題ない
builder.Services.AddSingleton<ProductRecordRepository>();
builder.Services.AddSingleton<SubstrateRecordRepository>();

var app = builder.Build();

// リポジトリのコンストラクタでDB疎通確認を行うため、ここで強制的に生成し起動時に失敗させる
var productRepository = app.Services.GetRequiredService<ProductRecordRepository>();
app.Services.GetRequiredService<SubstrateRecordRepository>();

// DB側のビュー定義がこのアプリの想定と一致するか、起動時に一度だけ検証する
ViewDefinitionVerifier.Verify(productRepository.ConnectionString);

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
