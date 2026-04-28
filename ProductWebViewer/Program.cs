using ProductWebViewer.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Host.UseWindowsService();

builder.Services.AddRazorPages();
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
