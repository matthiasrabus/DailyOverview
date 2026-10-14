using TeamsMessageFetcher;
using TeamsMessageFetcher.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

var config = AppConfig.Load();
builder.Services.AddSingleton(config);

// Level is evaluated on every call, so changes in Settings apply immediately
builder.Logging.SetMinimumLevel(LogLevel.Trace);
builder.Logging.AddFilter((_, level) => level >= config.MinLogLevel);
builder.Logging.AddFilter("Microsoft", LogLevel.Warning);
builder.Logging.AddFilter("System", LogLevel.Warning);
builder.Logging.AddProvider(new FileLoggerProvider(Path.Combine(AppContext.BaseDirectory, "logs")));

builder.Services.AddSingleton<DailyActivityService>();

var app = builder.Build();
FetchLog.Initialize(app.Services.GetRequiredService<ILoggerFactory>());

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
