using FlexKitTester.Components;
using FlexKitTester.Services;
using Microsoft.AspNetCore.SignalR;
using Radzen;
using Fx.ControlKit.Reports;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    // /pm-entry bench: simulated uplink latency applied to every inbound
    // circuit message (see BenchLatency).
    .AddHubOptions(o => o.AddFilter(new BenchLatencyHubFilter()));

// The two exports are parsed once per process, not once per circuit.
builder.Services.AddSingleton<BenchDataStore>();

// Created lazily on first use of the isolated SQLite provider bench. The
// original Edit Items bench never opens or initializes this database.
builder.Services.AddSingleton<SqliteProviderBenchDatabase>();

// Optional diagnostics used only when the operator presses a latency-probe
// button. SQL/TCP probes remain unavailable until explicitly configured through
// user-secrets or environment variables; no endpoint or credential is built in.
builder.Services.Configure<LatencyProbeOptions>(
    builder.Configuration.GetSection(LatencyProbeOptions.SectionName));
builder.Services.AddScoped<LatencyProbeService>();

builder.Services.AddSingleton(new ReportOptions { MaxRowsPerReport = CrystalBenchDataExecutor.MaxRows });
builder.Services.AddScoped<CrystalXmlReportLoader>();
builder.Services.AddSingleton<CrystalBenchSamples>();
builder.Services.AddScoped<CrystalBenchDataExecutor>();
builder.Services.AddScoped<IReportDataExecutor>(services => services.GetRequiredService<CrystalBenchDataExecutor>());
builder.Services.AddScoped<IReportSessionContext, CrystalBenchSessionContext>();
builder.Services.AddScoped<IReportExporter, CrystalBenchNoRuntimeExporter>();
builder.Services.AddScoped<IReportViewerSettings, InMemoryReportViewerSettings>();
builder.Services.AddScoped<IReportPickListProvider, EmptyReportPickListProvider>();

// Application-wide text zoom, one instance per circuit.
builder.Services.AddScoped<Fx.ControlKit.ZoomService>();

// Radzen Blazor component services (Dialog, Notification, Tooltip, ContextMenu).
builder.Services.AddRadzenComponents();

var app = builder.Build();

var benchData = app.Services.GetRequiredService<BenchDataStore>();
benchData.Load();
app.Logger.LogInformation("Bench data: {Report}", benchData.LoadReport);
Console.WriteLine($"Bench data: {benchData.LoadReport}");

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAntiforgery();

// Same-origin, zero-payload endpoint for browser <-> bench HTTP RTT. It never
// contacts an external service and deliberately cannot return configuration.
app.MapGet("/bench-latency/ping", (HttpContext context) =>
{
    context.Response.Headers.CacheControl = "no-store, no-cache, max-age=0";
    return Results.NoContent();
});

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
