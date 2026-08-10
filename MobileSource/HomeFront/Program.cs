using HomeFront.Components;
using HomeFront.Data;
using Fx.ControlKit.Notifications;
using Fx.ControlKit.Reports;
using HomeFront.Services.Reports;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.Extensions.Caching.Memory;
using System.Security.Claims;
using System.Security.Cryptography;

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseStaticWebAssets();

// SA password is NOT in appsettings.json (which ships in deployment zips
// as plaintext). It comes from user-secrets in dev or from the
// `Database__Password` env var in prod. Composition lives in
// DbWrapperSqlServer.ComposeConnectionString — single source of truth
// shared with the DI-resolved (IConfiguration) constructor. See
// memory/sa_password_secret.md.
DbWrapperSqlServer.DefaultConnectionString =
    DbWrapperSqlServer.ComposeConnectionString(builder.Configuration);

// Add services to the container.

// Prominent +++++++ NEW CONNECTION / PAGE OPENED logging — see
// Infrastructure/Logging/ConnectionTrackingCircuitHandler.cs and
// PageOpenLogger.cs for the rationale + log format. IHttpContextAccessor
// is needed by CircuitClientInfo to snapshot the IP from the initial
// HTTP request that bootstrapped the circuit.
builder.Services.AddHttpContextAccessor();
builder.Services.AddScoped<HomeFront.Infrastructure.Logging.CircuitClientInfo>();
builder.Services.AddScoped<Microsoft.AspNetCore.Components.Server.Circuits.CircuitHandler,
                           HomeFront.Infrastructure.Logging.ConnectionTrackingCircuitHandler>();
builder.Services.AddScoped<HomeFront.Infrastructure.Logging.PageOpenLogger>();

builder.Services.AddScoped<NotificationService>();
// App-wide modal dialog service (FlexKit) — replaces native confirm/alert/prompt JS.
// DialogHostControl (in MainLayout) wires up the concrete instance; forms inject IDialogService.
builder.Services.AddScoped<Fx.ControlKit.Dialogs.DialogService>();
builder.Services.AddScoped<Fx.ControlKit.Dialogs.IDialogService>(sp => sp.GetRequiredService<Fx.ControlKit.Dialogs.DialogService>());
builder.Services.AddScoped<DbWrapperSqlServer>();
builder.Services.AddScoped<HomeFront.Services.ISessionStateService, HomeFront.Services.SessionStateService>();
// SINGLETON (not Scoped) — security toggles must be application-global,
// otherwise opt 5 (EnforceAuth) is bypassable just by opening a fresh
// browser tab that gets a new circuit with toggle defaults (all OFF).
// All six toggles share the same instance across every request /
// circuit. Production lockdown (Security:HideDevPanel=true) still
// force-sets every toggle ON in the constructor. See
// memory/homefrontpoc_security_hardening.md.
builder.Services.AddSingleton<HomeFront.Services.ISecurityOptionsService, HomeFront.Services.SecurityOptionsService>();
// ASP.NET Identity PasswordHasher — backs the strong-hashing toggle
// (review opt 4). Singleton because it holds no per-user state; the
// hasher itself is thread-safe and reused across all callers.
builder.Services.AddSingleton<Microsoft.AspNetCore.Identity.IPasswordHasher<string>,
    Microsoft.AspNetCore.Identity.PasswordHasher<string>>();
builder.Services.AddScoped<HomeFront.Services.IGridLayoutService, HomeFront.Services.GridLayoutService>();
builder.Services.AddScoped<HomeFront.Services.IUserPreferencesService, HomeFront.Services.UserPreferencesService>();
// IGridSettingsStore (the new in-GridControl cache layer) is intentionally
// NOT registered. HomeFront persists grid customizations through the existing
// GridLayoutService flow that writes to dbo.AppGridLayout — wired per-page in
// OnXxxColumnsChosenAsync handlers. Re-register either AppGridLayoutSettingsStore
// or JsonFileGridSettingsStore here when a future page wants the auto-load /
// auto-save path on top of GridControl.
builder.Services.AddScoped<HomeFront.Services.ITaskService, HomeFront.Services.TaskService>();
// Short-lived server-side stash for built export files, so the browser pulls
// them via a real HTTP attachment URL (/export/file/{token}) instead of a JS
// blob triggered over SignalR (flaky in Blazor Server → "exported but no file").
builder.Services.AddSingleton<HomeFront.Services.ExportFileStash>();
// Non-UI engine behind FImportExportAssemblies (Data Import/Export Wizard):
// workbook/CSV readers, export table/template builders, ImportedAssemblies /
// ImportedAssemblyTakeoffs staging + Purch_* validate/commit wrappers.
// Scoped (uses the scoped DbWrapperSqlServer); stateless per call.
builder.Services.AddScoped<HomeFront.Services.AssemblyImportExportService>();
builder.Services.AddScoped<HomeFront.Services.DbUpgradeService>();
// Per-circuit drop-box for pre-filling Crystal-report parameters on the next
// /report-viewer navigation (e.g. SPW Preview seeds the current Worksheet so
// FRptViewer's parameter dialog skips that prompt). One-shot consume.
builder.Services.AddScoped<HomeFront.Services.IReportSeedParameters, HomeFront.Services.ReportSeedParameters>();
// === Reports (FxControlKit.Reports) ==========================================
// Core FxControlKit pieces (singleton options + scoped XML loader). HomeFront
// configures the application-specific session-key list here — DivisionID is
// HomeFront's multi-tenancy key and the report auto-injection step looks it
// up via SessionReportContext below.
builder.Services.AddSingleton(new ReportOptions
{
    SessionAutoInjectParameters = { "DivisionID" }
});
builder.Services.AddScoped<CrystalXmlReportLoader>();

// HomeFront-specific Crystal Reports engine wrapper (uses SAP DLLs).
builder.Services.AddScoped<HomeFront.Services.CrystalReportService>();
// RptToXmlConverter — FlexKit's native C# converter for uploaded .rpt
// files. Supersedes the prior HomeFront.Services.Reports.RptToXmlConverter
// adapter (deleted) — FlexKit now ships the converter built-in, so the
// host-side wrapper became redundant. FCrystalReports.razor's @inject
// resolves to Fx.ControlKit.Reports.RptToXmlConverter.
builder.Services.AddScoped<Fx.ControlKit.Reports.RptToXmlConverter>();
builder.Services.AddScoped<HomeFront.Services.ReportParameterPickListService>();
builder.Services.AddScoped<HomeFront.Services.POStatusByJobReportRequestService>();
builder.Services.AddScoped<HomeFront.Services.POStatusByVendorReportRequestService>();

// FxControlKit interface implementations — adapters wired to the HomeFront
// services above. Registered as the interface type so ReportWriterControl /
// ReportParamDialogControl resolve them via @inject.
builder.Services.AddScoped<IReportDataExecutor, DbReportDataExecutor>();
builder.Services.AddScoped<IReportSessionContext, SessionReportContext>();
builder.Services.AddScoped<IReportExporter, CrystalReportExporterAdapter>();
builder.Services.AddScoped<IReportPickListProvider, ReportPickListProviderAdapter>();
// Persisted per-user viewer preferences (zoom level). Default in-memory
// fallback; swap for a tblAppOptions-backed adapter when persistence
// matters across app restarts.
builder.Services.AddScoped<IReportViewerSettings, InMemoryReportViewerSettings>();

// CustomReportRouter — XML-inferred routing for reports with parameter shapes
// that match the dual-list-box page flow. Reads the XML's existing attributes
// (no host-injected elements) and consults IReportPickListProvider to decide.
builder.Services.AddScoped<HomeFront.Services.CustomReportRouter>();

// ── Cookie-backed authentication (review opt 5) ──
// Cookie issued by /auth/signin after FLogin validates credentials, then
// consumed by FMain's auth gate. Survives across tabs / refresh /
// browser-close-and-reopen, unlike the legacy per-circuit
// Session.IsAuthenticated flag. See memory/opt5_cookie_auth.md.
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.Cookie.Name = "HomeFront.Auth";
        options.Cookie.HttpOnly = true;
        options.Cookie.SameSite = SameSiteMode.Lax;
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
        options.ExpireTimeSpan = TimeSpan.FromHours(12);
        options.SlidingExpiration = true;
        options.LoginPath = "/login";
        options.AccessDeniedPath = "/login";
    });
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddMemoryCache();

// Feedback / Testing (F8 screen-capture → annotate → ticket).
builder.Services.AddScoped<HomeFront.Services.IFeedbackService, HomeFront.Services.FeedbackService>();
builder.Services.AddScoped<HomeFront.Services.FeedbackStateService>();
builder.Services.AddHttpClient();   // registers IHttpClientFactory (required by JiraService)
builder.Services.AddScoped<HomeFront.Services.IJiraService, HomeFront.Services.JiraService>();

// Amazon Cognito Authentication Service Setup
builder.Services.Configure<HomeFront.Auth.Cognito.CognitoOptions>(builder.Configuration.GetSection(HomeFront.Auth.Cognito.CognitoOptions.SectionName));
builder.Services.AddSingleton<HomeFront.Auth.Cognito.ICognitoAuthService, HomeFront.Auth.Cognito.CognitoAuthService>();
builder.Services.AddSingleton<HomeFront.Auth.Cognito.ICognitoTokenValidator, HomeFront.Auth.Cognito.CognitoTokenValidator>();

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddHubOptions(options =>
    {
        // Feedback/Testing capture (F8) sends a full-screen base64 PNG from JS → C#
        // (FMain.OpenFeedbackWithScreenshot) and reads the annotated image back. Both cross the
        // Blazor circuit's SignalR hub as INBOUND messages, and the default MaximumReceiveMessageSize
        // is only 32 KB — a screen capture is hundreds of KB to several MB, so without this the
        // interop call exceeds the cap and silently tears the circuit down ("connection lost").
        options.MaximumReceiveMessageSize = 10 * 1024 * 1024;
    });

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);

// Serve files written to wwwroot at RUNTIME (feedback ticket screenshots + imported Jira attachments
// under /tickets/...). MapStaticAssets() below only serves the build-time manifest, so without this
// any ticket image saved/imported after the last build returns 404 (thumbnails + the image editor
// show nothing). UseStaticFiles serves whatever is currently on disk.
app.UseStaticFiles();

// Authentication/Authorization MUST come before UseAntiforgery and the
// component mappings so the auth cookie is parsed into HttpContext.User
// before Blazor's authorization cascade reads it.
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

// ── Auth endpoints (review opt 5 — cookie-backed authentication) ──
app.MapGet("/auth/signin", async (HttpContext ctx, string? token, IMemoryCache cache) =>
{
    if (string.IsNullOrWhiteSpace(token))
        return Results.Redirect("/login");
    var cacheKey = $"auth:signin:{token}";
    if (!cache.TryGetValue<ValueTuple<string, int, string?, string>>(cacheKey, out var payload))
        return Results.Redirect("/login");
    cache.Remove(cacheKey); // one-time use

    var (userId, divisionId, cognitoIdToken, dbName) = payload;
    var claims = new List<Claim>
    {
        new Claim(ClaimTypes.Name, userId),
        new Claim("DivisionId", divisionId.ToString(System.Globalization.CultureInfo.InvariantCulture)),
    };
    if (!string.IsNullOrEmpty(cognitoIdToken))
    {
        claims.Add(new Claim("CognitoIdToken", cognitoIdToken));
    }
    if (!string.IsNullOrEmpty(dbName))
    {
        claims.Add(new Claim("DatabaseName", dbName));
    }
    var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    await ctx.SignInAsync(
        CookieAuthenticationDefaults.AuthenticationScheme,
        new ClaimsPrincipal(identity));
    return Results.Redirect("/main");
});

app.MapGet("/auth/signout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/login");
});

app.MapGet("/api/export/download", async (HttpContext ctx, string? token, IMemoryCache cache, DbWrapperSqlServer db) =>
{
    if (string.IsNullOrEmpty(token) || !cache.TryGetValue<ExportRequestPayload>(token, out var payload) || payload == null)
    {
        ctx.Response.StatusCode = 404;
        await ctx.Response.WriteAsync("Export request expired or invalid.");
        return;
    }
    cache.Remove(token); // Single-use token

    try
    {
        var dt = db.SqlExec(payload.Sql);
        ctx.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        ctx.Response.Headers.ContentDisposition = $"attachment; filename=\"{payload.FileName}\"";
        
        using (var stream = new System.IO.MemoryStream())
        {
            Fx.ControlKit.Data.FilerControl.WriteExcel(stream, dt, "Exported Data");
            var bytes = stream.ToArray();
            await ctx.Response.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
    catch (Exception ex)
    {
        ctx.Response.StatusCode = 500;
        await ctx.Response.WriteAsync($"Export failed: {ex.Message}");
    }
});

// Feedback ticket download — the service streams the saved annotation/document for a ticket as an
// attachment, so the UI is plain <a href> links (no JS). Formats: png / pdf / doc / eml / jira.
app.MapGet("/tickets/{ticket:int}/download/{format}", async (int ticket, string format, HomeFront.Services.IFeedbackService svc) =>
{
    var d = await svc.ExportAsync(ticket, format);
    return d is null ? Results.NotFound() : Results.File(d.Bytes, d.ContentType, d.FileName);
});

// Export download — streams a stashed export file as a real HTTP attachment
// (Content-Disposition: attachment via Results.File), so the Import/Export
// wizard's download is reliable instead of a SignalR-triggered JS blob. Token
// is an opaque GUID from ExportFileStash.Put; entries expire after 10 minutes.
app.MapGet("/export/file/{token:guid}", (Guid token, HomeFront.Services.ExportFileStash stash) =>
    stash.TryGet(token, out var bytes, out var name, out var mime)
        ? Results.File(bytes, mime, name)
        : Results.NotFound());

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
