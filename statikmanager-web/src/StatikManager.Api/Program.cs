using System.Text.Json;
using System.Text.Json.Serialization;
using StatikManager.Api.Contracts;
using StatikManager.Api.Contracts.Session;
using StatikManager.Api.Infrastructure;
using StatikManager.Api.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.ConfigureHttpJsonOptions(o =>
{
    o.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
    o.SerializerOptions.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
});

builder.Services.AddSingleton<IFileSystemService, FileSystemService>();

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins("http://localhost:5173", "https://localhost:5173")
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseCors();
}

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/api/health", () => Results.Json(new { ok = true }));

app.MapGet("/api/session/root", (IFileSystemService fs) =>
    Results.Json(fs.GetSession()));

app.MapPost("/api/session/root", (SetRootRequest? req, IFileSystemService fs) =>
{
    if (req is null)
        return Results.BadRequest(new ErrorResponse("Anfrage fehlt oder ist ungültig."));

    var fehler = fs.TrySetRoot(req.RootPath ?? "");
    return fehler is null
        ? Results.Json(fs.GetSession())
        : Results.BadRequest(new ErrorResponse(fehler));
});

app.MapPost("/api/session/pick-root", () =>
{
    try
    {
        var path = WindowsFolderPicker.PickFolder();
        return Results.Json(new PickRootResponse(path));
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new ErrorResponse($"Ordnerdialog fehlgeschlagen: {ex.Message}"));
    }
});

app.MapGet("/api/browse", (string? path, IFileSystemService fs) =>
{
    var fehler = fs.TryBrowse(path, out var resp);
    return fehler is null
        ? Results.Json(resp!)
        : Results.BadRequest(new ErrorResponse(fehler));
});

app.MapGet("/api/file/meta", (string? path, IFileSystemService fs) =>
{
    var fehler = fs.TryGetFileMeta(path, out var meta);
    return fehler is null
        ? Results.Json(meta!)
        : Results.BadRequest(new ErrorResponse(fehler));
});

app.MapGet("/api/preview/stream", (string? path, IFileSystemService fs) =>
{
    var fehler = fs.TryOpenPreviewRead(path, out var stream, out var contentType);
    return fehler is null
        ? Results.File(stream!, contentType: contentType, enableRangeProcessing: true)
        : Results.BadRequest(new ErrorResponse(fehler));
});

app.MapFallbackToFile("index.html");

app.Run();
