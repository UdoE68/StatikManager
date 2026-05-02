using System.Text.Json;
using System.Text.Json.Serialization;
using StatikManager.Api.Contracts;
using StatikManager.Api.Contracts.Projects;
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
builder.Services.AddSingleton<ProjectListStore>();

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

app.MapPost("/api/session/pick-root", (IFileSystemService fs) =>
{
    try
    {
        var path = WindowsFolderPicker.PickFolderOnStaThread();
        if (path is null)
            return Results.NoContent();

        var fehler = fs.TrySetRoot(path);
        return fehler is null
            ? Results.Json(fs.GetSession())
            : Results.BadRequest(new ErrorResponse(fehler));
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

app.MapGet("/api/projects", (ProjectListStore store, ILoggerFactory loggerFactory) =>
{
    try
    {
        return Results.Json(new ProjectsResponse(store.GetAll()));
    }
    catch (Exception ex)
    {
        loggerFactory.CreateLogger("ProjectsApi").LogError(ex, "GET /api/projects");
        return Results.Json(new ProjectsResponse(Array.Empty<SavedProjectDto>()));
    }
});

app.MapPost("/api/projects", (AddProjectRequest? req, ProjectListStore store, ILogger<Program> logger) =>
{
    try
    {
        if (req is null)
            return Results.BadRequest(new ErrorResponse("Anfrage fehlt oder ist ungültig."));

        var fehler = store.TryAdd(req.Path, req.Name);
        return fehler is null
            ? Results.Json(new ProjectsResponse(store.GetAll()))
            : Results.BadRequest(new ErrorResponse(fehler));
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "POST /api/projects");
        return Results.BadRequest(new ErrorResponse($"Projektliste (Server): {ex.Message}"));
    }
});

app.MapDelete("/api/projects", (string? path, ProjectListStore store, ILogger<Program> logger) =>
{
    try
    {
        var fehler = store.TryRemove(path);
        return fehler is null
            ? Results.Json(new ProjectsResponse(store.GetAll()))
            : Results.BadRequest(new ErrorResponse(fehler));
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "DELETE /api/projects");
        return Results.BadRequest(new ErrorResponse($"Projektliste (Server): {ex.Message}"));
    }
});

app.MapFallbackToFile("index.html");

app.Run();
