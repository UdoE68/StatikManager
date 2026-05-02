using StatikManager.Api.Contracts;
using StatikManager.Api.Contracts.Session;
using StatikManager.Api.Services;

var builder = WebApplication.CreateBuilder(args);

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

app.MapFallbackToFile("index.html");

app.Run();
