using CertMgmtWeb.Services;
using Microsoft.AspNetCore.Http.Json;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddSingleton<CertService>();
builder.Services.Configure<JsonOptions>(options => options.SerializerOptions.WriteIndented = true);

var app = builder.Build();
app.UseDefaultFiles();
app.UseStaticFiles();
app.UseCors(policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod());

app.MapGet("/api/ca/list", async (CertService service) => await service.GetEnrollmentCAsAsync());
app.MapGet("/api/ca/{config}/templates", async (CertService service, string config) => await service.GetCATemplatesAsync(config));
app.MapGet("/api/ca/{config}/status", async (CertService service, string config) => await service.GetCAStatusAsync(config));
app.MapPost("/api/request/dump", async (HttpRequest request, CertService service) =>
{
    if (!request.HasFormContentType)
    {
        return Results.BadRequest("Expected multipart/form-data");
    }

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("csr");
    if (file == null)
    {
        return Results.BadRequest("CSR file is required.");
    }

    var tempFile = await service.SaveUploadedFileAsync(file);
    var result = await service.DumpRequestAsync(tempFile);
    return result is null
        ? Results.BadRequest("Could not parse CSR file.")
        : Results.Ok(new { dump = result, requestFilePath = tempFile });
});

app.MapPost("/api/request/submit", async (SubmitRequestDto request, CertService service) =>
{
    if (string.IsNullOrWhiteSpace(request.TargetCaConfig) || string.IsNullOrWhiteSpace(request.TemplateName) || string.IsNullOrWhiteSpace(request.RequestFilePath))
    {
        return Results.BadRequest("Target CA, template, and CSR file path are required.");
    }

    var result = await service.SubmitRequestAsync(request.TargetCaConfig, request.TemplateName, request.RequestFilePath);
    return Results.Ok(new { message = result });
});

app.MapGet("/api/certs", async (string targetCaConfig, string? requesterName, int timeWindow, CertService service) =>
{
    var result = await service.GetIssuedCertListAsync(targetCaConfig, requesterName, timeWindow);
    return Results.Ok(result);
});

app.MapPost("/api/cert/retrieve", async (RetrieveCertDto request, CertService service) =>
{
    if (string.IsNullOrWhiteSpace(request.TargetCaConfig) || string.IsNullOrWhiteSpace(request.RequestId) || string.IsNullOrWhiteSpace(request.FolderPath))
    {
        return Results.BadRequest("Target CA, request ID, and folder path are required.");
    }

    var result = await service.RetrieveCertAsync(request.TargetCaConfig, request.RequestId, request.FolderPath);
    return Results.Ok(new { message = result });
});

app.MapGet("/api/pending", async (string targetCaConfig, CertService service) =>
{
    var result = await service.RetrievePendingCertsAsync(targetCaConfig);
    return Results.Ok(result);
});

app.MapPost("/api/pending/{requestId}/issue", async (string requestId, string targetCaConfig, CertService service) =>
{
    var result = await service.IssueCertRequestAsync(targetCaConfig, requestId);
    return Results.Ok(new { message = result });
});

app.MapPost("/api/pending/{requestId}/deny", async (string requestId, string targetCaConfig, CertService service) =>
{
    var result = await service.DenyCertRequestAsync(targetCaConfig, requestId);
    return Results.Ok(new { message = result });
});

app.MapGet("/api/raw/{requestId}", async (string requestId, string targetCaConfig, string mode, CertService service) =>
{
    var result = await service.DisplayRawDataAsync(targetCaConfig, requestId, mode);
    return result is null ? Results.BadRequest("Could not retrieve raw data.") : Results.Ok(new { rawText = result });
});

app.Run();

public record SubmitRequestDto(string TargetCaConfig, string TemplateName, string RequestFilePath);
public record RetrieveCertDto(string TargetCaConfig, string RequestId, string FolderPath);
