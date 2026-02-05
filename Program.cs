using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Net.Http.Headers;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

// ✅ 加這段：讓你可以用 IHttpClientFactory
builder.Services.AddHttpClient("excel", client =>
{
    client.Timeout = TimeSpan.FromSeconds(120);
});

// ✅ FinMind HttpClient（Named client）
builder.Services.AddHttpClient("finmind", client =>
{
    client.Timeout = TimeSpan.FromSeconds(120);

    var baseUrl = Environment.GetEnvironmentVariable("FinMind__BaseUrl");
    var token   = Environment.GetEnvironmentVariable("FinMind__Token");

    if (!string.IsNullOrWhiteSpace(baseUrl))
        client.BaseAddress = new Uri(baseUrl);

    if (!string.IsNullOrWhiteSpace(token))
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
});

// ✅ FinMind HttpClient（Named client）
builder.Services.AddHttpClient("line", client =>
{
    client.Timeout = TimeSpan.FromSeconds(120);

    var baseUrl = Environment.GetEnvironmentVariable("Line__BaseUrl");
    var token   = Environment.GetEnvironmentVariable("Line__Token");

    if (!string.IsNullOrWhiteSpace(baseUrl))
        client.BaseAddress = new Uri(baseUrl);

    if (!string.IsNullOrWhiteSpace(token))
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
});

builder.Build().Run();