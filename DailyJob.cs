using System.Globalization;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

public class ShareSubscriptionNotifyJob
{
    private readonly ILogger _logger;
    private readonly HttpClient _http;
    private readonly HttpClient _finmindHttp;
    private readonly HttpClient _lineHttp;

    private static readonly Regex StockCode4Digits = new(@"^\d{4}$", RegexOptions.Compiled);

    // 中文 CSV Header -> 英文 DTO 欄位
    private static readonly Dictionary<string, string> HeaderMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["序號"] = nameof(PublicFormRow.Seq),
            ["抽籤日期"] = nameof(PublicFormRow.DrawDate),
            ["證券名稱"] = nameof(PublicFormRow.SecurityName),
            ["證券代號"] = nameof(PublicFormRow.StockCode),
            ["發行市場"] = nameof(PublicFormRow.Market),

            ["申購開始日"] = nameof(PublicFormRow.SubscribeStartDate),
            ["申購結束日"] = nameof(PublicFormRow.SubscribeEndDate),

            ["承銷股數"] = nameof(PublicFormRow.UnderwriteShares),
            ["實際承銷股數"] = nameof(PublicFormRow.ActualUnderwriteShares),

            ["承銷價(元)"] = nameof(PublicFormRow.UnderwritePrice),
            ["實際承銷價(元)"] = nameof(PublicFormRow.ActualUnderwritePrice),

            ["撥券日期(上市、上櫃日期)"] = nameof(PublicFormRow.AllocateDate),
            ["主辦券商"] = nameof(PublicFormRow.LeadUnderwriter),

            ["申購股數"] = nameof(PublicFormRow.SubscribeShares),
            ["總承銷金額(元)"] = nameof(PublicFormRow.TotalAmount),

            ["總合格件"] = nameof(PublicFormRow.TotalQualified),
            ["中籤率(%)"] = nameof(PublicFormRow.WinningRate),

            ["取消公開抽籤"] = nameof(PublicFormRow.CancelRemark)
        };

    public ShareSubscriptionNotifyJob(ILoggerFactory loggerFactory, IHttpClientFactory httpClientFactory)
    {
        _logger = loggerFactory.CreateLogger<ShareSubscriptionNotifyJob>();
        _http = httpClientFactory.CreateClient("excel");
        _finmindHttp = httpClientFactory.CreateClient("finmind");
        _lineHttp = httpClientFactory.CreateClient("line");
    }

    [Function("ShareSubscriptionNotify")]
    public async Task Run([TimerTrigger("0 0 18 * * *")] TimerInfo timer)
    {
        var url = Environment.GetEnvironmentVariable("ExcelUrl"); // 你環境變數名稱沿用也行
        if (string.IsNullOrWhiteSpace(url))
        {
            _logger.LogError("ExcelUrl is missing in application settings.");
            return;
        }

        // 台北今年 → {year} 取代
        var taipeiTz = GetTaipeiTimeZone();
        int year = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, taipeiTz).Year;
        url = url.Replace("{year}", year.ToString(CultureInfo.InvariantCulture));

        _logger.LogInformation("Downloading CSV from: {Url}", url);

        await using var stream = await DownloadToStreamAsync(url);

        var rows = ParsePublicFormCsv(stream);

        if (rows.Count == 0)
        {
            _logger.LogInformation("No subscription stocks found today.");
            return;
        }

        var msg = $"今日公開抽籤申購中股票";
        var incomingMsg = $"即將公開抽籤申購股票";

        _logger.LogInformation("Parsed rows (4-digit StockCode only): {Count}", rows.Count);

        foreach (var r in rows)
        {
            _logger.LogInformation("Stock={StockCode} Name={Name} DrawDate={DrawDate} Rate={Rate}",
                r.StockCode, r.SecurityName, r.DrawDate, r.WinningRate);

            var startDate = DateTime.Today.AddDays(-5);
            var endDate = DateTime.Today;

            var response = await _finmindHttp.GetAsync(
                $"/api/v4/data?dataset=TaiwanStockPrice" +
                $"&data_id={r.StockCode}" +
                $"&start_date={startDate:yyyy-MM-dd}" +
                $"&end_date={endDate:yyyy-MM-dd}"
            );
            var body = await response.Content.ReadAsStringAsync();

            var result = JsonSerializer.Deserialize<FinMindResponse<TaiwanStockPriceDto>>(
                body,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

            if (result?.Data == null || result.Data.Count == 0)
            {
                _logger.LogWarning("FinMind no data. Stock={StockCode}", r.StockCode);
                continue;
            }

            var latest = result.Data
                .OrderByDescending(x => DateTime.Parse(x.Date))
                .First();

            var roi = (latest.Close - r.ActualUnderwritePrice) / r.ActualUnderwritePrice * 100m;

            if (roi < 30)
                continue;

            if (r.IsInSubscribePeriod) {
                msg += $"\n\n{r.StockCode} {r.SecurityName}\n抽籤日 {r.DrawDate:MM/dd}   價差 +{roi:F0}%\n最新價 {latest.Close}    承銷價 {r.ActualUnderwritePrice}";
            } else {
                incomingMsg += $"\n\n{r.StockCode} {r.SecurityName}\n申購開始日 {r.SubscribeStartDate:MM/dd}\n申購結束日 {r.SubscribeEndDate:MM/dd}\n抽籤日 {r.DrawDate:MM/dd}    價差 +{roi:F0}%\n最新價 {latest.Close}    承銷價 {r.ActualUnderwritePrice}";
            }
        }

        await BroadcastTextAsync(msg);
        await BroadcastTextAsync(incomingMsg);
        
        _logger.LogInformation("Job done.");
    }

    // ---------------------------
    // CSV Parse (Stream -> DTO)
    // ---------------------------
    private List<PublicFormRow> ParsePublicFormCsv(Stream csvStream)
    {
        // 如果你的 CSV 是 Big5/CP950，改 Encoding.GetEncoding(950)
        using var reader = new StreamReader(csvStream, Encoding.GetEncoding(950), detectEncodingFromByteOrderMarks: true, leaveOpen: true);

        string? headerLine;
        do
        {
            headerLine = reader.ReadLine();
        }
        while (headerLine != null && !headerLine.Contains("證券代號"));

        if (headerLine == null) return new List<PublicFormRow>();

        var headers = ParseCsvLine(headerLine);
        var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Count; i++)
        {
            var h = headers[i]?.Trim();
            if (!string.IsNullOrWhiteSpace(h) && !headerIndex.ContainsKey(h))
                headerIndex[h] = i;
        }

        _logger.LogInformation("CSV Headers: {Headers}", string.Join(", ", headerIndex.Keys));

        // 必備：證券代號
        if (!headerIndex.TryGetValue("證券代號", out var stockCodeIdx))
        {
            _logger.LogError("CSV missing required header: 證券代號");
            return new List<PublicFormRow>();
        }

        var today = DateOnly.FromDateTime(DateTime.Now);
        var results = new List<PublicFormRow>();
        int rowNumber = 1; // CSV 行號（含 header）；資料行從 2 開始

        while (!reader.EndOfStream)
        {
            var line = reader.ReadLine();
            rowNumber++;

            if (string.IsNullOrWhiteSpace(line)) continue;

            var fields = ParseCsvLine(line);

            string? stockCode = GetField(fields, stockCodeIdx);
            DateOnly? drawDate = GetRocDate(fields, headerIndex, "抽籤日期");
            DateOnly? subscribeStartDate = GetRocDate(fields, headerIndex, "申購開始日");
            DateOnly? subscribeEndDate = GetRocDate(fields, headerIndex, "申購結束日");

            if (
                string.IsNullOrWhiteSpace(stockCode) || 
                !StockCode4Digits.IsMatch(stockCode) ||
                (subscribeEndDate.HasValue && subscribeEndDate.Value <= today)
            )
                continue;

            var item = new PublicFormRow
            {
                RowNumber = rowNumber,
                StockCode = stockCode,

                Seq = GetInt(fields, headerIndex, "序號"),
                DrawDate = drawDate,
                SecurityName = GetString(fields, headerIndex, "證券名稱"),
                Market = GetString(fields, headerIndex, "發行市場"),

                SubscribeStartDate = subscribeStartDate,
                SubscribeEndDate = subscribeEndDate,

                UnderwriteShares = GetLong(fields, headerIndex, "承銷股數"),
                ActualUnderwriteShares = GetLong(fields, headerIndex, "實際承銷股數"),

                UnderwritePrice = GetDecimal(fields, headerIndex, "承銷價(元)"),
                ActualUnderwritePrice = GetDecimal(fields, headerIndex, "實際承銷價(元)"),

                AllocateDate = GetRocDate(fields, headerIndex, "撥券日期(上市、上櫃日期)"),
                LeadUnderwriter = GetString(fields, headerIndex, "主辦券商"),

                SubscribeShares = GetLong(fields, headerIndex, "申購股數"),
                TotalAmount = GetDecimal(fields, headerIndex, "總承銷金額(元)"),

                TotalQualified = GetLong(fields, headerIndex, "總合格件"),
                WinningRate = GetDecimal(fields, headerIndex, "中籤率(%)"),

                CancelRemark = GetString(fields, headerIndex, "取消公開抽籤"),
                IsInSubscribePeriod = subscribeStartDate.HasValue &&
                    subscribeEndDate.HasValue &&
                    today >= subscribeStartDate.Value &&
                    today <= subscribeEndDate.Value
            };

            results.Add(item);
        }

        return results.OrderBy(x => x.DrawDate).ToList();
    }

    // ---------------------------
    // CSV helpers (RFC4180-ish)
    // ---------------------------

    private static List<string?> ParseCsvLine(string line)
    {
        // 支援：
        // - 用逗號分隔
        // - "..." 引號包住欄位
        // - 引號內的 "" 代表一個 "
        var result = new List<string?>();
        var sb = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];

            if (inQuotes)
            {
                if (ch == '"')
                {
                    // 連續兩個引號 -> 代表字元 "
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    sb.Append(ch);
                }
            }
            else
            {
                if (ch == ',')
                {
                    result.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else if (ch == '"')
                {
                    inQuotes = true;
                }
                else
                {
                    sb.Append(ch);
                }
            }
        }

        result.Add(sb.ToString().Trim());
        return result;
    }

    private static string? GetField(IReadOnlyList<string?> fields, int idx)
        => (idx >= 0 && idx < fields.Count) ? fields[idx]?.Trim() : null;

    private static string? GetString(IReadOnlyList<string?> fields, Dictionary<string, int> headerIndex, string header)
        => headerIndex.TryGetValue(header, out var idx) ? GetField(fields, idx) : null;

    private static int? GetInt(IReadOnlyList<string?> fields, Dictionary<string, int> headerIndex, string header)
    {
        var s = GetString(fields, headerIndex, header);
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "").Trim();
        return int.TryParse(s, out var v) ? v : null;
    }

    private static long? GetLong(IReadOnlyList<string?> fields, Dictionary<string, int> headerIndex, string header)
    {
        var s = GetString(fields, headerIndex, header);
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "").Trim();
        return long.TryParse(s, out var v) ? v : null;
    }

    private static decimal? GetDecimal(IReadOnlyList<string?> fields, Dictionary<string, int> headerIndex, string header)
    {
        var s = GetString(fields, headerIndex, header);
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "").Trim();
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    /// <summary>
    /// 解析民國日期字串，例如 "115/02/25" -> 2026-02-25
    /// 也容忍 "115-02-25"
    /// </summary>
    private static DateOnly? GetRocDate(IReadOnlyList<string?> fields, Dictionary<string, int> headerIndex, string header)
    {
        var s = GetString(fields, headerIndex, header);
        if (string.IsNullOrWhiteSpace(s)) return null;

        s = s.Replace("-", "/").Trim();
        var parts = s.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 3) return null;

        if (!int.TryParse(parts[0], out var rocYear)) return null;
        if (!int.TryParse(parts[1], out var month)) return null;
        if (!int.TryParse(parts[2], out var day)) return null;

        var year = rocYear + 1911;
        try { return new DateOnly(year, month, day); }
        catch { return null; }
    }

    // ---------------------------
    // HTTP Download
    // ---------------------------

    private async Task<Stream> DownloadToStreamAsync(string url)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/csv"));

        using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
        resp.EnsureSuccessStatusCode();

        await using var networkStream = await resp.Content.ReadAsStreamAsync();
        var ms = new MemoryStream();
        await networkStream.CopyToAsync(ms);
        ms.Position = 0;
        return ms;
    }

    private static TimeZoneInfo GetTaipeiTimeZone()
    {
        try { return TimeZoneInfo.FindSystemTimeZoneById("Asia/Taipei"); }
        catch { return TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time"); }
    }

    public async Task BroadcastTextAsync(string text)
    {
        var payload = new
        {
            messages = new object[]
            {
                new { type = "text", text }
            }
        };

        var resp = await _lineHttp.PostAsJsonAsync("/v2/bot/message/broadcast", payload);

        var body = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode)
            throw new Exception($"LINE broadcast failed: {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}");
    }

    // ---------------------------
    // DTO (English fields)
    // ---------------------------

    public sealed class PublicFormRow
    {
        public int RowNumber { get; init; }

        public int? Seq { get; init; }
        public DateOnly? DrawDate { get; init; }

        public string? SecurityName { get; init; }
        public string? StockCode { get; init; } // 4-digit only
        public string? Market { get; init; }

        public DateOnly? SubscribeStartDate { get; init; }
        public DateOnly? SubscribeEndDate { get; init; }

        public long? UnderwriteShares { get; init; }
        public long? ActualUnderwriteShares { get; init; }

        public decimal? UnderwritePrice { get; init; }
        public decimal? ActualUnderwritePrice { get; init; }

        public DateOnly? AllocateDate { get; init; }
        public string? LeadUnderwriter { get; init; }

        public long? SubscribeShares { get; init; }
        public decimal? TotalAmount { get; init; }

        public long? TotalQualified { get; init; }
        public decimal? WinningRate { get; init; }

        public string? CancelRemark { get; init; }

        public bool IsInSubscribePeriod { get; init; }
    }

    public class FinMindResponse<T>
    {
        public int Status { get; set; }
        public string Msg { get; set; }
        public List<T> Data { get; set; }
    }

    public class TaiwanStockPriceDto
    {
        public string Date { get; set; }
        public string Stock_Id { get; set; }
        public decimal Open { get; set; }
        public decimal Max { get; set; }
        public decimal Min { get; set; }
        public decimal Close { get; set; }
        public long Trading_Volume { get; set; }
    }
}