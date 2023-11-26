using AngleSharp;
using AngleSharp.Html.Parser;
using NPOI.HSSF;
using NPOI.SS.UserModel;
using System.IO.Compression;
using System.Net;
using System.Text;

const string CachePath = "../../.cache/";
const string PageLastModifiedPath = CachePath + "pageLastModified.txt";
const string ZipLastModifiedPath = CachePath + "zipLastModified.txt";
string basePath = $"../../csv/";

if (!Directory.Exists(CachePath))
    Directory.CreateDirectory(CachePath);

using var client = new HttpClient(new HttpClientHandler()
{
    AutomaticDecompression = DecompressionMethods.All
})
{
    Timeout = TimeSpan.FromSeconds(10),
};

bool isModified = false;

DateTimeOffset? pageLastModified = null;
if (File.Exists(PageLastModifiedPath))
{
    if (DateTimeOffset.TryParse(await File.ReadAllTextAsync(PageLastModifiedPath), out var c))
        pageLastModified = c;
    else
        pageLastModified = null;
}

DateTimeOffset? zipLastModified = null;
if (File.Exists(ZipLastModifiedPath))
{
    if (DateTimeOffset.TryParse(await File.ReadAllTextAsync(ZipLastModifiedPath), out var c))
        zipLastModified = c;
    else
        zipLastModified = null;
}

using var pageResponse = await client.SendAsync(new(HttpMethod.Get, "http://xml.kishou.go.jp/tec_material.html")
{
    Headers = { IfModifiedSince = pageLastModified },
});
if (pageResponse.StatusCode == HttpStatusCode.NotModified)
{
    Console.WriteLine("Page not modified.");
    Environment.Exit(isModified ? 1 : 0);
    return;
}
Console.WriteLine("Page modified.");

if (pageResponse.Content.Headers.LastModified is not DateTimeOffset responsePageLastModified)
    throw new InvalidOperationException("Last-Modified header not found.");

await File.WriteAllTextAsync(PageLastModifiedPath, responsePageLastModified.ToString("o"));
isModified = true;

using var context = BrowsingContext.New(Configuration.Default);
var parser = new HtmlParser();

var zipUrl = "http://xml.kishou.go.jp/" + (await parser.ParseDocumentAsync(await pageResponse.Content.ReadAsStreamAsync(), CancellationToken.None))
    .QuerySelectorAll("a").Where(a => a.TextContent.Contains("個別コード表")).Select(a => a.GetAttribute("href")).First();

using var zipResponse = await client.SendAsync(new(HttpMethod.Get, zipUrl)
{
    Headers = { IfModifiedSince = zipLastModified },
});
if (zipResponse.StatusCode == HttpStatusCode.NotModified)
{
    Console.WriteLine("Zip not modified.");
    Environment.Exit(isModified ? 1 : 0);
    return;
}
Console.WriteLine("Zip modified.");

if (zipResponse.Content.Headers.LastModified is not DateTimeOffset responseZipLastModified)
    throw new InvalidOperationException("Last-Modified header not found.");

await File.WriteAllTextAsync(ZipLastModifiedPath, responseZipLastModified.ToString("o"));
isModified = true;

using var zipStream = await zipResponse.Content.ReadAsStreamAsync();
using var zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Read, false, Encoding.GetEncoding("Shift_JIS"));

if (!Directory.Exists(basePath))
    Directory.CreateDirectory(basePath);

foreach (var entry in zipArchive.Entries)
{
    Console.WriteLine(entry.FullName);
    if (!entry.FullName.EndsWith(".xls"))
        continue;

    // 直接 Zip から読み込むとエラーになるので一旦メモリに展開
    using var memStream = new MemoryStream();
    using (var cStream = entry.Open())
        cStream.CopyTo(memStream);
    memStream.Seek(0, SeekOrigin.Begin);
    try
    {
        using var workbook = WorkbookFactory.Create(memStream, true);

        if (entry.FullName.EndsWith("AreaMarineAJ.xls"))
        {
            // 海上
            await ExtractFromSheetAsync(workbook, "AreaMarineA", "AreaMarineA", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "AreaMarineJ", "AreaMarineJ", (4, 0), 5);
        }
        else if (entry.FullName.EndsWith("AreaForecast.xls"))
        {
            // 全国・地方予報区等
            await ExtractFromSheetAsync(workbook, "Sheet1", "AreaForecast", (2, 1), 4);
        }
        else if (entry.FullName.EndsWith("RiverOffice.xls"))
        {
            // 河川事務所
            await ExtractFromSheetAsync(workbook, "RiverOffice", "RiverOffice", (3, 0), 2);
        }
        else if (entry.FullName.EndsWith("WmoObservingStations.xls"))
        {
            // 国際地点番号
            await ExtractFromSheetAsync(workbook, "WmoObservingStations.", "WmoObservingStations", (3, 0), 15);
        }
        else if (entry.FullName.EndsWith("AreaFloodForecast.xls"))
        {
            // 指定河川洪水予報（予報区域）
            await ExtractFromSheetAsync(workbook, "AreaFloodForecast", "AreaFloodForecast", (3, 0), 3);
        }
        else if (entry.FullName.EndsWith("AreaRiver.xls"))
        {
            // 指定河川洪水予報
            await ExtractFromSheetAsync(workbook, "AreaRiver", "AreaRiver", (3, 0), 3);
        }
        else if (entry.FullName.EndsWith("WaterLevelStation.xls"))
        {
            // 指定河川洪水予報 水位観測所
            await ExtractFromSheetAsync(workbook, "WaterLevelStation", "WaterLevelStation", (3, 0), 3);
        }
        else if (entry.FullName.EndsWith("AreaInformationCity-AreaForecastLocalM.xls"))
        {
            // AreaInformationCity
            await ExtractFromSheetAsync(workbook, "AreaInformationCity", "AreaInformationCity", (3, 0), 19);
            // AreaForecastLocalM
            await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（コード表）", "AreaForecastLocalM", (4, 0), 13);
            await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（関係表　警報・注意報", "AreaForecastLocalM_WarningTable", (3, 0), 6);
            await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（関係表　竜巻注意情報", "AreaForecastLocalM_TornadoTable", (3, 0), 6);
        }
        else if (entry.FullName.EndsWith("PointAmedas.xls"))
        {
            await ExtractAmedasFromSheetAsync(workbook, "ame_master", "AmedasRainPoint", (2, 0), 16);
            await ExtractFromSheetAsync(workbook, "snow_master", "AmedasSnowPoint", (2, 0), 12);
        }
        else if (entry.FullName.EndsWith("地震火山関連コード表.xls"))
        {
            await ExtractFromSheetAsync(workbook, "11", "EarthquakeWarning", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "12", "EarthquakeForecast", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "14", "TsunamiWarning", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "21", "AreaForecastEEW", (3, 0), 4);
            await ExtractFromSheetAsync(workbook, "22", "AreaForecastLocalEEW", (3, 0), 4);
            await ExtractFromSheetAsync(workbook, "23", "AreaInformationPrefectureEarthquake", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "24", "AreaForecastLocalE_AreaInformationCity_PointSeismicIntensity", (3, 0), 9);
            await ExtractFromSheetAsync(workbook, "25", "AreaForecastLocalE_AreaInformationCity_PointRealtimeIntensity", (3, 0), 6);
            await ExtractFromSheetAsync(workbook, "26", "AreaForecastLocalE_PointSeismicLgIntensity", (3, 0), 6);
            await ExtractFromSheetAsync(workbook, "31", "AreaTsunami", (3, 0), 4);
            await ExtractFromSheetAsync(workbook, "34", "CoastTsunami", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "35 ", "PointTsunami", (3, 0), 6);
            await ExtractFromSheetAsync(workbook, "41", "AreaEpicenter", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "42", "AreaEpicenterAbbreviation", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "43", "AreaEpicenterDetail", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "44", "AreaEpicenterSuppliment", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "51", "TokaiInformation", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "52", "EarthquakeInformation", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "62", "AdditionalCommentEarthquake", (3, 0), 2);
            await ExtractFromSheetAsync(workbook, "81", "VolcanicWarning", (3, 0), 3);
            await ExtractFromSheetAsync(workbook, "82 ", "PointVolcano", (3, 0), 4);
        }
    }
    catch (OldExcelFormatException)
    {
        // PhenologicalType.xls が古すぎるフォーマットのせいで怒られる…。
        Console.WriteLine("旧Excelフォーマットのためスキップされました: " + entry.FullName);
    }
}
Environment.Exit(1);

// https://www.data.jma.go.jp/developer/jma_multilingual.xlsx

async Task ExtractFromSheetAsync(IWorkbook workbook, string sheetName, string name, (int Row, int Col) startPosition, int columnCount)
{
    var sheet = workbook.GetSheet(sheetName) ?? throw new InvalidOperationException($"Sheet {sheetName} not found.");
    var csvPath = basePath + name + ".csv";
    using var csvWriter = new StreamWriter(csvPath, false) { NewLine = "\n" };

    for (int i = startPosition.Row; i <= sheet.LastRowNum; i++)
    {
        var row = sheet.GetRow(i);
        if (row is null || string.IsNullOrWhiteSpace(row.GetCell(startPosition.Col).StringCellValue()))
            continue;

        var values = new List<string>();
        for (int j = startPosition.Col; j < startPosition.Col + columnCount; j++)
        {
            var cell = row.GetCell(j);
            var value = (cell?.StringCellValue() ?? "").Trim();
            if (value.Contains('\n'))
                value = value.Replace("\n", "\\n");
            if (value.Contains(','))
            {
                value = value.Replace("\"", "\\\"");
                value = "\"" + value + "\"";
            }
            values.Add(value);
        }
        await csvWriter.WriteLineAsync(string.Join(',', values));
    }
}
async Task ExtractAmedasFromSheetAsync(IWorkbook workbook, string sheetName, string name, (int Row, int Col) startPosition, int columnCount)
{
    var sheet = workbook.GetSheet(sheetName) ?? throw new InvalidOperationException($"Sheet {sheetName} not found.");
    var csvPath = basePath + name + ".csv";
    using var csvWriter = new StreamWriter(csvPath, false, Encoding.UTF8) { NewLine = "\n" };
    string? managedMOName = null;

    for (int i = startPosition.Row; i <= sheet.LastRowNum; i++)
    {
        var row = sheet.GetRow(i);
        if (row is null || string.IsNullOrWhiteSpace(row.GetCell(startPosition.Col).StringCellValue()))
            continue;

        if (row.GetCell(startPosition.Col).StringCellValue()?.EndsWith("気象台管理") ?? false)
        {
            managedMOName = row.GetCell(startPosition.Col).StringCellValue();
            continue;
        }

        if (managedMOName is null)
            throw new InvalidOperationException("管理気象台が見つかりませんでした。");

        var values = new List<string>() { managedMOName };
        for (int j = startPosition.Col; j < startPosition.Col + columnCount; j++)
        {
            var cell = row.GetCell(j);
            var value = cell?.StringCellValue() ?? "";
            if (value.Contains('\n'))
                value = value.Replace("\n", "\\n");
            if (value.Contains(','))
            {
                value = value.Replace("\"", "\\\"");
                value = "\"" + value + "\"";
            }
            values.Add(value);
        }
        await csvWriter.WriteLineAsync(string.Join(',', values));
    }
}

public static class CellExtensions
{
    public static string? StringCellValue(this ICell cell)
        => cell?.CellType switch
        {
            CellType.String => cell.StringCellValue,
            CellType.Numeric => cell.NumericCellValue.ToString(),
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Formula => cell.CellFormula,
            _ => null,
        };
}
