using AngleSharp;
using AngleSharp.Html.Parser;
using System.IO.Compression;
using System.Net;
using System.Text;

const string BaseUrl = "https://xml.kishou.go.jp/";
const string CachePath = "../../.cache/";
const string LastModifiedPath = CachePath + "lastModified.json";
var basePath = $"../../csv/";

if (!Directory.Exists(CachePath))
	Directory.CreateDirectory(CachePath);

using var client = new HttpClient(new HttpClientHandler()
{
	AutomaticDecompression = DecompressionMethods.All
})
{
	Timeout = TimeSpan.FromMinutes(1),
};
client.DefaultRequestHeaders.UserAgent.ParseAdd("jma-code-dictionary (+https://github.com/ingen084/jma-code-dictionary)");

var isModified = false;
var lastModifiedCache = await LastModifiedCache.LoadAsync(LastModifiedPath);

// コード表以外の公開データ
await AdditionalSources.ExportAsync(GetIfModifiedAsync, basePath);

using var pageResponse = await GetIfModifiedAsync("techMaterialPage", BaseUrl + "tec_material.html");
if (pageResponse is null)
{
	await ExitAsync();
	return;
}

using var context = BrowsingContext.New(Configuration.Default);
var parser = new HtmlParser();
var document = await parser.ParseDocumentAsync(await pageResponse.Content.ReadAsStreamAsync(), CancellationToken.None);

var xlsUrl = BaseUrl + document.QuerySelectorAll("a").Where(a => a.TextContent.Contains("コード管理表一式")).Select(a => a.GetAttribute("href")).First();

using var xlsResponse = await GetIfModifiedAsync("codeManagementTable", xlsUrl);
if (xlsResponse is not null)
{
	using var stream = await xlsResponse.Content.ReadAsStreamAsync();
	using var workbook = XlsxWorkbook.Open(stream);
	await ProcessCodesAsync(workbook);
}

var zipUrl = BaseUrl + document.QuerySelectorAll("a").Where(a => a.TextContent.Contains("個別コード表")).Select(a => a.GetAttribute("href")).First();

using var zipResponse = await GetIfModifiedAsync("individualCodeTables", zipUrl);
if (zipResponse is null)
{
	await ExitAsync();
	return;
}

using var zipStream = await zipResponse.Content.ReadAsStreamAsync();
// エントリ名が UTF-8 になったため
using var zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Read, false, Encoding.UTF8);

if (!Directory.Exists(basePath))
	Directory.CreateDirectory(basePath);

foreach (var entry in zipArchive.Entries)
{
	Console.WriteLine(entry.FullName);
	if (!entry.FullName.EndsWith(".xlsx"))
		continue;

	// 直接 Zip から読み込むとエラーになるので一旦メモリに展開
	using var memStream = new MemoryStream();
	using (var cStream = entry.Open())
		cStream.CopyTo(memStream);
	memStream.Seek(0, SeekOrigin.Begin);
	using var workbook = XlsxWorkbook.Open(memStream);

	if (entry.FullName.EndsWith("PhenologicalType.xlsx"))
	{
		// 生物季節観測
		await ExtractFromSheetAsync(workbook, "XML", "PhenologicalType", (2, 0), 3, "01");
	}
	else if (entry.FullName.EndsWith("AreaMarineAJ.xlsx"))
	{
		// 海上
		await ExtractFromSheetAsync(workbook, "AreaMarineA", "AreaMarineA", (3, 0), 3, "9010");
		await ExtractFromSheetAsync(workbook, "AreaMarineJ", "AreaMarineJ", (3, 0), 5, "1000");
	}
	else if (entry.FullName.EndsWith("AreaForecast.xlsx"))
	{
		// 全国・地方予報区等
		await ExtractFromSheetAsync(workbook, "Sheet1", "AreaForecast", (2, 1), 4, "全国・地方予報区等");
	}
	else if (entry.FullName.EndsWith("RiverOffice.xlsx"))
	{
		// 河川事務所
		await ExtractFromSheetAsync(workbook, "RiverOffice", "RiverOffice", (3, 0), 2, "23297");
	}
	else if (entry.FullName.EndsWith("WmoObservingStations.xlsx"))
	{
		// 国際地点番号
		await ExtractFromSheetAsync(workbook, "WmoObservingStations.", "WmoObservingStations", (3, 0), 15, "47401");
	}
	else if (entry.FullName.EndsWith("AreaFloodForecast.xlsx"))
	{
		// 指定河川洪水予報（予報区域）
		await ExtractFromSheetAsync(workbook, "AreaFloodForecast", "AreaFloodForecast", (3, 0), 3, "010175000100");
	}
	else if (entry.FullName.EndsWith("AreaRiver.xlsx"))
	{
		// 指定河川洪水予報
		await ExtractFromSheetAsync(workbook, "AreaRiver", "AreaRiver", (3, 0), 8, "0101750001");
	}
	else if (entry.FullName.EndsWith("WaterLevelStation.xlsx"))
	{
		// 指定河川洪水予報 水位観測所
		await ExtractFromSheetAsync(workbook, "WaterLevelStation", "WaterLevelStation", (3, 0), 4, "81010100010000004");
	}
	else if (entry.FullName.EndsWith("AreaInformationCity-AreaForecastLocalM.xlsx"))
	{
		// AreaInformationCity
		await ExtractFromSheetAsync(workbook, "AreaInformationCity", "AreaInformationCity", (3, 0), 19, "0110000");
		// AreaForecastLocalM
		await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（コード表）", "AreaForecastLocalM", (4, 0), 13, "011000");
		await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（関係表　警報・注意報", "AreaForecastLocalM_WarningTable", (3, 0), 6, "011011");
		await ExtractFromSheetAsync(workbook, "AreaForecastLocalM（関係表　竜巻注意情報", "AreaForecastLocalM_TornadoTable", (3, 0), 6, "011011");
	}
	else if (entry.FullName.EndsWith("PointAmedas.xlsx"))
	{
		await ExtractAmedasFromSheetAsync(workbook, "ame_master", "AmedasRainPoint", (2, 0), 16);
		await ExtractFromSheetAsync(workbook, "snow_master", "AmedasSnowPoint", (2, 0), 12, "宗谷");
	}
	else if (entry.FullName.EndsWith("地震火山関連コード表.xlsx"))
	{
		await ExtractFromSheetAsync(workbook, "11", "EarthquakeWarning", (3, 0), 2, "00");
		await ExtractFromSheetAsync(workbook, "12", "EarthquakeForecast", (3, 0), 3, "00");
		await ExtractFromSheetAsync(workbook, "14", "TsunamiWarning", (3, 0), 3, "00");
		await ExtractFromSheetAsync(workbook, "21", "AreaForecastEEW", (3, 0), 4, "9910");
		await ExtractFromSheetAsync(workbook, "22", "AreaForecastLocalEEW", (3, 0), 4, "9011");
		await ExtractFromSheetAsync(workbook, "23", "AreaInformationPrefectureEarthquake", (3, 0), 2, "01");
		await ExtractFromSheetAsync(workbook, "24", "AreaForecastLocalE_AreaInformationCity_PointSeismicIntensity", (3, 0), 9, "100");
		await ExtractFromSheetAsync(workbook, "25", "AreaForecastLocalE_AreaInformationCity_PointRealtimeIntensity", (3, 0), 6, "101");
		await ExtractFromSheetAsync(workbook, "26", "AreaForecastLocalE_PointSeismicLgIntensity", (3, 0), 6, "101");
		await ExtractFromSheetAsync(workbook, "31", "AreaTsunami", (3, 0), 4, "100");
		await ExtractFromSheetAsync(workbook, "34", "CoastTsunami", (3, 0), 3, "100");
		await ExtractFromSheetAsync(workbook, "35", "PointTsunami", (3, 0), 6, "00305");
		await ExtractFromSheetAsync(workbook, "41", "AreaEpicenter", (3, 0), 2, "011");
		await ExtractFromSheetAsync(workbook, "42", "AreaEpicenterAbbreviation", (3, 0), 3, "9011");
		await ExtractFromSheetAsync(workbook, "43", "AreaEpicenterDetail", (3, 0), 2, "1001");
		await ExtractFromSheetAsync(workbook, "44", "AreaEpicenterSuppliment", (3, 0), 2, "100");
		await ExtractFromSheetAsync(workbook, "51", "TokaiInformation", (3, 0), 2, "000");
		await ExtractFromSheetAsync(workbook, "52", "EarthquakeInformation", (3, 0), 3, "111");
		await ExtractFromSheetAsync(workbook, "62", "AdditionalCommentEarthquake", (3, 0), 2, "0101");
		await ExtractFromSheetAsync(workbook, "81", "VolcanicWarning", (3, 0), 3, "01");
		await ExtractFromSheetAsync(workbook, "82 ", "PointVolcano", (3, 0), 4, "101");
	}
}
await ExitAsync();

// 取得元ごとの Last-Modified を保存して終了する（終了コード 0: 更新なし, 1: 更新あり）
async Task ExitAsync()
{
	await lastModifiedCache.SaveAsync();
	Environment.Exit(isModified ? 1 : 0);
}

// 前回の取得から更新されていれば応答を、更新されていなければ null を返す
async Task<HttpResponseMessage?> GetIfModifiedAsync(string key, string url)
{
	var response = await SendAsync(url, lastModifiedCache.Get(key, url));
	if (response.StatusCode == HttpStatusCode.NotModified)
	{
		Console.WriteLine($"Not modified: {key}");
		response.Dispose();
		return null;
	}
	response.EnsureSuccessStatusCode();
	Console.WriteLine($"Modified: {key}");

	if (response.Content.Headers.LastModified is not DateTimeOffset lastModified)
		throw new InvalidOperationException($"Last-Modified header not found: {url}");

	// 途中で失敗したときに更新済みとして残らないよう、保存は ExitAsync でまとめて行う
	lastModifiedCache.Set(key, url, lastModified);
	isModified = true;
	return response;
}

async Task<HttpResponseMessage> SendAsync(string url, DateTimeOffset? ifModifiedSince)
{
	const int MaxAttempts = 3;
	// 一時的な通信障害やタイムアウトに備えて数回リトライする
	for (var attempt = 1; ; attempt++)
	{
		try
		{
			return await client.SendAsync(new(HttpMethod.Get, url)
			{
				Headers = { IfModifiedSince = ifModifiedSince },
			});
		}
		catch (Exception e) when (attempt < MaxAttempts && e is HttpRequestException or TaskCanceledException)
		{
			Console.WriteLine($"Request failed ({attempt}/{MaxAttempts}): {url} {e.Message}");
			await Task.Delay(TimeSpan.FromSeconds(10 * attempt));
		}
	}
}

async Task ExtractFromSheetAsync(XlsxWorkbook workbook, string sheetName, string name, (int Row, int Col) startPosition, int columnCount, string firstClumnsCheckForRegressionCheck)
{
	var sheet = workbook.GetSheet(sheetName) ?? throw new InvalidOperationException($"Sheet {sheetName} not found.");
	var csvPath = basePath + name + ".csv";
	using var csvWriter = new StreamWriter(csvPath, false) { NewLine = "\n" };

	for (var i = startPosition.Row; i <= sheet.LastRowNum; i++)
	{
		var row = sheet.GetRow(i);
		if (row is null || string.IsNullOrWhiteSpace(row.GetCell(startPosition.Col).StringCellValue()))
			continue;

		var values = new List<string>();
		for (var j = startPosition.Col; j < startPosition.Col + columnCount; j++)
		{
			var cell = row.GetCell(j);
			var value = (cell?.StringCellValue() ?? "").Trim();
			if (value.Contains('\n'))
				value = value.Replace("\n", "\\n");
			if (i == startPosition.Row && j == startPosition.Col && !string.IsNullOrWhiteSpace(firstClumnsCheckForRegressionCheck))
			{
				if (value != firstClumnsCheckForRegressionCheck)
					throw new InvalidOperationException($"Regression check failed. Expected: {firstClumnsCheckForRegressionCheck}, Actual: {value}");
            }
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
async Task ExtractAmedasFromSheetAsync(XlsxWorkbook workbook, string sheetName, string name, (int Row, int Col) startPosition, int columnCount)
{
	var sheet = workbook.GetSheet(sheetName) ?? throw new InvalidOperationException($"Sheet {sheetName} not found.");
	var csvPath = basePath + name + ".csv";
	using var csvWriter = new StreamWriter(csvPath, false, Encoding.UTF8) { NewLine = "\n" };
	string? managedMOName = null;

	for (var i = startPosition.Row; i <= sheet.LastRowNum; i++)
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
		for (var j = startPosition.Col; j < startPosition.Col + columnCount; j++)
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

async Task ProcessCodesAsync(XlsxWorkbook workbook)
{
	var sheet = workbook.GetSheet("警報等情報要素コード管理表") ?? throw new InvalidOperationException("Sheet 警報等情報要素コード管理表 not found.");

	for (var i = 0; i < sheet.LastRowNum; i++)
	{
		var row = sheet.GetRow(i);

		// D 列から とりうる値 を探す
		if (row?.GetCell(3).StringCellValue() != "とりうる値")
			continue;

		// とりうる値 があればそこから C 列をたどっていって、名前を取得する
		string? name = null;
		for (var j = i; j >= 0; j--)
		{
			var nameRow = sheet.GetRow(j);
			if (!string.IsNullOrWhiteSpace(nameRow?.GetCell(2).StringCellValue()))
			{
				name = nameRow.GetCell(2).StringCellValue();
				break;
			}
		}
		if (name is null)
			throw new InvalidOperationException("名前が見つかりませんでした。");

		// とりうる値のセルから下に向かって値を取得していく
		var csvPath = basePath + name + ".csv";
		using var csvWriter = new StreamWriter(csvPath, false) { NewLine = "\n" };
		for (var j = i; j < sheet.LastRowNum; j++)
		{
			var valueRow = sheet.GetRow(j);
			var cell = valueRow?.GetCell(4);
			var key = cell.StringCellValue();
			if (string.IsNullOrWhiteSpace(key) || !int.TryParse(key, out _) || cell is { IsDateFormatted: true })
				break;
			var value = valueRow?.GetCell(5).StringCellValue();
			string? annotation = null;
			if (value?.Contains('（') ?? false)
			{
				annotation = value[value.IndexOf('（')..];
				value = value[..value.IndexOf('（')];
			}
			await csvWriter.WriteLineAsync(key + "," + value + "," + annotation);
		}
	}
}

public static class CellExtensions
{
	public static string? StringCellValue(this XlsxCell? cell)
		=> cell?.Value;
}
