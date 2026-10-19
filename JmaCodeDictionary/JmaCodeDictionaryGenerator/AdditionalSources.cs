using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using System.Globalization;
using System.IO.Compression;
using System.Text.Json;

// 前回の取得から更新されていれば応答を、更新されていなければ null を返す
public delegate Task<HttpResponseMessage?> GetIfModified(string key, string url);

// 防災情報XMLのコード表以外の公開データから CSV を生成する
public static class AdditionalSources
{
	private const string BosaiUrl = "https://www.jma.go.jp/bosai/";
	private const string BulletinUrl = "https://www.data.jma.go.jp/eqev/data/bulletin/";
	private const string DeveloperUrl = "https://www.data.jma.go.jp/developer/";
	private const string NwptacUrl = "https://www.data.jma.go.jp/eqev/data/nwptac/";

	// 北西太平洋津波情報の予報地点の表の見出し
	private static readonly string[] ForecastPointColumns = ["Coastal Block", "Forecast Point", "Latitude", "Longitude", "Number in Fig."];

	// 多言語辞書の見出し。CSV にはヘッダー行を出力しないので、列の並びが変わっていないか確認する
	private static readonly string[] MultilingualColumns =
	[
		"No.", "種別", "情報種別等", "日本語",
		"英語", "中国語(簡体字)", "中国語(繁体字)", "韓国語", "ポルトガル語", "スペイン語", "インドネシア語",
		"ベトナム語", "タガログ語", "タイ語", "ネパール語", "クメール語", "ビルマ語", "モンゴル語",
		"備考", "XMLコード表", "コード", "更新履歴",
	];

	public static async Task ExportAsync(GetIfModified getIfModified, string basePath)
	{
		// 防災情報 (bosai) の JSON
		var webJsonPath = basePath + "web-json/";
		Directory.CreateDirectory(webJsonPath);
		await ExportVolcanoAsync(getIfModified, webJsonPath + "Volcano.csv");
		await ExportTideLevelStationAsync(getIfModified, webJsonPath + "TideLevelStation.csv");
		await ExportWindProfilerStationAsync(getIfModified, webJsonPath + "WindProfilerStation.csv");

		// 地震月報（カタログ編）
		var bulletinPath = basePath + "bulletin/";
		Directory.CreateDirectory(bulletinPath);
		await ExportSeismicStationAsync(getIfModified, bulletinPath + "SeismicStation.csv");

		// 多言語辞書
		var multilingualPath = basePath + "multilingual/";
		Directory.CreateDirectory(multilingualPath);
		await ExportMultilingualDictionaryAsync(getIfModified, multilingualPath + "MultilingualDictionary.csv");

		// 北西太平洋津波情報センター
		var nwptacPath = basePath + "nwptac/";
		Directory.CreateDirectory(nwptacPath);
		await ExportForecastPointAsync(getIfModified, nwptacPath + "ForecastPoint.csv");
	}

	// 火山
	private static async Task ExportVolcanoAsync(GetIfModified getIfModified, string csvPath)
	{
		using var json = await GetJsonAsync(getIfModified, "volcano", BosaiUrl + "volcano/const/volcano_list.json");
		if (json is null)
			return;
		await WriteCsvAsync(csvPath, json.RootElement.EnumerateArray().Select(volcano =>
		{
			var latLon = volcano.Property("latlon");
			return new[]
			{
				volcano.Property("code").Text(),
				volcano.Property("name_jp").Text(),
				volcano.Property("name_en").Text(),
				latLon?[0].Text(),
				latLon?[1].Text(),
				volcano.Property("levelOperation").Text(),
			};
		}));
	}

	// 潮位観測点
	private static async Task ExportTideLevelStationAsync(GetIfModified getIfModified, string csvPath)
	{
		using var json = await GetJsonAsync(getIfModified, "tideLevelStation", BosaiUrl + "tidelevel/const/tide_area.json");
		if (json is null)
			return;
		var stations = json.RootElement.EnumerateObject()
			.SelectMany(area => area.Value.GetProperty("class30s").EnumerateArray())
			.SelectMany(class30 => class30.Property("stations")?.EnumerateArray() ?? Enumerable.Empty<JsonElement>());
		await WriteCsvAsync(csvPath, stations.Select(station => new[]
		{
			station.Property("code").Text(),
			station.Property("name").Text(),
			station.Property("lat").Text(),
			station.Property("lon").Text(),
			station.Property("addr").Text(),
			station.Property("type").Text(),
			station.Property("typeName").Text(),
			station.Property("reference").Text(),
			station.Property("detail").Property("type").Text(),
			station.Property("detail").Property("height1").Text(),
			station.Property("detail").Property("height2").Text(),
			station.Property("detail").Property("height3").Text(),
			station.Property("detail").Property("note").Text(),
			station.Property("max").Property("level").Text(),
			station.Property("max").Property("datetime").Text(),
			station.Property("max").Property("period").Text(),
			station.Property("max").Property("description").Text(),
			station.Property("missing").Text(),
			station.Property("parents").Property("class20").Text(),
			station.Property("parents").Property("class30").Text(),
		}));
	}

	// ウィンドプロファイラ観測局
	private static async Task ExportWindProfilerStationAsync(GetIfModified getIfModified, string csvPath)
	{
		using var json = await GetJsonAsync(getIfModified, "windProfilerStation", BosaiUrl + "windprofiler/const/station.json");
		if (json is null)
			return;
		await WriteCsvAsync(csvPath, json.RootElement.EnumerateObject().Select(station => new[]
		{
			station.Name,
			station.Value.Property("jpname").Text(),
			station.Value.Property("enname").Text(),
			station.Value.Property("lat").Text(),
			station.Value.Property("lon").Text(),
			station.Value.Property("altitude").Text(),
		}));
	}

	// 地震観測地点ファイル (固定長: https://www.data.jma.go.jp/eqev/data/bulletin/data/format/stnfmt_j.html)
	private static async Task ExportSeismicStationAsync(GetIfModified getIfModified, string csvPath)
	{
		using var response = await getIfModified("seismicStation", BulletinUrl + "data/deck/stations.zip");
		if (response is null)
			return;
		using var zipArchive = new ZipArchive(await response.Content.ReadAsStreamAsync(), ZipArchiveMode.Read);
		var entry = zipArchive.Entries.Single();
		using var reader = new StreamReader(entry.Open());

		var rows = new List<string?[]>();
		while (await reader.ReadLineAsync() is { } line)
		{
			if (string.IsNullOrWhiteSpace(line))
				continue;
			if (line.Length < 31)
				throw new InvalidDataException($"Unexpected line in {entry.FullName}: {line}");

			rows.Add(
			[
				line[0..6].Trim(),
				ParseFixedInteger(line[14..17]),
				ParseFixedDecimal(line[17..21], 2),
				ParseFixedInteger(line[21..23]),
				ParseFixedDecimal(line[23..27], 2),
				ParseFixedInteger(line[27..31]),
			]);
		}
		await WriteCsvAsync(csvPath, rows);
	}

	// 多言語辞書 (https://www.data.jma.go.jp/developer/multilingual.html)
	private static async Task ExportMultilingualDictionaryAsync(GetIfModified getIfModified, string csvPath)
	{
		using var response = await getIfModified("multilingual", DeveloperUrl + "jma_multilingual.xlsx");
		if (response is null)
			return;
		using var stream = await response.Content.ReadAsStreamAsync();
		using var workbook = XlsxWorkbook.Open(stream);
		const string SheetName = "多言語辞書（本体）";
		var sheet = workbook.GetSheet(SheetName) ?? throw new InvalidOperationException($"Sheet {SheetName} not found.");

		// 列の追加・並べ替えで値がずれたまま出力されないようにする
		var header = sheet.GetRow(0);
		var actualColumns = Enumerable.Range(0, MultilingualColumns.Length + 1).Select(j => header?.GetCell(j)?.Value?.Trim() ?? "");
		if (!actualColumns.SequenceEqual(MultilingualColumns.Append("")))
			throw new InvalidOperationException($"Unexpected columns in {SheetName}: {string.Join(", ", actualColumns)}");

		var rows = new List<string?[]>();
		for (var i = 1; i <= sheet.LastRowNum; i++)
		{
			var row = sheet.GetRow(i);
			if (row is null || string.IsNullOrWhiteSpace(row.GetCell(0)?.Value))
				continue;
			// コード表の CSV と同じく前後の空白を除く。コードは先頭の 0 が欠けていても元データのまま出力する
			rows.Add([.. Enumerable.Range(0, MultilingualColumns.Length).Select(j => row.GetCell(j)?.Value?.Trim())]);
		}
		await WriteCsvAsync(csvPath, rows);
	}

	// 北西太平洋津波情報の予報地点 (HTML の表)
	private static async Task ExportForecastPointAsync(GetIfModified getIfModified, string csvPath)
	{
		const string Url = NwptacUrl + "forecastpoint.html";
		using var response = await getIfModified("nwptacForecastPoint", Url);
		if (response is null)
			return;
		var document = await new HtmlParser().ParseDocumentAsync(await response.Content.ReadAsStreamAsync(), CancellationToken.None);
		var table = document.QuerySelectorAll<IHtmlTableElement>("table")
			.FirstOrDefault(t => t.Rows.FirstOrDefault()?.Cells.Any(c => CellText(c) == "Forecast Point") == true)
			?? throw new InvalidOperationException($"Forecast point table not found: {Url}");

		var actualColumns = table.Rows[0].Cells.Select(CellText);
		if (!actualColumns.SequenceEqual(ForecastPointColumns))
			throw new InvalidOperationException($"Unexpected columns in {Url}: {string.Join(", ", actualColumns)}");

		// 海岸ブロックの列は rowspan でまとめられているので、各行に展開する
		var rows = new List<string?[]>();
		string? block = null;
		var blockRows = 0;
		foreach (var row in table.Rows.Skip(1))
		{
			var cells = row.Cells.ToArray();
			if (blockRows == 0)
			{
				if (cells.Length != ForecastPointColumns.Length)
					throw new InvalidDataException($"Unexpected row in {Url}: {row.TextContent}");
				block = CellText(cells[0]);
				blockRows = cells[0].RowSpan;
				cells = cells[1..];
			}
			else if (cells.Length != ForecastPointColumns.Length - 1)
				throw new InvalidDataException($"Unexpected row in {Url}: {row.TextContent}");
			blockRows--;
			rows.Add([block, .. cells.Select(CellText)]);
		}
		if (rows.Count == 0 || blockRows != 0)
			throw new InvalidDataException($"Unexpected rows in {Url}");
		await WriteCsvAsync(csvPath, rows);
	}

	// 改行やタグによる空白をまとめる
	private static string CellText(IElement cell)
		=> string.Join(' ', cell.TextContent.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

	private static string ParseFixedInteger(string field)
		=> int.Parse(field, NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);

	// Fortran の F 形式。小数点が省略されている場合は末尾 decimals 桁を小数部として扱う
	private static string ParseFixedDecimal(string field, int decimals)
	{
		var text = field.Trim();
		var value = text.Contains('.')
			? decimal.Parse(text, CultureInfo.InvariantCulture)
			: int.Parse(text, CultureInfo.InvariantCulture) / (decimal)Math.Pow(10, decimals);
		return value.ToString("F" + decimals, CultureInfo.InvariantCulture);
	}

	private static async Task<JsonDocument?> GetJsonAsync(GetIfModified getIfModified, string key, string url)
	{
		using var response = await getIfModified(key, url);
		return response is null ? null : await JsonDocument.ParseAsync(await response.Content.ReadAsStreamAsync());
	}

	// 既存のコード表の CSV と同じ規則で書き出す
	private static async Task WriteCsvAsync(string csvPath, IEnumerable<IEnumerable<string?>> rows)
	{
		using var csvWriter = new StreamWriter(csvPath, false) { NewLine = "\n" };
		foreach (var row in rows)
			await csvWriter.WriteLineAsync(string.Join(',', row.Select(EscapeCsvValue)));
	}

	private static string EscapeCsvValue(string? value)
	{
		value ??= "";
		if (value.Contains('\n'))
			value = value.Replace("\r\n", "\n").Replace("\n", "\\n");
		if (value.Contains(','))
			value = "\"" + value.Replace("\"", "\\\"") + "\"";
		return value;
	}

	private static JsonElement? Property(this JsonElement element, string name)
		=> element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value) ? value : null;

	private static JsonElement? Property(this JsonElement? element, string name)
		=> element?.Property(name);

	// 数値は元の表記のまま出力する
	private static string? Text(this JsonElement? element)
		=> element?.ValueKind switch
		{
			JsonValueKind.String => element.Value.GetString(),
			JsonValueKind.Number => element.Value.GetRawText(),
			JsonValueKind.True => "true",
			JsonValueKind.False => "false",
			_ => null,
		};

	private static string? Text(this JsonElement element)
		=> ((JsonElement?)element).Text();
}
