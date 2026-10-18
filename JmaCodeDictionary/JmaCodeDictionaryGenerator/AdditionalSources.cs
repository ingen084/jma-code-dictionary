using System.Globalization;
using System.IO.Compression;
using System.Text.Json;

// 防災情報XMLのコード表以外の公開データから CSV を生成する
public static class AdditionalSources
{
	private const string BosaiUrl = "https://www.jma.go.jp/bosai/";
	private const string BulletinUrl = "https://www.data.jma.go.jp/eqev/data/bulletin/";

	public static async Task ExportAsync(Func<string, Task<HttpResponseMessage>> getAsync, string basePath)
	{
		// 防災情報 (bosai) の JSON
		var webJsonPath = basePath + "web-json/";
		Directory.CreateDirectory(webJsonPath);
		await ExportVolcanoAsync(getAsync, webJsonPath + "Volcano.csv");
		await ExportTideLevelStationAsync(getAsync, webJsonPath + "TideLevelStation.csv");
		await ExportWindProfilerStationAsync(getAsync, webJsonPath + "WindProfilerStation.csv");

		// 地震月報（カタログ編）
		var bulletinPath = basePath + "bulletin/";
		Directory.CreateDirectory(bulletinPath);
		await ExportSeismicStationAsync(getAsync, bulletinPath + "SeismicStation.csv");
	}

	// 火山
	private static async Task ExportVolcanoAsync(Func<string, Task<HttpResponseMessage>> getAsync, string csvPath)
	{
		using var json = await GetJsonAsync(getAsync, BosaiUrl + "volcano/const/volcano_list.json");
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
	private static async Task ExportTideLevelStationAsync(Func<string, Task<HttpResponseMessage>> getAsync, string csvPath)
	{
		using var json = await GetJsonAsync(getAsync, BosaiUrl + "tidelevel/const/tide_area.json");
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
	private static async Task ExportWindProfilerStationAsync(Func<string, Task<HttpResponseMessage>> getAsync, string csvPath)
	{
		using var json = await GetJsonAsync(getAsync, BosaiUrl + "windprofiler/const/station.json");
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
	private static async Task ExportSeismicStationAsync(Func<string, Task<HttpResponseMessage>> getAsync, string csvPath)
	{
		using var response = await GetAsync(getAsync, BulletinUrl + "data/deck/stations.zip");
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

	private static async Task<HttpResponseMessage> GetAsync(Func<string, Task<HttpResponseMessage>> getAsync, string url)
	{
		var response = await getAsync(url);
		response.EnsureSuccessStatusCode();
		return response;
	}

	private static async Task<JsonDocument> GetJsonAsync(Func<string, Task<HttpResponseMessage>> getAsync, string url)
	{
		using var response = await GetAsync(getAsync, url);
		return await JsonDocument.ParseAsync(await response.Content.ReadAsStreamAsync());
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
