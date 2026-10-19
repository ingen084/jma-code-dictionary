using System.Globalization;
using System.Text.RegularExpressions;

// 北西太平洋津波情報の予報地点の緯度経度（58.8N, 164.5E など）を数値に変換する
public static partial class NwptaForecastPoint
{
	public static async Task ExportAsync(string basePath, string csvPath)
	{
		var sourcePath = basePath + "nwptac/ForecastPoint.csv";
		if (!File.Exists(sourcePath))
		{
			// 初回の実行で元の CSV がまだ無い場合
			Console.WriteLine("Skipped: NWPTA forecast point (source CSV not found)");
			return;
		}

		var rows = new List<string?[]>();
		foreach (var row in await CsvFile.ReadAsync(sourcePath))
		{
			if (row.Length != 5)
				throw new InvalidDataException($"Unexpected row in {sourcePath}: {string.Join(',', row)}");
			rows.Add([row[0], row[1], ParseCoordinate(row[2], 'N', 'S'), ParseCoordinate(row[3], 'E', 'W'), row[4]]);
		}
		Console.WriteLine($"NWPTA forecast point: {rows.Count} points");
		await CsvFile.WriteAsync(csvPath, rows);
	}

	// 北緯・東経を正、南緯・西経を負とする。小数点以下の桁数は元の表記のまま
	private static string ParseCoordinate(string value, char positive, char negative)
	{
		var match = CoordinateRegex().Match(value);
		var hemisphere = match.Success ? match.Groups[2].Value[0] : default;
		if (hemisphere != positive && hemisphere != negative)
			throw new InvalidDataException($"Unexpected coordinate: {value}");
		var number = decimal.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
		return (hemisphere == negative && number != 0 ? -number : number).ToString(CultureInfo.InvariantCulture);
	}

	[GeneratedRegex(@"^(\d+(?:\.\d+)?)([NSEW])$")]
	private static partial Regex CoordinateRegex();
}
