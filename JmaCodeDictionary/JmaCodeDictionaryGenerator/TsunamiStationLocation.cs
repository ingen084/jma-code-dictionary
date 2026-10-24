using System.Globalization;
using System.Text.RegularExpressions;

// 津波観測点の地図（画像上の位置）から津波観測点の緯度経度を推定する
// このリポジトリの CSV だけから毎回作り直す
public static partial class TsunamiStationLocation
{
	// 防災科学技術研究所の観測点は地震観測点と同じ場所にあるので、その座標を使う
	private const string NiedName = "防災科学技術研究所";
	private const int MinControlCount = 3;
	// 基準点の誤差がこれを超えたら、地図が変わったか対応付けが誤っているとみなして失敗させる
	private const double MaxControlErrorKm = 2;
	// 推定位置の近くに地震観測点が 1 つだけあるときに限り、その観測点とみなす
	private const double MaxSnapDistanceKm = 2;
	private const double MinSecondSnapDistanceKm = 3;

	// 地図と潮位観測点で表記の異なる所属機関
	private static readonly Dictionary<string, string> OrganizationNames = new()
	{
		["国土交通省港湾局"] = "港湾局",
	};

	public static async Task ExportAsync(string basePath, string csvPath)
	{
		var pointTsunamiPath = basePath + "PointTsunami.csv";
		var tideLevelStationPath = basePath + "web-json/TideLevelStation.csv";
		var seismicStationPath = basePath + "bulletin/SeismicStation.csv";
		var mapPaths = Directory.Exists(basePath + "tsunamimap/")
			? Directory.GetFiles(basePath + "tsunamimap/", "*.csv").Order(StringComparer.Ordinal).ToArray()
			: [];
		if (!File.Exists(pointTsunamiPath) || !File.Exists(tideLevelStationPath) || !File.Exists(seismicStationPath) || mapPaths.Length == 0)
		{
			// 初回の実行で元の CSV がまだ無い場合
			Console.WriteLine("Skipped: tsunami station location (source CSV not found)");
			return;
		}

		// ヘッダ部でのみ使用する観測点名には、同じ名前の観測点があるものがあるので除く
		var stations = (await CsvFile.ReadAsync(pointTsunamiPath))
			.Where(row => !row[5].Contains("ヘッダ部でのみ使用"))
			.Select(row => (Code: row[0], Name: row[1]))
			.ToArray();
		var duplicateName = stations.GroupBy(s => s.Name).FirstOrDefault(g => g.Count() > 1);
		if (duplicateName is not null)
			throw new InvalidDataException($"Duplicate station name in {pointTsunamiPath}: {duplicateName.Key}");

		// 括弧の中を除いた名前と所属機関が 1 つの潮位観測点にだけ一致する場合に、基準点として使う
		var tideLevelStations = (await CsvFile.ReadAsync(tideLevelStationPath))
			.Where(row => row[2] != "" && row[3] != "")
			.Select(row => (Key: (StripParentheses(row[1]), StripParentheses(row[6])), Lat: Parse(row[2]), Lon: Parse(row[3])))
			.GroupBy(s => s.Key)
			.Where(g => g.Count() == 1)
			.ToDictionary(g => g.Key, g => g.Single());

		var seismicStations = (await CsvFile.ReadAsync(seismicStationPath))
			.Select(row => (Code: row[0], Lat: Parse(row[3]) + Parse(row[4]) / 60, Lon: Parse(row[1]) + Parse(row[2]) / 60))
			.ToArray();

		// 観測点ごとに、縮尺の最も細かい地図での推定位置を使う
		var estimates = new Dictionary<string, (string Organization, double Lat, double Lon, string Page, double Scale)>();
		foreach (var mapPath in mapPaths)
		{
			var page = Path.GetFileNameWithoutExtension(mapPath);
			var markers = (await CsvFile.ReadAsync(mapPath))
				.Select(row => (Name: row[0], Organization: row[1], X: (Parse(row[2]) + Parse(row[4])) / 2, Y: (Parse(row[3]) + Parse(row[5])) / 2))
				.ToArray();

			var controls = markers
				.GroupBy(m => (m.Name, NormalizeOrganization(m.Organization)))
				.Where(g => g.Count() == 1 && tideLevelStations.ContainsKey(g.Key))
				.Select(g => (Marker: g.Single(), Station: tideLevelStations[g.Key]))
				.ToArray();
			if (controls.Length < MinControlCount)
			{
				Console.WriteLine($"Skipped: tsunami map {page} ({controls.Length} control points)");
				continue;
			}

			// 地図は緯度・経度がそれぞれ画像の縦・横に比例しているので、軸ごとに一次式で当てはめる
			var (lonScale, lonOffset) = FitLine(controls.Select(c => (c.Marker.X, c.Station.Lon)));
			var (latScale, latOffset) = FitLine(controls.Select(c => (c.Marker.Y, c.Station.Lat)));
			var errors = controls
				.Select(c => (c.Marker.Name, Error: DistanceKm(latScale * c.Marker.Y + latOffset, lonScale * c.Marker.X + lonOffset, c.Station.Lat, c.Station.Lon)))
				.ToArray();
			var maxError = errors.MaxBy(e => e.Error);
			Console.WriteLine($"Tsunami map {page}: {controls.Length} control points, max error {maxError.Error:F2} km ({maxError.Name})");
			if (maxError.Error > MaxControlErrorKm)
				throw new InvalidDataException($"Control point error too large in tsunami map {page}: {maxError.Name} {maxError.Error:F2} km");

			foreach (var marker in markers)
			{
				if (estimates.TryGetValue(marker.Name, out var current) && current.Scale <= Math.Abs(lonScale))
					continue;
				estimates[marker.Name] = (marker.Organization, latScale * marker.Y + latOffset, lonScale * marker.X + lonOffset, page, Math.Abs(lonScale));
			}
		}

		var rows = new List<string?[]>();
		var snappedStations = new HashSet<string>();
		foreach (var station in stations)
		{
			if (!estimates.TryGetValue(station.Name, out var estimate))
				continue;

			if (estimate.Organization == NiedName)
			{
				var nearest = seismicStations
					.Select(s => (Station: s, Distance: DistanceKm(estimate.Lat, estimate.Lon, s.Lat, s.Lon)))
					.OrderBy(s => s.Distance)
					.Take(2)
					.ToArray();
				if (nearest[0].Distance <= MaxSnapDistanceKm && nearest[1].Distance >= Math.Max(MinSecondSnapDistanceKm, nearest[0].Distance * 3))
				{
					if (!snappedStations.Add(nearest[0].Station.Code))
						throw new InvalidDataException($"Seismic station {nearest[0].Station.Code} matched more than one tsunami station");
					rows.Add([station.Code, station.Name, estimate.Organization, Format(nearest[0].Station.Lat, 4), Format(nearest[0].Station.Lon, 4), "SeismicStation", nearest[0].Station.Code]);
					continue;
				}
			}
			rows.Add([station.Code, station.Name, estimate.Organization, Format(estimate.Lat, 3), Format(estimate.Lon, 3), "TsunamiMap", estimate.Page]);
		}
		Console.WriteLine($"Tsunami station location: {rows.Count} stations ({snappedStations.Count} from seismic stations)");
		await CsvFile.WriteAsync(csvPath, rows);
	}

	// 最小二乗法で y = a * x + b を求める
	private static (double A, double B) FitLine(IEnumerable<(double X, double Y)> points)
	{
		var array = points.ToArray();
		var meanX = array.Average(p => p.X);
		var meanY = array.Average(p => p.Y);
		var a = array.Sum(p => (p.X - meanX) * (p.Y - meanY)) / array.Sum(p => (p.X - meanX) * (p.X - meanX));
		return (a, meanY - a * meanX);
	}

	// 狭い範囲の距離なので、緯度経度の差をそのまま距離に換算する
	private static double DistanceKm(double lat1, double lon1, double lat2, double lon2)
	{
		const double KmPerDegree = 111.2;
		var dy = (lat1 - lat2) * KmPerDegree;
		var dx = (lon1 - lon2) * KmPerDegree * Math.Cos((lat1 + lat2) / 2 * Math.PI / 180);
		return Math.Sqrt(dx * dx + dy * dy);
	}

	private static string NormalizeOrganization(string name)
		=> OrganizationNames.GetValueOrDefault(name, name);

	private static string StripParentheses(string value)
		=> ParenthesesRegex().Replace(value, "").Trim();

	private static double Parse(string value)
		=> double.Parse(value, CultureInfo.InvariantCulture);

	private static string Format(double value, int decimals)
		=> value.ToString("F" + decimals, CultureInfo.InvariantCulture);

	[GeneratedRegex(@"（[^（）]*）|\([^()]*\)")]
	private static partial Regex ParenthesesRegex();
}
