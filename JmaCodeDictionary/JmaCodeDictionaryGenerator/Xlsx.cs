using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

// コード表の xlsx からセルの値を読み出すためだけの最小限のリーダー
public sealed partial class XlsxWorkbook : IDisposable
{
	private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

	// 一般的な数値として扱う場合の書式
	private const string GeneralFormat = "0.##########";

	// 組み込みの表示形式 (ECMA-376 Part 1 18.8.30)
	private static readonly Dictionary<int, string> BuiltinFormats = new()
	{
		[0] = "General",
		[1] = "0",
		[2] = "0.00",
		[3] = "#,##0",
		[4] = "#,##0.00",
		[9] = "0%",
		[10] = "0.00%",
		[11] = "0.00E+00",
		[37] = "#,##0 ;(#,##0)",
		[38] = "#,##0 ;[Red](#,##0)",
		[39] = "#,##0.00;(#,##0.00)",
		[40] = "#,##0.00;[Red](#,##0.00)",
		[48] = "##0.0E+0",
		[49] = "@",
	};

	private readonly ZipArchive archive;
	private readonly Dictionary<string, string> sheetPaths;
	private readonly string[] sharedStrings;
	private readonly Dictionary<int, string> customFormats;
	private readonly int[] cellFormatIds;
	private readonly Dictionary<int, NumberFormat> numberFormats = [];

	private XlsxWorkbook(ZipArchive archive)
	{
		this.archive = archive;

		var relationships = Load("xl/_rels/workbook.xml.rels")!.Root!.Elements(PackageRelationships + "Relationship")
			.ToDictionary(r => (string)r.Attribute("Id")!, r => ResolvePath((string)r.Attribute("Target")!));
		sheetPaths = Load("xl/workbook.xml")!.Root!.Element(Main + "sheets")!.Elements(Main + "sheet")
			.ToDictionary(s => (string)s.Attribute("name")!, s => relationships[(string)s.Attribute(Relationships + "id")!]);

		sharedStrings = Load("xl/sharedStrings.xml", optional: true)?.Root!.Elements(Main + "si").Select(ReadText).ToArray() ?? [];

		var styles = Load("xl/styles.xml", optional: true)?.Root;
		customFormats = styles?.Element(Main + "numFmts")?.Elements(Main + "numFmt")
			.ToDictionary(f => (int)f.Attribute("numFmtId")!, f => (string)f.Attribute("formatCode")!) ?? [];
		cellFormatIds = styles?.Element(Main + "cellXfs")?.Elements(Main + "xf")
			.Select(xf => (int?)xf.Attribute("numFmtId") ?? 0).ToArray() ?? [];
	}

	public static XlsxWorkbook Open(Stream stream)
		=> new(new ZipArchive(stream, ZipArchiveMode.Read));

	public XlsxSheet? GetSheet(string name)
	{
		if (!sheetPaths.TryGetValue(name, out var path))
			return null;

		var rows = new Dictionary<int, XlsxRow>();
		var rowIndex = -1;
		foreach (var row in Load(path)!.Root!.Element(Main + "sheetData")!.Elements(Main + "row"))
		{
			rowIndex = (int?)row.Attribute("r") - 1 ?? rowIndex + 1;

			var cells = new Dictionary<int, XlsxCell>();
			var columnIndex = -1;
			foreach (var cell in row.Elements(Main + "c"))
			{
				columnIndex = (string?)cell.Attribute("r") is { } reference ? ParseColumnIndex(reference) : columnIndex + 1;
				cells[columnIndex] = ReadCell(cell);
			}
			rows[rowIndex] = new XlsxRow(cells);
		}
		return new XlsxSheet(rows);
	}

	public void Dispose()
		=> archive.Dispose();

	private XDocument? Load(string path, bool optional = false)
	{
		var entry = archive.GetEntry(path);
		if (entry is null)
			return optional ? null : throw new InvalidDataException($"{path} not found.");

		using var stream = entry.Open();
		return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
	}

	private XlsxCell ReadCell(XElement cell)
	{
		var value = cell.Element(Main + "v")?.Value;
		switch ((string?)cell.Attribute("t"))
		{
			case "s":
				return new(value is null ? null : sharedStrings[int.Parse(value, CultureInfo.InvariantCulture)], false);
			case "inlineStr":
				return new(cell.Element(Main + "is") is { } inlineString ? ReadText(inlineString) : null, false);
			case "str":
			case "d":
				return new(value, false);
			case "b":
				return new(value is null ? null : (value == "1").ToString(), false);
			case "e":
				return new(null, false);
		}

		if (value is null)
			return new(null, false);

		var number = double.Parse(value, CultureInfo.InvariantCulture);
		var format = GetNumberFormat((int?)cell.Attribute("s") ?? 0);
		if (format.IsDate)
			return new(DateTime.FromOADate(number).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture), true);
		// 先頭の 0 が欠損しないよう、セルの表示形式に従って文字列化する
		return new(number.ToString(format.Format, CultureInfo.InvariantCulture), false);
	}

	private NumberFormat GetNumberFormat(int styleIndex)
	{
		var id = styleIndex < cellFormatIds.Length ? cellFormatIds[styleIndex] : 0;
		if (numberFormats.TryGetValue(id, out var format))
			return format;

		if ((customFormats.GetValueOrDefault(id) ?? BuiltinFormats.GetValueOrDefault(id)) is { } code)
			format = ConvertFormat(code);
		// 日付・時刻の組み込み表示形式 (日本語ロケールのものを含む)
		else if (id is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58))
			format = new(true, "");
		else
			throw new NotSupportedException($"Unsupported number format id: {id}");

		numberFormats[id] = format;
		return format;
	}

	private static NumberFormat ConvertFormat(string code)
	{
		// 文字列リテラル・エスケープ・[色] などを除いた部分で書式の種類を判定する
		var tokens = FormatTokenRegex().Replace(code, "");
		if (tokens.Equals("General", StringComparison.OrdinalIgnoreCase) || tokens.Contains('@'))
			return new(false, GeneralFormat);
		if (tokens.IndexOfAny(['y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 's', 'S']) >= 0)
			return new(true, "");
		if (tokens.IndexOfAny(['?', '/']) >= 0)
			throw new NotSupportedException($"Unsupported number format: {code}");

		// Excel の表示形式は .NET のカスタム数値書式とほぼ互換なので、互換でない部分だけ置き換える
		return new(false, FormatTokenRegex().Replace(code, m => m.Value[0] switch
		{
			'[' or '*' => "",
			'_' => " ",
			_ => m.Value,
		}));
	}

	private static string ResolvePath(string target)
		=> target.StartsWith('/') ? target[1..] : "xl/" + target;

	private static int ParseColumnIndex(string reference)
	{
		var index = 0;
		foreach (var c in reference.TakeWhile(char.IsAsciiLetter))
			index = index * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
		return index - 1;
	}

	// ふりがな (rPh) は含めず、本文と書式付き文字列の各ランだけを連結する
	private static string ReadText(XElement element)
		=> EscapedCharRegex().Replace(
			element.Element(Main + "t")?.Value ?? string.Concat(element.Elements(Main + "r").Select(r => r.Element(Main + "t")?.Value)),
			m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());

	[GeneratedRegex("\"[^\"]*\"|\\\\.|\\[[^\\]]*\\]|_.|\\*.")]
	private static partial Regex FormatTokenRegex();

	[GeneratedRegex("_x([0-9A-Fa-f]{4})_")]
	private static partial Regex EscapedCharRegex();

	private readonly record struct NumberFormat(bool IsDate, string Format);
}

public sealed class XlsxSheet(Dictionary<int, XlsxRow> rows)
{
	public int LastRowNum { get; } = rows.Count == 0 ? 0 : rows.Keys.Max();

	public XlsxRow? GetRow(int index)
		=> rows.GetValueOrDefault(index);
}

public sealed class XlsxRow(Dictionary<int, XlsxCell> cells)
{
	public XlsxCell? GetCell(int index)
		=> cells.GetValueOrDefault(index);
}

public sealed record XlsxCell(string? Value, bool IsDateFormatted);
