using System.Text;

// 既存のコード表の CSV と同じ規則で読み書きする
// ヘッダー行なし、カンマを含む値だけ " で囲み、値の中の " は \"、改行は \n とする
public static class CsvFile
{
	public static async Task WriteAsync(string csvPath, IEnumerable<IEnumerable<string?>> rows)
	{
		using var csvWriter = new StreamWriter(csvPath, false) { NewLine = "\n" };
		foreach (var row in rows)
			await csvWriter.WriteLineAsync(string.Join(',', row.Select(EscapeValue)));
	}

	public static string EscapeValue(string? value)
	{
		value ??= "";
		if (value.Contains('\n'))
			value = value.Replace("\r\n", "\n").Replace("\n", "\\n");
		if (value.Contains(','))
			value = "\"" + value.Replace("\"", "\\\"") + "\"";
		return value;
	}

	public static async Task<List<string[]>> ReadAsync(string csvPath)
		=> [.. (await File.ReadAllLinesAsync(csvPath)).Select(ParseLine)];

	// 改行を表す \n は元に戻さずそのまま返す
	public static string[] ParseLine(string line)
	{
		var values = new List<string>();
		var i = 0;
		while (true)
		{
			if (i < line.Length && line[i] == '"')
			{
				var value = new StringBuilder();
				var j = i + 1;
				while (true)
				{
					if (j >= line.Length)
						throw new InvalidDataException($"Unterminated quoted value: {line}");
					if (line[j] == '\\' && j + 1 < line.Length && line[j + 1] == '"')
					{
						value.Append('"');
						j += 2;
					}
					else if (line[j] == '"' && (j + 1 == line.Length || line[j + 1] == ','))
						break;
					else
						value.Append(line[j++]);
				}
				values.Add(value.ToString());
				i = j + 1;
			}
			else
			{
				var end = line.IndexOf(',', i);
				if (end < 0)
					end = line.Length;
				values.Add(line[i..end]);
				i = end;
			}
			if (i >= line.Length)
				break;
			// 区切りのカンマを読み飛ばす
			i++;
		}
		return [.. values];
	}
}
