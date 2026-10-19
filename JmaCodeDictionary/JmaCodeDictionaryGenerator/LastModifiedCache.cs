using System.Text.Encodings.Web;
using System.Text.Json;

// 取得元ごとの Last-Modified を 1 つの JSON ファイルにまとめて保存する
public sealed class LastModifiedCache
{
	private static readonly JsonSerializerOptions JsonOptions = new()
	{
		PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
		WriteIndented = true,
		// 実行環境によって改行コードが変わらないようにする
		NewLine = "\n",
		// 日時のオフセットの + がエスケープされないようにする
		Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
	};

	private readonly string path;
	// 差分が安定するようキーの順に保存する
	private readonly SortedDictionary<string, Entry> entries;

	private LastModifiedCache(string path, SortedDictionary<string, Entry> entries)
	{
		this.path = path;
		this.entries = entries;
	}

	public static async Task<LastModifiedCache> LoadAsync(string path)
	{
		if (!File.Exists(path))
			return new(path, new(StringComparer.Ordinal));

		using var stream = File.OpenRead(path);
		var entries = await JsonSerializer.DeserializeAsync<Dictionary<string, Entry>>(stream, JsonOptions) ?? [];
		return new(path, new(entries, StringComparer.Ordinal));
	}

	// URL が前回と変わっていれば別のファイルとみなし、日時を返さない
	public DateTimeOffset? Get(string key, string url)
		=> entries.TryGetValue(key, out var entry) && entry.Url == url ? entry.LastModified : null;

	public void Set(string key, string url, DateTimeOffset lastModified)
		=> entries[key] = new(url, lastModified);

	public Task SaveAsync()
		=> File.WriteAllTextAsync(path, JsonSerializer.Serialize(entries, JsonOptions) + "\n");

	public sealed record Entry(string Url, DateTimeOffset LastModified);
}
