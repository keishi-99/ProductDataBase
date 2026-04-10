using ProductDatabase.Models;

namespace ProductDatabase.Services {
    internal static class SettingsLoader {
        private static readonly System.Text.Json.JsonSerializerOptions _jsonOptions = new() { PropertyNameCaseInsensitive = true };

        // JSON設定ファイルを読み込んでGeneralSettingsとして返す
        public static GeneralSettings Load(string jsonFilePath) {
            if (!File.Exists(jsonFilePath)) {
                throw new DirectoryNotFoundException($"'{jsonFilePath}'\nが見つかりません。");
            }
            return System.Text.Json.JsonSerializer.Deserialize<GeneralSettings>(
                File.ReadAllText(jsonFilePath),
                _jsonOptions
            ) ?? new GeneralSettings();
        }
    }
}
