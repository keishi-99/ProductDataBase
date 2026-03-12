namespace ProductDatabase.Models {
    public class QrSettings {
        public List<string> CategoryItemNumber { get; set; } = [];
        public List<string> CategoryProductName { get; set; } = [];
        public List<string> CategorySubstrateName { get; set; } = [];
        public List<string> CategoryProductType { get; set; } = [];
        public List<string> CategoryType { get; set; } = [];
    }

    public class AppSettings {
        public List<string> PersonList { get; set; } = [];
        public string FontName { get; set; } = "Meiryo UI";
        public float FontSize { get; set; } = 9;
        public bool IsAdministrator { get; set; } = false;
        public bool IsAuthorizedUser { get; set; } = false;
    }

    public class Config {
        public string BackupFolderPath { get; set; } = string.Empty;
        public string[] Administrators { get; set; } = [];
        public string[] AuthorizedUsers { get; set; } = [];
    }

    /// <summary>
    /// appsettings.json 全体に対応する設定クラス。IConfiguration でバインドして使用します。
    /// </summary>
    public class GeneralSettings {
        public string BackupFolderPath { get; set; } = string.Empty;
        public List<string> Persons { get; set; } = [];
        public string[] AuthorizedUsers { get; set; } = [];
        public string[] Administrators { get; set; } = [];
        public string DSN { get; set; } = string.Empty;
        public string UID { get; set; } = string.Empty;
        public string PWD { get; set; } = string.Empty;
    }
}
