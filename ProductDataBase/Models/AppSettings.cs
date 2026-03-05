namespace ProductDatabase {
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
}
