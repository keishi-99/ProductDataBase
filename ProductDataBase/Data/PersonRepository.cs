using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;

namespace ProductDatabase.Data {
    internal static class PersonRepository {

        // 有効な担当者一覧を "XX.名前" 形式で返す（IDは2桁ゼロパディング）
        public static List<string> GetActivePersonDisplayList() {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            return con.Query<PersonDef>(
                "SELECT PersonID, PersonName FROM M_Person WHERE IsActive = 1 ORDER BY PersonID")
                .Select(p => $"{p.PersonID:D2}.{p.PersonName}")
                .ToList();
        }

        // 全担当者を返す（管理画面用）
        public static List<PersonDef> GetAll() {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            return con.Query<PersonDef>(
                "SELECT * FROM M_Person ORDER BY PersonID")
                .ToList();
        }

        // 有効な担当者一覧を返す（ComboBox用）
        public static List<PersonDef> GetActivePersons() {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            return con.Query<PersonDef>(
                "SELECT * FROM M_Person WHERE IsActive = 1 ORDER BY PersonID")
                .ToList();
        }

        public static void Insert(PersonDef personInfo) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            con.Execute(
                "INSERT INTO M_Person (PersonName, IsActive) VALUES (@PersonName, @IsActive)",
                new { personInfo.PersonName, IsActive = personInfo.IsActive });
        }

        public static void Update(PersonDef personInfo) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            con.Execute(
                "UPDATE M_Person SET PersonName=@PersonName, IsActive=@IsActive WHERE PersonID=@PersonID",
                new { personInfo.PersonName, IsActive = personInfo.IsActive, personInfo.PersonID });
        }

        // 同一名前が既存レコードに存在するか確認する（excludeId は編集時に自身を除外するために使用）
        public static bool ExistsName(string name, long excludeId = 0) {
            if (string.IsNullOrWhiteSpace(name)) { return false; }
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            con.Open();
            return con.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM M_Person WHERE PersonName=@Name AND PersonID != @ExcludeId",
                new { Name = name.Trim(), ExcludeId = excludeId }) > 0;
        }
    }
}
