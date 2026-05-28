using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Models;

namespace ProductDatabase.Data {
    internal static class PersonRepository {

        // 有効な担当者一覧を "XX.名前" 形式で返す（IDは2桁ゼロパディング）
        public static List<string> GetActivePersonDisplayList() {
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                return con.Query<PersonDef>(
                    "SELECT PersonID, PersonName FROM M_Person WHERE IsActive = 1 ORDER BY PersonID")
                    .Select(p => $"{p.PersonID:D2}.{p.PersonName}")
                    .ToList();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(GetActivePersonDisplayList), ex, null);
                throw;
            }
        }

        // 全担当者を返す（管理画面用）
        public static List<PersonDef> GetAll() {
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                return con.Query<PersonDef>(
                    "SELECT * FROM M_Person ORDER BY PersonID")
                    .ToList();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(GetAll), ex, null);
                throw;
            }
        }

        // 有効な担当者一覧を返す（ComboBox用）
        public static List<PersonDef> GetActivePersons() {
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                return con.Query<PersonDef>(
                    "SELECT * FROM M_Person WHERE IsActive = 1 ORDER BY PersonID")
                    .ToList();
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(GetActivePersons), ex, null);
                throw;
            }
        }

        public static void Insert(PersonDef personInfo) {
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                con.Execute(
                    "INSERT INTO M_Person (PersonName, IsActive) VALUES (@PersonName, @IsActive)",
                    new { personInfo.PersonName, IsActive = personInfo.IsActive });
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(Insert), ex, $"PersonName: {personInfo.PersonName}");
                throw;
            }
        }

        public static void Update(PersonDef personInfo) {
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                con.Execute(
                    "UPDATE M_Person SET PersonName=@PersonName, IsActive=@IsActive WHERE PersonID=@PersonID",
                    new { personInfo.PersonName, IsActive = personInfo.IsActive, personInfo.PersonID });
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(Update), ex, $"PersonID: {personInfo.PersonID}, PersonName: {personInfo.PersonName}");
                throw;
            }
        }

        // 同一名前が既存レコードに存在するか確認する（excludeId は編集時に自身を除外するために使用）
        public static bool ExistsName(string name, long excludeId = 0) {
            if (string.IsNullOrWhiteSpace(name)) { return false; }
            try {
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                con.Open();
                return con.ExecuteScalar<int>(
                    "SELECT COUNT(*) FROM M_Person WHERE PersonName=@Name AND PersonID != @ExcludeId",
                    new { Name = name.Trim(), ExcludeId = excludeId }) > 0;
            } catch (Exception ex) {
                Logger.AppendErrorLog(nameof(ExistsName), ex, $"PersonName: {name}");
                throw;
            }
        }
    }
}
