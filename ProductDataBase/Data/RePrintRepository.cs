using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Models;
using ProductDatabase.Other;

namespace ProductDatabase.Data {
    // 再印刷テーブルへのINSERT操作を担当するリポジトリクラス
    internal static class RePrintRepository {

        // 再印刷テーブルにレコードをINSERTして生成されたROWIDを返す
        public static int InsertRePrintRecord(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());

            var sql =
                $"""
                INSERT INTO {Constants.TRePrintTableName}
                (
                    ProductID,
                    ProductName,
                    ProductType,
                    ProductModel,
                    SerialPrintType,
                    OrderNumber,
                    ProductNumber,
                    OLesNumber,
                    Quantity,
                    Person,
                    RegDate,
                    Revision,
                    RevisionGroup,
                    SerialFirst,
                    SerialLast,
                    Comment
                )
                VALUES
                (
                    @ProductID,
                    @ProductName,
                    @ProductType,
                    @ProductModel,
                    @SerialPrintType,
                    @OrderNumber,
                    @ProductNumber,
                    @OLesNumber,
                    @Quantity,
                    @Person,
                    @RegDate,
                    @Revision,
                    @RevisionGroup,
                    @SerialFirst,
                    @SerialLast,
                    @Comment
                );
                """;

            con.Execute(sql, new {
                productMaster.ProductID,
                ProductName   = productMaster.ProductName.NullIfWhiteSpace(),
                ProductType   = productMaster.ProductType.NullIfWhiteSpace(),
                ProductModel  = productMaster.ProductModel.NullIfWhiteSpace(),
                productMaster.SerialPrintType,
                OrderNumber   = productRegisterWork.OrderNumber.NullIfWhiteSpace(),
                ProductNumber = productRegisterWork.ProductNumber.NullIfWhiteSpace(),
                OLesNumber    = productRegisterWork.OLesNumber.NullIfWhiteSpace(),
                productRegisterWork.Quantity,
                Person        = productRegisterWork.Person.NullIfWhiteSpace(),
                RegDate       = productRegisterWork.RegDate.NullIfWhiteSpace(),
                Revision      = productRegisterWork.Revision.NullIfWhiteSpace(),
                productMaster.RevisionGroup,
                SerialFirst   = productRegisterWork.SerialFirst.NullIfWhiteSpace(),
                SerialLast    = productRegisterWork.SerialLast.NullIfWhiteSpace(),
                Comment       = productRegisterWork.Comment.NullIfWhiteSpace()
            });

            return con.ExecuteScalar<int>("SELECT last_insert_rowid();");
        }
    }
}
