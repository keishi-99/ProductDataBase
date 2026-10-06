using ProductDatabase.Services;

namespace ProductDataBase.Tests {
    public class QrCodeParserTests {

        // ===== Parse =====

        [Fact]
        public void Parse_区切りなし_入力全体がProductModelになる() {
            var result = QrCodeParser.Parse("ABC-100");

            Assert.Equal("ABC-100", result.ProductModel);
            Assert.Equal(string.Empty, result.ProductNumber);
            Assert.Equal(0, result.Quantity);
            Assert.Equal(string.Empty, result.OrderNumber);
        }

        [Fact]
        public void Parse_4分割_各項目に分解される() {
            var result = QrCodeParser.Parse("12345//ABC-100//10//ORD-1");

            Assert.Equal("12345", result.ProductNumber);
            Assert.Equal("ABC-100", result.ProductModel);
            Assert.Equal(10, result.Quantity);
            Assert.Equal("ORD-1", result.OrderNumber);
        }

        // 現状の仕様を固定する: 空文字は例外にならず、空のProductModelを返す
        [Fact]
        public void Parse_空文字_空のProductModelを返す() {
            var result = QrCodeParser.Parse("");

            Assert.Equal(string.Empty, result.ProductModel);
        }

        [Theory]
        [InlineData("A//B")]
        [InlineData("A//B//10")]
        [InlineData("A//B//10//C//D")]
        public void Parse_項目数が4でも1でもない_例外になる(string input) {
            var ex = Assert.Throws<Exception>(() => QrCodeParser.Parse(input));

            Assert.Equal("QRコードが正しくありません。", ex.Message);
        }

        [Fact]
        public void Parse_数量が数値以外_例外になる() {
            var ex = Assert.Throws<Exception>(() => QrCodeParser.Parse("A//B//十//C"));

            Assert.Equal("数量に数値以外が入力されています。", ex.Message);
        }

        // 現状の仕様を固定する: 0や負数でも例外にならない
        [Theory]
        [InlineData("0", 0)]
        [InlineData("-5", -5)]
        public void Parse_数量が0や負数_そのまま受け入れる(string quantityText, int expected) {
            var result = QrCodeParser.Parse($"A//B//{quantityText}//C");

            Assert.Equal(expected, result.Quantity);
        }

        // ===== NormalizeProductModel =====

        [Theory]
        [InlineData("ABC", "ABC")]
        [InlineData("ABC-H1", "ABC")]
        [InlineData("ABC-SMT", "ABC")]
        [InlineData("ABC-GH", "ABC")]
        [InlineData("ABC-ACGH", "ABC-AC")]
        [InlineData("ABC-DCGH", "ABC-DC")]
        public void NormalizeProductModel_正規化される(string input, string expected) {
            var result = QrCodeParser.NormalizeProductModel(input);

            Assert.Equal(expected, result);
        }
    }
}
