using ProductDatabase.Services;

namespace ProductDataBase.Tests {
    public class SerialCodeFormatterTests {

        private static readonly DateTime March15 = new(2026, 3, 15);

        // 書式以外は固定値にして、テストしたい引数だけ差し替えられるようにする
        private static string Format(
            string format, string? initial = "AB", string? oLesInitial = "OL", DateTime? regDate = null,
            string revision = "A", int serialNumber = 7, int serialDigit = 3, string oLesSuffix = "C") =>
            SerialCodeFormatter.Format(format, initial, oLesInitial, regDate ?? March15, revision, serialNumber, serialDigit, oLesSuffix);

        // ===== Format =====

        [Fact]
        public void Format_複数のプレースホルダー_全て置換される() {
            Assert.Equal("AB263-007", Format("{T}{Y}{M}-{S}"));
        }

        [Fact]
        public void Format_プレースホルダーなし_そのまま返る() {
            Assert.Equal("FIXED", Format("FIXED"));
        }

        [Fact]
        public void Format_空の書式_空文字を返す() {
            Assert.Equal(string.Empty, Format(""));
        }

        [Fact]
        public void Format_同じプレースホルダーが複数_全て置換される() {
            Assert.Equal("AB-AB", Format("{T}-{T}"));
        }

        [Theory]
        [InlineData(1, "1", "01")]
        [InlineData(9, "9", "09")]
        [InlineData(10, "X", "10")]
        [InlineData(11, "Y", "11")]
        [InlineData(12, "Z", "12")]
        public void Format_月に応じて月コードと2桁の月が入る(int month, string expectedMonthCode, string expectedMonth) {
            var regDate = new DateTime(2026, month, 1);

            Assert.Equal(expectedMonthCode, Format("{M}", regDate: regDate));
            Assert.Equal(expectedMonth, Format("{MM}", regDate: regDate));
        }

        [Fact]
        public void Format_製造年_西暦下2桁になる() {
            Assert.Equal("26", Format("{Y}", regDate: new DateTime(2026, 3, 15)));
            Assert.Equal("05", Format("{Y}", regDate: new DateTime(2005, 3, 15)));
        }

        [Fact]
        public void Format_リビジョン_そのまま入る() {
            Assert.Equal("B", Format("{R}", revision: "B"));
        }

        [Theory]
        [InlineData(7, 3, "007")]
        [InlineData(7, 4, "0007")]
        [InlineData(999, 3, "999")]
        [InlineData(1234, 4, "1234")]
        public void Format_シリアルは桁数でゼロ埋めされる(int serialNumber, int serialDigit, string expected) {
            Assert.Equal(expected, Format("{S}", serialNumber: serialNumber, serialDigit: serialDigit));
        }

        [Fact]
        public void Format_O_Lesの接頭と接尾_入る() {
            Assert.Equal("OLABC", Format("{OT}{T}{SA}"));
        }

        [Fact]
        public void Format_接頭がnull_空文字として扱う() {
            Assert.Equal("-26", Format("{T}-{Y}", initial: null));
            Assert.Equal("-", Format("{OT}-", oLesInitial: null));
        }

        // ===== GetNextOLesSuffix =====

        [Theory]
        [InlineData(null, "A")]
        [InlineData("", "A")]
        [InlineData("  ", "A")]
        [InlineData("A", "B")]
        [InlineData("Y", "Z")]
        [InlineData("b", "C")]
        [InlineData("AB", "B")]
        public void GetNextOLesSuffix_次の文字を返す(string? current, string expected) {
            Assert.Equal(expected, SerialCodeFormatter.GetNextOLesSuffix(current));
        }

        [Theory]
        [InlineData("Z")]
        [InlineData("1")]
        [InlineData("あ")]
        public void GetNextOLesSuffix_Zまたは範囲外の文字_Aに戻る(string current) {
            Assert.Equal("A", SerialCodeFormatter.GetNextOLesSuffix(current));
        }
    }
}
