using ProductDatabase.Common;

namespace ProductDataBase.Tests {
    public class CommonUtilsTests {

        [Theory]
        [InlineData(1, "01")]
        [InlineData(9, "09")]
        [InlineData(10, "X")]
        [InlineData(11, "Y")]
        [InlineData(12, "Z")]
        public void ToMonthCode_月に応じたコードを返す(int month, string expected) {
            var date = new DateTime(2026, month, 1);

            Assert.Equal(expected, CommonUtils.ToMonthCode(date));
        }
    }
}
