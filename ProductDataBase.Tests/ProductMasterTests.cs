using ProductDatabase.Models;

namespace ProductDataBase.Tests {
    public class ProductMasterTests {

        [Theory]
        [InlineData(3, 1, 999, 3)]
        [InlineData(4, 1, 9999, 4)]
        [InlineData(101, 1, 899, 3)]
        [InlineData(102, 901, 999, 3)]
        public void GetSerialRange_種別ごとの範囲と桁数を返す(int serialDigitType, int min, int max, int digit) {
            var master = new ProductMaster { SerialDigitType = serialDigitType };

            var range = master.GetSerialRange();

            Assert.Equal((min, max, digit), range);
        }

        [Fact]
        public void GetSerialRange_不明な種別_例外になる() {
            var master = new ProductMaster { SerialDigitType = 999 };

            Assert.Throws<InvalidOperationException>(() => master.GetSerialRange());
        }

        [Theory]
        [InlineData(3, 3)]
        [InlineData(101, 3)]
        [InlineData(102, 3)]
        [InlineData(4, 4)]
        [InlineData(0, 0)]
        public void SerialDigit_種別に応じた桁数を返す(int serialDigitType, int expected) {
            var master = new ProductMaster { SerialDigitType = serialDigitType };

            Assert.Equal(expected, master.SerialDigit);
        }
    }
}
