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

        // ===== RegType のフラグ展開 =====

        [Theory]
        [InlineData(0, false)]
        [InlineData(1, true)]
        [InlineData(2, true)]
        [InlineData(3, true)]
        [InlineData(4, false)]
        [InlineData(8, false)]
        [InlineData(9, true)]
        public void RegType_シリアル生成フラグが立つのは1_2_3_9だけ(int regType, bool expected) {
            var master = new ProductMaster { RegType = regType };

            Assert.Equal(expected, master.IsSerialGeneration);
        }

        [Theory]
        [InlineData(8, false)]
        [InlineData(9, true)]
        public void RegType_9のときだけIsRegType9が立つ(int regType, bool expected) {
            var master = new ProductMaster { RegType = regType };

            Assert.Equal(expected, master.IsRegType9);
        }

        // ===== SerialPrintType のフラグ展開 =====

        // 値: Label=1 / Barcode=2 / Nameplate=4 / Underline=8 / Last4Digits=16 / OLesSerial=32
        [Theory]
        [InlineData(0, false, false, false, false, false, false)]
        [InlineData(1, true, false, false, false, false, false)]
        [InlineData(2, false, true, false, false, false, false)]
        [InlineData(4, false, false, true, false, false, false)]
        [InlineData(8, false, false, false, true, false, false)]
        [InlineData(16, false, false, false, false, true, false)]
        [InlineData(32, false, false, false, false, false, true)]
        [InlineData(3, true, true, false, false, false, false)]
        [InlineData(63, true, true, true, true, true, true)]
        public void SerialPrintType_ビットに対応したフラグが立つ(
            int serialPrintType, bool label, bool barcode, bool nameplate, bool underline, bool last4Digits, bool oLesLabel) {
            var master = new ProductMaster { SerialPrintType = serialPrintType };

            Assert.Equal(label, master.IsLabelPrint);
            Assert.Equal(barcode, master.IsBarcodePrint);
            Assert.Equal(nameplate, master.IsNameplatePrint);
            Assert.Equal(underline, master.IsUnderlinePrint);
            Assert.Equal(last4Digits, master.IsLast4Digits);
            Assert.Equal(oLesLabel, master.IsOLesLabelPrint);
        }

        // ===== SheetPrintType のフラグ展開 =====

        [Theory]
        [InlineData(0, false, false)]
        [InlineData(1, true, false)]
        [InlineData(2, false, true)]
        [InlineData(3, true, true)]
        public void SheetPrintType_ビットに対応したフラグが立つ(int sheetPrintType, bool checkSheet, bool list) {
            var master = new ProductMaster { SheetPrintType = sheetPrintType };

            Assert.Equal(checkSheet, master.IsCheckSheetPrint);
            Assert.Equal(list, master.IsListPrint);
        }
    }
}
