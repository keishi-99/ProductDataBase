using System.Runtime.CompilerServices;

namespace ProductDatabase.Common {

    // catchした例外をMessageBoxでユーザーに表示するための共通ヘルパー
    internal static class ErrorMessageHelper {

        // memberNameは[CallerMemberName]により呼び出し元のメソッド名が自動的に渡される
        public static void Show(Exception ex, [CallerMemberName] string memberName = "") {
            MessageBox.Show(SqliteBusyErrorHelper.GetUserMessage(ex), $"[{memberName}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
