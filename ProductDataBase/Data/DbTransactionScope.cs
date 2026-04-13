using Microsoft.Data.Sqlite;

namespace ProductDatabase.Data {
    // SQLite接続とトランザクションのライフサイクルを管理するクラス
    internal sealed class DbTransactionScope : IDisposable {
        private SqliteTransaction? _transaction;
        private SqliteConnection? _connection;
        private bool _committed;
        private bool _disposed;

        public bool IsCommitted => _committed;

        public SqliteConnection Connection => _connection
            ?? throw new InvalidOperationException("接続が初期化されていません。");

        public SqliteTransaction Transaction => _transaction
            ?? throw new InvalidOperationException("トランザクションが初期化されていません。");

        // 接続を開きトランザクションを開始する
        public void Begin() {
            _connection = new SqliteConnection(ProductRepository.GetConnectionRegistration());
            _connection.Open();
            _transaction = _connection.BeginTransaction();
            _committed = false;
        }

        // トランザクションをコミットする（既にコミット/ロールバック済みの場合は何もしない）
        public void Commit() {
            if (_committed) return;
            _transaction?.Commit();
            _committed = true;
        }

        // トランザクションをロールバックする（未コミットの場合のみ）
        public void Rollback() {
            if (_committed) return;
            _transaction?.Rollback();
            _committed = true;
        }

        public void Dispose() {
            if (_disposed) { return; }
            // usingブロックを抜けた際にコミットされていない場合は自動ロールバック
            if (!_committed) {
                _transaction?.Rollback();
            }
            _transaction?.Dispose();
            _connection?.Close();
            _connection?.Dispose();
            _disposed = true;
        }
    }
}
