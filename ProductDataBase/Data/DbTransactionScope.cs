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

        // トランザクションをコミットする
        public void Commit() {
            _transaction?.Commit();
            _committed = true;
        }

        // トランザクションをロールバックする（未コミットの場合のみ）
        public void Rollback() {
            if (!_committed) {
                _transaction?.Rollback();
                _committed = true;
            }
        }

        public void Dispose() {
            if (_disposed) { return; }
            _transaction?.Dispose();
            _connection?.Close();
            _connection?.Dispose();
            _disposed = true;
        }
    }
}
