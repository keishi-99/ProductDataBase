import { useEffect, useState } from 'react'
import { deleteProduct, fetchAuditLogs, fetchCategories, fetchProducts, login, logout, updateProduct } from './api'
import type { ProductEditFields } from './api'
import type { AuditLog, Product } from './types'
import './App.css'

const emptyEditForm: ProductEditFields = { orderNumber: '', productNumber: '', olesNumber: '', comment: '' }

function App() {
  const [products, setProducts] = useState<Product[]>([])
  const [categories, setCategories] = useState<string[]>([])
  const [category, setCategory] = useState('')
  const [keyword, setKeyword] = useState('')
  const [errorMessage, setErrorMessage] = useState<string | null>(null)

  const [isLoggedIn, setIsLoggedIn] = useState(false)
  const [password, setPassword] = useState('')
  const [loginError, setLoginError] = useState<string | null>(null)

  const [editingId, setEditingId] = useState<number | null>(null)
  const [editForm, setEditForm] = useState<ProductEditFields>(emptyEditForm)

  const [showAuditLogs, setShowAuditLogs] = useState(false)
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([])

  const loadProducts = async (currentCategory: string, currentKeyword: string) => {
    try {
      const data = await fetchProducts(currentCategory, currentKeyword)
      setProducts(data)
      setErrorMessage(null)
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
  }

  useEffect(() => {
    fetchCategories().then(setCategories).catch(() => {})
    loadProducts('', '')
  }, [])

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault()
    loadProducts(category, keyword)
  }

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault()
    const success = await login(password)
    if (success) {
      setIsLoggedIn(true)
      setLoginError(null)
      setPassword('')
    } else {
      setLoginError('パスワードが正しくありません。')
    }
  }

  const handleLogout = async () => {
    await logout()
    setIsLoggedIn(false)
  }

  const startEdit = (p: Product) => {
    setEditingId(p.id)
    setEditForm({
      orderNumber: p.orderNumber ?? '',
      productNumber: p.productNumber ?? '',
      olesNumber: p.olesNumber ?? '',
      comment: p.comment ?? '',
    })
  }

  const cancelEdit = () => {
    setEditingId(null)
    setEditForm(emptyEditForm)
  }

  const saveEdit = async (id: number) => {
    const success = await updateProduct(id, editForm)
    if (success) {
      cancelEdit()
      loadProducts(category, keyword)
    } else {
      setErrorMessage('更新に失敗しました。')
    }
  }

  const handleDelete = async (p: Product) => {
    if (!window.confirm(`「${p.productName}」を削除しますか？`)) return
    const success = await deleteProduct(p.id)
    if (success) {
      loadProducts(category, keyword)
    } else {
      setErrorMessage('削除に失敗しました。')
    }
  }

  const toggleAuditLogs = async () => {
    if (!showAuditLogs) {
      try {
        setAuditLogs(await fetchAuditLogs())
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err))
        return
      }
    }
    setShowAuditLogs(!showAuditLogs)
  }

  return (
    <div className="page">
      <header className="header">
        <h1>ProductWebViewer (Spring)</h1>
        {isLoggedIn ? (
          <button type="button" onClick={handleLogout}>ログアウト</button>
        ) : (
          <form className="login-form" onSubmit={handleLogin}>
            <input
              type="password"
              placeholder="管理者パスワード"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
            <button type="submit">ログイン</button>
            {loginError && <span className="error">{loginError}</span>}
          </form>
        )}
      </header>

      <form className="search-form" onSubmit={handleSearch}>
        <select value={category} onChange={(e) => setCategory(e.target.value)}>
          <option value="">すべてのカテゴリ</option>
          {categories.map((c) => (
            <option key={c} value={c}>{c}</option>
          ))}
        </select>
        <input
          type="text"
          placeholder="製品名・注文番号・製造番号で検索"
          value={keyword}
          onChange={(e) => setKeyword(e.target.value)}
        />
        <button type="submit">検索</button>
      </form>

      {errorMessage && <p className="error">{errorMessage}</p>}

      <table className="product-table">
        <thead>
          <tr>
            <th>ID</th>
            <th>カテゴリ</th>
            <th>製品名</th>
            <th>型式</th>
            <th>注文番号</th>
            <th>製造番号</th>
            <th>OLES番号</th>
            <th>数量</th>
            <th>コメント</th>
            {isLoggedIn && <th>操作</th>}
          </tr>
        </thead>
        <tbody>
          {products.map((p) =>
            editingId === p.id ? (
              <tr key={p.id}>
                <td>{p.id}</td>
                <td>{p.categoryName}</td>
                <td>{p.productName}</td>
                <td>{p.productModel}</td>
                <td>
                  <input
                    value={editForm.orderNumber}
                    onChange={(e) => setEditForm({ ...editForm, orderNumber: e.target.value })}
                  />
                </td>
                <td>
                  <input
                    value={editForm.productNumber}
                    onChange={(e) => setEditForm({ ...editForm, productNumber: e.target.value })}
                  />
                </td>
                <td>
                  <input
                    value={editForm.olesNumber}
                    onChange={(e) => setEditForm({ ...editForm, olesNumber: e.target.value })}
                  />
                </td>
                <td>{p.quantity}</td>
                <td>
                  <input
                    value={editForm.comment}
                    onChange={(e) => setEditForm({ ...editForm, comment: e.target.value })}
                  />
                </td>
                <td>
                  <button type="button" onClick={() => saveEdit(p.id)}>保存</button>
                  <button type="button" onClick={cancelEdit}>キャンセル</button>
                </td>
              </tr>
            ) : (
              <tr key={p.id}>
                <td>{p.id}</td>
                <td>{p.categoryName}</td>
                <td>{p.productName}</td>
                <td>{p.productModel}</td>
                <td>{p.orderNumber}</td>
                <td>{p.productNumber}</td>
                <td>{p.olesNumber}</td>
                <td>{p.quantity}</td>
                <td>{p.comment}</td>
                {isLoggedIn && (
                  <td>
                    <button type="button" onClick={() => startEdit(p)}>編集</button>
                    <button type="button" onClick={() => handleDelete(p)}>削除</button>
                  </td>
                )}
              </tr>
            )
          )}
        </tbody>
      </table>

      <button type="button" className="audit-log-toggle" onClick={toggleAuditLogs}>
        {showAuditLogs ? '操作ログを隠す' : '操作ログを表示'}
      </button>

      {showAuditLogs && (
        <table className="product-table">
          <thead>
            <tr>
              <th>日時</th>
              <th>操作</th>
              <th>製品名</th>
              <th>詳細</th>
            </tr>
          </thead>
          <tbody>
            {auditLogs.map((log) => (
              <tr key={log.id}>
                <td>{log.createdAt}</td>
                <td>{log.action === 'DELETE' ? '削除' : '編集'}</td>
                <td>{log.productName}</td>
                <td>{log.detail}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  )
}

export default App
