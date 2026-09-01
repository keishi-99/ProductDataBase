import { useEffect, useState } from 'react'
import { fetchCategories, fetchProducts, login, logout } from './api'
import type { Product } from './types'
import './App.css'

function App() {
  const [products, setProducts] = useState<Product[]>([])
  const [categories, setCategories] = useState<string[]>([])
  const [category, setCategory] = useState('')
  const [keyword, setKeyword] = useState('')
  const [errorMessage, setErrorMessage] = useState<string | null>(null)

  const [isLoggedIn, setIsLoggedIn] = useState(false)
  const [password, setPassword] = useState('')
  const [loginError, setLoginError] = useState<string | null>(null)

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
          </tr>
        </thead>
        <tbody>
          {products.map((p) => (
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
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export default App
