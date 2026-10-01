import { useEffect, useState } from 'react'
import {
  deleteProduct, deleteSubstrate, fetchAuditLogs, fetchCategories, fetchProductNames, fetchProducts, fetchProductTypes,
  fetchSubstrateCategories, fetchSubstrateNames, fetchSubstrateProductNames, fetchSubstrates, login, logout, updateProduct,
} from './api'
import type { ProductEditFields } from './api'
import type { AuditLog, Product, Substrate } from './types'
import './App.css'

const emptyEditForm: ProductEditFields = { orderNumber: '', productNumber: '', olesNumber: '', comment: '' }

function App() {
  const [products, setProducts] = useState<Product[]>([])
  const [categories, setCategories] = useState<string[]>([])
  const [category, setCategory] = useState('')
  const [productNames, setProductNames] = useState<string[]>([])
  const [productNameFilter, setProductNameFilter] = useState('')
  const [productTypes, setProductTypes] = useState<string[]>([])
  const [productTypeFilter, setProductTypeFilter] = useState('')
  const [keyword, setKeyword] = useState('')
  const [errorMessage, setErrorMessage] = useState<string | null>(null)

  const [isLoggedIn, setIsLoggedIn] = useState(false)
  const [password, setPassword] = useState('')
  const [loginError, setLoginError] = useState<string | null>(null)

  const [editingId, setEditingId] = useState<number | null>(null)
  const [editForm, setEditForm] = useState<ProductEditFields>(emptyEditForm)

  const [activeTab, setActiveTab] = useState<'product' | 'substrate' | 'auditLog'>('product')
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([])

  const [substrates, setSubstrates] = useState<Substrate[]>([])
  const [subCategories, setSubCategories] = useState<string[]>([])
  const [subCategory, setSubCategory] = useState('')
  const [subProductNames, setSubProductNames] = useState<string[]>([])
  const [subProductNameFilter, setSubProductNameFilter] = useState('')
  const [substrateNames, setSubstrateNames] = useState<string[]>([])
  const [substrateNameFilter, setSubstrateNameFilter] = useState('')

  const loadProducts = async (currentCategory: string, currentProductName: string, currentProductType: string, currentKeyword: string) => {
    try {
      const data = await fetchProducts(currentCategory, currentProductName, currentProductType, currentKeyword)
      setProducts(data)
      setErrorMessage(null)
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
  }

  const loadSubstrates = async (currentCategory: string, currentProductName: string, currentSubstrateName: string) => {
    try {
      setSubstrates(await fetchSubstrates(currentCategory, currentProductName, currentSubstrateName))
      setErrorMessage(null)
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
  }

  useEffect(() => {
    fetchCategories().then(setCategories).catch(() => {})
    fetchProductNames('').then(setProductNames).catch(() => {})
    fetchProductTypes('', '').then(setProductTypes).catch(() => {})
    loadProducts('', '', '', '')

    fetchSubstrateCategories().then(setSubCategories).catch(() => {})
    fetchSubstrateProductNames('').then(setSubProductNames).catch(() => {})
    fetchSubstrateNames('', '').then(setSubstrateNames).catch(() => {})
    loadSubstrates('', '', '')
  }, [])

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault()
    loadProducts(category, productNameFilter, productTypeFilter, keyword)
  }

  // カスケードリストボックス: カテゴリを選ぶと、製品名・種別の候補を絞り込んで即座に検索する
  const handleCategoryChange = async (value: string) => {
    setCategory(value)
    setProductNameFilter('')
    setProductTypeFilter('')
    try {
      setProductNames(await fetchProductNames(value))
      setProductTypes(await fetchProductTypes(value, ''))
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
    loadProducts(value, '', '', keyword)
  }

  // カスケードリストボックス: 製品名を選ぶと、種別の候補を絞り込んで即座に検索する
  const handleProductNameChange = async (value: string) => {
    setProductNameFilter(value)
    setProductTypeFilter('')
    try {
      setProductTypes(await fetchProductTypes(category, value))
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
    loadProducts(category, value, '', keyword)
  }

  // カスケードリストボックス: 種別を選ぶと即座に検索する
  const handleProductTypeChange = (value: string) => {
    setProductTypeFilter(value)
    loadProducts(category, productNameFilter, value, keyword)
  }

  // 基板タブ用カスケードリストボックス: カテゴリを選ぶと、製品名・基板名の候補を絞り込んで即座に検索する
  const handleSubCategoryChange = async (value: string) => {
    setSubCategory(value)
    setSubProductNameFilter('')
    setSubstrateNameFilter('')
    try {
      setSubProductNames(await fetchSubstrateProductNames(value))
      setSubstrateNames(await fetchSubstrateNames(value, ''))
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
    loadSubstrates(value, '', '')
  }

  // 基板タブ用カスケードリストボックス: 製品名を選ぶと、基板名の候補を絞り込んで即座に検索する
  const handleSubProductNameChange = async (value: string) => {
    setSubProductNameFilter(value)
    setSubstrateNameFilter('')
    try {
      setSubstrateNames(await fetchSubstrateNames(subCategory, value))
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : String(err))
    }
    loadSubstrates(subCategory, value, '')
  }

  // 基板タブ用カスケードリストボックス: 基板名を選ぶと即座に検索する
  const handleSubstrateNameChange = (value: string) => {
    setSubstrateNameFilter(value)
    loadSubstrates(subCategory, subProductNameFilter, value)
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
      loadProducts(category, productNameFilter, productTypeFilter, keyword)
    } else {
      setErrorMessage('更新に失敗しました。')
    }
  }

  const handleDelete = async (p: Product) => {
    if (!window.confirm(`「${p.productName}」を削除しますか？`)) return
    const success = await deleteProduct(p.id)
    if (success) {
      loadProducts(category, productNameFilter, productTypeFilter, keyword)
    } else {
      setErrorMessage('削除に失敗しました。')
    }
  }

  const handleDeleteSubstrate = async (s: Substrate) => {
    if (!window.confirm(`「${s.substrateName}」を削除しますか？`)) return
    const success = await deleteSubstrate(s.id)
    if (success) {
      loadSubstrates(subCategory, subProductNameFilter, substrateNameFilter)
    } else {
      setErrorMessage('削除に失敗しました。')
    }
  }

  const switchTab = async (tab: 'product' | 'substrate' | 'auditLog') => {
    // タブを開くたびに最新化する（削除操作の直後でも古いログが見えないように）
    if (tab === 'auditLog') {
      try {
        setAuditLogs(await fetchAuditLogs())
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : String(err))
        return
      }
    }
    setActiveTab(tab)
  }

  return (
    <div className="page">
      <header className="header">
        <h1>ProductWebViewer (Spring)</h1>
        {isLoggedIn ? (
          <button type="button" className="btn btn-secondary" onClick={handleLogout}>ログアウト</button>
        ) : (
          <form className="login-form" onSubmit={handleLogin}>
            <input
              type="password"
              placeholder="管理者パスワード"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
            <button type="submit" className="btn btn-primary">ログイン</button>
            {loginError && <span className="error">{loginError}</span>}
          </form>
        )}
      </header>

      {errorMessage && <p className="error">{errorMessage}</p>}

      <ul className="nav nav-tabs">
        <li className="nav-item">
          <a className={`nav-link ${activeTab === 'product' ? 'active' : ''}`}
             href="#product" onClick={(e) => { e.preventDefault(); switchTab('product') }}>製品一覧</a>
        </li>
        <li className="nav-item">
          <a className={`nav-link ${activeTab === 'substrate' ? 'active' : ''}`}
             href="#substrate" onClick={(e) => { e.preventDefault(); switchTab('substrate') }}>基板一覧</a>
        </li>
        <li className="nav-item">
          <a className={`nav-link ${activeTab === 'auditLog' ? 'active' : ''}`}
             href="#auditLog" onClick={(e) => { e.preventDefault(); switchTab('auditLog') }}>操作ログ</a>
        </li>
      </ul>

      {activeTab === 'product' && (
        <form className="search-form-listbox" onSubmit={handleSearch}>
          <div className="listbox-group">
            <div className="listbox-label">カテゴリ</div>
            <div className="list-box">
              <select size={10} value={category} onChange={(e) => handleCategoryChange(e.target.value)}>
                <option value="">（すべて）</option>
                {categories.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="listbox-group">
            <div className="listbox-label">製品名</div>
            <div className="list-box">
              <select size={10} value={productNameFilter} onChange={(e) => handleProductNameChange(e.target.value)}>
                <option value="">（すべて）</option>
                {productNames.map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="listbox-group">
            <div className="listbox-label">種別</div>
            <div className="list-box">
              <select size={10} value={productTypeFilter} onChange={(e) => handleProductTypeChange(e.target.value)}>
                <option value="">（すべて）</option>
                {productTypes.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="search-form">
            <input
              type="text"
              placeholder="製品名・注文番号・製造番号で検索"
              value={keyword}
              onChange={(e) => setKeyword(e.target.value)}
            />
            <button type="submit" className="btn btn-primary">検索</button>
          </div>
        </form>
      )}

      {activeTab === 'substrate' && (
        <div className="search-form-listbox">
          <div className="listbox-group">
            <div className="listbox-label">カテゴリ</div>
            <div className="list-box">
              <select size={10} value={subCategory} onChange={(e) => handleSubCategoryChange(e.target.value)}>
                <option value="">（すべて）</option>
                {subCategories.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="listbox-group">
            <div className="listbox-label">製品名</div>
            <div className="list-box">
              <select size={10} value={subProductNameFilter} onChange={(e) => handleSubProductNameChange(e.target.value)}>
                <option value="">（すべて）</option>
                {subProductNames.map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </div>
          </div>
          <div className="listbox-group">
            <div className="listbox-label">基板名</div>
            <div className="list-box">
              <select size={10} value={substrateNameFilter} onChange={(e) => handleSubstrateNameChange(e.target.value)}>
                <option value="">（すべて）</option>
                {substrateNames.map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
      )}

      {activeTab === 'product' && (
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
                    <button type="button" className="btn btn-primary" onClick={() => saveEdit(p.id)}>保存</button>
                    <button type="button" className="btn btn-secondary" onClick={cancelEdit}>キャンセル</button>
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
                      <button type="button" className="btn btn-outline-warning" onClick={() => startEdit(p)}>編集</button>
                      <button type="button" className="btn btn-outline-danger" onClick={() => handleDelete(p)}>削除</button>
                    </td>
                  )}
                </tr>
              )
            )}
          </tbody>
        </table>
      )}

      {activeTab === 'substrate' && (
        <table className="product-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>カテゴリ</th>
              <th>製品名</th>
              <th>基板名</th>
              <th>型式</th>
              <th>注文番号</th>
              <th>製造番号</th>
              <th>入庫/出庫/不良</th>
              <th>コメント</th>
              {isLoggedIn && <th>操作</th>}
            </tr>
          </thead>
          <tbody>
            {substrates.map((s) => (
              <tr key={s.id}>
                <td>{s.id}</td>
                <td>{s.categoryName}</td>
                <td>{s.productName}</td>
                <td>{s.substrateName}</td>
                <td>{s.substrateModel}</td>
                <td>{s.orderNumber}</td>
                <td>{s.substrateNumber}</td>
                <td>{s.increase} / {s.decrease} / {s.defect}</td>
                <td>{s.comment}</td>
                {isLoggedIn && (
                  <td>
                    <button type="button" className="btn btn-outline-danger" onClick={() => handleDeleteSubstrate(s)}>削除</button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {activeTab === 'auditLog' && (
        <table className="product-table">
          <thead>
            <tr>
              <th>日時</th>
              <th>種別</th>
              <th>操作</th>
              <th>対象名</th>
              <th>詳細</th>
            </tr>
          </thead>
          <tbody>
            {auditLogs.map((log) => (
              <tr key={log.id}>
                <td>{log.createdAt}</td>
                <td>{log.targetType === 'SUBSTRATE' ? '基板' : '製品'}</td>
                <td>{log.action === 'DELETE' ? '削除' : '編集'}</td>
                <td>{log.targetName}</td>
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
