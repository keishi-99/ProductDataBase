import type { Product } from './types'

const API_BASE = 'http://localhost:8080'

export async function fetchProducts(category: string, keyword: string): Promise<Product[]> {
  const params = new URLSearchParams()
  if (category) params.set('category', category)
  if (keyword) params.set('keyword', keyword)

  const res = await fetch(`${API_BASE}/api/products?${params}`, { credentials: 'include' })
  if (!res.ok) throw new Error('製品一覧の取得に失敗しました。')
  return res.json()
}

export async function fetchCategories(): Promise<string[]> {
  const res = await fetch(`${API_BASE}/api/products/categories`, { credentials: 'include' })
  if (!res.ok) throw new Error('カテゴリ一覧の取得に失敗しました。')
  return res.json()
}

export async function login(password: string): Promise<boolean> {
  const res = await fetch(`${API_BASE}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    credentials: 'include',
    body: JSON.stringify({ password }),
  })
  return res.ok
}

export async function logout(): Promise<void> {
  await fetch(`${API_BASE}/api/auth/logout`, { method: 'POST', credentials: 'include' })
}
