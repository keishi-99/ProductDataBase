import type { AuditLog, Product, Substrate } from './types'

const API_BASE = ''

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

export interface ProductEditFields {
  orderNumber: string
  productNumber: string
  olesNumber: string
  comment: string
}

export async function updateProduct(id: number, data: ProductEditFields): Promise<boolean> {
  const res = await fetch(`${API_BASE}/api/products/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    credentials: 'include',
    body: JSON.stringify(data),
  })
  return res.ok
}

export async function deleteProduct(id: number): Promise<boolean> {
  const res = await fetch(`${API_BASE}/api/products/${id}`, {
    method: 'DELETE',
    credentials: 'include',
  })
  return res.ok
}

export async function fetchAuditLogs(): Promise<AuditLog[]> {
  const res = await fetch(`${API_BASE}/api/audit-logs`, { credentials: 'include' })
  if (!res.ok) throw new Error('操作ログの取得に失敗しました。')
  return res.json()
}

export async function fetchSubstrates(): Promise<Substrate[]> {
  const res = await fetch(`${API_BASE}/api/substrates`, { credentials: 'include' })
  if (!res.ok) throw new Error('基板一覧の取得に失敗しました。')
  return res.json()
}

export async function deleteSubstrate(id: number): Promise<boolean> {
  const res = await fetch(`${API_BASE}/api/substrates/${id}`, {
    method: 'DELETE',
    credentials: 'include',
  })
  return res.ok
}
