export interface Product {
  id: number
  categoryName: string | null
  productName: string | null
  productModel: string | null
  productType: string | null
  orderNumber: string | null
  productNumber: string | null
  olesNumber: string | null
  quantity: number | null
  regDate: string | null
  comment: string | null
  createdAt: string | null
}

export interface AuditLog {
  id: number
  action: string
  productId: number
  productName: string | null
  detail: string | null
  createdAt: string
}
