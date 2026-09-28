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

export interface Substrate {
  id: number
  categoryName: string | null
  productName: string | null
  substrateName: string | null
  substrateModel: string | null
  orderNumber: string | null
  substrateNumber: string | null
  increase: number | null
  decrease: number | null
  defect: number | null
  personInfo: string | null
  regDate: string | null
  comment: string | null
  createdAt: string | null
}

export interface AuditLog {
  id: number
  action: string
  targetType: 'PRODUCT' | 'SUBSTRATE'
  targetId: number
  targetName: string | null
  detail: string | null
  createdAt: string
}
