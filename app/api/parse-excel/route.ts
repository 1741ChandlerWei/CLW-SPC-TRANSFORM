import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'

export async function POST(req: NextRequest) {
  const formData = await req.formData()
  const file = formData.get('file')

  if (!file || !(file instanceof File)) {
    return new NextResponse('No file uploaded', { status: 400 })
  }

  const buffer = Buffer.from(await file.arrayBuffer())
  const workbook = XLSX.read(buffer, { type: 'buffer' })

  // 先簡單回傳成功，下一步再加解析邏輯
  const result = {
    sheetCount: workbook.SheetNames.length,
    sheets: workbook.SheetNames,
  }

  return NextResponse.json(result)
}
