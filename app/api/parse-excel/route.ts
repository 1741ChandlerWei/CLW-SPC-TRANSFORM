import { NextRequest, NextResponse } from 'next/server'
import { parseSpecFile } from '../../lib/parser'
import { buildOutputExcel } from '../../lib/builder'

export const runtime = 'nodejs'

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const file = formData.get('file') as File | null

    if (!file) {
      return new NextResponse(JSON.stringify({ error: '未收到檔案' }), { status: 400 })
    }

    const buffer = Buffer.from(await file.arrayBuffer())
    const spec = parseSpecFile(buffer)
    spec.fileName = file.name

    if (spec.models.length === 0) {
      return new NextResponse(
        JSON.stringify({ error: '無法從此檔案中找到型號資料，請確認是否為正確的規格表格式' }),
        { status: 422 }
      )
    }

    const outputBuffer = await buildOutputExcel(spec)

    const baseName = file.name.replace(/\.[^.]+$/, '')
    const outName = `整理_${baseName}.xlsx`

    return new NextResponse(outputBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(outName)}`,
      },
    })
  } catch (e) {
    console.error(e)
    return new NextResponse(JSON.stringify({ error: '解析失敗，請確認檔案格式' }), { status: 500 })
  }
}
