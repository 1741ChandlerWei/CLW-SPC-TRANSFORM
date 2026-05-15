import { NextRequest, NextResponse } from 'next/server'
import { parseSpecFile } from '../../lib/parser'
import { buildCompareExcel } from '../../lib/builder'

export const runtime = 'nodejs'

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const fileA = formData.get('fileA') as File | null
    const fileB = formData.get('fileB') as File | null

    if (!fileA || !fileB) {
      return new NextResponse(JSON.stringify({ error: '需要上傳兩個檔案' }), { status: 400 })
    }

    const bufA = Buffer.from(await fileA.arrayBuffer())
    const bufB = Buffer.from(await fileB.arrayBuffer())

    const specA = parseSpecFile(bufA)
    specA.fileName = fileA.name

    const specB = parseSpecFile(bufB)
    specB.fileName = fileB.name

    const outputBuffer = await buildCompareExcel(specA, specB, fileA.name, fileB.name)
    const outName = `比對_${fileA.name.replace(/\.[^.]+$/, '')}_vs_${fileB.name.replace(/\.[^.]+$/, '')}.xlsx`

    return new NextResponse(outputBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(outName)}`,
      },
    })
  } catch (e) {
    console.error(e)
    return new NextResponse(JSON.stringify({ error: '比對失敗' }), { status: 500 })
  }
}
