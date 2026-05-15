export const metadata = { title: 'CLW 規格表轉換工具', description: 'CLW-SPC-TRANSFORM' }

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-TW">
      <body style={{ margin: 0 }}>{children}</body>
    </html>
  )
}
