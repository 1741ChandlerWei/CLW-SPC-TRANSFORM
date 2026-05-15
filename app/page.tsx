export default function Page() {
  return (
    <main style={{ padding: '40px', fontFamily: 'sans-serif' }}>
      <h1>CLW 規格表轉換工具</h1>

      <form
        action="/api/parse-excel"
        method="post"
        encType="multipart/form-data"
      >
        <input
          type="file"
          name="file"
          accept=".xls,.xlsx,.xlsm"
          required
        />
        <br />
        <br />
        <button type="submit">上傳並產出整理後 Excel</button>
      </form>
    </main>
  )
}
