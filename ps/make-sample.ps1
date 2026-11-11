$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force 'D:\tmp\control-test' | Out-Null

try {
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $d = $w.Documents.Add()
  $d.Content.Text = "Control module test document.`r`nline two: verify docx to pdf. date 2026-11-11."
  $d.SaveAs('D:\tmp\control-test\sample.docx')
  $d.Close($false)
  $w.Quit()
  Write-Output 'WORD OK'
} catch {
  Write-Output ('WORD ERR ' + $_.Exception.Message)
}

try {
  $e = New-Object -ComObject Excel.Application
  $e.Visible = $false
  $wb = $e.Workbooks.Add()
  $sh = $wb.Sheets.Item(1)
  $sh.Range('A1').Value2 = 'name'
  $sh.Range('B1').Value2 = 'count'
  $sh.Range('A2').Value2 = 'yuzai'
  $sh.Range('B2').Value2 = 42
  $wb.SaveAs('D:\tmp\control-test\sample.xlsx')
  $wb.Close($false)
  $e.Quit()
  Write-Output 'EXCEL OK'
} catch {
  Write-Output ('EXCEL ERR ' + $_.Exception.Message)
}
