$ErrorActionPreference = 'Stop'
try {
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $d = $w.Documents.Add()
  $d.Content.Text = 'hello'
  $d.SaveAs('D:\tmp\com-test.docx')
  $d.Close($false)
  $w.Quit()
  Write-Output 'OK'
} catch {
  Write-Output ('ERR ' + $_.Exception.GetType().Name + ': ' + $_.Exception.Message)
  exit 1
}
