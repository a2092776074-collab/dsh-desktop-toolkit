$ErrorActionPreference = 'Stop'
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $pres = $ppt.Presentations.Add()
  $slide = $pres.Slides.Add(1, 1)
  $slide.Shapes.Title.TextFrame.TextRange.Text = "Control module test"
  $slide.Shapes.AddTextbox(1, 80, 140, 480, 120).TextFrame.TextRange.Text = "pptx -> pdf via LibreOffice"
  $pres.SaveAs('D:\tmp\control-test\sample.pptx')
  $pres.Close()
  $ppt.Quit()
  Write-Output 'PPT OK'
} catch {
  Write-Output ('PPT ERR ' + $_.Exception.Message)
}
