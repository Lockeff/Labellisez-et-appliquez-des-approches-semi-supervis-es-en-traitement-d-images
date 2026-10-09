$ErrorActionPreference='Stop'
$workspace=(Get-Location).Path
$source=Join-Path $workspace 'output/presentations/BrainScanAI_presentation_V1.pptx'
$renderDir=Join-Path $workspace '.build/brainscanai/audit-v1/ppt-renders'
[void][IO.Directory]::CreateDirectory($renderDir)
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,-1,0,0)
try {
 for($i=1;$i -le $deck.Slides.Count;$i++){
  $deck.Slides.Item($i).Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)
 }
 Write-Output ('Rendered '+$deck.Slides.Count+' slides without saving.')
} finally {$deck.Close()}
