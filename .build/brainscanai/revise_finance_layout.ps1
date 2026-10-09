param([string]$Workspace=(Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/aws-clear'
$newDraft=Join-Path $build 'candidate-v2.pptx'
if(Test-Path -LiteralPath $newDraft){throw 'New draft already exists.'}
$helperSource=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$helperStart=$helperSource.IndexOf('function RGB(')
$helperEnd=$helperSource.IndexOf('function Picture(')
Invoke-Expression $helperSource.Substring($helperStart,$helperEnd-$helperStart)
$revision=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/rewrite_finance.ps1') -Raw -Encoding UTF8
$functionsStart=$revision.IndexOf('function ClearContent(')
$functionsEnd=$revision.IndexOf('try {')
Invoke-Expression $revision.Substring($functionsStart,$functionsEnd-$functionsStart)
$bodyStart=$revision.IndexOf('    $s=$deck.Slides.Item(11)')
$bodyEnd=$revision.IndexOf('    if($deck.Slides.Count -ne 13)')
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open((Join-Path $build 'candidate.pptx'),0,0,0)
try{
    Invoke-Expression $revision.Substring($bodyStart,$bodyEnd-$bodyStart)
    $deck.SaveAs($newDraft,24)
    $renderDir=Join-Path $build 'candidate-v2-renders'
    [void][IO.Directory]::CreateDirectory($renderDir)
    for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
}finally{$deck.Close()}
