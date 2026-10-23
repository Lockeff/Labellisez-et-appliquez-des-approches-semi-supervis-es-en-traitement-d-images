$ErrorActionPreference = 'Stop'

$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$htmlPath = (Resolve-Path (Join-Path $PSScriptRoot 'guide.html')).Path
$docxPath = Join-Path $workspace 'Ronan_Menguy_Guide_soutenance_BrainScanAI.docx'
$pdfPath = Join-Path $PSScriptRoot 'Ronan_Menguy_Guide_soutenance_BrainScanAI.pdf'

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Write-Output 'STEP=open-html'
    $document = $word.Documents.Open($htmlPath, $false, $true)
    Write-Output 'STEP=format-document'

    foreach ($section in $document.Sections) {
        $section.PageSetup.PaperSize = 7
        $section.PageSetup.TopMargin = $word.CentimetersToPoints(1.7)
        $section.PageSetup.BottomMargin = $word.CentimetersToPoints(1.7)
        $section.PageSetup.LeftMargin = $word.CentimetersToPoints(1.8)
        $section.PageSetup.RightMargin = $word.CentimetersToPoints(1.8)

        $header = $section.Headers.Item(1).Range
        $header.Text = 'BrainScanAI  |  Guide de soutenance'
        $header.Font.Name = 'Aptos'
        $header.Font.Size = 8
        $header.Font.Color = 6974058
        $header.ParagraphFormat.Alignment = 2

        $footer = $section.Footers.Item(1).Range
        $footer.Text = 'Ronan Menguy  |  Page '
        $footer.Font.Name = 'Aptos'
        $footer.Font.Size = 8
        $footer.Font.Color = 6974058
        $footer.ParagraphFormat.Alignment = 1
        $footer.Collapse(0)
        $null = $section.Footers.Item(1).Range.Fields.Add($footer, 33)
    }

    try {
        $document.BuiltInDocumentProperties.Item('Title').Value = 'Guide oral BrainScanAI et notions de machine learning supervise'
        $document.BuiltInDocumentProperties.Item('Author').Value = 'Ronan Menguy'
        $document.BuiltInDocumentProperties.Item('Subject').Value = 'Guide slide par slide et glossaire pour la soutenance BrainScanAI'
    }
    catch {
        Write-Warning 'Document properties could not be set; document creation continues.'
    }

    Write-Output 'STEP=save-docx'
    $document.SaveAs2($docxPath, 12)
    Write-Output 'STEP=export-pdf'
    $document.ExportAsFixedFormat($pdfPath, 17)

    Write-Output "DOCX=$docxPath"
    Write-Output "PDF=$pdfPath"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($null -ne $word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
