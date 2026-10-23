$ErrorActionPreference = "Stop"

function Convert-ToOfficeRgb {
    param(
        [int]$Red,
        [int]$Green,
        [int]$Blue
    )

    return $Red + (256 * $Green) + (65536 * $Blue)
}

function Set-TextStyle {
    param(
        $Shape,
        [string]$Text,
        [double]$FontSize,
        [int]$Color,
        [bool]$Bold = $false
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Name = "Arial"
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $Shape.TextFrame.TextRange.Font.Bold = if ($Bold) { -1 } else { 0 }
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $Shape.TextFrame.WordWrap = -1
    $Shape.TextFrame.AutoSize = 0
}

$workspace = (Get-Location).Path
$sourcePath = Join-Path $workspace "Ronan_Menguy_1_Presentation_092026.pptx"
$buildDirectory = Join-Path $workspace ".build\presentation-coherence"
$draftPath = Join-Path $buildDirectory "Ronan_Menguy_1_Presentation_092026_draft.pptx"
$previewDirectory = Join-Path $buildDirectory "after"

New-Item -ItemType Directory -Force -Path $buildDirectory | Out-Null
New-Item -ItemType Directory -Force -Path $previewDirectory | Out-Null
Copy-Item -LiteralPath $sourcePath -Destination $draftPath -Force

$navy = Convert-ToOfficeRgb 23 47 67
$teal = Convert-ToOfficeRgb 0 127 130
$blue = Convert-ToOfficeRgb 57 120 169
$gray = Convert-ToOfficeRgb 83 100 115

$powerpoint = New-Object -ComObject PowerPoint.Application

try {
    $presentation = $powerpoint.Presentations.Open(
        $draftPath,
        $false,
        $false,
        $false
    )

    # Diapositive 3 : rendre visible le contrôle des images corrompues.
    $slide3 = $presentation.Slides.Item(3)
    $integrityText = $slide3.Shapes.Item("TextBox 16")
    $integrityText.Height = 66
    Set-TextStyle `
        -Shape $integrityText `
        -Text ("Contrôle d’intégrité : 1 410 images vérifiées avec Pillow et Torchvision, aucune image corrompue détectée." + [char]13 + "Égalisation des histogrammes avant la normalisation pour ajuster le contraste.") `
        -FontSize 18 `
        -Color $navy
    $integrityText.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $integrityText.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $teal

    # Diapositive 8 : diagnostic initial sur le train et la validation.
    $slide8 = $presentation.Slides.Item(8)
    Set-TextStyle `
        -Shape $slide8.Shapes.Item("TextBox 1") `
        -Text "Entraînements et diagnostic sur la validation" `
        -FontSize 32 `
        -Color $navy `
        -Bold $true

    $validationBox = $slide8.Shapes.Item("Rectangle 11")
    Set-TextStyle `
        -Shape $validationBox `
        -Text ("Validation forte" + [char]13 + "16 images") `
        -FontSize 21 `
        -Color $navy
    $validationBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $validationBox.TextFrame.VerticalAnchor = 3

    $semiDiagnostic = $slide8.Shapes.Item("TextBox 8")
    $semiDiagnostic.Left = 48
    $semiDiagnostic.Top = 278
    $semiDiagnostic.Width = 620
    $semiDiagnostic.Height = 63
    Set-TextStyle `
        -Shape $semiDiagnostic `
        -Text ("F2 train 0,969, validation 0,854, écart +11,5 points." + [char]13 + "Surapprentissage possible pour le semi-supervisé.") `
        -FontSize 16.5 `
        -Color $navy
    $semiDiagnostic.TextFrame.TextRange.Paragraphs(2).Font.Bold = -1
    $semiDiagnostic.TextFrame.TextRange.Paragraphs(2).Font.Color.RGB = $teal

    $supervisedDiagnostic = $slide8.Shapes.Item("TextBox 18")
    $supervisedDiagnostic.Left = 48
    $supervisedDiagnostic.Top = 438
    $supervisedDiagnostic.Width = 864
    $supervisedDiagnostic.Height = 62
    Set-TextStyle `
        -Shape $supervisedDiagnostic `
        -Text ("Supervisé : F2 train 0,943, validation 0,976, écart -3,2 points. Aucun surapprentissage observé." + [char]13 + "Le dropout et la L2 sont donc testés uniquement sur le semi-supervisé. Le test fort reste réservé à l’évaluation finale.") `
        -FontSize 14.5 `
        -Color $gray
    $supervisedDiagnostic.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $supervisedDiagnostic.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $blue

    # Diapositive 9 : décision sur validation, puis résultat final sur le test.
    $slide9 = $presentation.Slides.Item(9)
    Set-TextStyle `
        -Shape $slide9.Shapes.Item("TextBox 1") `
        -Text "Sélection sur validation et résultats finaux sur test" `
        -FontSize 32 `
        -Color $navy `
        -Bold $true

    foreach ($shapeName in "TextBox 11", "TextBox 12") {
        $errorLabel = $slide9.Shapes.Item($shapeName)
        $errorLabel.Top = 338
        $errorLabel.Height = 43
        Set-TextStyle `
            -Shape $errorLabel `
            -Text ("1 faux positif" + [char]13 + "1 faux négatif") `
            -FontSize 14 `
            -Color $(if ($shapeName -eq "TextBox 11") { $teal } else { $blue }) `
            -Bold $true
    }

    $validationDecision = $slide9.Shapes.Item("TextBox 13")
    $validationDecision.Top = 398
    $validationDecision.Height = 54
    Set-TextStyle `
        -Shape $validationDecision `
        -Text "Validation semi-supervisée : initial 0,854, dropout 0,526, L2 0,641. Les deux régularisations sont écartées." `
        -FontSize 16.5 `
        -Color $navy `
        -Bold $true

    $finalTest = $slide9.Shapes.Item("TextBox 14")
    $finalTest.Top = 458
    $finalTest.Height = 43
    Set-TextStyle `
        -Shape $finalTest `
        -Text "Test final : F2 cancer = 0,900 pour les deux modèles initiaux. Chaque matrice contient 1 faux positif et 1 faux négatif." `
        -FontSize 15.5 `
        -Color $gray

    $presentation.Save()

    foreach ($number in 3, 8, 9) {
        $pngPath = Join-Path $previewDirectory ("slide-{0}.png" -f $number)
        $presentation.Slides.Item($number).Export($pngPath, "PNG", 1920, 1080)
    }

    $presentation.Close()
}
finally {
    $powerpoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
}

Write-Output $draftPath
