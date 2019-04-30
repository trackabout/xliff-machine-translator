# Accepts an XLIFF file (1.2 schema), which is the preferred output format for the
# Microsoft Multilingual App Toolkit that integrates with Visual Studio.
# Processes the Xliff through the Microsoft Translator API v3.0
# Supports a custom translator trained Neural Machiine Translation model
# through the setting of the CategoryId parameter.

param(
    [ValidateScript({Test-Path -Path $_ -PathType leaf})]
    [Parameter(ValueFromPipeline=$true, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [string]$XlfInputFile,
    [Parameter(Mandatory=$true)]
    [string]$XlfOutputFile,
    [Parameter(Mandatory=$true)]
    [string]$ApiKey,
    [Parameter(Mandatory=$true)]
    [string]$CategoryId = "General",
    # Our company uses square brackets to denote terms that are replaced
    # on the fly in our application. For translations, we remove them.
    [switch]$RemoveSquareBrackets

)

function Translate-String {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ApiKey,
        [Parameter(Mandatory=$true)]
        [string]$sourceLanguage,
        [Parameter(Mandatory=$true)]
        [string]$targetLanguage,
        [Parameter(Mandatory=$true)]
        [string]$textToConvert
    )
    # Translation API
    $translateBaseURI = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&category=$CategoryId"
    # API Auth Headers
    $headers = @{}
    $headers.Add("Ocp-Apim-Subscription-Key",$apiKey)
    $headers.Add("Content-Type","application/json")
    # Conversion URI
    $convertURI = "$($translateBaseURI)&from=$($sourceLanguage)&to=$($targetLanguage)"
    # Build Conversion Body
    $text = @{'Text' = $($textToConvert)}
    $text = $text | ConvertTo-Json
    # Convert
    $conversionResult = Invoke-RestMethod -Method POST -Uri $convertURI -Headers $headers -Body "[$($text)]"
    return [string]$conversionResult.translations[0].text
}

$ErrorActionPreference = "Stop"
$sw = [Diagnostics.Stopwatch]::StartNew()
$sourceByteSum = 0
$targetByteSum = 0

Write-Output "Opening $($XlfInputFile)"
[System.Xml.XmlDocument]$xml = new-object System.Xml.XmlDocument
$xml.load($XlfInputFile)

$sourceLanguage = $xml.xliff.file.'source-language'
$targetLanguage = $xml.xliff.file.'target-language'

Write-Host -ForegroundColor Yellow "Translating $($XlfInputFile) from $($sourceLanguage) to $($targetLanguage)."

$count = 0
foreach ($node in $xml.xliff.file.body.group.'trans-unit') {
    if ($node.target.state -ne "new") {
        continue
    }
    # According to the XLIFF XSD http://docs.oasis-open.org/xliff/v1.2/cs02/xliff-core.html#it,
    # there can be sub-elements of the <source> tag, like 'g', 'x', 'bx', 'it', etc.
    # Check for them, and if they exist, change our node pointer to the sub-element.
    $sourceElementContainingText = $node.source
    while ($sourceElementContainingText -isnot [String] -and $sourceElementContainingText -isnot [System.Xml.XmlText])
    {
        $sourceElementContainingText = $sourceElementContainingText.ChildNodes[0]
    }

    if ($sourceElementContainingText -is [String]) {
        $sourceText = $sourceElementContainingText
    }
    elseif ($sourceElementContainingText -is [System.Xml.XmlText]) {
        $sourceText = $sourceElementContainingText.Value
    }
    if ($RemoveSquareBrackets.IsPresent) {
        # Remove square brackets from the source text.
        # We do not currently support localization of Term Lookups.
        $sourceText = ($sourceText -replace '[][]','')
    }

    # Farm it out to the great AI in the Clouds!
    $translatedText = Translate-String -ApiKey $ApiKey -sourceLanguage $sourceLanguage -targetLanguage $targetLanguage -textToConvert $sourceText

    $sourceByteSum += [System.Text.Encoding]::UTF8.GetByteCount([string]$sourceText)
    $targetByteSum += [System.Text.Encoding]::UTF8.GetByteCount([string]$translatedText)

    Write-Output "Translated '$($sourceText)' to '$($translatedText)'."

    # Set the same attributes on the 'trans-unit.target' element that the Multilingual App Toolkit would.
    $node.target.SetAttribute("state", "needs-review-translation")
    $node.target.SetAttribute("state-qualifier", "tm-suggestion")

    # Put the translation into the target.
    $targetElementContainingText = $node.target
    while ($targetElementContainingText -isnot [String] -and $targetElementContainingText -isnot [System.Xml.XmlText])
    {
        $targetElementContainingText = $targetElementContainingText.ChildNodes[0]
    }
    if ($targetElementContainingText -is [String]) {
        $targetElementContainingText = $translatedText
    }
    elseif ($targetElementContainingText -is [System.Xml.XmlText]) {
        $targetElementContainingText.InnerText = $translatedText
    }

    $count = $count + 1
}

$xml.Save($XlfOutputFile)

Write-Host -ForegroundColor Yellow "Done. File written to $($XlfOutputFile)"
Write-Output "Translated $($count) strings."
Write-Output "Total bytes of source text: $({0:N}-f $sourceByteSum)"
Write-Output "Total bytes of target text: $({0:N} -f $targetByteSum)"

$sw.Stop()
Write-Host -ForegroundColor Green "Time taken: $($sw.Elapsed)"
