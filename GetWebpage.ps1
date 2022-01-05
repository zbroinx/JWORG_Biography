function getWebsiteText ($url){

    $response = Invoke-WebRequest –Uri $url
    #$webpage = $response.ParsedHtml.getElementsByTagName('article')[0].InnerText
    $webpage = $response.ParsedHtml.getElementById('article').InnerText

    return $webpage -replace '(?ms)(?:\r|\n)^\s*$'
}

#Get file name from narrator
function GetFileName
{
    $t = ($webp -split '\n') | Select-String -Pattern "Narrad" | Select -First 1
    $title = $t -split "por " | Select -Skip 1 #-First 1

    return $title
}

#Convert input (file) into a PDF document (requires Word installed).
Function ConvertTo-PDFFile
{
    Param
    (
        [string]$Source,
        [string]$Destionation
    )

    #Get the content of the file.
    #$Source = Get-Content $Source -Encoding UTF8;
    $Source = Get-Content $Source -Encoding UTF8 -Raw;

    #Required Word Variables.
    $ExportFormat = 17;
    $SaveOption = 0

    #Create a hidden Word window.
    $WordObject = New-Object -ComObject word.application;
    $WordObject.Visible = $false;

    #Add a Word document.
    $DcoumentObject = $WordObject.Documents.Add();

    #Put the text into the Word document.
    $WordSelection = $WordObject.Selection;
    $WordSelection.TypeText($Source);

    #Set the page orientation to landscape.
    $DcoumentObject.PageSetup.Orientation = 1;

    #Export the PDF file and close without saving a Word document.
    $DcoumentObject.ExportAsFixedFormat($Destionation,$ExportFormat);
    $DcoumentObject.close([ref]$SaveOption);
    $WordObject.Quit();
}


Get-Content 'C:\temp\WatchtowerUrls.txt' | ForEach-Object {
    
    $website = $_

    #Get website
    $webp = getWebsiteText($website)

    #Ignore 'Our Readers' articles
    $readers = ($webp | Select-String -Pattern 'Nossos Leitores').Matches.Value
    if ($readers -notcontains "*Nossos Leitores*"){
    
        # Get filename
        $file = GetFileName
        $fileName = $file | Select -First 1 | Out-String
        #Remove carriage return
        $fileName = $fileName -replace "`n|`r"
        #Set path to save file
        $txtFilePath="C:\temp\txtArticles\" + $fileName.ToString() + ".txt"
        $pdfFilePath="C:\temp\pdfArticles\" + $fileName.ToString() + ".pdf"
    
        #Save webpage text as txt file
        $webp | Out-File -FilePath $txtFilePath

        ConvertTo-PDFFile -Source $txtFilePath -Destionation $pdfFilePath;
    }
    
}



#$website = 'https://wol.jw.org/pt/wol/d/r5/lp-t/102007329' #De Nossos Leitores
$website = 'https://wol.jw.org/pt/wol/d/r5/lp-t/2015649' # Normal


