<#
    .Synopsis
    Extracts exported '.IMSCC' and repairs media/content locations in html documents. Exports to 

    .Description
    Extracts exported '.IMSCC' and repairs media/content locations in html documents.

    .Parameter IMSCCPath
    Full path to IMSCC file to be extracted/converted.

    .Parameter DestinationPath
    Full path to folder for extracted/converted content.

    .Parameter Clean
    Boolean parameter that specifies the original imscc and html structure should be removed after conversion. Default true.

    .Example
    Export-CanvasCourse -IMSCCPath C:\Users\Person\Desktop\example_course.imscc -DestinationPath C:\Users\Person\Desktop\example_course

    .Example
    Export-CanvasCourse -IMSCCPath C:\Users\Person\Desktop\example_course.imscc -DestinationPath C:\Users\Person\Desktop\example_course -Clean $true
#>
function Export-CanvasCourse{
    param(
        [Parameter(Mandatory=$true)]
        $IMSCCPath,
        [Parameter(Mandatory=$true)]
        $DestinationPath,
        [Parameter(Mandatory=$false)]
        [Bool]$Clean = $true
        )

    #Check for wkhtmltopdf installation
    $Check32 = Test-Path -Path 'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe'
    $check64 = Test-Path -Path 'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    
    #Renames imscc file to .zip, and expands archive in destination file path
    $NewArchiveName = $IMSCCPath.replace('.imscc', '.zip')
    Rename-Item -Path $IMSCCPath -NewName $IMSCCPath.replace('.imscc', '.zip')
    Expand-Archive -Path $NewArchiveName -DestinationPath $DestinationPath -Force

    #Gather HTML documents
    $htmlObjects = Get-ChildItem -Path $DestinationPath -Recurse -Include "*.html"

    #Repair HTML to point all content to the appropriate directory 'web_resources'
    #Export to pdf is wkhtmltopdf is installed
    foreach ($htmlObject in $htmlObjects){
        (Get-Content $htmlObject).Replace('%24IMS-CC-FILEBASE%24', '../web_resources') | Set-Content $htmlObject -Force
        (Get-Content $htmlObject).Replace('%20', ' ') | Set-Content $htmlObject -Force
        if ($Check32 -eq $true){
            $temp = $htmlObject.Name
            $temp = $temp.Replace(".html", ".pdf")
            $temp = "$DestinationPath\$temp"
            & 'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe' $htmlObject $temp
        }
        if ($Check64 -eq $true){
            $temp = $htmlObject.Name
            $temp = $temp.Replace(".html", ".pdf")
            $temp = "$DestinationPath\$temp"
            & 'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe' $htmlObject $temp
        }
    }
    
    #If wkhtmltopdf is not installed, let user know where html structure is located
    if ($Check32 -eq $false -and $check64 -eq $false){
        Write-Host "WKHTMLTOPDF is not installed. PDF's not generated.`nPlease see $DestinationPath for html version."
    }
    #If wkhtmltopdf is installed, and clean is true, remove everything but remaining pdf's
    if ($Check32 -eq $true -or $check64 -eq $true -and $Clean -eq $true){
        Remove-Item $DestinationPath -Exclude "*.pdf" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $NewArchiveName -Force -ErrorAction SilentlyContinue
    }
    #If clean was set to false, rename zip back to imscc
    if ($Clean -eq $false){
        Rename-Item -Path $NewArchiveName -NewName $IMSCCPath
    }
}