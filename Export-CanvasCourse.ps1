<#
    .Synopsis
    Extracts exported '.IMSCC' and repairs media/content locations in html or pdf documents. 

    .Description
    Extracts exported '.IMSCC' and repairs media/content locations in html documents. Can export
    to PDF if wkhtmltopdf is installed.

    .Parameter IMSCCPath
    Full path to IMSCC file to be extracted/converted.

    .Parameter DestinationPath
    Full path to folder for extracted/converted content.

    .Parameter Clean
    Default value is true. 'Clean' will default to $true unless otherwise specified. This means that if wkhtmltopdf is
    installed, you will only be left with PDF documents. If 'Clean' is set to $false, then you will be left with both
    the PDF documents as well as the HTML file structure with your raw course content (images, videos, etc.).

    .Example
    The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive in the destination directory and will do one of two things:
    If wkhtmltopdf is installed in it's default installation path, you will be left with PDF's for each module in your Canvas Course.
    If wkhtmltopdf is not installed, you will be left with a file structure including HTML documents with your course content.

    Export-CanvasCourse -IMSCCPath C:\Users\Person\Desktop\example_course.imscc -DestinationPath C:\Users\Person\Desktop\example_course

    .Example
    The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive in the destination directory. The -Clean parameter only makes a difference when wkhtmltopdf is installed, and if set to false, this function will leave behind the raw folder structure, and move all PDF documents to a directory within that folder structure (.\PDF\)

    Export-CanvasCourse -IMSCCPath C:\Users\Person\Desktop\example_course.imscc -DestinationPath C:\Users\Person\Desktop\example_course -Clean $false
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
    }
    #If wkhtmltopdf is installed, and clean is false, store PDF's in PDF directory under the destination root
    if ($Check32 -eq $true -or $check64 -eq $true -and $Clean -eq $false){
        New-Item -ItemType Directory -Path $DestinationPath\PDF -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $DestinationPath -Include "*.pdf" -Recurse -Depth 0 | Move-Item -Destination "$DestinationPath\PDF"
    }
    #rename zip back to imscc
    Rename-Item -Path $NewArchiveName -NewName $IMSCCPath
}