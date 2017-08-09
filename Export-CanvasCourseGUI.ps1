<#
    .Synopsis
    Extracts exported '.IMSCC' and repairs media/content locations in html or pdf documents. 

    .Description
    Extracts exported '.IMSCC' and repairs media/content locations in html documents. Can export
    to PDF if wkhtmltopdf is installed.

    .Parameter Path
    Full path to IMSCC file to be extracted/converted.

    .Parameter Destination
    Full path to folder for extracted/converted content.

    .Parameter Clean
    Default value is true. 'Clean' will default to $true unless otherwise specified. This means that if wkhtmltopdf is
    installed, you will only be left with PDF documents. If 'Clean' is set to $false, then you will be left with both
    the PDF documents as well as the HTML file structure with your raw course content (images, videos, etc.).

    .Example
    The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive in the destination directory and will do one of two things:
    If wkhtmltopdf is installed in it's default installation path, you will be left with PDF's for each module in your Canvas Course.
    If wkhtmltopdf is not installed, you will be left with a file structure including HTML documents with your course content.

    Export-CanvasCourse -Path C:\Users\Person\Desktop\example_course.imscc -Destination C:\Users\Person\Desktop\example_course

    .Example
    The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive in the destination directory. The -Clean parameter only makes a difference when wkhtmltopdf is installed, and if set to false, this function will leave behind the raw folder structure, and move all PDF documents to a directory within that folder structure (.\PDF\)

    Export-CanvasCourse -Path C:\Users\Person\Desktop\example_course.imscc -Destination C:\Users\Person\Desktop\example_course -Clean $false
#>

Function Get-FileName($initialDirectory)
{
    #[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.Title = "IMSCC File Location..."
    $OpenFileDialog.filter = "IMSCC (*.imscc)| *.imscc"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Set-FolderName($initialDirectory){
    $SaveFileDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $SaveFileDialog.ShowNewFolderButton = $true
    $SaveFileDialog.Description = "Select destination location for content"
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.SelectedPath
}

function Export-CanvasCourse{
    param(
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $Destination,
        [Parameter(Mandatory=$false)]
        $Clean
        )

    #Check for wkhtmltopdf installation
    $script:Check32 = Test-Path -Path 'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe'
    $script:check64 = Test-Path -Path 'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    
    #Renames imscc file to .zip, and expands archive in destination file path
    $NewArchiveName = $Path.replace('.imscc', '.zip')
    Rename-Item -Path $Path -NewName $NewArchiveName
    Expand-Archive -Path $NewArchiveName -Destination $Destination -Force
    #rename zip back to imscc
    Rename-Item -Path $NewArchiveName -NewName $Path

    #Gather HTML documents
    $htmlObjects = Get-ChildItem -Path $Destination -Recurse -Include "*.html"

    #Create PDF folder
    New-Item -ItemType Directory -Path "$Destination\PDF" -ErrorAction SilentlyContinue

    #Repair HTML to point all content to the appropriate directory 'web_resources'
    #Export to pdf if wkhtmltopdf is installed
    foreach ($htmlObject in $htmlObjects){
        (Get-Content $htmlObject).Replace('%24IMS-CC-FILEBASE%24', '../web_resources') | Set-Content $htmlObject -Force
        (Get-Content $htmlObject).Replace('%20', ' ') | Set-Content $htmlObject -Force
        if ($Check32 -eq $true){
            $temp = $htmlObject.Name
            $temp = $temp.Replace(".html", ".pdf")
            $temp = "$Destination\pdf\$temp"
            & 'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe' $htmlObject $temp | Out-Null
        }
        if ($Check64 -eq $true){
            $temp = $htmlObject.Name
            $temp = $temp.Replace(".html", ".pdf")
            $temp = "$Destination\pdf\$temp"
            & 'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe' $htmlObject $temp | Out-Null
        }
    }
    
    #If wkhtmltopdf is not installed, let user know where html structure is located
    if ($Check32 -eq $false -and $check64 -eq $false){
        [System.Windows.forms.MessageBox]::Show("No PDFs were generated because WKHTMLTOPDF is not installed. Please see $Destination for html based course content.", "Information", "OK")
    }
    #If wkhtmltopdf is installed, and clean is true, remove everything but remaining pdf's
    if (($Check32 -eq $true -or $check64 -eq $true) -and $Clean -eq "Yes"){
        Get-ChildItem -Path $Destination -Exclude "*.pdf" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    #If wkhtmltopdf is installed, and clean is false, store PDF's in PDF directory under the destination root
    if (($Check32 -eq $true -or $check64 -eq $true) -and $Clean -eq "No"){
        Get-ChildItem -Path $Destination\* -Include "*.pdf" | Move-Item -Destination "$Destination\PDF"
    }
}

#The script bit :P########################################################################

#Get input file
$inputFile = Get-FileName "$env:USERPROFILE"
if ($inputFile -eq ""){exit}

#Get output directory
$outputDirectory = Set-FolderName "$env:USERPROFILE"
if ($outputDirectory -eq ""){exit}

if ($Check32 -eq $true -or $Check64 -eq $true){
    $CleanContent = [System.Windows.forms.MessageBox]::Show('Remove extra course matterial? Selecting yes removes all of the original course matterial, such as videos and picture (Leaves only PDF files) Selecting leaves all matterial, as well as the new PDFs.','Remove extra content...','YesNo')
}

Export-CanvasCourse $inputFile $outputDirectory $CleanContent