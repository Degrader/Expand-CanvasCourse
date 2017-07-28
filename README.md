# Export-CanvasCourse
Takes a Canvas '.imscc' file, and generates either a repaired HTML structure, or exports to pdf with wkhtmltopdf.


A parameter for the function called 'Clean' will default to $true unless otherwise specified. This means that if wkhtmltopdf is
installed, you will only be left with PDF documents. If 'Clean' is set to $false, then you will be left with both the PDF documents
as well as the HTML file structure with your raw course content (images, videos, etc.).


Example:

The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive
in the destination directory and will do one of two things:
  If wkhtmltopdf is installed in it's default installation path, you will be left with PDF's for each module in your Canvas Course.
  If wkhtmltopdf is not installed, you will be left with a file structure including HTML documents with your course content.

Export-CanvasCourse -IMSCCPath C:\Users\Person\Desktop\example_course.imscc -DestinationPath "C:\temp\Example Course"




Example:

The following will take our imscc file located at C:\Users\Person\Desktop and rename the file to .zip. It then expands the archive in the destination directory. The -Clean parameter only makes a difference when wkhtmltopdf is installed, and if set to false, this function will leave behind the raw folder structure, and move all PDF documents to a directory within that folder structure (.\PDF\)

Export-CanvasCourse -Path C:\Users\Person\Desktop\example_course.imscc -Destination C:\Users\Person\Desktop\example_course -Clean $false
