Imports System.IO
Imports System.IO.Packaging


Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim path_str As String = "D:\"
        Dim zip_name As String = "1111"


        Dim Tab As String
        Dim FI As IO.FileInfo = New IO.FileInfo(path_str + "sheet1.xml")
        If FI.Exists Then
            Tab = System.IO.File.ReadAllText(path_str + "sheet1.xml") 'My.Resources.sheet1
        Else
            Tab = My.Resources.sheet1
        End If

        Tab = Replace(Tab, "_stroka", TextBox1.Text)
        System.IO.File.WriteAllText(path_str + "sheet1.xml", Tab)
        TO_ZIP(path_str, zip_name)
    End Sub


    Private Sub TO_ZIP(ByVal path_str As String, ByVal file_str As String)
        Dim zip_name As String = file_str & "___" & Replace(Replace(Now, ":", "_"), ".", "_")
        System.IO.File.WriteAllBytes(path_str + zip_name + ".zip.xlsx", My.Resources.ved)

        Call ZipFiles(path_str + zip_name + ".zip.xlsx", path_str + "sheet1.xml")

        System.IO.File.Delete(path_str + "sheet1.xml")
        System.Diagnostics.Process.Start(path_str + zip_name + ".zip.xlsx")
    End Sub


    Private Sub ZipFiles(ByVal path_str_zip As String, ByVal path_str_xml As String)
        Dim zipPath As String = path_str_zip
        'Open the zip file if it exists, else create a new one 
        Dim zip As Package = ZipPackage.Open(zipPath, _
             IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
        'Add as many files as you like:
        AddToArchive(zip, path_str_xml)
        zip.Close() 'Close the zip file
    End Sub


    Private Sub AddToArchive(ByVal zip As Package, _
                         ByVal fileToAdd As String)
        'Replace spaces with an underscore (_) 
        Dim uriFileName As String = fileToAdd.Replace(" ", "_")
        'A Uri always starts with a forward slash "/" 
        Dim zipUri As String = String.Concat("/xl/worksheets/", _
                   IO.Path.GetFileName(uriFileName))
        Dim partUri As New Uri(zipUri, UriKind.Relative)
        Dim contentType As String = _
                   Net.Mime.MediaTypeNames.Application.Zip
        'The PackagePart contains the information: 
        ' Where to extract the file when it's extracted (partUri) 
        ' The type of content stream (MIME type):  (contentType) 
        ' The type of compression:  (CompressionOption.Normal)   
        Dim pkgPart As PackagePart = zip.CreatePart(partUri, _
                   contentType, CompressionOption.Normal)
        'Read all of the bytes from the file to add to the zip file 
        Dim bites As Byte() = File.ReadAllBytes(fileToAdd)
        'Compress and write the bytes to the zip file 
        pkgPart.GetStream().Write(bites, 0, bites.Length)
    End Sub
End Class
