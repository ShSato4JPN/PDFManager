Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class PDFManager

    Private _listOfPdfs As List(Of PDFInfo)
    Private _work As String = ".\PDFManagerWork"
    Private _output As String = ".\PDFManagerWork"

    Private Structure PDFInfo
        Dim PATH As String
        Dim SAVE_PAGE As String
    End Structure

    Sub New()
        Me._listOfPdfs = New List(Of PDFInfo)
    End Sub

    Public Sub add(ByVal srcPath As String)
        Dim d As New PDFInfo
        d.PATH = srcPath
        d.SAVE_PAGE = String.Empty

        _listOfPdfs.Add(d)
    End Sub

    Public Sub add(ByVal srcPath As String, ByVal pages As String())
        Dim d As New PDFInfo

        If 0 < pages.Length Then
            d.PATH = srcPath
            d.SAVE_PAGE = ConvertStringArrayToString(pages)
        Else
            d.PATH = srcPath
            d.SAVE_PAGE = String.Empty
        End If

        _listOfPdfs.Add(d)
    End Sub

    Public Function Merge() As Boolean
        If _listOfPdfs.Count < 1 Then
            Return True
        End If

        Dim result As Boolean = False
        Try
            CreateDirectory()

            If Not CopyPDFToWork() Then
                result = False
                Exit Try
            End If

            Using doc As New Document()
                Using pdfCopy As New PdfCopy(doc, New FileStream(_output, FileMode.Create))
                    doc.Open()

                    ' create merge pdf
                    Dim files As List(Of String) = GetAllPDFsInWork()
                    For Each file In files
                        Dim pdfReader As New PdfReader(file)
                        pdfCopy.AddDocument(pdfReader)
                        pdfReader.Close()
                    Next
                End Using
            End Using

            result = True

        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            result = False

        Finally
            DeleteDirectory()

        End Try

        Return result

    End Function

    Public Function Clear()
        _listOfPdfs = New List(Of PDFInfo)
    End Function

    Public Sub SetWorkPath(ByVal work As String)
        Me._work = work
    End Sub

    Public Sub SetOutputPath(ByVal output As String)
        Me._output = output
    End Sub

    Private Function GetAllPDFsInWork() As List(Of String)
        Dim files As IEnumerable(Of String) = Directory.EnumerateFiles(_work, "*.pdf", SearchOption.TopDirectoryOnly)
        Return files.ToList
    End Function

    Private Function CopyPDFToWork() As Boolean
        For Each file In _listOfPdfs
            ' delete pages
            If String.IsNullOrEmpty(file.SAVE_PAGE) Then
                If Not CopyNoRemovedPDFToWork(file.PATH) Then
                    Return False
                End If
            Else
                If Not CopyRemovedPDFToWork(file.PATH, file.SAVE_PAGE) Then
                    Return False
                End If
            End If
        Next

        Return True

    End Function

    Private Function CopyRemovedPDFToWork(ByVal srcFile As String, ByVal pages As String) As Boolean
        Dim result As Boolean = False
        Dim fileName As String = Path.GetFileNameWithoutExtension(srcFile)
        Try
            Using pdfReader As New PdfReader(srcFile)
                pdfReader.SelectPages(pages)
                Using pdfStamper As New PdfStamper(pdfReader, New FileStream($"{_work}\{fileName}_after.pdf", FileMode.Create))
                    result = True
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            result = False

        End Try

        Return result

    End Function

    Private Function CopyNoRemovedPDFToWork(ByVal srcFile As String) As Boolean
        Dim result As Boolean = False
        Dim fileName As String = Path.GetFileNameWithoutExtension(srcFile)
        Try
            My.Computer.FileSystem.MoveFile(srcFile, $"{_work}\{fileName}_after.pdf")
            result = True
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            result = False

        End Try

        Return result

    End Function

    Private Sub CreateDirectory()
        My.Computer.FileSystem.CreateDirectory(_work)
    End Sub

    Private Sub DeleteDirectory()
        My.Computer.FileSystem.DeleteDirectory(_work, FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

    Private Function ConvertStringArrayToString(ByVal strArray As String()) As String
        Dim result As String = String.Empty
        For Each str As String In strArray
            result = result & str & ","
        Next

        Return result.TrimEnd(","c)

    End Function

End Class
