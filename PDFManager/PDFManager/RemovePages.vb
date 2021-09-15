'
' https://ktts.hatenablog.com/entry/2018/07/17/013829 のソースをVB.net に書き換えただけ。
'
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Module RemovePages

    Sub Main()
        Debug.WriteLine("RemovePages start")

        Dim targetSrcPath As String = ".\output\Merge.pdf"

        Dim outSrcName As String = "RemovePages.pdf"
        Dim outSrcDir As String = ".\output"
        Dim outSrcPath As String = $"{outSrcDir}\{outSrcName}"

        Try
            Using pdfReader As New PdfReader(targetSrcPath)
                pdfReader.SelectPages("1,3,100")
                Using pdfStamper As New PdfStamper(pdfReader, New FileStream(outSrcPath, FileMode.Create))
                    Debug.WriteLine("delete success!")
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine(ex.Message)

        End Try

        Debug.WriteLine("RemovePages end")
    End Sub

End Module
