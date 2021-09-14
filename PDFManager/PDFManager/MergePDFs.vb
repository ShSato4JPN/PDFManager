'
' https://zero0nine.com/archives/2866 のソースをVB.net に書き換えただけ。
'
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Module MergePDFs

    Sub Main()
        Debug.WriteLine("MergePDFs start")

        Dim pdfsDir As String = ".\..\..\PDFs"
        Dim outSrcName As String = "Merge.pdf"
        Dim outSrcDir As String = ".\output"
        Dim outSrcPath As String = $"{outSrcDir}\{outSrcName}"

        If Not Directory.Exists(outSrcDir) Then
            Directory.CreateDirectory(outSrcDir)
        End If

        If File.Exists(outSrcPath) Then
            File.Delete(outSrcPath)
        End If

        ' PDFファイルを取得
        Dim files As IEnumerable(Of String) = Directory.EnumerateFiles(pdfsDir, "*.pdf", SearchOption.TopDirectoryOnly)

        ' PDF集約
        Dim doc As Document = Nothing
        Dim pdfCopy As PdfCopy = Nothing

        Try
            doc = New Document()
            pdfCopy = New PdfCopy(doc, New FileStream(outSrcPath, FileMode.Create))

            doc.Open()

            ' 出力するPDFのプロパティを設定(必要なさそう)
            doc.AddKeywords("a")
            doc.AddAuthor("b")
            doc.AddTitle("c")
            doc.AddCreator("d")
            doc.AddSubject("e")

            For Each file In files
                Dim pdfReader As New PdfReader(file)
                pdfCopy.AddDocument(pdfReader)
                pdfReader.Close()
            Next

        Catch ex As Exception
            Debug.WriteLine(ex.Message)

        Finally
            doc.Close()
            pdfCopy.Close()

        End Try

        Debug.WriteLine("MergePDFs end")
    End Sub

End Module
