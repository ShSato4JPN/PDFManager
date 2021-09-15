Module Main
    Sub Main()
        Dim pdm As New PDFManager("C:\Users\satos\Desktop\output\output.pdf")
        pdm.add("C:\Users\satos\Desktop\input\Merge1.pdf", {"1", "3"})
        pdm.add("C:\Users\satos\Desktop\input\Merge2.pdf", {"2"})
        pdm.Merge()

    End Sub
End Module
