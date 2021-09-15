Module Main
    Sub Main()
        Dim pdm As New PDFManager()
        pdm.SetOutputPath("C:\Users\satos\Desktop\output\output.pdf")
        pdm.add("C:\Users\satos\Desktop\input\Merge1.pdf")
        pdm.add("C:\Users\satos\Desktop\input\Merge2.pdf", {"2", "3"})
        pdm.add("C:\Users\satos\Desktop\input\Merge3.pdf", {"1", "3"})
        pdm.Merge()
        pdm.Clear()

    End Sub
End Module
