Attribute VB_Name = "Module1"
Sub Printable_MarginSet_Papersize()
With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(1.5)
        .RightMargin = Application.InchesToPoints(1)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PaperSize = xlPaperA4
End With
End Sub
Sub save_as_pdf()

        ChDir "C:\Users\Ashish\Desktop\new pdfs"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\Ashish\Desktop\new pdfs\" + ActiveSheet.Name(), Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

End Sub
Sub Column_Width_Fly_Levelling()

Columns("A").ColumnWidth = 8
Columns("B:C").ColumnWidth = 7
Columns("D:G").ColumnWidth = 5.71
Columns("H").ColumnWidth = 7
Columns("I:M").ColumnWidth = 5.86
Columns("N").ColumnWidth = 8.57
Columns("O").ColumnWidth = 7.14
Columns("P").ColumnWidth = 6.57
Columns("Q").ColumnWidth = 13
End Sub

Sub Column_Width_Detailing()
Rows("13:43").RowHeight = 18
    Columns("A").ColumnWidth = 5
    Columns("B").ColumnWidth = 4
    Columns("C:C").ColumnWidth = 5.29
    Columns("D:F").ColumnWidth = 6.71
    Columns("G:G").ColumnWidth = 8.57
    Columns("H:H").ColumnWidth = 5.5
    Columns("I:I").ColumnWidth = 7.43
    Columns("J:J").ColumnWidth = 11
End Sub


