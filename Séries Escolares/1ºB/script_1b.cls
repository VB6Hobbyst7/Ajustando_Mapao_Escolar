Sub Ajusta_Mapao()
    
    ' Ajusta Espaçamento de Colunas
    Range("F12:F49,K12:K49,P12:P49,U12:U49,Z12:Z49,AE12:AE49,AJ12:AJ49,AO12:AO49,AT12:AT49,AY12:AY49,BD12:BD49,BI12:BI49,BN12:BN49").Select
    Selection.Delete

    ' Ajusta Fonte de Linha Aulas Dadas
    Rows("49:49").Select
    With Selection.Font
        .Size = 10
    End With

    Rows("49:49").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = True
        .ReadingOrder = xlContext
    End With

    ' Remove Linhas
    Range("50:50,15:15,11:11,10:10,5:5,4:4,2:2,1:1").Select
    Selection.Delete

    ' Remove a Primeira Coluna
    Columns("A:A").Select
    Selection.Delete

    ' Ajusta Fonte Cabeçalho
    Rows("1:8").Select
    With Selection.Font
        .Size = 10
    End With

    ' Ajusta Fonte Nomes Alunos
    Rows("9:49").Select
    With Selection.Font
        .Size = 12
    End With
    
    ' Ajusta Espaçamento do Nome do Aluno
    Columns("C:C").ColumnWidth = 34.29

    ' Ajusta Altura Linhas 
    Rows("1:67").EntireRow.AutoFit

    ' Ajusta Espaçamento das Colunas Notas/Faltas
    Columns("D:BD").EntireColumn.AutoFit

    Range("A2:AC2,A3:AC3,A4:AC4,A5:AC5").Select
    Selection.UnMerge

    Range("A2:D2,A3:D3,A4:D4,A5:D5").Select
    Selection.Merge True

    Range("E2:AD2,E3:AD3,E4:AD4,E5:AD5").Select
    Selection.Delete Shift:=xlToLeft

    Range("E2:AL2,E3:AL3,E4:AL4,E5:AL5").Select
    Selection.UnMerge

    Range("E2:R2,E3:R3,E4:R4,E5:R5").Select
    Selection.Merge True

    Range("S2:AL2,S3:AL3,S4:AL4,S5:AL5").Select
    Selection.Delete Shift:=xlToLeft

    Columns("T:T").ColumnWidth = 3.86
    
    Columns("AN:AN").ColumnWidth = 3.86
    
    Columns("BC:BC").ColumnWidth = 8.71

    Rows("8:8").EntireRow.AutoFit

    Rows("6:6").Select
    Selection.RowHeight = 37.5
    
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.787401575)
        .RightMargin = Application.InchesToPoints(0.787401575)
        .TopMargin = Application.InchesToPoints(0.984251969)
        .BottomMargin = Application.InchesToPoints(0.984251969)
        .HeaderMargin = Application.InchesToPoints(0.4921259845)
        .FooterMargin = Application.InchesToPoints(0.4921259845)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 48
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    
    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1

End Sub