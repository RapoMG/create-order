Sub Zamówienie()
'
' Zamówienie Makro
'

'Deklaracja dla usuwania i daty

Dim PustePap, PusteTyt, data, Kopia, Linia

'Zm. Kopia służy do przenoszenia wartości komórki bez formatowania _
 Zm. Linia do rysowania dołu tabeli

PustePap = 109 'KoniecPap
PusteTyt = 124 'KoniecTyt
PusteGilzy = 134 'KoniecGil
data = Date
data = Format(data, "dd.mm.yy")
Linia = 4

' Kontrola duplikatu nazwy

For Nazwy = 1 To Sheets.Count
    If Sheets(Nazwy).Name = data Then
        msg = "Dziś utworzono już zamówienie."
        Style = vbOKOnly
        Title = "Zamówienie"
        Response = MsgBox(msg, Style, Title)
            If Response = 1 Then
                Exit Sub
            End If
    End If
Next Nazwy

' Brak migotania - skoków między kartami
Application.ScreenUpdating = False

ActiveWorkbook.Unprotect Password:="onomatopeja"

' Zliczanie wierszy dla poszczególnych grup

For IlePap = 4 To Rows.Count
    If Cells(IlePap, 2) = 0 Then
        IlePap = IlePap - 1
        Exit For
    End If
Next IlePap

For IleTyt = IlePap + 4 To Rows.Count
    If Cells(IleTyt, 2) = 0 Then
        IleTyt = IleTyt - 1
        Exit For
    End If
Next IleTyt

For IleGil = IleTyt + 4 To Rows.Count
    If Cells(IleGil, 2) = 0 Then
        IleGil = IleGil - 1
        Exit For
    End If
Next IleGil

' Nowa katra - szer. kolumn i nazwa

Set NewSheet = Sheets.Add(After:=Sheets(1))
With NewSheet
   .Name = data
   .Columns("A:A").ColumnWidth = 3.67
   .Columns("B:B").ColumnWidth = 8.56
   .Columns("C:C").ColumnWidth = 15
   .Columns("D:D").ColumnWidth = 35
   .Columns("E:E").ColumnWidth = 8
   End With

' Połączenie komórek dla pola daty zam.

NewSheet.Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
 
' Data zamówienia
    
NewSheet.Cells(1, 1) = "Zamówione:"
NewSheet.Cells(1, 3) = data


'    PAPIEROSY
' Test nagłówka zamówienie pap.

For CzyPap = 4 To IlePap
If Sheets("Zamówienia").Cells(CzyPap, 5) > 0 Then
    Sheets("Zamówienia").Activate
    Range("A2:E3").Select
    Selection.Copy
    NewSheet.Activate
    Range("A2:E3").Select
    ActiveSheet.Paste
End If
Next CzyPap

' Kopiowanie zamówionych pap.

For pap = 4 To IlePap
If Sheets("Zamówienia").Cells(pap, 5) > 0 Then
    Sheets("Zamówienia").Activate
    Range(Cells(pap, 1), Cells(pap, 4)).Select
    Selection.Copy
    ' Pobranie danych zamiast kopii
    Cells(pap, 5).Select
    Kopia = Cells(pap, 5)
    ' Wklejanie
    NewSheet.Activate
    Range(Cells(pap, 1), Cells(pap, 4)).Select 'o tu 5
    ActiveSheet.Paste
    ' Wpisywanie
    Cells(pap, 5).Select
    Cells(pap, 5) = Kopia
    ' Prawa krawędź tabeli
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    ' Oczyszczenie pola w formularzu
    Sheets("Zamówienia").Activate
    Cells(pap, 5) = ""
End If
Next pap


'   TYTOŃ
' Test nagłówka zamówienie tyt.

For CzyTyt = IlePap + 4 To IleTyt
If Sheets("Zamówienia").Cells(CzyTyt, 5) > 0 Then
    Sheets("Zamówienia").Activate
    'Range("A111:E112").Select
    Range(Cells(IlePap + 2, 1), Cells(IlePap + 3, 5)).Select
    Selection.Copy
    NewSheet.Activate
    Range(Cells(IlePap + 2, 1), Cells(IlePap + 3, 5)).Select
    ActiveSheet.Paste
End If
Next CzyTyt

' Kopiowanie zamówionego Tyt.

For Tyt = IlePap + 4 To IleTyt
If Sheets("Zamówienia").Cells(Tyt, 5) > 0 Then
    Sheets("Zamówienia").Activate
    Range(Cells(Tyt, 1), Cells(Tyt, 4)).Select
    Selection.Copy
    ' Pobranie danych zamiast kopii
    Cells(Tyt, 5).Select
    Kopia = Cells(Tyt, 5)
    ' Wklejanie
    NewSheet.Activate
    Range(Cells(Tyt, 1), Cells(Tyt, 4)).Select 'o tu 5
    ActiveSheet.Paste
    ' Wpisywanie
    Cells(Tyt, 5).Select
    Cells(Tyt, 5) = Kopia
    ' Prawa krawędź tabeli
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    ' Oczyszczenie pola w formularzu
    Sheets("Zamówienia").Activate
    Cells(Tyt, 5) = ""
End If
Next Tyt

'   GILZY
' Test nagłówka zamówienie gilz

For CzyGilzy = IleTyt + 4 To IleGil
If Sheets("Zamówienia").Cells(CzyGilzy, 5) > 0 Then
    Sheets("Zamówienia").Activate
    'Range("A126:E127").Select
    Range(Cells(IleTyt + 2, 1), Cells(IleTyt + 3, 5)).Select
    Selection.Copy
    NewSheet.Activate
    'Range("A126:E127").Select
    Range(Cells(IleTyt + 2, 1), Cells(IleTyt + 3, 5)).Select
    ActiveSheet.Paste
End If
Next CzyGilzy

' Kopiowanie zamówionych gilz

For Gilzy = IleTyt + 4 To IleGil
If Sheets("Zamówienia").Cells(Gilzy, 5) > 0 Then
    Sheets("Zamówienia").Activate
    Range(Cells(Gilzy, 1), Cells(Gilzy, 4)).Select
    Selection.Copy
    ' Pobranie danych zamiast kopii
    Cells(Gilzy, 5).Select
    Kopia = Cells(Gilzy, 5)
    ' Wklejanie
    NewSheet.Activate
    Range(Cells(Gilzy, 1), Cells(Gilzy, 4)).Select 'o tu 5
    ActiveSheet.Paste
    ' Wpisywanie
    Cells(Gilzy, 5).Select
    Cells(Gilzy, 5) = Kopia
    ' Prawa krawędź tabeli
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    ' Oczyszczenie pola w formularzu
    Sheets("Zamówienia").Activate
    Cells(Gilzy, 5) = ""
End If
Next Gilzy

'   PUSTE LINIE
' Usunięcie pustych lini gilz

Do
If NewSheet.Cells(IleGil, 5) = "" Then
    NewSheet.Activate
    Rows(IleGil).Select
    Selection.Delete Shift:=xlUp
    IleGil = IleGil - 1
    Else
    IleGil = IleGil - 1
    End If
Loop Until IleGil = IleTyt + 3


' Usunięcie pustych lini tytoniu
Do
If NewSheet.Cells(IleTyt, 5) = "" Then
    NewSheet.Activate
    Rows(IleTyt).Select
    Selection.Delete Shift:=xlUp
    IleTyt = IleTyt - 1
    Else
    IleTyt = IleTyt - 1
    End If
Loop Until IleTyt = IlePap + 3


' Usunięcie pustych lini papierosów

Do
If NewSheet.Cells(IlePap, 5) = "" Then
    NewSheet.Activate
    Rows(IlePap).Select
    Selection.Delete Shift:=xlUp
    IlePap = IlePap - 1
    Else
    IlePap = IlePap - 1
    End If
Loop Until IlePap = 2

' Rysowanie dołu tabeli
Do
If NewSheet.Cells(Linia, 5) = "" And NewSheet.Cells(Linia - 1, 5) <> "" Then
    NewSheet.Activate
    Range(Cells(Linia - 1, 1), Cells(Linia - 1, 5)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Linia = Linia + 1
    Else
    Linia = Linia + 1
    End If
Loop Until Linia = Gilzy ' zmienna pochodzi z kopiowania gilz

' Okno wydruku

NewSheet.Activate
Application.Dialogs(xlDialogPrint).Show
Sheets("Zamówienia").Activate

' Ochrona karty

NewSheet.Activate
Cells.Locked = True
ActiveSheet.Protect Password:="onomatopeja"
ActiveWorkbook.Protect Password:="onomatopeja", Structure:=True, Windows:=False
       
End Sub
