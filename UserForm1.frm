VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Dodaj produkt"
   ClientHeight    =   3465
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5580
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload UserForm1
End Sub

Private Sub CommandButton2_Click()

Dim Wiersz ' Zmienna wype�nienia w�a�ciwego wiersza
Dim NaLiczbe As Double ' Aby kody towaru zapisa�o jako liczb�


' Spr. czy wype�niono pole kodu
If TextBox1.Text = "" Then
        MsgBox ("Brak wewn�trznego kodu towaru.")
        Exit Sub
End If

' Spr. czy wype�niono pole nazwy
If TextBox3.Text = "" Then
        MsgBox ("Brak nazwy towaru.")
        Exit Sub
End If

' Spr. czy wype�niono pole kodu
'If TextBox2.text = "" Then
'        MsgBox, kt�ry zapyta czy kontynuowa� mimo braku kodu
'
'End If

' Sprawdzenie duplikatu w bazie
For dupl = 4 To Rows.Count
    If Cells(dupl, 2).Value = TextBox1.Text Then
        MsgBox ("Produkt o takim kodzie ju� jest wprowadzony.")
        Exit Sub
    End If
Next dupl

' Licznik_wierszy

For IlePap = 4 To Rows.Count
    If Cells(IlePap, 2) = 0 Then
        Exit For
    End If
Next IlePap

For IleTyt = IlePap + 3 To Rows.Count
    If Cells(IleTyt, 2) = 0 Then
        Exit For
    End If
Next IleTyt

For IleGil = IleTyt + 3 To Rows.Count
    If Cells(IleGil, 2) = 0 Then
        Exit For
    End If
Next IleGil

' Wyb�r kategorii

If OptionButton2 = True Then
    Wiersz = IleTyt
ElseIf OptionButton3 = True Then
    Wiersz = IleGil
Else
    Wiersz = IlePap
End If



' Usuni�cie dolnej ramki wiersza wybranej kategorii

Rows(Wiersz - 1).Select
Selection.Borders(xlEdgeBottom).LineStyle = xlNone

' Wstawianie wiersza

Rows(Wiersz).Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
NaLiczbe = TextBox1.Text
Worksheets("Zam�wienia").Cells(Wiersz, 2).Value = NaLiczbe 'TextBox1
NaLiczbe = TextBox2.Text
Worksheets("Zam�wienia").Cells(Wiersz, 3).Value = NaLiczbe 'TextBox2
Worksheets("Zam�wienia").Cells(Wiersz, 4).Value = TextBox3.Text

' Formatowanie dolnej linii tabeli
Worksheets("Zam�wienia").Range(Cells(Wiersz, 1), Cells(Wiersz, 5)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

' Formatowanie prawej linii tabeli
For Kolumna = 1 To 5
Worksheets("Zam�wienia").Cells(Wiersz, Kolumna).Select
With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
End With
Next Kolumna
    
' Formatowanie lewej linii tabeli
Worksheets("Zam�wienia").Cells(Wiersz, 1).Select
With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
End With


Unload UserForm1

End Sub


Private Sub TextBox1_Change()
If IsNumeric(TextBox1.Value) = False Then
    MsgBox ("Kod wewn�trzny towaru musi by� liczb�.")

End If
End Sub

Private Sub TextBox2_Change()
If IsNumeric(TextBox2.Value) = False Then
MsgBox ("Kod kreskowy towaru musi by� liczb�.")
End If
End Sub

Sub dziala()
MsgBox ("Dzia�a!")
End Sub
