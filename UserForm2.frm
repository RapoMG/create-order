VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Usuwanie i edycja"
   ClientHeight    =   2832
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6435
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public szuk As Long

Private Sub cmd_Zakoncz_Click() 'zako�cz

Unload UserForm2

End Sub

Public Sub cmd_Znajdz_Click() ' Znajd� +

Dim Jest
For szuk = 4 To Rows.Count
    If Cells(szuk, 2).Value = tbx_KodWew.Text Then
        lbl_KodTowPole = Cells(szuk, 3).Value
        lbl_NazwaTowPole = Cells(szuk, 4).Value
        If Cells(szuk, 1).Value = "" Then
            Me.chk_Dostepny.Value = True
        Else
            Me.chk_Dostepny.Value = False
        End If
        Call Aktyw_Klaw_Gl
        Jest = 1
        Exit For
        
    End If
Next szuk
If Jest <> 1 Then
    MsgBox ("Towaru nie znaleziono.")
    Jest = 0
End If

End Sub

Private Sub cmd_Edytuj_Click() 'Edytuj+

Me.tbx_KodTow.Visible = True
Me.tbx_NazwaTow.Visible = True

Me.chk_Dostepny.Enabled = True
Me.tbx_KodTow.Enabled = True
Me.tbx_NazwaTow.Enabled = True

tbx_KodTow.Text = lbl_KodTowPole
tbx_NazwaTow.Text = lbl_NazwaTowPole

Call Deakt_Klaw_Gl
Call Aktyw_Klaw_Dod

End Sub

Private Sub cmd_Zapisz_Click() ' Zapisz

For dupl = 4 To Rows.Count
    If Cells(dupl, 2).Value = tbx_KodWew.Text And dupl <> szuk Then
        MsgBox ("Towar o takim kodzie ju� jest zapisany")
        Exit Sub
    End If
Next dupl

If szuk > 0 Then
    Cells(szuk, 2).Value = tbx_KodWew.Text
    Cells(szuk, 3).Value = tbx_KodTow.Text
    Cells(szuk, 4).Value = tbx_NazwaTow.Text
Else
    MsgBox ("Wyst�pi� b��d. Zamknij okno formularza i spr�buj ponownie.")
End If

Call Dostepnosc

'Call cmd_Anuluj_Click
'Me.tbx_KodWew.SetFocus

Call cmd_Zakoncz_Click
Worksheets("Zam�wienia").Cells(1, 9).Activate

End Sub

Private Sub cmd_Anuluj_Click() 'anuluj

Me.tbx_KodTow.Enabled = False
Me.tbx_NazwaTow.Enabled = False
Me.chk_Dostepny.Enabled = False

Me.tbx_KodTow.Visible = False
Me.tbx_NazwaTow.Visible = False

Call Deaktyw_Klaw_Dod
Call Aktyw_Klaw_Gl

Me.lbl_KodTowPole = Cells(szuk, 3).Value
Me.lbl_NazwaTowPole = Cells(szuk, 4).Value

End Sub

Private Sub tbx_KodWew_Change() 'kod wew
cmd_Usun.Enabled = False
cmd_Edytuj.Enabled = False
Me.chk_Dostepny.Enabled = False
Me.chk_Dostepny.Value = False
lbl_KodTowPole = ""
lbl_NazwaTowPole = ""
End Sub

Private Sub cmd_Usun_Click() 'Usu� produkt+
If szuk > 0 Then
    Rows(szuk).Select
    Selection.Delete Shift:=xlUp
    MsgBox ("Towar usuni�ty z listy")
Else
    MsgBox ("Usuni�cie niemo�liwe.")
End If

If Cells(szuk, 2) = "" Then
        Range(Cells(szuk - 1, 1), Cells(szuk - 1, 5)).Select
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
End If
Cells(4, 5).Select
Call Pola_Txt_Stop
Call Deakt_Klaw_Gl


End Sub


Sub Test_wywo�ania()
MsgBox ("A jednak si� kr�ci!")
End Sub

''Procedury wsp�dzielone


Sub Aktyw_Klaw_Gl() ' cz w��

Me.cmd_Edytuj.Enabled = True
Me.cmd_Usun.Enabled = True
Me.cmd_Znajdz.Enabled = True

Me.cmd_Edytuj.Visible = True
Me.cmd_Usun.Visible = True

End Sub

Sub Pola_Txt_Stop() ' cz linia

Me.tbx_KodTow.Enabled = False
Me.tbx_NazwaTow.Enabled = False
Me.chk_Dostepny.Enabled = False

Me.tbx_KodTow.Visible = False
Me.tbx_NazwaTow.Visible = False

Me.tbx_KodWew = ""
Me.tbx_KodTow = ""
Me.tbx_NazwaTow = ""
Me.lbl_KodTowPole = ""
Me.lbl_NazwaTowPole = ""
Me.chk_Dostepny.Value = False

End Sub

Sub Deaktyw_Klaw_Dod() 'ziel w��

Me.cmd_Zapisz.Enabled = False
Me.cmd_Anuluj.Enabled = False

Me.cmd_Zapisz.Visible = False
Me.cmd_Anuluj.Visible = False

End Sub

Sub Deakt_Klaw_Gl() ' ziel lin

Me.cmd_Edytuj.Enabled = False
Me.cmd_Usun.Enabled = False
Me.cmd_Znajdz.Enabled = False

Me.cmd_Edytuj.Visible = False
Me.cmd_Usun.Visible = False

End Sub

Sub Aktyw_Klaw_Dod()
Me.cmd_Zapisz.Enabled = True
Me.cmd_Anuluj.Enabled = True

Me.cmd_Zapisz.Visible = True
Me.cmd_Anuluj.Visible = True
End Sub

Sub Dostepnosc()
If chk_Dostepny.Value = False Then
    Worksheets("Zam�wienia").Cells(szuk, 1).Value = "N.D."
    Worksheets("Zam�wienia").Cells(szuk, 5).Select
       With Selection.Borders(xlDiagonalDown)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
       End With
Else
    Worksheets("Zam�wienia").Cells(szuk, 1).Value = ""
    Worksheets("Zam�wienia").Cells(szuk, 5).Borders(xlDiagonalDown).LineStyle = xlNone


End If

End Sub
Sub ochr_wl()
Worksheets("Zam�wienia").Activate
Cells.Locked = True
ActiveSheet.Protect Password:="asdf"
ActiveWorkbook.Protect Password:="asdf", Structure:=True, Windows:=False
End Sub

Sub ochr_wyl()
Worksheets("Zam�wienia").Activate

ActiveSheet.Protect Password:="asdf"

ActiveWorkbook.Protect Password:="asdf", Structure:=True, Windows:=False
'Worksheets("Zam�wienia").Cells.Locked = False
End Sub
