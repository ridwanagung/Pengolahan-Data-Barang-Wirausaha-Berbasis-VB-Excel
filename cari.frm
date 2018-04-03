VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cari 
   Caption         =   "UserForm7"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7995
   OleObjectBlob   =   "cari.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub
Private Sub CommandButton1_Click()
barang = Me.ComboBox1.Value
With Worksheets("LOGIN").Range("b6:b999")
Set kuro = .Find(barang, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Me.TextBox1.Value = Worksheets("LOGIN").Cells(baris, 2).Value
Me.TextBox2.Value = Worksheets("LOGIN").Cells(baris, 5).Value
Me.TextBox3.Value = Worksheets("LOGIN").Cells(baris, 7).Value
Me.TextBox4.Value = Worksheets("LOGIN").Cells(baris, 6).Value

Else
MsgBox "Maaf barang yang anda butuhkan tidak tersedia."
End If
End With
    
End Sub

Private Sub CommandButton2_Click()
Me.Hide
pilihan.Show
End Sub

Private Sub TextBox2_Change()

End Sub
