VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
barang = Me.ComboBox1.Value
With Worksheets("Sheet2").Range("b4:b999")
Set kuro = .Find(barang, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Me.TextBox1.Value = Worksheets("Sheet2").Cells(baris, 2).Value
Me.TextBox2.Value = Worksheets("Sheet2").Cells(baris, 3).Value
Me.TextBox3.Value = Worksheets("Sheet2").Cells(baris, 4).Value
Me.TextBox4.Value = Worksheets("Sheet2").Cells(baris, 5).Value
Me.TextBox5.Value = Worksheets("Sheet2").Cells(baris, 6).Value
Me.TextBox6.Value = Worksheets("Sheet2").Cells(baris, 7).Value

Else
MsgBox "Maaf barang yang anda yang butuhkan tidak tersedia."
End If
End With
End Sub

Private Sub CommandButton2_Click()
barang = Me.ComboBox1.Value
With Worksheets("Sheet2").Range("b4:b999")
Set kuro = .Find(barang, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Me.TextBox1.Value = Worksheets("Sheet2").Cells(baris, 2).Value
Me.TextBox2.Value = Worksheets("Sheet2").Cells(baris, 3).Value
Me.TextBox3.Value = Worksheets("Sheet2").Cells(baris, 4).Value
Me.TextBox4.Value = Worksheets("Sheet2").Cells(baris, 5).Value
Me.TextBox5.Value = Worksheets("Sheet2").Cells(baris, 6).Value
Me.TextBox6.Value = Worksheets("Sheet2").Cells(baris, 7).Value

Else
MsgBox "Maaf barang yang anda yang butuhkan tidak tersedia."
End If
End With
End Sub
Private Sub CommandButton3_Click()
nama = ComboBox1.Value
With Worksheets("Sheet2").Range("B4:B999")
Set kuro = .Find(nama, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Worksheets("Sheet2").Cells(baris, 2).EntireRow.Delete
TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""
TextBox6.Value = ""
End If
End With
End Sub

Private Sub CommandButton4_Click()
Me.Hide
pilihan.Show
End Sub

Private Sub CommandButton5_Click()
End
End Sub
