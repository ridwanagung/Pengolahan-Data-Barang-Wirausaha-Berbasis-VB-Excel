VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} carinamabarang 
   Caption         =   "UserForm5"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   OleObjectBlob   =   "carinamabarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "carinamabarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ScrollBar1_Change()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()
barang = Me.ComboBox1.Value
With Worksheets("LOGIN").Range("b6:b999")
Set kuro = .Find(barang, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Me.TextBox1.Value = Worksheets("LOGIN").Cells(baris, 2).Value
Me.TextBox2.Value = Worksheets("LOGIN").Cells(baris, 3).Value
Me.TextBox3.Value = Worksheets("LOGIN").Cells(baris, 4).Value
Me.TextBox4.Value = Worksheets("LOGIN").Cells(baris, 5).Value
Me.TextBox5.Value = Worksheets("LOGIN").Cells(baris, 6).Value
Me.TextBox6.Value = Worksheets("LOGIN").Cells(baris, 7).Value

Else
MsgBox "Maaf barang yang anda butuhkan tidak tersedia."
End If
End With
    
    
End Sub


Private Sub CommandButton3_Click()
Me.Hide
pilihan.Show

End Sub

Private Sub CommandButton4_Click()
End
End Sub

Private Sub CommandButton5_Click()
nama = ComboBox1.Value
With Worksheets("LOGIN").Range("B6:B999")
Set kuro = .Find(nama, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Worksheets("LOGIN").Cells(baris, 2).Value = ComboBox1.Value
Worksheets("LOGIN").Cells(baris, 2).Value = TextBox1.Value
Worksheets("LOGIN").Cells(baris, 3).Value = TextBox2.Value
Worksheets("LOGIN").Cells(baris, 4).Value = TextBox3.Value
Worksheets("LOGIN").Cells(baris, 5).Value = TextBox4.Value
Worksheets("LOGIN").Cells(baris, 6).Value = TextBox5.Value
Worksheets("LOGIN").Cells(baris, 7).Value = TextBox6.Value
End If
End With
TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""
TextBox6.Value = ""
ComboBox1.Value = ""

TextBox1.SetFocus
End Sub
Private Sub CommandButton6_Click()
nama = ComboBox1.Value
With Worksheets("LOGIN").Range("B6:B999")
Set kuro = .Find(nama, LookIn:=xlValues)
If Not kuro Is Nothing Then
baris = kuro.Row
Worksheets("LOGIN").Cells(baris, 2).EntireRow.Delete
TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""
TextBox6.Value = ""
End If
End With
End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub
