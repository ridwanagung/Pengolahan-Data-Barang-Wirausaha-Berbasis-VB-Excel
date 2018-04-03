VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} inputstokbrg 
   Caption         =   "UserForm8"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   OleObjectBlob   =   "inputstokbrg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "inputstokbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheet2.Activate
Range("a10000").Select
ActiveCell.End(xlUp).Offset(1, 0).Select
ActiveCell.Value = "=ROW()-5"
ActiveCell.Offset(0, 1).Value = TextBox1.Text
ActiveCell.Offset(0, 2).Value = TextBox2.Text
ActiveCell.Offset(0, 3).Value = TextBox3.Text
ActiveCell.Offset(0, 4).Value = TextBox4.Text
ActiveCell.Offset(0, 5).Value = TextBox5.Text
Dim Pemasukan, Pengeluaran, Sisa As Single
Pemasukan = TextBox4.Text
Pengeluaran = TextBox5.Text
Sisa = Pemasukan - Pengeluaran
ActiveCell.Offset(0, 6).Value = Sisa
Nama_Barang = TextBox1.Text
Tanggal = TextBox2.Text
Satuan = TextBox3.Text
Pemasukan = TextBox4.Text
Pengeluaran = TextBox5.Text
End Sub

Private Sub CommandButton2_Click()
TextBox1.Text = ""
TextBox2.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
End Sub


Private Sub CommandButton3_Click()
End
End Sub

Private Sub CommandButton4_Click()
Me.Hide
pilihan.Show
End Sub
