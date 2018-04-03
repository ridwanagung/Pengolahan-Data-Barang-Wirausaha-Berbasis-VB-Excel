VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} inputdatabarang 
   Caption         =   "UserForm3"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11010
   OleObjectBlob   =   "inputdatabarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "inputdatabarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheet1.Activate
Range("a10000").Select
ActiveCell.End(xlUp).Offset(1, 0).Select
ActiveCell.Value = "=ROW()-5"
ActiveCell.Offset(0, 1).Value = TextBox1.Text
ActiveCell.Offset(0, 2).Value = TextBox2.Text
ActiveCell.Offset(0, 3).Value = TextBox3.Text
ActiveCell.Offset(0, 4).Value = TextBox4.Text
ActiveCell.Offset(0, 5).Value = TextBox5.Text
ActiveCell.Offset(0, 6).Value = TextBox6.Text
ActiveCell.Offset(0, 7).Value = TextBox7.Text
Nama_Barang = TextBox1.Text
Kode_Barang = TextBox2.Text
Harga_Beli = TextBox3.Text
Harga_jual = TextBox4.Text
Jenis_Barang = TextBox5.Text
Tanggal_Kadaluarsa = TextBox6.Text
Jumlah_Barang = TextBox7.Text

End Sub

Private Sub CommandButton2_Click()
TextBox1.Text = ""
TextBox2.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox6.Text = ""
TextBox7.Text = ""
End Sub

Private Sub CommandButton3_Click()
Me.Hide
pilihan.Show
End Sub

Private Sub CommandButton4_Click()
End
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox1_Change()

End Sub
