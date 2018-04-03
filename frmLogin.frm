VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "LOGIN"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
If TextBox1 = "admin" And TextBox2 = "admin" Then
    Me.Hide
   pilihan.Show
   Unload Me '.Visible = False 'Menyembunyikan Form 1
  'Menutup Form 1
  Else
  MsgBox "User Name atau Password yang Anda Masukkan salah" & vbNewLine & "Silahkan Coba lagi !!", vbCritical, "Warning!!"
  TextBox1 = " "
  TextBox2 = " "
  TextBox1.SetFocus
  End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub
