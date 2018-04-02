VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub sheets1()
frmLogin.Show
End Sub
Private Sub CommandButton2_Click()
Set kuro = sheets("sheet1")
If kuro.FilterMode Then
    kuro.ShowAllData
End If
End Sub

Private Sub CommandButton3_Click()
pilihan.Show
End Sub

Private Sub CommandButton1_Click()
frmLogin.Show
End Sub
