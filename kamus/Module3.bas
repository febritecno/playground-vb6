Attribute VB_Name = "Module3"
Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&


Public Sub RemoveCancelMenuItem(frm As Form)
Dim hSysMenu As Long
  'Ambil menu system untuk form ini
  hSysMenu = GetSystemMenu(frm.hWnd, 0)
  'Hilangkan tombol Close (X)
  Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
  'Hilangkan pemisah yang melalui tombol Close tersebut
  Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

'Walaupun tombol "Close" di pojok kanan atas form tidak 'dapat diklik karena sudah disabled, Anda masih bisa 'menutup form dengan menggunakan tombol Alt-F4. Agar 'form juga tidak dapat ditutup dengan menggunakan
'Alt-'F4, Anda harus menahannya di event procedure 'Form_QueryUnload dengan meng-assignment nilai 'parameter Cancel = -1.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = -1 'Jadi, Alt-F4 juga tidak berfungsi!
End Sub



