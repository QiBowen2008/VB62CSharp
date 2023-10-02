Attribute VB_Name = "SaveFileDialog"
Public Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long
Public Type BROWSEINFO
hOwner As Long
lpfn As Long
lParam As Long
iImage As Long
lpszTitle As String
ulFlags As Long
pszDisplayName As String
pidlRoot As Long
End Type
