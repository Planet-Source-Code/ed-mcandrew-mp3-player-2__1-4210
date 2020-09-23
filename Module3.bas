Attribute VB_Name = "Module3"
Public Const MF_BYPOSITION = &H400
Public Const MF_REMOVE = &H1000
Public Const MF_ENABLED = &H0&
Public Const MF_INSERT = &H0&


Public Declare Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, _
ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long) As Long





