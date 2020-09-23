Attribute VB_Name = "Mod_XButton"
'This code i found on the internet...
'changed only some long names...
'it's to disable the X button on a form....
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
     
Private Const POS = &H400&

Public Function WisX(formName As Form) As Boolean
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
        lHndSysMenu = GetSystemMenu(formName.hwnd, 0)
        'remove the X
        lAns1 = RemoveMenu(lHndSysMenu, 6, POS)
        'remove seperator bar
        lAns2 = RemoveMenu(lHndSysMenu, 5, POS)
        'Return True if both calls were successful
        WisX = (lAns1 <> 0 And lAns2 <> 0)
End Function
