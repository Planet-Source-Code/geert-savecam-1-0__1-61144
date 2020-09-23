VERSION 5.00
Begin VB.Form Frm_Systray 
   Caption         =   "SaveCam 1.0"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "nobody can see this"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "SaveCam 1.0 Systray..."
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Frm_Systray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Bericht As Long
    'What happend...
    Bericht = X / Screen.TwipsPerPixelX
    
    'Make sure the caption of the menu is pointing his function...
    If Frm_main.Zichtbaar = True Then
        Frm_main.mnu_tray.Caption = "Hide"
    Else
        Frm_main.mnu_tray.Caption = "View"
    End If
    
    Select Case Bericht
        'If mouse is right clicked then popup the menu
      Case WM_RBUTTONUP
        PopupMenu Frm_main.mnu_main
        'or if mouse is double clicked, hide/show the form
      Case WM_LBUTTONDBLCLK
        If Frm_main.Zichtbaar = False Then
            'do a little annimation
            ST_ANI Frm_main, False
            'change the menu and show the form
            With Frm_main
                .mnu_tray.Caption = "Verbergen"
                .Visible = True
                .Zichtbaar = True
            End With
        Else
            'do a litte animation
            ST_ANI Frm_main, True
            'change the menu and hide the form
            With Frm_main
                .mnu_tray.Caption = "Openen"
                .Visible = False
                .Zichtbaar = False
            End With
         End If
        
    End Select
End Sub
