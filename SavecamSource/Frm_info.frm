VERSION 5.00
Begin VB.Form Frm_info 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WickChat Server, Informatie"
   ClientHeight    =   2565
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4845
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770.409
   ScaleMode       =   0  'User
   ScaleWidth      =   4549.705
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Dodo2479 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3480
      Picture         =   "Frm_info.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox Logo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "Frm_info.frx":141D
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Text 
      Caption         =   "By : Dodo2479@hotmail.com"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Text 
      Caption         =   "Version : 1.00"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Text 
      Caption         =   "SaveCam 1.0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4507.448
      Y1              =   1408.045
      Y2              =   1408.045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4507.448
      Y1              =   1408.045
      Y2              =   1408.045
   End
End
Attribute VB_Name = "Frm_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A little information about Me..lol
Private Sub cmdOK_Click()
ST_ANI Me, True
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = Frm_main.Icon
Logo.Picture = Frm_main.Icon
End Sub
