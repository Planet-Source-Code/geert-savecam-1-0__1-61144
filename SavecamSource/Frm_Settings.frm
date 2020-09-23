VERSION 5.00
Begin VB.Form Frm_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaveCam 1.0, Settings..."
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CB_Time 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox Logo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   120
      Picture         =   "Frm_Settings.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   29
      Top             =   5280
      Width           =   480
   End
   Begin VB.CommandButton BGCOLOR 
      Caption         =   "Color"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton FONTCOLOR 
      Caption         =   "Color"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5160
      Width           =   975
   End
   Begin VB.CheckBox CB_C2 
      Caption         =   "Display my e-mail adres"
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CheckBox CB_C1 
      Caption         =   "Display my website adres"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox TMPdata 
      Height          =   1455
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "Frm_Settings.frx":0E42
      Top             =   7080
      Width           =   5655
   End
   Begin VB.ComboBox CB_Year 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox CB_Month 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox CB_Day 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox CB_County 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox CB_Email 
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox CB_Homepage 
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox CB_Port 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      MaxLength       =   4
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox CB_Title 
      Height          =   285
      Left            =   840
      MaxLength       =   25
      TabIndex        =   9
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox CB_Name 
      Height          =   285
      Left            =   840
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label txt_label 
      Caption         =   "seconds"
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   32
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label txt_label 
      Caption         =   "Interval :"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   30
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label txt_label 
      Caption         =   "HTML Font Color"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   28
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label txt_label 
      Caption         =   "HTML Background color"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   27
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label txt_label 
      Caption         =   "SaveCam 1.0 Settings..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label txt_label 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label txt_label 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1800
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label txt_label 
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   960
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label txt_label 
      Caption         =   "E-Mail :"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label txt_label 
      Caption         =   "Website :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label txt_label 
      Caption         =   "Country :"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label txt_label 
      Caption         =   "Birthdate :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label txt_label 
      Caption         =   "Name :"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label txt_label 
      Caption         =   "Port :"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label txt_label 
      Caption         =   "Title :"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Frm_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB_City_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
          Case 65 To 90
          'are ok to type
          Case 97 To 122
          'are ok to type
          Case 32
          Case 8
          
          Case Else
          'are not oke
               KeyAscii = 0
     End Select
End Sub

Private Sub BGCOLOR_Click()
'select the background color of the html file
Frm_main.CD.ShowColor
BGCOLOR.BackColor = Frm_main.CD.color
End Sub

Private Sub CB_Email_Change()
Select Case KeyAscii
          Case 32 To 122
          'Are ok to type
          Case 8
          'ok to type
          Case Else
          'are not oke
               KeyAscii = 0
     End Select
End Sub

Private Sub CB_Name_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
          Case 65 To 90
          'are ok to type
          Case 97 To 122
          'are ok to type
          Case 32
          Case 8
          
          Case Else
          'are not oke
               KeyAscii = 0
     End Select
End Sub

Private Sub CB_Port_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 48 To 57
    Case Else
       KeyAscii = 0
End Select
End Sub

'the general settings for the server app.
Private Sub Cmd_Cancel_Click()
'hide this form without saving the settings
ST_ANI Me, True
Me.Hide
End Sub

Private Sub Cmd_Save_Click()
'save the settings to a file
SaveSettings
End Sub

Private Sub FONTCOLOR_Click()
'select the font color of the html file
Frm_main.CD.ShowColor
FONTCOLOR.BackColor = Frm_main.CD.color
End Sub

Private Sub Form_Load()
    Me.Icon = Frm_main.Icon
End Sub
