VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaveCam 1.0"
   ClientHeight    =   2715
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3855
   Icon            =   "Frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox iplist 
      Height          =   1815
      Left            =   7800
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_settings 
      Caption         =   "Settings"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin MSWinsockLib.Winsock WSServer 
      Index           =   0
      Left            =   2520
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSListen 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_snap 
      Caption         =   "SnapShot"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Timer CamTimer 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2040
      Top             =   960
   End
   Begin MSComctlLib.Slider Quality 
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Min             =   1
      Max             =   100
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.CommandButton Cmd_StartStop 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_display 
      Caption         =   "Display"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmd_format 
      Caption         =   "Format"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmd_source 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label viewers 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label txt_lbl 
      Caption         =   "Viewers:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Lbl 
      Caption         =   "WebCam Quality :"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image WebCam 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnu_main 
      Caption         =   "File"
      Begin VB.Menu mnu_tray 
         Caption         =   "View"
      End
      Begin VB.Menu line0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_info 
         Caption         =   "Info"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Zichtbaar is set boolean to tell the
'program it visible of not (true/false)
'see MOD_START (Sub Main)
Public Zichtbaar As Boolean
Public client As Integer


Private Sub CamTimer_Timer()
'take a picture and display it...
Maak_JPG Quality.Value
WebCam.Picture = LoadPicture(App.Path & "\Data\Webcam.jpg")
End Sub

Private Sub Cmd_display_Click()
'if clicked the show the display setting of the webcam.
'if this function isn't suported then
'tell the user
If WebCam_WeerGave = False Then MsgBox "SaveCam 1.0" & vbCrLf & "This Function isn't supported by your webcam.", vbInformation, "SaveCam 1.0"

End Sub

Private Sub cmd_format_Click()
'if clicked the show the format setting of the webcam.
'if this function isn't suported then
'tell the user
If WebCam_ForMaat = False Then MsgBox "SaveCam 1.0" & vbCrLf & "This Function isn't supported by your webcam.", vbInformation, "SaveCam 1.0"

End Sub

Private Sub Cmd_settings_Click()
'to view/edit the settings of Savecam
ST_ANI Frm_Settings, False
Frm_Settings.Visible = True
End Sub

Private Sub Cmd_snap_Click()
'Take a snapshot
If Frm_snap.Visible = True Then
Exit Sub
Else
Frm_snap.CamSnap.Picture = WebCam.Picture
ST_ANI Frm_snap, False
Frm_snap.Visible = True
End If
End Sub

Private Sub cmd_source_Click()
'if clicked the show the source setting of the webcam.
'if this function isn't suported then
'tell the user
If WebCam_Bron = False Then MsgBox "SaveCam 1.0" & vbCrLf & "This Function isn't supported by your webcam.", vbInformation, "SaveCam 1.0"

End Sub

Private Sub Cmd_StartStop_Click()
'This is to enable/disable the timer
'look at the timer
Select Case Cmd_StartStop.Caption
    Case Is = "Start"
        Cmd_StartStop.Caption = "Stop"
        CamTimer.Enabled = True
        'enable the snapshot button
        Cmd_snap.Enabled = True
        'disable the settings button
        Cmd_settings.Enabled = False
        'make the html page for the viewers to see
        MakeHtml
        'just in case
        WSListen.Close
        'set port to listen to
        WSListen.LocalPort = Frm_Settings.CB_Port.Text
        'listen for connecttions
        WSListen.Listen
        
    Case Is = "Stop"
        'set correct caption
        Cmd_StartStop.Caption = "Start"
        'disable the timer
        CamTimer.Enabled = False
        'display the logo
        WebCam.Picture = Frm_main.Icon
        'disable the snapshot button
        Cmd_snap.Enabled = False
        'enable the settings button
        Cmd_settings.Enabled = True
        'close winsock
        WSListen.Close
        'set viewers to zero
        viewers.Caption = "0"
End Select
End Sub

Private Sub mnu_exit_Click()
'close the app
'look at Mod_start (Function close_app)
Close_App
End Sub

Private Sub mnu_info_Click()
'show a handsom man..lol
ST_ANI Frm_info, False
Frm_info.Show
End Sub

Private Sub mnu_tray_Click()
'here we select what to do...
'show the server or hide it...
'and we change the caption...
Zichtbaar = Not (Zichtbaar)
    If Zichtbaar Then
        mnu_tray.Caption = "Hide"
        ST_ANI Me, False
        Frm_main.Visible = True
    Else
        mnu_tray.Caption = "View"
        ST_ANI Me, True
        Frm_main.Visible = False
    End If
End Sub

Private Sub WSListen_ConnectionRequest(ByVal requestID As Long)
    Load WSServer(WSServer.UBound + 1)
    WSServer(WSServer.UBound).Accept requestID
If CheckIP(WSListen.RemoteHostIP) = False Then
    viewers.Caption = viewers.Caption + 1
    iplist.AddItem WSListen.RemoteHostIP
End If
End Sub


Private Sub WSServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim tempstr As String
'get data
WSServer(Index).GetData tempstr
DoEvents

Dim pos1, POS As Long, pos2
Dim file_req, file_path, data As String
Dim FF As Long

'extract name of file
pos1 = InStr(1, tempstr, "/", vbTextCompare)
pos2 = InStr(1, tempstr, "HTTP/", vbTextCompare)
file_req = Mid(tempstr, pos1 + 1, (pos2 - pos1) - 1)

FF = FreeFile

'if no filename, set the filename...
If file_req = " " Then
    file_req = "Savecam.dat"
End If

'set file path
file_path = App.Path & "\Data\" & file_req

'read file to temperary string
Open file_path For Binary As #FF
    data = Input(FileLen(file_path), #FF)
Close #FF

'Send the file
WSServer(Index).SendData (data)
DoEvents

'unload teh winsock
Unload WSServer(Index)

End Sub
