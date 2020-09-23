VERSION 5.00
Begin VB.Form Frm_snap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaveCam 1.0, Snapshot"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Op_J 
      Caption         =   "Jpg File"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton Op_B 
      Caption         =   "Bmp File"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Close"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_new 
      Caption         =   "New"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image CamSnap 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Frm_snap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Cancel_Click()
ST_ANI Me, True
Unload Me
End Sub

Private Sub Cmd_new_Click()
'display a new picture of the webcam
    CamSnap.Picture = Frm_main.WebCam.Picture
End Sub

Private Sub Cmd_Save_Click()
'let's save the picture file
StartSave:
Dim FileName As String
'aks the user for a name
FileName = InputBox("Please give for the snapshot." & vbCrLf & vbCrLf & "Only the first name please, No extencion", "Save SnapShot", "SnapShot_from_" & Format(Now, "DD_MM_YYYY"))
'check if user told us a name to save to...
    If FileName <> "" Then
            GoTo SaveFile
    Else
            MsgBox "You didn't give me a name to save your picture...", vbInformation, "SaveCam 1.0, Save SnapShot"
        Exit Sub
    End If

SaveFile:
'check if we need to save if to a jpg file...
If Op_J.Value = True Then
    'check if jpg file exists
    If CheckFile(App.Path & "/SnapShot/" & FileName & ".jpg") = False Then
        GoTo JpgFile
    Else
        MsgBox "The file " & filenam & ".jpg already exists..." & vbCrLf & "Please give up a new name...", vbInformation, "SaveCam 1.0, Save SnapShot"
        GoTo StartSave
        Exit Sub
    End If
Else
  'no we didn't have to save it to a jpg file...
  'so we gona save it as bmp file
    'check if bmp file exists
    If CheckFile(App.Path & "/SnapShot/" & FileName & ".bmp") = False Then
        GoTo BmpFile
    Else
        MsgBox "The file " & filenam & ".bmp already exists..." & vbCrLf & "Please give up a new name...", vbInformation, "SaveCam 1.0, Save SnapShot"
        GoTo StartSave
        Exit Sub
    End If
End If

JpgFile: 'to save the picture as jpg....
    Dim TmpBmp As String, SaveJPG As String
        'set the files you need
        TmpBmp = App.Path & "/SnapShot/Tmpfile.bmp"

        'make firts a bmp file
        SavePicture CamSnap.Picture, TmpBmp

        'set filename
        SaveJPG = App.Path & "/Snapshot/" & FileName & ".jpg"

        'save the file to a jpg
        C2P TmpBmp, SaveJPG, Frm_main.Quality.Value

        'tell the user the snapshot has been saved...
        MsgBox "Your SnapShot has been saved to: " & SaveJPG

        'kill the tmp file
        Kill TmpBmp

    Exit Sub

BmpFile: 'to save the picture a bmp file
    Dim SaveBMP As String
        'set filename to save to
        SaveBMP = App.Path & "/SnapShot/" & FileName & ".bmp"
        
        'save the bmpfile
        SavePicture CamSnap.Picture, SaveBMP
        
        'tell the user the snapshot has been saved...
        MsgBox "Your SnapShot has been saved to: " & SaveBMP
    Exit Sub
End Sub

Private Sub Form_Load()
Me.Icon = Frm_main.Icon
End Sub
