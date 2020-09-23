Attribute VB_Name = "Mod_WebCam"
Option Explicit
Public Const WM_CAP_DRIVER_CONNECT As Long = 1034
Public Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Public Const WM_CAP_GRAB_FRAME As Long = 1084
Public Const WM_CAP_EDIT_COPY As Long = 1054
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Public Const WM_CAP_DLG_VIDEODISPLAY As Long = 1067
Public Const WM_CAP_GET_STATUS As Long = 1078
Public Const WM_CLOSE = &H10

Public Declare Function C2P Lib "WickPic.dll" (ByVal B_Bestand As String, ByVal J_Bestand As String, ByVal K As Long) As Boolean

Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public WHwnd As Long


'WebCam Functions

'loads the webcam
Public Sub StartWebCam()
On Error GoTo Fout
    WHwnd = capCreateCaptureWindow("SaveCam", 1, 0, 0, Frm_main.WebCam.Width, Frm_main.WebCam.Height, 0, 0)
    SendMessage WHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
    Exit Sub
Fout:
    MsgBox "SaveCam v1.0" & vbCrLf * vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Description : " & Err.Description
End Sub

'unload the webcam
Public Sub StopWebCam()
    SendMessage WHwnd, WM_CLOSE, 0, 0
    SendMessage WHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub

'See frm_main for these function. (look at the buttons))
Public Function WebCam_Bron() As Boolean
    WebCam_Bron = SendMessage(WHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
End Function

Function WebCam_ForMaat() As Boolean
   WebCam_ForMaat = SendMessage(WHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
End Function

Function WebCam_WeerGave() As Boolean
    WebCam_WeerGave = SendMessage(WHwnd, WM_CAP_DLG_VIDEODISPLAY, 0, 0)
End Function

Public Function DelData()
'delete the server files...
'the server file will be made every time a user starts the server...
'this function is to prevent users to edit the server files...

'close all file to be sure...
Close
'check if the map exist...
CheckMap "Data"
Kill App.Path & "\Data\*.*"
RmDir App.Path & "\Data"
End Function

'make the bmpfile
Public Function Maak_BMP(BMP_BestandsNaam As String)
On Error GoTo Fout
    Maak_BMP = 0
        SendMessage WHwnd, WM_CAP_GRAB_FRAME, 0, 0
        SendMessage WHwnd, WM_CAP_EDIT_COPY, 0, 0
        
        SavePicture Clipboard.GetData, BMP_BestandsNaam
        Clipboard.Clear
        Exit Function
Fout:
    Maak_BMP = 1
End Function

'convert the bmp to a jpg file
Public Function Maak_JPG(Kwa As Integer)
'On Error GoTo Fout
If Kwa > 100 Or Kwa < 1 Then Kwa = 50
Maak_BMP App.Path & "\Data\Webcam.bmp"
C2P App.Path & "\Data\Webcam.bmp", App.Path & "\Data\Webcam.jpg", Kwa
End Function

