Attribute VB_Name = "Mod_Start"
'This is where the App starts...
'the main module for savecam....

Sub Main()
'This is where the app. first starts...
'0> Check if the map Snapshot and Data exists.. (if not the make them....)
CheckMap "SnapShot"
CheckMap "Data"

'1> Check/Load the settings...
'   if the settings file dosn't exists the make new settings...
LoadFSet

'2> Add the program to the systray
ST_ON Frm_Systray, "SaveCam 1.0 (testversion)", Frm_main.Icon

'3> Is form_main viseble or not at the start ?
Frm_main.Zichtbaar = True  'frm_main = visible

'4> Set the menu caption to the correct caption
Frm_main.mnu_tray.Caption = "Hide"

'5> Do some annimation
ST_ANI Frm_main, False

'6> Show the main form
Frm_main.Visible = True

'6> Load the webcam
StartWebCam

'8> Disable the X button for all forms
WisX Frm_main
WisX Frm_snap
WisX Frm_info
WisX Frm_Settings

'9> Display a icon in place of a black webcam picture
Frm_main.WebCam.Picture = Frm_main.Icon

End Sub

Public Function Close_App()
'This function is to close the app. correct

'0> Delete the data map.. to prevent users to edit the serverdata....
DelData

'1> remove the icon from the systray
ST_OFF

'2> unload the webcam
StopWebCam

'3> be sure the icon is gone
    ST_OFF
    
'4> unload the forms
Unload Frm_info
Unload Frm_main
Unload Frm_Settings
Unload Frm_snap
Unload Frm_Systray

'5> and check the icon again
    ST_OFF
    
'6> close the app for sure
End
End Function

'This function is to check if a Map exist
'use : CheckMap "MAP"
'checks if the map app.path/MAP exists
Function CheckMap(Map As String)
Dim DMap, CMap
'check if Map exist...
DMap = Dir(App.Path & "\" & Map, vbDirectory)
'if not exist then make the Map
    If DMap = "" Then
        MkDir App.Path & "\" & Map
    Exit Function
'If Map exist then check if it's realy a directory
   Else

'if it's a directory the exit
CMap = GetAttr(App.Path & "\" & Map)
      If CMap = vbDirectory Then
      Exit Function

'else kill the "Map" and make it a directory
      Else
        'kill the file
        Kill App.Path & "\" & Map
        'make the map
        MkDir App.Path & "\" & Map
      End If
   End If
End Function

'this function is to check if a file exist. (use : CheckFile app.path & "/Map/Filename.ext")
Public Function CheckFile(FuLLFilePath As String) As Boolean
'if filename exists then checkfile = true, else checkfile = false
    If Dir(FuLLFilePath) = "" Then
        CheckFile = False
    Else
        CheckFile = True
    End If
End Function

Public Function CheckMail(MAIL As String) As Boolean
'check if the e-mail adres is valid
If InStr(1, MAIL, "@") = 0 Or InStr(1, MAIL, ".") = 0 Or Len(MAIL) < 7 Then
    CheckMail = False
Else
    CheckMail = True
End If
End Function

Public Function LoadFSet()
'Here we're loading the settings.dat
'first add some stuff to the frm_settings...

'For birthday
Dim A As Integer, B As Integer, C As Integer, D As Integer
For A = 1 To 31
 Frm_Settings.CB_Day.AddItem A
Next A

'for birthmonth
For B = 1 To 12
 Frm_Settings.CB_Month.AddItem B
Next B

'for birthyear
For C = 1930 To Format(Now, "YYYY")
 Frm_Settings.CB_Year.AddItem C
Next C

'for the country selection
'a few country to add to the combo..
'glad that this was finished..lol...
With Frm_Settings.CB_County
    .AddItem "United States"
    .AddItem "Afghanistan"
    .AddItem "Albania"
    .AddItem "Algeria"
    .AddItem "Andorra"
    .AddItem "Angola"
    .AddItem "Anguilla"
    .AddItem "Antarctica"
    .AddItem "Antigua and Barbuda"
    .AddItem "Argentina"
    .AddItem "Armenia"
    .AddItem "Aruba"
    .AddItem "Australia"
    .AddItem "Austria"
    .AddItem "Azerbaijan"
    .AddItem "Bahamas"
    .AddItem "Bahrain"
    .AddItem "Bangladesh"
    .AddItem "Barbados"
    .AddItem "Belarus"
    .AddItem "Belgium"
    .AddItem "Belize"
    .AddItem "Benin"
    .AddItem "Bermuda"
    .AddItem "Bhutan"
    .AddItem "Bolivia"
    .AddItem "Bosnia-Herzegovina"
    .AddItem "Botswana"
    .AddItem "Bouvet Island"
    .AddItem "Br Indian Ocean Ter"
    .AddItem "Brazil"
    .AddItem "Brit Virgin Islands"
    .AddItem "Brunei Darussalam"
    .AddItem "Bulgaria"
    .AddItem "Burkina Faso"
    .AddItem "Burundi"
    .AddItem "Cambodia"
    .AddItem "Cameroon"
    .AddItem "Canada"
    .AddItem "Cape Verde"
    .AddItem "Cayman Islands"
    .AddItem "Central Africa"
    .AddItem "Chad"
    .AddItem "Chile"
    .AddItem "China"
    .AddItem "Christmas Island"
    .AddItem "Cocos Islands"
    .AddItem "Colombia"
    .AddItem "Comoros"
    .AddItem "Congo"
    .AddItem "Cook Islands"
    .AddItem "Costa Rica"
    .AddItem "Cote D'ivoire"
    .AddItem "Croatia"
    .AddItem "Cuba"
    .AddItem "Cyprus"
    .AddItem "Czech Republic"
    .AddItem "Denmark"
    .AddItem "Djibouti"
    .AddItem "Dominica"
    .AddItem "Dominican Republic"
    .AddItem "East Timor"
    .AddItem "Ecuador"
    .AddItem "Egypt"
    .AddItem "El Salvador"
    .AddItem "Equatorial Guinea"
    .AddItem "Eritrea"
    .AddItem "Estonia"
    .AddItem "Ethiopia"
    .AddItem "Falkland Islands"
    .AddItem "Faroe Islands"
    .AddItem "Fiji"
    .AddItem "Finland"
    .AddItem "France"
    .AddItem "France-metropolitan"
    .AddItem "French Guiana"
    .AddItem "French Polynesia"
    .AddItem "French So Territorie"
    .AddItem "Gabon Republic"
    .AddItem "Gambia"
    .AddItem "Georgia"
    .AddItem "Germany"
    .AddItem "Ghana"
    .AddItem "Gibraltar"
    .AddItem "Greece"
    .AddItem "Greenland"
    .AddItem "Grenada"
    .AddItem "Guadeloupe"
    .AddItem "Guatemala"
    .AddItem "Guinea"
    .AddItem "Guinea-bissau"
    .AddItem "Guyana"
    .AddItem "Haiti"
    .AddItem "Heard Mcdonald Isl"
    .AddItem "Honduras"
    .AddItem "Hong Kong"
    .AddItem "Hungary"
    .AddItem "Iceland"
    .AddItem "India"
    .AddItem "Indonesia"
    .AddItem "Iran"
    .AddItem "Iraq"
    .AddItem "Ireland"
    .AddItem "Israel"
    .AddItem "Italy"
    .AddItem "Jamaica"
    .AddItem "Japan"
    .AddItem "Jordan"
    .AddItem "Kazakhstan"
    .AddItem "Kenya"
    .AddItem "Kiribati"
    .AddItem "Kuwait"
    .AddItem "Kyrgyzstan"
    .AddItem "Laos"
    .AddItem "Latvia"
    .AddItem "Lebanon"
    .AddItem "Lesotho"
    .AddItem "Liberia"
    .AddItem "Libyan Arab Jamahi"
    .AddItem "Liechtenstein"
    .AddItem "Lithuania"
    .AddItem "Luxembourg"
    .AddItem "Macau"
    .AddItem "Macedonia"
    .AddItem "Madagascar"
    .AddItem "Malawi"
    .AddItem "Malaysia"
    .AddItem "Maldives"
    .AddItem "Mali"
    .AddItem "Malta"
    .AddItem "Martinique"
    .AddItem "Mauritania"
    .AddItem "Mauritius"
    .AddItem "Mayotte"
    .AddItem "Mexico"
    .AddItem "Moldova"
    .AddItem "Monaco"
    .AddItem "Mongolia"
    .AddItem "Montserrat"
    .AddItem "Morocco"
    .AddItem "Mozambique"
    .AddItem "Myanmar"
    .AddItem "Namibia"
    .AddItem "Nauru"
    .AddItem "Nepal"
    .AddItem "Nether Antilles"
    .AddItem "Netherlands"
    .AddItem "New Caledonia"
    .AddItem "New Zealand"
    .AddItem "Nicaragua"
    .AddItem "Niger"
    .AddItem "Nigeria"
    .AddItem "Niue"
    .AddItem "Norfolk Island"
    .AddItem "Norway"
    .AddItem "Oman"
    .AddItem "Other"
    .AddItem "Pakistan"
    .AddItem "Panama"
    .AddItem "Papua New Guinea"
    .AddItem "Paraguay"
    .AddItem "People 's Rep Korea"
    .AddItem "Peru"
    .AddItem "Philippines"
    .AddItem "Pitcairn"
    .AddItem "Poland"
    .AddItem "Portugal"
    .AddItem "Qatar"
    .AddItem "Reunion"
    .AddItem "Romania"
    .AddItem "Russian Federation"
    .AddItem "Rwanda"
    .AddItem "Saint Helena"
    .AddItem "Saint Kitts Nevis"
    .AddItem "Saint Lucia"
    .AddItem "Samoa"
    .AddItem "San Marino"
    .AddItem "Saudi Arabia"
    .AddItem "Senegal"
    .AddItem "Serbia and Montenegro"
    .AddItem "Seychelles"
    .AddItem "Sierra Leone"
    .AddItem "Singapore"
    .AddItem "Slovakia"
    .AddItem "Slovenia"
    .AddItem "Soa Tome Pincipe"
    .AddItem "Solomon Islands"
    .AddItem "Somalia"
    .AddItem "South Africa"
    .AddItem "South Georgia"
    .AddItem "South Korea"
    .AddItem "Spain"
    .AddItem "Sri Lanka"
    .AddItem "St Vincent Grenadine"
    .AddItem "Sudan"
    .AddItem "Suriname"
    .AddItem "Svalbard Jan Mayen I"
    .AddItem "Swaziland"
    .AddItem "Sweden"
    .AddItem "Switzerland"
    .AddItem "Syrian Arab Republic"
    .AddItem "Taiwan, Roc"
    .AddItem "Tajikistan"
    .AddItem "Tanzania"
    .AddItem "Thailand"
    .AddItem "Togo"
    .AddItem "Tokelau"
    .AddItem "Tonga"
    .AddItem "Trinidad And Tobago"
    .AddItem "Tunisia"
    .AddItem "Turkey"
    .AddItem "Turkmenistan"
    .AddItem "Turks Caicos Islands"
    .AddItem "Tuvalu"
    .AddItem "Uganda"
    .AddItem "Ukraine"
    .AddItem "United Arab Emirates"
    .AddItem "United Kingdom"
    .AddItem "Uruguay"
    .AddItem "Us Minor Islands"
    .AddItem "Uzbekistan"
    .AddItem "Vanuatu"
    .AddItem "Vatican City"
    .AddItem "Venezuela"
    .AddItem "Viet NAM"
    .AddItem "Wallis Futuna Isl"
    .AddItem "Western Sahara"
    .AddItem "Yemen"
    .AddItem "Zaire"
    .AddItem "Zambia"
    .AddItem "Zimbabwe"
End With

'for interval time for htmlpage
For D = 1 To 60
 Frm_Settings.CB_Time.AddItem D
Next D


'let's have a look to the settiings file...
'first check if it's exists..
    If CheckFile(App.Path & "\Settings.dat") = False Then
        'if the file isn't there do function NoSettings
        NoSettings
        Exit Function
    ElseIf CheckFile(App.Path & "\Settings.dat") = True Then
        'if the file exist then load it
        LoadSettings
        Exit Function
    End If
    
End Function

Public Function NoSettings()
'if the settings file wasn't found... load this function...
'first welcome the user...
MsgBox "Welcome..." & vbCrLf & vbCrLf & "The Settings file couldn't be found..." & vbCrLf & "So this maybe the first time you start SaveCam 1.0." & vbCrLf & vbCrLf & "The default settings will be loaded." & vbCrLf & "Please take a look at the settings and correct them if nessesary." & vbCrLf & vbCrLf & "I'm a dutch guy with a bad english type... So Sorry for that...:)", vbInformation, "Settings file not found... (Maybe firts time...)"

'fill out the boxes of th frm_settings...

With Frm_Settings
.CB_Title.Text = "My SaveCam 1.0"
.CB_Port.Text = "80"
.CB_Name.Text = "SAVECAMTESTER"
.CB_Day.Text = "3"
.CB_Month.Text = "8"
.CB_Year.Text = "1979"
.CB_County.Text = "Netherlands"
.CB_Homepage.Text = "Http://www.planet-source-code.com"
.CB_Email.Text = "YourEmail@Adress.nl"
.CB_C1.Value = 1
.CB_C2.Value = 1
.CB_Time = 5
.FONTCOLOR.BackColor = vbYellow
.BGCOLOR.BackColor = vbBlack

'save these settings to the settings file
SaveSettings
End With

End Function

Public Function LoadSettings()
Dim TmpStr As String, LST() As String
Dim BR, BT, I As Integer, CT As Long

'oke.. now we know the settingsfile exists,
'lets check the length
CT = FileLen(App.Path & "\Settings.dat")
'if the length smallet then 10 the file isn't valid...
'kill if... and load the default settings
If CT < 10 Then
    Kill App.Path & "\Settings.dat"
    NoSettings
    Exit Function
End If


With Frm_Settings
'clear the tmpdatafield
.TMPdata.Text = ""

'open the settings encoded file...
Open App.Path & "\Settings.dat" For Input As #2
  Do While Not EOF(2)
        Line Input #2, TmpStr
        'display the encoded text to the tmpdatafield
            .TMPdata.Text = .TMPdata.Text & TmpStr
        Loop
Close #2

'decode the tmpdatafield to text
.TMPdata.Text = Code_D(.TMPdata.Text)

'find the end
BR = Split(.TMPdata.Text, "<END>")
BT = UBound(BR)
'now split the text fields.. with the <SC>
For I = 0 To BT - 1

 LST() = Split(BR(I), "<SC>")
 
 'write te text to it's correct text/datafield

.CB_Title.Text = LST(0)
.CB_Port.Text = LST(1)
.CB_Name.Text = LST(2)
.CB_Day.Text = LST(3)
.CB_Month.Text = LST(4)
.CB_Year.Text = LST(5)
.CB_County.Text = LST(6)
.CB_Homepage.Text = LST(7)
.CB_Email.Text = LST(8)
.CB_C1.Value = LST(9)
.CB_C2.Value = LST(10)
.FONTCOLOR.BackColor = LST(11)
.BGCOLOR.BackColor = LST(12)
.CB_Time = LST(13)

Next I

.TMPdata.Text = ""

End With
End Function

'a function to encode text files
Public Function Code_E(Text2Code As String)
Dim TC As Integer, tn As String, TW As String
    For TC = 1 To Len(Text2Code)
        tn$ = Asc(Mid(Text2Code, TC, Len(Text2Code))) + 120
    
        TW$ = TW$ & Chr(tn$)
    Next TC
    
    Code_E = TW$
End Function

'a function to decode the encoded files...
Public Function Code_D(Code2Text As String)
Dim CC As Integer, CN As String, CW As String
    For CC = 1 To Len(Code2Text)
        CN$ = Asc(Mid(Code2Text, CC, Len(Code2Text))) - 120
        
        CW$ = CW$ & Chr(CN$)
    Next CC
    
    Code_D = CW$
End Function

Public Function SaveSettings()

With Frm_Settings
'clear the old textfield to be sure it's empty...
.TMPdata.Text = ""

'check if all fields for data...

'first check the title...
If Len(.CB_Title.Text) < 5 Or Len(.CB_Title.Text) > 25 Then
 MsgBox "The Tilte you've submitted is not valid." & vbCrLf & vbCrLf & "A valid title most be at least 5 charactars and less then 25 charactars...", vbInformation, "SavaCam 1.0, Settings, Title not valid."
 .CB_Title.SetFocus
 Exit Function
End If

'second check the portnumber
Dim CT As Integer
CT = .CB_Port.Text
If CT < 80 Or CT > 8080 Then
    MsgBox "No valid port number..." & vbCrLf & vbCrLf & "A valid port number sits between port 80 and port 8080." & vbCrLf & "Please enter a valid port number.", vbInformation, "SaveCam 1.0, Settings, Port number not valid."
    .CB_Port.SetFocus
    Exit Function
End If

'third check the name...
If Len(.CB_Name.Text) < 5 Or Len(.CB_Name.Text) > 20 Then
 MsgBox "The Name you've submitted is not valid." & vbCrLf & vbCrLf & "A valid name most be at least 5 charactars and less then 20 charactars...", vbInformation, "SavaCam 1.0, Settings, Name not valid."
 .CB_Name.SetFocus
 Exit Function
End If

'fourth check the birthdate
'the day
If .CB_Day.List(.CB_Day.ListIndex) = .CB_Day.List(-1) Then
    MsgBox "Please give up your birthdate correctly.", vbInformation, "SaveCam 1.0, Settings, No correct birthdate."
    Exit Function
End If
'the month
If .CB_Month.List(.CB_Month.ListIndex) = .CB_Month.List(-1) Then
    MsgBox "Please give up your birthdate correctly.", vbInformation, "SaveCam 1.0, Settings, No correct birthdate."
    Exit Function
End If
'the year
If .CB_Year.List(.CB_Year.ListIndex) = .CB_Year.List(-1) Then
    MsgBox "Please give up your birthdate correctly.", vbInformation, "SaveCam 1.0, Settings, No correct birthdate."
    Exit Function
End If

'fifth check the country
'forgot the R in country.. sorry...
If .CB_County.List(.CB_County.ListIndex) = .CB_County.List(-1) Then
    MsgBox "Please give up your country.", vbInformation, "SaveCam 1.0, Settings, No country given."
    Exit Function
End If

'six check the website
If Len(.CB_Homepage.Text) < 5 Then
    MsgBox "You enterd no valid website adress..." & vbCrLf & vbCrLf & "Google will be your website adres...", vbInformation, "SaveCam 1.0, Settings, No valid Website given."
    .CB_Homepage.Text = "Http://www.google.com"
End If

'seven check the e-mail
If CheckMail(.CB_Email.Text) = False Then
    MsgBox "The E-Mail adress you enterd is not a valid e-mail adress," & vbCrLf & vbCrLf & "Please correct it...", vbInformation, "SaveCam 1.0, Settings, No valid Email Adress"
    Exit Function
End If

'eight check the interval
If .CB_Time.List(.CB_Time.ListIndex) = .CB_Time.List(-1) Then
    MsgBox "Please give up the webinterval.", vbInformation, "SaveCam 1.0, Settings, No correct interval."
    Exit Function
End If
'all important textfield have been checked...
'Now save then...

'first kill the old file.......
If CheckFile(App.Path & "\Settings.dat") = True Then
    Kill App.Path & "\Settings.dat"
End If

Dim TS As String, ES As String, T2C As String
TS = "<SC>"
ES = "<END>"

.TMPdata.Text = .CB_Title.Text & TS & .CB_Port.Text & TS & .CB_Name.Text & TS & .CB_Day.Text & TS & .CB_Month.Text & TS & .CB_Year.Text & TS & .CB_County.Text & TS & .CB_Homepage.Text & TS & .CB_Email.Text & TS & .CB_C1.Value & TS & .CB_C2.Value & TS & .FONTCOLOR.BackColor & TS & .BGCOLOR.BackColor & TS & .CB_Time.Text & ES

'set the string to encode
T2C = .TMPdata.Text

'clear the tmpfield
.TMPdata.Text = ""

'encode de string to the tmpfield
.TMPdata.Text = Code_E(T2C)

'save the data to the settings file
Open App.Path & "\Settings.dat" For Output As #1
    Print #1, .TMPdata.Text
Close #1

'clear the tmp field
.TMPdata.Text = ""

'wooowwwwwwww... it's saved... tell the user...
MsgBox "Settings succesfully saved...", vbInformation, "SaveCam 1.0"
End With
End Function

Function CheckIP(Ipadress As String) As Boolean
'to check the total of different viewers...
'look at the wslisten_connectionrequest
Dim C As Integer
    For C = 0 To Frm_main.iplist.ListCount - 1
        If Ipadress = Frm_main.iplist.List(C) Then
            CheckIP = True
            Exit Function
        End If
    Next C
CheckIP = False
End Function

Function HtmlColor(ByVal color As Long) As String
'here we convert the visuab basic color to web color
    Dim tmp As String
    tmp = Right$("00000" & Hex$(color), 6)
    HtmlColor = "#" & Right$(tmp, 2) & Mid$(tmp, 3, 2) & Left$(tmp, 2)
End Function

Public Function MakeHtml()
'here we make the html file for the viewers to look at...

'first kill the old file.. if it exists...
'this because a user could change the settings and we want the new settings...
If CheckFile(App.Path & "\SaveCam.dat") = True Then
    Kill App.Path & "\Savecam.dat"
End If

'open the html file to write to...
Open App.Path & "\Data\Savecam.dat" For Output As #3
'print some html data....
Print #3, "<!-- SaveCam 1.0 - By Dodo2479@Hotmail.com >"
Print #3, "<Html>" & vbCrLf & "<Head>"
Print #3, "<Title>SaveCam 1.0 - " & Frm_Settings.CB_Name.Text & "'s SaveCam Online</Title>"

'this is the refresh function of the html file....
Print #3, "<script language='JavaScript'>"
Print #3, "var refreshinterval=" & Frm_Settings.CB_Time.Text  'interval !!!!!!!!!!!!
Print #3, "var displaycountdown='yes'"
Print #3, "var starttime"
Print #3, "var nowtime"
Print #3, "var reloadseconds=0"
Print #3, "var secondssinceloaded=0"
Print #3, "function starttime() {"
Print #3, "starttime=new Date()"
Print #3, "starttime=starttime.getTime()"
Print #3, "countdown()"
Print #3, "}"
Print #3, "function countdown() {"
Print #3, "nowtime= new Date()"
Print #3, "nowtime=nowtime.getTime()"
Print #3, "secondssinceloaded=(nowtime-starttime)/1000"
Print #3, "reloadseconds=Math.round(refreshinterval-secondssinceloaded)"
Print #3, "if (refreshinterval>=secondssinceloaded) {"
Print #3, "var timer=setTimeout('countdown()',1000)"
Print #3, "if (displaycountdown=='yes') {"
Print #3, "window.status='SaveCam 1.0: A new picture will be visseble in '+reloadseconds+ ' seconds'"
Print #3, "}"
Print #3, "}"
Print #3, "else {"
Print #3, "clearTimeout(timer)"
Print #3, "window.location.reload(true)"
Print #3, "}"
Print #3, "}"
Print #3, "window.onload=starttime"
Print #3, "</script>"

Print #3, "</Head>"

'this is what the viewers will see...
'just html with data from the frm_settings
Print #3, "<Body bgcolor='" & HtmlColor(Frm_Settings.BGCOLOR.BackColor) & "'>"
Print #3, "<HR color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>"
Print #3, "<Center><B><I><U><Font color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "' size='8'>SaveCam 1.0</U></I></B></Font>"
Print #3, "<HR color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>"
Print #3, "<TABLE BORDER='0'>"
Print #3, "<TR><TD><img alt='" & Frm_Settings.CB_Title.Text & "' border='0' src='Webcam.jpg'></TD><TD></TD>"
Print #3, "<TD><Font size='5' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>"
Print #3, "This is....<BR><BR>Name : " & Frm_Settings.CB_Name.Text & "<BR>"
Print #3, "Birthdate : " & Frm_Settings.CB_Day.Text & "-" & Frm_Settings.CB_Month.Text & "-" & Frm_Settings.CB_Year.Text & "<BR>"
Print #3, "Country : " & Frm_Settings.CB_County.Text & "<BR><BR></Font>"

'check if the user wants to display his homepage
If Frm_Settings.CB_C1.Value <> 0 Then
    Print #3, "<Font size='4' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>"
    Print #3, "Visit my homepage at :<br> <A Href='" & Frm_Settings.CB_Homepage.Text & "' target='_Blank'><Font Size='4' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>" & Frm_Settings.CB_Homepage.Text & "</A><BR><BR></Font>"
End If

'check if the user wants to display his email adres...
If Frm_Settings.CB_C2.Value <> 0 Then
    Print #3, "<Font size='4' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>"
    Print #3, "Click <A Href='Mailto:" & Frm_Settings.CB_Email.Text & "'><Font Size='4' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>here</A><Font Size='4' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'> to send me an e-mail...</Font>"
End If

Print #3, "</TD></TR></TABLE>"
Print #3, "<HR color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'><br><br>"
Print #3, "<font size='2' color='" & HtmlColor(Frm_Settings.FONTCOLOR.BackColor) & "'>SaveCam 1.0<br>by<br>Dodo2479@hotmail.com</font>"
Print #3, "</Body></Html>"
Print #3, "<!-- SaveCam 1.0 by Dodo2479@hotmail.com >"

'close the file...
Close #3
End Function
