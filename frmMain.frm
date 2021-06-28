VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Begin VB.Form frmMain 
   Caption         =   "Website Downloader                                                                      Step  4  of 4"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   612
      Left            =   4680
      OleObjectBlob   =   "frmMain.frx":000C
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   3720
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   252
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   252
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<<Back"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtWebsite 
      Height          =   288
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.TextBox txtDir 
      Height          =   288
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3252
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtMessages 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":004E
      Top             =   720
      Width           =   5172
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Website Downloader"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Boolean
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAcessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nServerPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, byValReferer As String, ByVal lpszAcceptTypes As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal lpszheaders As String, ByVal dwHeadersLenght As Long, ByVal lpOptional As String, ByVal dwOptionalLength As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long) As Boolean
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal address As String, ByVal headers As String, ByVal headlenght As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Dim url(100) As String
Dim levu(1000) As String
Dim xz As Integer
Dim oo As Integer
Dim opt1 As Boolean
Dim opt2 As Boolean
Dim opt3 As Boolean
Dim o As Integer
Dim strDurl As String
Dim exitproc As Boolean
Dim msize As Long
Dim b As Boolean
Dim f As Boolean
Dim files As Integer
Dim hInternet As Long
Dim hConnect As Long
Dim strServer As String
Dim iPort As Integer
Dim bRes As Boolean
Dim lFlags As Long
Dim hRequest As Long
Dim strURL As String
Dim strBuffer As String * 1
Dim strDir As String
Dim strFile As String
Dim strMurl As String
Dim appdir As String
Dim files1 As Integer
Const INTERNET_FLAG_NO_COOKIES = &H80000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const INTERNET_SERVICE_HTTP = 3
Private Sub cmdConnect_Click()
cmdHangup.Enabled = True
b = InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0)
End Sub

Private Sub cmdHangup_Click()
f = InternetAutodialHangup(0)

End Sub

Private Sub cmdBack_Click()
Load frmUrl
frmUrl.Text1.Text = txtWebsite.Text
frmUrl.Text2.Text = txtDir.Text
Unload Me
frmUrl.Show

End Sub

Private Sub cmdStart_Click()
txtMessages.Text = ""

On Error Resume Next
exitproc = False
Gif89a1.Visible = True
Gif89a1.FileName = App.Path & "\mov1.gif"
xz = 0
o = 1
oo = 0
Dim a As Integer
Dim c As Integer
Dim er As Integer
Dim br As Integer
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim s As String
files1 = Text1.Text
If Check1.Value = 1 Then opt1 = True Else opt1 = False
If Check2.Value = 1 Then opt2 = True Else opt2 = False
If Check3.Value = 1 Then opt3 = True Else opt3 = False
appdir = txtDir.Text
br = Len(appdir)
er = InStrRev(appdir, "\")
If Not fso.folderexists(appdir) Then
MsgBox "Invalid destination directory"
Exit Sub
End If
If br = er Then appdir = Left(appdir, br)
stryyy = txtWebsite.Text
files = 0
c = Len(stryyy)
a = InStr(stryyy, "/")
If a = 0 Then stryyy = stryyy & "/"
a = InStr(stryyy, "/")
strServer = Left(stryyy, a - 1)
strURL = Right(stryyy, c - a + 1)
strTryurl = strURL
er = InStr(strTryurl, ".htm")
If er = 0 Then
a = InStrRev(strTryurl, "/")
If Not a = Len(strTryurl) Then strTryurl = strTryurl & "/"
a = InStrRev(strTryurl, "/")
c = Len(strTryurl)
strMurl = Left(strTryurl, a - 1)
Call getsize
Call urltry
Else
a = InStrRev(strTryurl, "/")
c = Len(strTryurl)
strMurl = Left(strTryurl, a - 1)
End If
iPort = 80
Call process(strServer, strURL)
Call stripurl
txtMessages.Text = txtMessages.Text & vbCrLf & " Starting to download links in file"
txtMessages.SelStart = Len(txtMessages.Text)
Call dotry
For jj = 1 To o
Call level1(url(jj))
Next jj
Call downlevel
MsgBox "Finished downloading"
Command1.Caption = "Exit"
Set frmMain = Nothing
Set frmstart = Nothing
Set frmUrl = Nothing

End Sub




Private Sub download(strSServer As String, strUURL As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim sServer As String
Dim sUrl As String
Dim x As String
Dim y As String
Dim z, f
If files > files1 Then Exit Sub
iPort = 80
sServer = strSServer
sUrl = strUURL
iFlags = INTERNET_FLAG_NO_COOKIES
iFlags = iFlags Or INTERNET_FLAG_NO_CACHE_WRITE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.fileexists(appdir & sUrl) Then Exit Sub
hInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
If hInternet <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Open Successfull"
hConnect = InternetConnect(hInternet, sServer, iPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)
If hConnect <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Connect Succesfull"
hRequest = HttpOpenRequest(hConnect, "GET", sUrl, "HTTP/1.0", vbNullString, vbNullString, iFlags, 0)
If hRequest <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Http Open Request succesfull"
bRes = HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0)
If bRes = True Then txtMessages.Text = txtMessages.Text & vbCrLf & "Request successfull"
strDir = Dir(appdir & sUrl)
If Len(strDir) > 0 Then
Kill appdir & sUrl
End If
iFile = FreeFile()
Call makedire(sUrl)
Open appdir & sUrl For Binary Access Write As iFile
Do
bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
If lBytesRead > 0 Then
Put iFile, , strBuffer
End If
Loop While lBytesRead > 0
Close iFile
files = files + 1
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished downloading " & sServer & sUrl
txtMessages.SelStart = Len(txtMessages.Text)
DoEvents
If exitproc = True Then Unload Me
End Sub


Private Sub makedire(strYZ As String)
If exitproc = True Then Exit Sub
On Error Resume Next
strYZZ = strYZ
Dim sty As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim b As Integer
Dim a As Integer
Dim x(10) As Integer
b = 0
a = InStr(strYZZ, "/")
c = Len(strYZZ)
stree = strYZZ
x(0) = 0
While a <> 0
b = b + 1
x(b) = x(b - 1) + a
strYZZ = Right(strYZZ, c - a)
c = Len(strYZZ)
a = InStr(strYZZ, "/")
Wend
For s = 1 To b
stre = Left(stree, x(s))

y = appdir & stre
txtMessages.Text = txtMessages.Text & vbCrLf & "Creating local sub directory " & appdir & stre
txtMessages.SelStart = Len(txtMessages.Text)
If Not fso.folderexists(y) Then MkDir (y)
Next s
DoEvents
If exitproc = True Then Unload Me
End Sub

 
 

Private Sub subfiles(strBserver As String, strBurl As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim aer As Integer
Dim ber As Integer
Dim iFile As Integer
Dim strTry5 As String
Dim strbburl As String
strbburl = strBurl
If strbburl = "" Then Exit Sub
Dim strTry6 As String
strTry6 = ""
iFile = 1
Dim strCheck As String
Dim strTry3 As String
strTry3 = "src=" & Chr(34)
strTry4 = Chr(34)
Dim strtry9 As String
strtry9 = "SRC=" & Chr(34)
Open appdir & strbburl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1

ans = InStr(strCheck, strTry3)
If ans = 0 Then ans = InStr(strCheck, strtry9)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strTry3) - ans)
cns = InStr(strCheck, strTry4)
If cns > 0 Then
strTry5 = Left(strCheck, cns - 1)
aer = InStr(strTry5, "/")
If aer <> 0 Then
ber = InStr(strTry5, "../")
If ber <> 0 Then
strTry6 = Right(strTry5, Len(strTry5) - ber + 1)
GoTo 10
Else:
strTry6 = strMurl & "/" & strTry5
GoTo 10
End If
End If
strTry6 = strMurl & "/" & strTry5
10:
Dim mz As Integer
Dim mx As Integer
Dim my As Integer
Dim mw As Integer
Dim ms As Integer
Dim mt As Integer
If opt2 = False Then ms = 0 Else ms = InStr(strTry6, ".gif")
If opt2 = False Then mt = 0 Else mt = InStr(strTry6, ".jpg")
If opt3 = True Then
ms = 1
mt = 1
End If
mz = InStr(strTry6, ".co")
mx = InStr(strTry6, ".net")
my = InStr(strTry6, ".org")
mw = InStr(strTry6, ".edu")
If mz = 0 And my = 0 And mx = 0 And mw = 0 And (mt <> 0 Or ms <> 0) Then
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading File " & strServer & strTry6
txtMessages.SelStart = Len(txtMessages.Text)
Call download(strServer, strTry6)
End If
End If

ans = InStr(strCheck, strTry3)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile

End Sub

Private Sub stripurl()
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim strseek99 As String
Dim g As Integer
Dim mep As Integer
Dim mpp As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
txtMessages.SelStart = Len(txtMessages.Text)
strTry = Chr(34)
strSeek = "href=" & Chr(34)
strseek99 = "HREF=" & Chr(34)
iFile = FreeFile()
Open appdir & strURL For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 Then
mep = InStr(strtry1, "../")
mpp = InStr(strtry1, "./")
If mep <> 0 Then
url(o) = strMurl & strtry1
ElseIf mpp <> 0 Then url(o) = strMurl & strtry1
Else: url(o) = strMurl & "/" & strtry1
End If
o = o + 1
End If
End If
ans = InStr(strCheck, strSeek)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"
txtMessages.SelStart = Len(txtMessages.Text)
End Sub

Private Sub Command4_Click()
Call stripurl
End Sub
Private Sub process(strsrv As String, stru As String)
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strDserv As String

strDserv = strsrv
strDurl = stru
txtMessages.Text = txtMessages & vbCrLf & "Starting to download " & strDserv & strDurl
txtMessages.SelStart = Len(txtMessages.Text)
Call download(strDserv, strDurl)
Call background(appdir & strDurl)
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading image files"
txtMessages.SelStart = Len(txtMessages.Text)
Call subfiles(strDserv, strDurl)
txtMessages.Text = ""
End Sub

Private Sub dotry()
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer

For jj = 1 To o
DoEvents
If exitproc = True Then Unload Me

If url(jj) = "" Then Exit Sub
Call process(strServer, url(jj))
Next jj
End Sub

Private Sub Command1_Click()
exitproc = True
Set frmMain = Nothing
Unload Me

End Sub






Private Sub getsize()
hInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
If hInternet <> 0 Then hRequest = InternetOpenUrl(hInternet, "http://" & txtWebsite.Text, vbNullString, 0, INTERNET_FLAG_NO_AUTO_REDIRECT, 0)
Open appdir & "/temp.log" For Binary Access Write As 1
Do
bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
If lBytesRead > 0 Then
Put #1, , strBuffer
End If
Loop While lBytesRead > 0
Close #1
Dim fso, y, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(appdir & "\temp.log")
y = f.Size
msize = y
End Sub

Private Sub background(strFilename As String)
Dim check As String
Dim muck As String
muck = "background=" & Chr(34)
muck1 = "BACKGROUND=" & Chr(34)
Dim a As Integer
Dim b As Integer
Open strFilename For Input As 1
Do While Not EOF(1)
Input #1, check
a = InStr(check, muck)
If a <> 0 Then
check = Right(check, Len(check) - 11 - a)
muck = Chr(34)
b = InStr(check, muck)
check = Left(check, b - 1)
cdz = InStr(check, "http://")
If cdz <> 0 Then Call download(strServer, strDurl)
Close #1
Exit Sub
End If
gr = InStr(check, muck1)
If gf <> 0 Then
check = Right(check, Len(check) - 11 - a)
muck1 = Chr(34)
b = InStr(check, muck1)
check = Left(check, b - 1)
cdz = InStr(check, "http://")
If cdz <> 0 Then Call download(strServer, strDurl)
Close #1
Exit Sub
End If
Loop
Close #1
End Sub

Private Sub Form_Load()
txtMessages.Text = "PRESS START TO BEGIN DOWNLOADING"
txtWebsite.Enabled = False
txtDir.Enabled = False
Unload frmUrl
Unload frmstart
End Sub

Private Sub urltry()
Call download(strServer, strMurl & "/index.htm")
Dim fso, f, s, t, y
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(appdir & strMurl & "/index.htm")
s = f.Size
If s = msize Then
strURL = strMurl & "/index.htm"
Exit Sub
End If
Call download(strServer, strMurl & "/index.html")
Set f = fso.GetFile(appdir & strMurl & "/index.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/index.html"
Kill (appdir & strMurl & "/index.htm")
Exit Sub
End If
Call download(strServer, strMurl & "/default.htm")
Set f = fso.GetFile(appdir & strMurl & "/default.htm")
t = f.Size
If t = msize Then
strURL = strMurl & "/default.htm"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Exit Sub
End If
Call download(strServer, strMurl & "/default.html")
Set f = fso.GetFile(appdir & strMurl & "/default.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/default.html"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Exit Sub
End If

Call download(strServer, strMurl & "/start.htm")
Set f = fso.GetFile(appdir & strMurl & "/start.htm")
t = f.Size
If t = msize Then
strURL = strMurl & "/start.htm"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Kill (appdir & strMurl & "/default.html")
Exit Sub
End If
Call download(strServer, strMurl & "/start.html")
Set f = fso.GetFile(appdir & strMurl & "/start.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/start.html"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Kill (appdir & strMurl & "/default.html")
Kill (appdir & strMurl & "/start.htm")
Exit Sub
End If
Kill (appdir & strMurl & "/start.html")
End Sub

Private Sub level1(uurl As String)
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim strseek99 As String
Dim mep As Integer
Dim mpp As Integer
Dim g As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
txtMessages.SelStart = Len(txtMessages.Text)
strTry = Chr(34)
strSeek = "href=" & Chr(34)
strseek99 = "HREF=" & Chr(34)
iFile = FreeFile()
If uurl = "" Then Exit Sub
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim asd As Boolean
asd = fso.fileexists(appdir & uurl)
If asd = True Then
Open appdir & uurl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
If Not strCheck = "" Then
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strSeek1)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
Dim bee As Integer
bee = InStr(strtry1, ".htm")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 And bee <> 0 Then
mep = InStr(strtry1, "../")
mpp = InStr(strtry1, "./")
If mep <> 0 Then
levu(oo) = strMurl & strtry1
ElseIf mpp <> 0 Then levu(oo) = strMurl & strtry1
Else: levu(oo) = strMurl & "/" & strtry1
End If
oo = oo + 1
End If
End If
ans = InStr(strCheck, strSeek)
DoEvents
If exitproc = True Then Unload Me
Wend
If exitproc = True Then Unload Me
End If
Loop
Close iFile
End If
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"
txtMessages.SelStart = Len(txtMessages.Text)

End Sub
Private Sub downlevel()
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer

For jj = 1 To oo
DoEvents
If exitproc = True Then Unload Me

If levu(jj) = "" Then Exit Sub
Call process(strServer, levu(jj))
Next jj
End Sub
