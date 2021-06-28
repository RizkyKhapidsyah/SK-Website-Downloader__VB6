VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Begin VB.Form frmUrl 
   Caption         =   "Website Downloader -venky_dude                                              Step  2  of 4"
   ClientHeight    =   4248
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6660
   Icon            =   "frmUrl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4248
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   732
      Left            =   4800
      OleObjectBlob   =   "frmUrl.frx":18FA
      TabIndex        =   10
      Top             =   1320
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>>>"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<<Back"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "ex: www.geocities.com/venky_dude/index.htm"
      Height          =   252
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Width           =   3732
   End
   Begin VB.Label Label4 
      Caption         =   "Enter the directory in which the files should be saved"
      Height          =   372
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   3132
   End
   Begin VB.Label Label3 
      Caption         =   "http://"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the website URL to download:"
      Height          =   252
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   3492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Website Downloader"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5892
   End
End
Attribute VB_Name = "frmUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
frmstart.Show
End Sub

Private Sub cmdCancel_Click()
Unload Me
Unload frmMain
Unload frmstart
End Sub

Private Sub cmdNext_Click()
Dim appdir As String
Dim stryyy As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim s As String
appdir = Text2.Text
br = Len(appdir)
er = InStrRev(appdir, "\")
If Not fso.folderexists(appdir) Then
MsgBox "Invalid destination directory"
Exit Sub
End If
If br = er Then appdir = Left(appdir, br - 1)

stryyy = Text1.Text
er = InStr(stryyy, ".htm")
If er = 0 Then
End If
Unload Me
Load frmMain
frmMain.txtWebsite.Text = stryyy
frmMain.txtDir.Text = appdir
frmOptions.Show
End Sub

Private Sub Form_Load()
Gif89a1.FileName = App.Path & "\mov3.gif"
End Sub
