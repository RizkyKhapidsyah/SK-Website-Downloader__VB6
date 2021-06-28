VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Begin VB.Form frmOptions 
   Caption         =   "Website Downloader -venky_dude                                              Step  3  of 4"
   ClientHeight    =   4248
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6660
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4248
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   1572
      Left            =   4560
      OleObjectBlob   =   "frmOptions.frx":18FA
      TabIndex        =   14
      Top             =   960
      Width           =   1572
   End
   Begin VB.OptionButton Option1 
      Caption         =   "No Limit"
      Height          =   372
      Index           =   3
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "50"
      Height          =   372
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   2400
      Width           =   492
   End
   Begin VB.OptionButton Option1 
      Caption         =   "25"
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   2400
      Width           =   612
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
   Begin VB.OptionButton Option1 
      Caption         =   "10"
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Value           =   -1  'True
      Width           =   492
   End
   Begin VB.CheckBox Check3 
      Caption         =   "All Types"
      Height          =   612
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   1092
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Gif/Jpeg"
      Height          =   612
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Text/Html"
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   492
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Width           =   12
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Maximum No of Files to be downloaded"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Select Types of files to be downloaded"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3372
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5892
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmUrl.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()
If Option1(0).Value = True Then frmMain.Text1.Text = 10
If Option1(1).Value = True Then frmMain.Text1.Text = 25
If Option1(2).Value = True Then frmMain.Text1.Text = 50
If Option1(3).Value = True Then frmMain.Text1.Text = 200
frmMain.Check1 = Check1
frmMain.Check2 = Check2
frmMain.Check3 = Check3
frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()
Gif89a1.FileName = App.Path & "\mov3.gif"
End Sub
