VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form PROGRESSFORM 
   BackColor       =   &H00404000&
   Caption         =   "Processing the details and generating your account nnumber"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   3
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Please wait while your account number is being generated"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "PROGRESSFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer

Private Sub Form_Load()
ProgressBar1.Value = 0
I = 0
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
I = I + 1
ProgressBar1.Value = I
If ProgressBar1.Value = 100 Then
    Timer1.Enabled = False
    PROGRESSFORM.Hide
    CTOACCOUNTNUMBER.Show
End If
End Sub

