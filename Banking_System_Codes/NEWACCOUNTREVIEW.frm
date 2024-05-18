VERSION 5.00
Begin VB.Form NEWACCOUNTREVIEW 
   BackColor       =   &H00404000&
   Caption         =   "Review your details"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CONFIRM 
      BackColor       =   &H0080C0FF&
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton EDIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton REFRESH 
      BackColor       =   &H0080C0FF&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox NQUALIFICATION2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   6
      Top             =   4200
      Width           =   5415
   End
   Begin VB.TextBox NGENDER2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox NNAME2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox NAGE2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox NAADHAR2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      MaxLength       =   12
      TabIndex        =   2
      Top             =   4680
      Width           =   5415
   End
   Begin VB.TextBox NPAN2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5160
      Width           =   5415
   End
   Begin VB.TextBox NMOBILE2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5640
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Revie the details of new account"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10575
   End
   Begin VB.Image SIGNATURE 
      Height          =   855
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image PHOTO 
      Height          =   2055
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Qualification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Aadhar number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Pan number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Date Of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Mobile number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Width           =   3015
   End
End
Attribute VB_Name = "NEWACCOUNTREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim D As Integer
Dim M As Integer
Dim Y As Integer


Private Sub BACK_Click()
NEWACCOUNTREVIEW.Hide
NEWACCOUNT.Show
End Sub

Private Sub CONFIRM_Click()
On Error GoTo E1
NEWACC.NEWNAME = NNAME2.Text
NEWACC.NEWAADHAR = NAADHAR2.Text
NEWACC.NEWPAN = NPAN2.Text
NEWACC.NEWMOBILE = NMOBILE2.Text

NNAME2.Text = ""
NAGE2.Text = ""
NGENDER2.Text = ""
NQUALIFICATION2.Text = ""
NAADHAR2.Text = ""
NPAN2.Text = ""
NMOBILE2.Text = ""
PHOTO.Picture = LoadPicture("")
SIGNATURE.Picture = LoadPicture("")

PROGRESSFORM.Show
Exit Sub
E1:
    MsgBox "All fields are mandatory", , "C t O Bank"

End Sub

Private Sub EDIT_Click()
NNAME2.Enabled = True
NAADHAR2.Enabled = True
NPAN2.Enabled = True
NMOBILE2.Enabled = True
End Sub

Private Sub END_Click()
End
End Sub

Private Sub Form_Load()
PHOTO.Picture = LoadPicture(NEWACC.NEWPHOTO)
SIGNATURE.Picture = LoadPicture(NEWACC.NEWSIGNATURE)

D = NEWACC.NEWDOBD
M = NEWACC.NEWDOBM
Y = NEWACC.NEWDOBY
NNAME2.Text = NEWACC.NEWNAME
NAGE2.Text = D & "/" & M & "/" & Y
NGENDER2.Text = NEWACC.NEWGENDER
NQUALIFICATION2.Text = NEWACC.NEWQUALIFICATION10 & ", " & NEWACC.NEWQUALIFICATION12 & ", " & NEWACC.NEWQUALIFICATIONG & ", " & NEWACC.NEWQUALIFICATIONPG
NAADHAR2.Text = NEWACC.NEWAADHAR
NPAN2.Text = NEWACC.NEWPAN
NMOBILE2.Text = NEWACC.NEWMOBILE
End Sub

Private Sub REFRESH_Click()
PHOTO.Picture = LoadPicture(NEWACC.NEWPHOTO)
SIGNATURE.Picture = LoadPicture(NEWACC.NEWSIGNATURE)
D = NEWACC.NEWDOBD
M = NEWACC.NEWDOBM
Y = NEWACC.NEWDOBY
NNAME2.Text = NEWACC.NEWNAME
NAGE2.Text = D & "/" & M & "/" & Y
NGENDER2.Text = NEWACC.NEWGENDER
NQUALIFICATION2.Text = NEWACC.NEWQUALIFICATION10 & ", " & NEWACC.NEWQUALIFICATION12 & ", " & NEWACC.NEWQUALIFICATIONG & ", " & NEWACC.NEWQUALIFICATIONPG
NAADHAR2.Text = NEWACC.NEWAADHAR
NPAN2.Text = NEWACC.NEWPAN
NMOBILE2.Text = NEWACC.NEWMOBILE
End Sub
