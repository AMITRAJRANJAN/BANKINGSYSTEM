VERSION 5.00
Begin VB.Form CTOHOME 
   BackColor       =   &H00FF00FF&
   Caption         =   "C t O Home"
   ClientHeight    =   10725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame ABOUTUSFRAME 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   4680
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton CLOSE 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox TEXTABOUTCOPYRIGHT 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "CTOHOME.frx":0000
         Top             =   1920
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.TextBox TEXTABOUTDEVELOPER 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "CTOHOME.frx":00B8
         Top             =   1920
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.TextBox TEXTABOUTOWNER 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "CTOHOME.frx":0141
         Top             =   1920
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.TextBox TEXTABOUTSOFTWARE 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "CTOHOME.frx":0173
         Top             =   1920
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.CommandButton ABOUTCOPYRIGHT 
         BackColor       =   &H000080FF&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton ABOUTDEVELOPER 
         BackColor       =   &H000080FF&
         Caption         =   "Developer"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton ABOUTOWNER 
         BackColor       =   &H000080FF&
         Caption         =   "Owner"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton ABOUTSOFTWARE 
         BackColor       =   &H000080FF&
         Caption         =   "Software"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10575
      End
   End
   Begin VB.CommandButton ABOUTUS 
      BackColor       =   &H0080C0FF&
      Caption         =   "About us"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton END 
      BackColor       =   &H008080FF&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   4680
      ScaleHeight     =   4665
      ScaleWidth      =   10905
      TabIndex        =   2
      Top             =   5760
      Width           =   10935
      Begin VB.CommandButton NEWUSER 
         BackColor       =   &H0080C0FF&
         Caption         =   "New user Click here"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3600
         Width           =   3975
      End
      Begin VB.CommandButton ULOGIN 
         BackColor       =   &H0080C0FF&
         Caption         =   "User Login"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   3975
      End
      Begin VB.CommandButton ELOGIN 
         BackColor       =   &H0080C0FF&
         Caption         =   "Employee Login"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   7440
         Picture         =   "CTOHOME.frx":0297
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   1320
         Picture         =   "CTOHOME.frx":104D7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "The Bank of Every One"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   24.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   10455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "  B a n k"
      BeginProperty Font 
         Name            =   "HoloLens MDL2 Assets"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1695
      Left            =   7200
      TabIndex        =   0
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      Height          =   1455
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   6000
      X2              =   13815
      Y1              =   1215
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   9720
      X2              =   9720
      Y1              =   1200
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   7440
      X2              =   6000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   6000
      X2              =   6000
      Y1              =   1680
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   7440
      X2              =   6000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404000&
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   4680
      Top             =   240
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404000&
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   10215
      Left            =   15840
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   10215
      Left            =   240
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "CTOHOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ABOUTCOPYRIGHT_Click()
TEXTABOUTSOFTWARE.Visible = True
TEXTABOUTDEVELOPER.Visible = False
TEXTABOUTOWNER.Visible = False
TEXTABOUTCOPYRIGHT.Visible = True
End Sub

Private Sub ABOUTDEVELOPER_Click()
TEXTABOUTSOFTWARE.Visible = False
TEXTABOUTDEVELOPER.Visible = True
TEXTABOUTOWNER.Visible = False
TEXTABOUTCOPYRIGHT.Visible = False

End Sub

Private Sub ABOUTOWNER_Click()
TEXTABOUTSOFTWARE.Visible = False
TEXTABOUTDEVELOPER.Visible = False
TEXTABOUTOWNER.Visible = True
TEXTABOUTCOPYRIGHT.Visible = False
End Sub

Private Sub ABOUTSOFTWARE_Click()
TEXTABOUTSOFTWARE.Visible = True
TEXTABOUTDEVELOPER.Visible = False
TEXTABOUTOWNER.Visible = False
TEXTABOUTCOPYRIGHT.Visible = False
End Sub

Private Sub ABOUTUS_Click()
ABOUTUSFRAME.Visible = True
End Sub

Private Sub CLOSE_Click()
TEXTABOUTSOFTWARE.Visible = False
TEXTABOUTDEVELOPER.Visible = False
TEXTABOUTOWNER.Visible = False
TEXTABOUTCOPYRIGHT.Visible = False
ABOUTUSFRAME.Visible = False
End Sub

Private Sub ELOGIN_Click()
EMPLOYEELOGIN.Show
End Sub

Private Sub END_Click()
End
End Sub

Private Sub NEWUSER_Click()
NEWUSERFORM.Show
End Sub

Private Sub OPENACCOUNT_Click()
EMPLOYEELOGIN.Show
End Sub

Private Sub ULOGIN_Click()
ACCOUNT.Show
End Sub
