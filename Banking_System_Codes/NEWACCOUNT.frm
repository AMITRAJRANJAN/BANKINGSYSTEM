VERSION 5.00
Begin VB.Form NEWACCOUNT 
   BackColor       =   &H00404000&
   Caption         =   "Enter the details for new account"
   ClientHeight    =   5880
   ClientLeft      =   5670
   ClientTop       =   3105
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton BACK 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton SUBBMIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Subbmit"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton RESET 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Frame DOBF 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   12015
      Begin VB.ComboBox Y 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "NEWACCOUNT.frx":0000
         Left            =   8760
         List            =   "NEWACCOUNT.frx":0043
         TabIndex        =   19
         Text            =   "Year"
         Top             =   0
         Width           =   1935
      End
      Begin VB.ComboBox M 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "NEWACCOUNT.frx":00C5
         Left            =   6240
         List            =   "NEWACCOUNT.frx":00ED
         TabIndex        =   18
         Text            =   "Month"
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox D 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "NEWACCOUNT.frx":0121
         Left            =   3960
         List            =   "NEWACCOUNT.frx":0182
         TabIndex        =   17
         Text            =   "Day"
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "*     Date of birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox NMOBILE 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   15
      ToolTipText     =   "Enter 10 digit mobile number"
      Top             =   4320
      Width           =   8055
   End
   Begin VB.Frame QUALIFICATIONF 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   12015
      Begin VB.CheckBox POSTGRADUATE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Post Graduate"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
      Begin VB.CheckBox GRADUATE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Graduate"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox TWELVE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Intermediate"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   11
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox TEN 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Matric"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   10
         Top             =   0
         Width           =   1335
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
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Frame GENDERF 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   12015
      Begin VB.OptionButton TRANSGENDER 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Transgender"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
      Begin VB.OptionButton FEMALE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   6
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton MALE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox NPAN 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "Enter 10 digit PAN number"
      Top             =   3720
      Width           =   8055
   End
   Begin VB.TextBox NAADHAR 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4200
      MaxLength       =   12
      TabIndex        =   2
      ToolTipText     =   "Enter 12 digit AADHAR number"
      Top             =   3120
      Width           =   8055
   End
   Begin VB.TextBox NNAME 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Enter name "
      Top             =   720
      Width           =   8055
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label LABEL6 
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label LABEL5 
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label LABEL4 
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Enter the id details for New Account"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "NEWACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BACK_Click()
NEWACCOUNT.Hide
EMPLOYEEOPTION.Show

End Sub

Private Sub END_Click()
NNAME.Text = ""
D.Text = "Day"
M.Text = "Month"
Y.Text = "Year"
MALE.Value = False
FEMALE.Value = False
TRANSGENDER.Value = False
TEN.Value = 0
TWELVE.Value = 0
GRADUATE.Value = 0
POSTGRADUATE.Value = 0
NAADHAR.Text = ""
NPAN.Text = ""
NMOBILE.Text = ""
NEWACCOUNT.Hide
EMPLOYEEOPTION.Show
End
End Sub

Private Sub HOME_Click()
NNAME.Text = ""
NAADHAR.Text = ""
NPAN.Text = ""
NMOBILE.Text = ""
MALE.Value = False
FEMALE.Value = False
TRANSGENDER.Value = False
TEN.Value = 0
TWELVE.Value = 0
GRADUATE.Value = 0
POSTGRADUATE.Value = 0
D.Text = ""
M.Text = ""
Y.Text = ""
NEWACCOUNT.Hide
EMPLOYEEOPTION.Hide
CTOHOME.Show
End Sub

Private Sub LOGOUT_Click()
NNAME.Text = ""
NAADHAR.Text = ""
NPAN.Text = ""
NMOBILE.Text = ""
MALE.Value = False
FEMALE.Value = False
TRANSGENDER.Value = False
TEN.Value = 0
TWELVE.Value = 0
GRADUATE.Value = 0
POSTGRADUATE.Value = 0
D.Text = ""
M.Text = ""
Y.Text = ""
NEWACCOUNT.Hide
EMPLOYEEOPTION.Hide
CTOHOME.Show
End Sub


Private Sub RESET_Click()
NNAME.Text = ""
NAADHAR.Text = ""
NPAN.Text = ""
NMOBILE.Text = ""
MALE.Value = False
FEMALE.Value = False
TRANSGENDER.Value = False
TEN.Value = 0
TWELVE.Value = 0
GRADUATE.Value = 0
POSTGRADUATE.Value = 0
D.Text = ""
M.Text = ""
Y.Text = ""
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1
NEWACC.NEWNAME = NNAME.Text
NEWACC.NEWAADHAR = NAADHAR.Text
NEWACC.NEWPAN = NPAN.Text
NEWACC.NEWMOBILE = NMOBILE.Text


If MALE.Value = True Then
    NEWACC.NEWGENDER = "Male"
ElseIf FEMALE.Value = True Then
    NEWACC.NEWGENDER = "Female"
ElseIf TRANSGENDER.Value = True Then
    NEWACC.NEWGENDER = "Transgender"
End If

If TEN.Value = 1 Then
    NEWACC.NEWQUALIFICATION10 = "Matriculation"
Else
    NEWACC.NEWQUALIFICATION10 = ""
End If
If TWELVE.Value = 1 Then
    NEWACC.NEWQUALIFICATION12 = "Intermediate"
Else
    NEWACC.NEWQUALIFICATION12 = ""
End If
If GRADUATE.Value = 1 Then
    NEWACC.NEWQUALIFICATIONG = "Graduated"
Else
    NEWACC.NEWQUALIFICATIONG = ""
End If
If POSTGRADUATE.Value = 1 Then
    NEWACC.NEWQUALIFICATIONPG = "Postgraduated"
Else
    NEWACC.NEWQUALIFICATIONPG = ""
End If

If D.Text > 31 Or M.Text > 12 Or Y.Text > 2020 Then
    MsgBox "Please enter a valid date of birth.", , "C t O Bank"
Else
    NEWACC.NEWDOBD = D.Text
    NEWACC.NEWDOBM = M.Text
    NEWACC.NEWDOBY = Y.Text
    NNAME.Text = ""
    NAADHAR.Text = ""
    NPAN.Text = ""
    NMOBILE.Text = ""
    MALE.Value = False
    FEMALE.Value = False
    TRANSGENDER.Value = False
    TEN.Value = 0
    TWELVE.Value = 0
    GRADUATE.Value = 0
    POSTGRADUATE.Value = 0
    D.Text = ""
    M.Text = ""
    Y.Text = ""

    NEWACCOUNT.Hide
    PHOTOSIGNATURE.Show
    
End If


Exit Sub

E1:
    MsgBox "All fields are mandatory.", , "C t O Bank"

End Sub
