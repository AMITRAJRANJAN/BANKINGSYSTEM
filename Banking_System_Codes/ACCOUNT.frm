VERSION 5.00
Begin VB.Form ACCOUNT 
   BackColor       =   &H00404000&
   Caption         =   "Login your account"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   16695
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
      Height          =   615
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox ACCOUNTDETAIL 
      BackColor       =   &H00404000&
      Height          =   2415
      Left            =   3000
      ScaleHeight     =   2355
      ScaleWidth      =   12795
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   12855
      Begin VB.TextBox ACCOUNTNUMBER 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MaxLength       =   9
         TabIndex        =   15
         ToolTipText     =   "Enter your 9 digit account number"
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox ACCOUNTPASSWORD 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "."
         TabIndex        =   14
         ToolTipText     =   "Enter your account password"
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton SHOWW 
         BackColor       =   &H0080C0FF&
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton HIDEE 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide Password"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "*     Account number"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "*     Account password"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   3735
      End
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
      Height          =   855
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
   End
   Begin VB.PictureBox DEBITCARDDETAIL 
      BackColor       =   &H00404000&
      Height          =   2415
      Left            =   3000
      ScaleHeight     =   2355
      ScaleWidth      =   12795
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   12855
      Begin VB.TextBox DEBITCARDNUMBER 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   7
         ToolTipText     =   "Enter your 16 digit Debit card number"
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox DEBITCARDPIN 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   4200
         MaxLength       =   4
         PasswordChar    =   "."
         TabIndex        =   6
         ToolTipText     =   "Enter your 4 digit Debit card PIN"
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton SHOWDPIN 
         BackColor       =   &H0080C0FF&
         Caption         =   "Show PIN"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton HIDEDPIN 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hide PIN"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "*     Debit card number"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "*     Debit card PIN"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   3735
      End
   End
   Begin VB.OptionButton DEBITCARDOPTION 
      BackColor       =   &H00FFFF00&
      Caption         =   "Start Banking via Debit Card"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      MaskColor       =   &H0000FF00&
      Picture         =   "ACCOUNT.frx":0000
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.OptionButton ACCOUNTOPTION 
      BackColor       =   &H00FFFF00&
      Caption         =   "Start Banking via Account Details"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      MaskColor       =   &H0000FF00&
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "If you are accessing your account for first time then you have to select ""Start Banking via Account details"" option"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16455
   End
End
Attribute VB_Name = "ACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim ANUMBER As Double
Dim APASSWORD As Double

Dim OP As Double
Dim CP As Double


Dim DNUMBER As Double
Dim OPIN As Integer
Dim CPIN As Integer



Private Sub ACCOUNTOPTION_Click()
If ACCOUNTOPTION.Value = True Then
    ACCOUNTDETAIL.Visible = True
    ACCOUNTNUMBER.Text = ""
    ACCOUNTPASSWORD.Text = ""
    SHOWW.Visible = True
    HIDEE.Visible = False
    DEBITCARDDETAIL.Visible = False
End If
End Sub


Private Sub DEBITCARDOPTION_Click()
If DEBITCARDOPTION.Value = True Then
    DEBITCARDDETAIL.Visible = True
    DEBITCARDNUMBER.Text = ""
    DEBITCARDPIN.Text = ""
    SHOWDPIN.Visible = True
    HIDEDPIN.Visible = False
    ACCOUNTDETAIL.Visible = False
End If
End Sub


Private Sub END_Click()
DEBITCARDNUMBER.Text = ""
DEBITCARDPIN.Text = ""
ACCOUNTNUMBER.Text = ""
ACCOUNTPASSWORD.Text = ""
ACCOUNTOPTION.Value = False
DEBITCARDOPTION.Value = False
ACCOUNTDETAIL.Visible = False
DEBITCARDDETAIL.Visible = False
ACCOUNTPASSWORD.PasswordChar = "."
DEBITCARDPIN.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
HIDEDPIN.Visible = False
SHOWDPIN.Visible = True

ACCOUNT.Hide
CTOHOME.Show

End Sub


Private Sub HIDEDPIN_Click()
DEBITCARDPIN.PasswordChar = "."
HIDEDPIN.Visible = False
SHOWDPIN.Visible = True
End Sub

Private Sub HIDEE_Click()
ACCOUNTPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
End Sub

Private Sub SHOWDPIN_Click()
DEBITCARDPIN.PasswordChar = ""
SHOWDPIN.Visible = False
HIDEDPIN.Visible = True
End Sub

Private Sub SHOWW_Click()
ACCOUNTPASSWORD.PasswordChar = ""
SHOWW.Visible = False
HIDEE.Visible = True
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1

If ACCOUNTOPTION.Value = True Then
    ANUMBER = ACCOUNTNUMBER.Text
    OP = ACCOUNTPASSWORD.Text
    CP = CONVERT_PASSWORD(OP)
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
    If RS.Fields(16).Value = -1 Then
        If RS.Fields(0).Value = ANUMBER And RS.Fields(7).Value = CP Then
            CTOACC.NAME = RS.Fields(1).Value
            CTOACC.ACCNUMBER = RS.Fields(0).Value
            ACCOUNTOPTION.Value = False
            DEBITCARDOPTION.Value = False
            ACCOUNTNUMBER.Text = ""
            ACCOUNTPASSWORD.Text = ""
            DEBITCARDNUMBER.Text = ""
            DEBITCARDPIN.Text = ""
            DEBITCARDDETAIL.Visible = False
            ACCOUNTDETAIL.Visible = False
        
            ACCOUNT.Hide
            TRANSACTIONS.Show
        Else
            MsgBox "OOPS!!! INVALID DATA", , "C t O Bank"
        End If
    ElseIf RS.Fields(16).Value = 0 Then
        MsgBox "Your account has been blocked. Please contact your branch.", , "C t O Bank"
    End If
        
End If

If DEBITCARDOPTION.Value = True Then
    DNUMBER = DEBITCARDNUMBER.Text
    DPIN = DEBITCARDPIN.Text
    CPIN = CONVERT_PIN(OPIN)
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_DEBITCARD=" & DNUMBER)
    If RS.Fields(16).Value = -1 Then
        If RS.Fields(12).Value = DNUMBER And RS.Fields(14).Value = CPIN Then
            CTOACC.NAME = RS.Fields(1).Value
            CTOACC.ACCNUMBER = RS.Fields(0).Value
            ACCOUNTOPTION.Value = False
            DEBITCARDOPTION.Value = False
            ACCOUNTNUMBER.Text = ""
            ACCOUNTPASSWORD.Text = ""
            DEBITCARDNUMBER.Text = ""
            DEBITCARDPIN.Text = ""
            DEBITCARDDETAIL.Visible = False
            ACCOUNTDETAIL.Visible = False
        
            ACCOUNT.Hide
            TRANSACTIONS.Show
        Else
            MsgBox "OOPS!!! INVALID DATA", , "C t O Bank"
        End If
    ElseIf RS.Fields(16).Value = 0 Then
        MsgBox "Your account has been blocked. Please contact your branch.", , "C t O Bank"
    End If
        
End If

Exit Sub
E1:
    MsgBox "Please enter a valid details", , "C t O Bank"
End Sub

Private Function CONVERT_PASSWORD(X As Double)
Dim RET As Double
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PASSWORD = RET
End Function

Private Function CONVERT_PIN(X As Integer)
Dim RET As Integer
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PIN = RET
End Function
