VERSION 5.00
Begin VB.Form NEWUSERFORM 
   BackColor       =   &H00404000&
   Caption         =   "Password status"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   9225
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
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
      TabIndex        =   10
      Top             =   3120
      Width           =   1575
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox NUSERACCOUNT 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MaxLength       =   9
      TabIndex        =   5
      ToolTipText     =   "Enter your 9- digitS account number"
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton PASSWORDSTATUS 
      BackColor       =   &H0080C0FF&
      Caption         =   "Check status of your account password"
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
      TabIndex        =   4
      Top             =   960
      Width           =   6975
   End
   Begin VB.TextBox NUSERPASSWORD 
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
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "."
      TabIndex        =   3
      ToolTipText     =   "Create a account password"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox NUSERMOBILE 
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
      Height          =   510
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Enter your 10-digits mobile number"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3375
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Create password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Mobile number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "NEWUSERFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim NEWACCOUNTNUMBER As Double
Dim NEWMOBILENUMBER As Double
Dim NEWPASSWORD As Double
Dim OP As Double
Dim CP As Double

Private Sub BACK_Click()
NUSERACCOUNT.Text = ""
NUSERMOBILE.Text = ""
NUSERPASSWORD.Text = ""
NUSERPASSWORD.PasswordChar = "."
SHOWW.Visible = True
HIDEE.Visible = False
Label4.Visible = False
NUSERMOBILE.Visible = False
NUSERPASSWORD.Visible = False
Label5.Visible = False
SUBBMIT.Visible = False

NEWUSERFORM.Hide
CTOHOME.Show

End Sub

Private Sub END_Click()
NUSERACCOUNT.Text = ""
NUSERMOBILE.Text = ""
NUSERPASSWORD.Text = ""
NUSERPASSWORD.PasswordChar = "."
SHOWW.Visible = True
HIDEE.Visible = False
Label4.Visible = False
NUSERMOBILE.Visible = False
NUSERPASSWORD.Visible = False
Label5.Visible = False
SUBBMIT.Visible = False

NEWUSERFORM.Hide
CTOHOME.Show

End Sub

Private Sub HIDEE_Click()
NUSERPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True

End Sub

Private Sub PASSWORDSTATUS_Click()
On Error GoTo E1
NEWACCOUNTNUMBER = NUSERACCOUNT.Text
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & NEWACCOUNTNUMBER)
If RS.Fields(7).Value = 0 And RS.Fields(16).Value = -1 Then
    MsgBox "You are a new user. Please create your password first.", , "C t O Bank"
    Label4.Visible = True
    NUSERMOBILE.Visible = True
    NUSERPASSWORD.Visible = True
    Label5.Visible = True
    SUBBMIT.Visible = True
ElseIf RS.Fields(7).Value <> 0 And RS.Fields(16).Value = -1 Then
    MsgBox "You are not a new user. Please click on New User user option.", , "C t O Bank"
ElseIf RS.Fields(7).Value <> 0 And RS.Fields(16).Value = 0 Then
    MsgBox "Your account has been blocked. Please contact your bank", , "C t O Bank"
End If
Exit Sub
E1:
    MsgBox "Please enter a valid account number.", , "C t O Bank"
End Sub

Private Sub SHOWW_Click()
NUSERPASSWORD.PasswordChar = ""
SHOWW.Visible = False
HIDEE.Visible = True
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E2
NEWACCOUNTNUMBER = NUSERACCOUNT.Text
NEWMOBILENUMBER = NUSERMOBILE.Text
OP = NUSERPASSWORD.Text
CP = CONVERT_PASSWORD(OP)
NEWPASSWORD = CP

Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & NEWACCOUNTNUMBER)
If RS.Fields(0).Value = NEWACCOUNTNUMBER And RS.Fields(5).Value = NEWMOBILENUMBER Then
    RS.EDIT
    RS.Fields(7).Value = NEWPASSWORD
    RS.UPDATE
    MsgBox "You have successfully created your account password. Please login using your account details.", , "C t O Bank"
    NUSERACCOUNT.Text = ""
    NUSERMOBILE.Text = ""
    NUSERPASSWORD.Text = ""
    NUSERPASSWORD.PasswordChar = "."
    SHOWW.Visible = True
    HIDEE.Visible = False
    Label4.Visible = False
    NUSERMOBILE.Visible = False
    NUSERPASSWORD.Visible = False
    Label5.Visible = False
    SUBBMIT.Visible = False

    NEWUSERFORM.Hide
    CTOHOME.Show
Else
    MsgBox "Please enter your valid account number and mobile number.", , "C t O Bank"
End If
Exit Sub
E2:
    MsgBox "Please enter a valid mobile number linked to your account and a password of your choice.", , "C t O Bank"

End Sub

Private Function CONVERT_PASSWORD(X As Double)
Dim RET As Double
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PASSWORD = RET
End Function

