VERSION 5.00
Begin VB.Form CHANGEPASSWORD 
   BackColor       =   &H00404000&
   Caption         =   "Change your account password"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CANCEL 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
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
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox CURRENTPASSWORD 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      MaxLength       =   9
      TabIndex        =   2
      ToolTipText     =   "Enter your current accoount password"
      Top             =   1080
      Width           =   3405
   End
   Begin VB.TextBox NEWPASSWORD 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "."
      TabIndex        =   1
      ToolTipText     =   "Enter your new password"
      Top             =   2040
      Width           =   3375
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "* Current password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "* New password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Change your account password"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "CHANGEPASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim OP As Double
Dim CP As Double

Dim NEWP As Double
Dim CNEWP As Double


Dim ANUMBER As Double


Private Sub CANCEL_Click()
NEWPASSWORD.Text = ""
NEWPASSWORD.PasswordChar = "."
CURRENTPASSWORD.Text = ""
HIDEE.Visible = False
SHOWW.Visible = True
CHANGEPASSWORD.Hide
TRANSACTIONS.Show

End Sub

Private Sub HIDEE_Click()
NEWPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
End Sub

Private Sub SHOWW_Click()
NEWPASSWORD.PasswordChar = ""
SHOWW.Visible = False
HIDEE.Visible = True
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1
ANUMBER = CTOACC.ACCNUMBER
OP = CURRENTPASSWORD.Text
NEWP = NEWPASSWORD.Text
CP = CONVERT_PASSWORD(OP)
CNEWP = CONVERT_PASSWORD(NEWP)
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
If RS.Fields(7).Value = CP Then
    RS.EDIT
    RS.Fields(7).Value = CNEWP
    RS.UPDATE
    MsgBox "Your password has been updated", , "C t O Bank"
    CURRENTPASSWORD.Text = ""
    NEWPASSWORD.Text = ""
    CHANGEPASSWORD.Hide
    TRANSACTIONS.Show
    
Else
    MsgBox "Please enter a valid current password", , "C t O Bank"
End If
Exit Sub
E1:
    MsgBox "All fields are mandatory", , "C t O Bank"
End Sub


Private Function CONVERT_PASSWORD(X As Double)
Dim RET As Double
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PASSWORD = RET
End Function

