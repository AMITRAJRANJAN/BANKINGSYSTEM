VERSION 5.00
Begin VB.Form CHANGEPIN 
   BackColor       =   &H00404000&
   Caption         =   "Change your debit card pin"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
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
   Begin VB.CommandButton SHOWPIN 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show Pin"
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
      Top             =   2280
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
   Begin VB.TextBox CURRENTPIN 
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
      ToolTipText     =   "Enter your current debit card pin"
      Top             =   1080
      Width           =   3405
   End
   Begin VB.TextBox NEWPIN 
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
      ToolTipText     =   "Enter your new pin "
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton HIDEPIN 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Pin"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "* Current pin"
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
      Caption         =   "* New pin"
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
      Caption         =   "Change your debit card pin"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "CHANGEPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim OPIN As Integer
Dim CPIN As Integer

Dim NEWP As Integer
Dim CNEWP As Integer


Dim ANUMBER As Double

Private Sub CANCEL_Click()
NEWPIN.Text = ""
CURRENTPIN.Text = ""
NEWPIN.PasswordChar = "."
HIDEPIN.Visible = False
SHOWPIN.Visible = True
CHANGEPIN.Hide
TRANSACTIONS.Show

End Sub



Private Sub HIDEPIN_Click()
NEWPIN.PasswordChar = "."
HIDEPIN.Visible = False
SHOWPIN.Visible = True
End Sub

Private Sub SHOWPIN_Click()
NEWPIN.PasswordChar = ""
HIDEPIN.Visible = True
SHOWPIN.Visible = False
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1
ANUMBER = CTOACC.ACCNUMBER
OPIN = CURRENTPIN.Text
NEWP = NEWPIN.Text
CPIN = CONVERT_PIN(OPIN)
CNEWP = CONVERT_PIN(NEWP)
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHER ACCOUNT_NUMBER=" & ANUMBER)
If RS.Fields(14).Value = CPIN Then
    RS.EDIT
    RS.Fields(14).Value = CNEWP
    RS.UPDATE
    MsgBox "Your pin has been updated", , "C t O Bank"
    CURRENTPIN.Text = ""
    NEWPIN.Text = ""
    CHANGEPIN.Hide
    TRANSACTIONS.Show
Else
    MsgBox "Please enter a valid current pin", , "C t O Bank"
End If
Exit Sub
E1:
    MsgBox "All fields are mandatory", , "C t O Bank"
End Sub


Private Function CONVERT_PIN(X As Integer)
Dim RET As Integer
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PIN = RET
End Function

