VERSION 5.00
Begin VB.Form LINKDEBITCARD 
   BackColor       =   &H00404000&
   Caption         =   "Debit card linking"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   9270
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
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
      Height          =   510
      Left            =   120
      MaxLength       =   9
      TabIndex        =   2
      ToolTipText     =   "Enter your 9 digit account number"
      Top             =   1440
      Width           =   4215
   End
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
      Height          =   510
      Left            =   4920
      MaxLength       =   16
      TabIndex        =   1
      ToolTipText     =   "Enter your 16 digit debit card number"
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton LINK 
      BackColor       =   &H0080C0FF&
      Caption         =   "            LINK           <::::::::::::>"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
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
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label5 
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
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Enter the detail for linking debit card"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "LINKDEBITCARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Double
Dim D As Double
Dim MSG As String


Private Sub CANCEL_Click()
ACCOUNTNUMBER.Text = ""
DEBITCARDNUMBER.Text = ""
LINKDEBITCARD.Hide
EMPLOYEEOPTION.Show
End Sub

Private Sub LINK_Click()
On Error GoTo E1

A = ACCOUNTNUMBER.Text
D = DEBITCARDNUMBER.Text

Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
RS.EDIT
RS.Fields(12).Value = D
RS.UPDATE
MSG = "DEBIT CARD " & D & " LINKED TO THE ACCOUNT NUMBER " & A
MsgBox MSG, , "C t O Bank"
ACCOUNTNUMBER.Text = ""
DEBITCARDNUMBER.Text = ""
LINKDEBITCARD.Hide
EMPLOYEEOPTION.Show
Exit Sub

E1:
    MsgBox "Invalid account number", , "C t O Bank"

End Sub
