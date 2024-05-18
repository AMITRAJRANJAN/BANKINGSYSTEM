VERSION 5.00
Begin VB.Form LINKCREDITCARD 
   BackColor       =   &H00404000&
   Caption         =   "Credit card linking"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   8970
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
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
      Height          =   615
      Left            =   120
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Enter your 9 digit account number"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox CREDITCARDNUMBER 
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
      Height          =   615
      Left            =   4800
      MaxLength       =   16
      TabIndex        =   0
      ToolTipText     =   "Enter your 16 digit credit number"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Enter the detail for linking credit card"
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
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   6735
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
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "*    Credit card number"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "LINKCREDITCARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Double
Dim C As Double
Dim MSG As String



Private Sub CANCEL_Click()
ACCOUNTNUMBER.Text = ""
CREDITCARDNUMBER.Text = ""
LINKCREDITCARD.Hide
EMPLOYEEOPTION.Show
End Sub

Private Sub LINK_Click()
On Error GoTo E1

A = ACCOUNTNUMBER.Text
C = CREDITCARDNUMBER.Text

Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
RS.EDIT
RS.Fields(13).Value = C
RS.UPDATE
MSG = "CREDIT CARD " & C & " LINKED TO THE ACCOUNT NUMBER " & A
MsgBox MSG, , "C t O Bank"
ACCOUNTNUMBER.Text = ""
CREDITCARDNUMBER.Text = ""
LINKCREDITCARD.Hide
EMPLOYEEOPTION.Show

Exit Sub

E1:
    MsgBox "Invalid account number", , "C t O Bank"
End Sub

