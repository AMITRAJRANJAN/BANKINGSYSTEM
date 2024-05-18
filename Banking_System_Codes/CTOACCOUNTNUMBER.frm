VERSION 5.00
Begin VB.Form CTOACCOUNTNUMBER 
   BackColor       =   &H00404000&
   Caption         =   "Account Number"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton CLICKTOADDAZERO 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to add a zero"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton VALIDATE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Validate"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox NEWACCOUNTNUMBER 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MaxLength       =   9
      TabIndex        =   0
      ToolTipText     =   "Enter a account number of  7-digits  according to your branch specified format"
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Validate this auto generated account number"
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
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Account number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "CTOACCOUNTNUMBER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NEWACNUMBER As Double
Dim ACCOUNTNUMBER As String
Dim CHECK1_ACCOUNTNUMBER As Double
Dim CHECK2_ACCOUNTNUMBER As String

Dim L As Integer

Dim D As Integer
Dim DD As String
Dim M As String
Dim MM As String
Dim A As String
Dim AA As String



Private Sub CLICKTOADDAZERO_Click()
NEWACCOUNTNUMBER.Enabled = True
End Sub

Private Sub Form_Load()
D = NEWACC.NEWDOBD
DD = Str(D)
M = Str(NEWACC.NEWMOBILE)
MM = Mid(M, 3, 3)
A = Str(NEWACC.NEWAADHAR)
AA = Mid(A, 3, 4)

ACCOUNTNUMBER = DD + MM + AA

NEWACNUMBER = CDbl(ACCOUNTNUMBER)

NEWACCOUNTNUMBER.Text = NEWACNUMBER

NEWACC.NEWACCTPASSWORD = 0

End Sub

Private Sub SUBBMIT_Click()
CHECK1_ACCOUNTNUMBER = NEWACCOUNTNUMBER.Text
CHECK2_ACCOUNTNUMBER = Str(CHECK1_ACCOUNTNUMBER)
L = Len(CHECK2_ACCOUNTNUMBER) - 1
If L < 9 Then
    MsgBox "Account Number must be of 9-digits.", , "C t O Bank"
Else
    NEWACC.NEWACCTNUMBER = NEWACCOUNTNUMBER.Text
    NEWACCOUNTNUMBER.Text = ""
    NEWACCOUNTNUMBER.Enabled = False
    CTOACCOUNTNUMBER.Hide
    NEWACCOUNTSTATUS.Show
End If

End Sub

Private Sub VALIDATE_Click()
CHECK1_ACCOUNTNUMBER = NEWACCOUNTNUMBER.Text
CHECK2_ACCOUNTNUMBER = Str(CHECK1_ACCOUNTNUMBER)
L = Len(CHECK2_ACCOUNTNUMBER) - 1
If L < 9 Then
    MsgBox "Account Number must be of 9-digits.", , "C t O Bank"
Else
    MsgBox "Click on SUBBMIT button", , "C t O Bank"
End If

End Sub
