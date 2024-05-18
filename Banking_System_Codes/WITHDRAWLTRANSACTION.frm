VERSION 5.00
Begin VB.Form WITHDRAWLTRANSACTION 
   BackColor       =   &H00404000&
   Caption         =   "Status of your current transaction"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
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
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
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
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      Caption         =   "Your transaction is successful"
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
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label BALANCE 
      BackColor       =   &H0080FF80&
      Caption         =   "Rs.   "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label WITHDRAWLAMOUNT 
      BackColor       =   &H0080FF80&
      Caption         =   "Rs.   "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "Withdrawl Amount"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "WITHDRAWLTRANSACTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim ANUMBER As Double


Private Sub Form_Load()
WITHDRAWLAMOUNT.Caption = "Rs. " & CTOACCT.WITHDRAW
ANUMBER = CTOACC.ACCNUMBER
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
BALANCE.Caption = "Rs. " & RS.Fields(8).Value
End Sub

Private Sub OK_Click()
WITHDRAWLAMOUNT.Caption = "Rs. "
BALANCE.Caption = "Rs. "
WITHDRAWLTRANSACTION.Hide
TRANSACTIONS.Show
End Sub

Private Sub REFRESH_Click()
WITHDRAWLAMOUNT.Caption = "Rs. " & CTOACCT.WITHDRAW
ANUMBER = CTOACC.ACCNUMBER
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
BALANCE.Caption = "Rs. " & RS.Fields(8).Value
End Sub
