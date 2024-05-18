VERSION 5.00
Begin VB.Form WITHDRAWL 
   BackColor       =   &H00404000&
   Caption         =   "Enter the amount you want to withdraw"
   ClientHeight    =   2745
   ClientLeft      =   3555
   ClientTop       =   2610
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   8835
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox WITHDRAWAMOUNT 
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
      Height          =   735
      Left            =   4560
      MaxLength       =   11
      TabIndex        =   1
      ToolTipText     =   "Enter your 11 digit account number"
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton WITHDRAW 
      BackColor       =   &H0080C0FF&
      Caption         =   "Withdraw"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Enter amont you want to withdraw"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label3 
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
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "WITHDRAWL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim ANUMBER As Double
Dim CURRENTBALANCE As Long
Dim WITHDRAWBALANCE As Long


Private Sub CANCEL_Click()
WITHDRAWAMOUNT.Text = ""
WITHDRAWL.Hide
TRANSACTIONS.Show

End Sub

Private Sub WITHDRAW_Click()
On Error GoTo E1
ANUMBER = CTOACC.ACCNUMBER
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
CURRENTBALANCE = RS.Fields(8).Value
WITHDRAWBALANCE = WITHDRAWAMOUNT.Text
If WITHDRAWBALANCE > CURRENTBALANCE Then
    MsgBox "Please enter a amount less than your current balance.", , "C t O Bank"
Else
CTOACCT.WITHDRAW = WITHDRAWAMOUNT.Text
CURRENTBALANCE = CURRENTBALANCE - WITHDRAWBALANCE
RS.EDIT
RS.Fields(8).Value = CURRENTBALANCE
RS.UPDATE
End If
WITHDRAWAMOUNT.Text = ""
WITHDRAWL.Hide
WITHDRAWLTRANSACTION.Show

Exit Sub

E1:
    MsgBox "Please enter a valid amount in rupees.", , "C t O Bank"

End Sub
