VERSION 5.00
Begin VB.Form DEPOSITOR 
   BackColor       =   &H00404000&
   Caption         =   "Enter the amount you want to deposit"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   8880
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
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox DEPOSITAMOUNT 
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
      ToolTipText     =   "Enter the amount you want to be deposited in your account"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton DEPOSIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Deposit"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Deposit Amount"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Enter amont you want to deposit"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "DEPOSITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim DBT As Database
Dim RST As Recordset

Dim N As Integer

Dim ANUMBER As Double
Dim CURRENTBALANCE As Long
Dim DEPOSITBALANCE As Long


Private Sub DEPOSIT_Click()
On Error GoTo E1
ANUMBER = CTOACC.ACCNUMBER
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
CURRENTBALANCE = RS.Fields(8).Value
DEPOSITBALANCE = DEPOSITAMOUNT.Text
CTOACCT.DEPOSIT = DEPOSITAMOUNT.Text
CURRENTBALANCE = CURRENTBALANCE + DEPOSITBALANCE
RS.EDIT
RS.Fields(8).Value = CURRENTBALANCE
RS.UPDATE

'Set DBT = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\TRANSACTIONS.mdb")
'Set RST = DBT.OpenRecordset("SELECT * FROM TRANSACTION WHERE ACCOUNT_NUMBER=" & ANUMBER)
'For N = 1 To 10
'    If RST.Fields(N + 1).Value = 0 Then
'       RST.Fields(N + 1).Value = DEPOSITBALANCE
'    End If
'Next N
    
    

DEPOSITAMOUNT.Text = ""
DEPOSITOR.Hide
DEPOSITTRANSACTION.Show


Exit Sub
E1:
    MsgBox "Please enter a valid amount in rupees.", , "C t O Bank"

End Sub

Private Sub CANCEL_Click()
DEPOSITAMOUNT.Text = ""
DEPOSITOR.Hide
TRANSACTIONS.Show
End Sub
