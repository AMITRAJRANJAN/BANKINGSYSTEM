VERSION 5.00
Begin VB.Form BENEFICIARYACCOUNT 
   BackColor       =   &H00404000&
   Caption         =   "Enter the details of Beneficiary"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8985
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton VALIDATE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Validate"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox AMOUNTTOBETRANSFERED 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
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
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Enter the amount you want to transfer"
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox BENEFICIARYMOBILENUMBER 
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
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Enter the beneficiary 10 digit mobile number"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox BENEFICIARYACCOUNTNUMBER 
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
      Left            =   4800
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Enter the beneficiary 9 digit account number"
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton TRANSFER 
      BackColor       =   &H0080C0FF&
      Caption         =   "Transfer"
      Enabled         =   0   'False
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label LABEL2 
      BackColor       =   &H00404000&
      Caption         =   "Enter the details of the beneficiary"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "*     Ammount to be transfered               "
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "* Beneficiary mobile number"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "* Beneficiary account number"
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
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "BENEFICIARYACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim DB2 As Database
Dim RS2 As Recordset
Dim RS22 As Recordset

Dim BNUMBER As Double

Dim VBNUMBER As Double

Dim BBALANCE As Long
Dim TAMOUNT As Integer
Dim ANUMBER As Double
Dim CBALANCE As Long
Dim MSG As VbMsgBoxResult


Private Sub CANCEL_Click()
BENEFICIARYACCOUNTNUMBER.Text = ""
BENEFICIARYMOBILENUMBER.Text = ""
AMOUNTTOBETRANSFERED.Text = ""
Label5.Enabled = False
AMOUNTTOBETRANSFERED.Enabled = False
TRANSFER.Visible = False
BENEFICIARYACCOUNT.Hide
TRANSACTIONS.Show

End Sub

Private Sub TRANSFER_Click()
On Error GoTo E2
ANUMBER = CTOACC.ACCNUMBER

CTOB.BACCOUNT = BENEFICIARYACCOUNTNUMBER.Text

BNUMBER = CTOB.BACCOUNT
Set DB2 = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS2 = DB2.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & BNUMBER)
Set RS22 = DB2.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
CBALANCE = RS22.Fields(8).Value
BBALANCE = RS2.Fields(8).Value
TAMOUNT = AMOUNTTOBETRANSFERED.Text
If TAMOUNT > CBALANCE Then
    MsgBox "You do not have sufficient funds to make this transaction.", , "C t O Bank"
Else
CTOB.BTAMOUNT = AMOUNTTOBETRANSFERED.Text
BBALANCE = BBALANCE + TAMOUNT
CBALANCE = CBALANCE - TAMOUNT
RS2.EDIT
RS2.Fields(8).Value = BBALANCE
RS2.UPDATE
RS22.EDIT
RS22.Fields(8).Value = CBALANCE
RS22.UPDATE
End If
BENEFICIARYACCOUNTNUMBER.Text = ""
BENEFICIARYMOBILENUMBER.Text = ""
AMOUNTTOBETRANSFERED.Text = ""
Label5.Enabled = False
AMOUNTTOBETRANSFERED.Enabled = False
BENEFICIARYACCOUNT.Hide
TRANSFERTRANSACTION.Show


Exit Sub
E2:
    MsgBox "Please enter the required details."

End Sub

Private Sub VALIDATE_Click()
On Error GoTo E1

VBNUMBER = BENEFICIARYACCOUNTNUMBER.Text
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & VBNUMBER)
MSG = MsgBox("You are transfering money to Mr." & RS.Fields(1).Value, vbYesNo + vbInformation + vbSystemModal, "C t O Bank: Beneficiary account validation")
If MSG = 6 Then
    Label5.Enabled = True
    AMOUNTTOBETRANSFERED.Enabled = True
    TRANSFER.Enabled = True
ElseIf MSG = 7 Then
    BENEFICIARYACCOUNTNUMBER.Text = ""
    BENEFICIARYMOBILENUMBER.Text = ""
    Label5.Enabled = False
    AMOUNTTOBETRANSFERED.Enabled = False
End If

Exit Sub
E1:
   MsgBox "Invalid beneficiary detail.", , "C t O Bank"

End Sub
