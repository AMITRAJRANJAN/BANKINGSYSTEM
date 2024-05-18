VERSION 5.00
Begin VB.Form TRANSACTIONS 
   BackColor       =   &H00FF00FF&
   Caption         =   "Start Banking"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CURRENTBALANCE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   15960
      Picture         =   "TRANSACTIONS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton MONEYTRANSFER 
      BackColor       =   &H0080C0FF&
      Caption         =   "Money Transfer"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   12120
      Picture         =   "TRANSACTIONS.frx":16E8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton DEPOSITMONEY 
      BackColor       =   &H0080C0FF&
      Caption         =   "Deposit Money"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      Picture         =   "TRANSACTIONS.frx":2C0B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton WITHDRAWMONEY 
      BackColor       =   &H0080C0FF&
      Caption         =   "Withdraw Money"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8280
      Picture         =   "TRANSACTIONS.frx":39F8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   2760
   End
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10080
      Width           =   1935
   End
   Begin VB.ComboBox ACCOUNTTRANSACTION 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "TRANSACTIONS.frx":46D3
      Left            =   13680
      List            =   "TRANSACTIONS.frx":46E3
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Account Transactions"
      Top             =   6240
      Width           =   6375
   End
   Begin VB.ComboBox ACCOUNTMANAGEMENT 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "TRANSACTIONS.frx":4727
      Left            =   4320
      List            =   "TRANSACTIONS.frx":4731
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Account Management"
      Top             =   6240
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "CTO"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label TDATE 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      Height          =   2775
      Left            =   4320
      Top             =   240
      Width           =   15735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      Height          =   10455
      Left            =   120
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label WELCOME 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
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
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3240
      Width           =   15735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Welcome  to  C  t  O  Bank"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   3840
      Width           =   15735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Start Banking"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   5160
      Width           =   15735
   End
End
Attribute VB_Name = "TRANSACTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TODAY As Variant
Dim I As Integer


Private Sub ACCOUNTMANAGEMENT_Click()
If ACCOUNTMANAGEMENT.ListIndex = 0 Then
    CHANGEPASSWORD.Show
End If
If ACCOUNTMANAGEMENT.ListIndex = 1 Then
    CHANGEPIN.Show

End If
End Sub

Private Sub ACCOUNTTRANSACTION_Click()
If ACCOUNTTRANSACTION.ListIndex = 0 Then
    BALANCE.Show
End If
If ACCOUNTTRANSACTION.ListIndex = 1 Then
    DEPOSITOR.Show
End If
If ACCOUNTTRANSACTION.ListIndex = 2 Then
    BENEFICIARYACCOUNT.Show

End If
If ACCOUNTTRANSACTION.ListIndex = 3 Then
    WITHDRAWL.Show
    
End If

End Sub

Private Sub BACK_Click()
TRANSACTIONS.Hide
ACCOUNT.Show
End Sub

Private Sub CURRENTBALANCE_Click()
BALANCE.Show
End Sub

Private Sub DEPOSITMONEY_Click()
DEPOSITOR.Show
End Sub

Private Sub END_Click()
WELCOME.Caption = ""
ACCOUNTMANAGEMENT.Text = "Account Management"
ACCOUNTTRANSACTION.Text = "Account Transaction"
TRANSACTIONS.Hide
CTOHOME.Show
End Sub

Private Sub Form_Load()
WELCOME.Caption = "Hello Mr." & CTOACC.NAME
End Sub


Private Sub MONEYTRANSFER_Click()
BENEFICIARYACCOUNT.Show

End Sub

Private Sub Timer1_Timer()
I = I + 1
TODAY = Now()
TDATE.Caption = Format(TODAY)

End Sub

Private Sub WITHDRAWMONEY_Click()
WITHDRAWL.Show
End Sub
