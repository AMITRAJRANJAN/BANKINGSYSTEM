VERSION 5.00
Begin VB.Form NEWACCOUNTSTATUS 
   BackColor       =   &H00404000&
   Caption         =   "Status of the new account"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OK 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox OPENINGBALANCE 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Enter your 11 digit account number"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox TEMPORARYPASSWORD 
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "0"
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Your Account Number has been generated successfully"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   3015
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label CTONEWACCOUNTNAME 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label CTONEWACCOUNTNUMBER 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label CTONEWACCOUNTMOBILE 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "Mobile number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   "Temporary Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
   End
End
Attribute VB_Name = "NEWACCOUNTSTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim DBT As Database
Dim RST As Recordset

Dim D As Integer
Dim M As Integer
Dim Y As Integer


Private Sub Form_Load()
CTONEWACCOUNTNAME.Caption = NEWACC.NEWNAME
CTONEWACCOUNTNUMBER.Caption = NEWACC.NEWACCTNUMBER
CTONEWACCOUNTMOBILE.Caption = NEWACC.NEWMOBILE
End Sub

Private Sub OK_Click()
D = NEWACC.NEWDOBD
M = NEWACC.NEWDOBM
Y = NEWACC.NEWDOBY
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA")
RS.AddNew
RS.Fields(0).Value = NEWACC.NEWACCTNUMBER
RS.Fields(1).Value = NEWACC.NEWNAME
RS.Fields(2).Value = D & "/" & M & "/" & Y
RS.Fields(3).Value = NEWACC.NEWGENDER
RS.Fields(4).Value = NEWACC.NEWAADHAR

RS.Fields(5).Value = NEWACC.NEWMOBILE
RS.Fields(8).Value = OPENINGBALANCE.Text
RS.Fields(6).Value = NEWACC.NEWQUALIFICATION10 & ", " & NEWACC.NEWQUALIFICATION12 & ", " & NEWACC.NEWQUALIFICATIONG & ", " & NEWACC.NEWQUALIFICATIONPG
RS.Fields(7).Value = NEWACC.NEWACCTPASSWORD
RS.Fields(9).Value = NEWACC.NEWPAN
RS.Fields(10).Value = NEWACC.NEWPHOTO
RS.Fields(11).Value = NEWACC.NEWSIGNATURE
RS.Fields(16).Value = -1
RS.UPDATE

Set DBT = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\TRANSACTIONS.mdb")
Set RST = DBT.OpenRecordset("SELECT * FROM TRANSACTION")
RST.AddNew
RST.Fields(0).Value = NEWACC.NEWACCTNUMBER
RST.Fields(1).Value = NEWACC.NEWNAME
RST.UPDATE
MsgBox "Remember these account detail for future access to your account"
CTONEWACCOUNTNAME.Caption = ""
CTONEWACCOUNTNUMBER.Caption = ""
CTONEWACCOUNTMOBILE.Caption = ""
NEWACCOUNTSTATUS.Hide
EMPLOYEEOPTION.Show

End Sub

Private Sub OPENINGBALANCE_Click()
OPENINGBALANCE.Enabled = True
End Sub
