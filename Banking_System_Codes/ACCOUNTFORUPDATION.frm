VERSION 5.00
Begin VB.Form ACCOUNTFORUPDATION 
   BackColor       =   &H00404000&
   Caption         =   "Search the account for updation"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form2"
   ScaleHeight     =   3495
   ScaleWidth      =   9060
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
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox MOBILENUMBERT 
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
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   4
      ToolTipText     =   "Enter your 5 digit Mobile number"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.OptionButton SEARCHVIAMOBILE 
      BackColor       =   &H00FFFF00&
      Caption         =   "Search via mobile number"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.OptionButton SEARCHVIAACCOUNT 
      BackColor       =   &H00FFFF00&
      Caption         =   "Search via account number"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox ACCOUNTNUMBERT 
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
      Left            =   4920
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Enter your 9 digit account number"
      Top             =   1440
      Width           =   3855
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
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label MOBILENUMBERL 
      BackColor       =   &H0080FF80&
      Caption         =   "* Enter the mobile number"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label ACCOUNTNUMBERL 
      BackColor       =   &H0080FF80&
      Caption         =   "* Enter the account number"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label LABEL2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Search the account for updation"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "ACCOUNTFORUPDATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Double
Dim M As Double


Private Sub CANCEL_Click()
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False

    
ACCOUNTFORUPDATION.Hide
EMPLOYEEOPTION.Show
End Sub

Private Sub SEARCHVIAACCOUNT_Click()
If SEARCHVIAACCOUNT.Value = True Then
    ACCOUNTNUMBERL.Enabled = True
    ACCOUNTNUMBERT.Enabled = True
    ACCOUNTNUMBERT.Text = ""
    SUBBMIT.Visible = True
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    MOBILENUMBERT.Text = ""
End If
End Sub

Private Sub SEARCHVIAMOBILE_Click()
If SEARCHVIAMOBILE.Value = True Then
    MOBILENUMBERL.Enabled = True
    MOBILENUMBERT.Enabled = True
    MOBILENUMBERT.Text = ""
    SUBBMIT.Visible = True
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    ACCOUNTNUMBERT.Text = ""

End If
End Sub

Private Sub Form_Load()
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False

End Sub


Private Sub SUBBMIT_Click()
On Error GoTo E1

If SEARCHVIAACCOUNT.Value = True Then
    A = ACCOUNTNUMBERT.Text
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
    CTOUP.ACCNUMBER = RS.Fields(0).Value
    CTOUP.NAME = RS.Fields(1).Value
    CTOUP.GENDER = RS.Fields(3).Value
    CTOUP.MOBILE = RS.Fields(5).Value
    CTOUP.PAN = RS.Fields(9).Value
    CTOUP.AADHAR = RS.Fields(4).Value
    CTOUP.PHOTO = RS.Fields(10).Value
    CTOUP.SIGNATURE = RS.Fields(11).Value
    CTOUP.BALANCE = RS.Fields(8).Value
    CTOUP.DOB = RS.Fields(2).Value
    
    ACCOUNTFORUPDATION.Hide
    ACCOUNTUPDATION.Show
    
ElseIf SEARCHVIAMOBILE.Value = True Then
    M = MOBILENUMBERT.Text
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_MOBILE=" & M)
    CTOUP.ACCNUMBER = RS.Fields(0).Value
    CTOUP.NAME = RS.Fields(1).Value
    CTOUP.GENDER = RS.Fields(3).Value
    CTOUP.MOBILE = RS.Fields(5).Value
    CTOUP.PAN = RS.Fields(9).Value
    CTOUP.AADHAR = RS.Fields(4).Value
    CTOUP.PHOTO = RS.Fields(10).Value
    CTOUP.SIGNATURE = RS.Fields(11).Value
    CTOUP.BALANCE = RS.Fields(8).Value
    CTOUP.DOB = RS.Fields(2).Value
    
    ACCOUNTFORUPDATION.Hide
    ACCOUNTUPDATION.Show
    
End If

Exit Sub
E1:
   MsgBox "Account does not exist.", , "C t O Bank"
End Sub


