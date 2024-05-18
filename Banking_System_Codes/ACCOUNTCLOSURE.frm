VERSION 5.00
Begin VB.Form ACCOUNTCLOSURE 
   BackColor       =   &H00404000&
   Caption         =   "Review and close your account"
   ClientHeight    =   7665
   ClientLeft      =   270
   ClientTop       =   465
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11175
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton CLOSE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close this account"
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
      TabIndex        =   18
      Top             =   6840
      Width           =   4335
   End
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox GENDER 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   15
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox ACCOUNT 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox BALANCE 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox NAME2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   4
      Top             =   3480
      Width           =   6855
   End
   Begin VB.TextBox AGE 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox AADHAR 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   2
      Top             =   4920
      Width           =   6855
   End
   Begin VB.TextBox PAN 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   6855
   End
   Begin VB.TextBox MOBILE 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   0
      Top             =   5880
      Width           =   6855
   End
   Begin VB.Image SIGNATURE 
      Height          =   855
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Image PHOTO 
      Height          =   2055
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Review and update your account detail"
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
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   10575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Balance Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   " Aadhar number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Pan number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   3255
   End
End
Attribute VB_Name = "ACCOUNTCLOSURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim ACC As Double
Dim MSG As VbMsgBoxResult


Private Sub CANCEL_Click()
NAME2.Text = ""
AGE.Text = ""
GENDER.Text = ""
MOBILE.Text = ""
PAN.Text = ""
AADHAR.Text = ""
PHOTO.Picture = LoadPicture("")
SIGNATURE.Picture = LoadPicture("")
BALANCE.Text = ""
MOBILE.Enabled = False
AADHAR.Enabled = False
PAN.Enabled = False
AGE.Enabled = False
GENDER.Enabled = False

ACCOUNTCLOSURE.Hide
EMPLOYEEOPTION.Show


End Sub

Private Sub CLOSE_Click()
ACC = ACCOUNT.Text
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ACC)
MSG = MsgBox("Are you sure you want to close this account", vbYesNoCancel + vbExclamation + vbDefaultButton3 + vbSystemModal, "C t O Bank: Account Closure")
If MSG = 6 Then
    RS.Delete
    NAME2.Text = ""
    GENDER.Text = ""
    AGE.Text = ""
    MOBILE.Text = ""
    PAN.Text = ""
    AADHAR.Text = ""
    PHOTO.Picture = LoadPicture("")
    SIGNATURE.Picture = LoadPicture("")
    MsgBox ("This account has been deleted permanently")
    ACCOUNTCLOSURE.Hide
    EMPLOYEEOPTION.Show
    
ElseIf MSG = 7 Then
    NAME2.Text = ""
    GENDER.Text = ""
    AGE.Text = ""
    MOBILE.Text = ""
    PAN.Text = ""
    AADHAR.Text = ""
    PHOTO.Picture = LoadPicture("")
    SIGNATURE.Picture = LoadPicture("")
    ACCOUNTCLOSUREDETAIL.Hide
    EMPLOYEEOPTION.Show
ElseIf MSG = 2 Then
    ACCOUNTCLOSUREDETAIL.Show
End If

End Sub

Private Sub Form_Load()
ACCOUNT.Text = CTOUP.ACCNUMBER
NAME2.Text = CTOUP.NAME
AGE.Text = CTOUP.DOB
GENDER.Text = CTOUP.GENDER
MOBILE.Text = CTOUP.MOBILE
PAN.Text = CTOUP.PAN
AADHAR.Text = CTOUP.AADHAR
BALANCE.Text = CTOUP.BALANCE
PHOTO.Picture = LoadPicture(CTOUP.PHOTO)
SIGNATURE.Picture = LoadPicture(CTOUP.SIGNATURE)
End Sub

Private Sub REFRESH_Click()
ACCOUNT.Text = CTOUP.ACCNUMBER
NAME2.Text = CTOUP.NAME
AGE.Text = CTOUP.DOB
GENDER.Text = CTOUP.GENDER
MOBILE.Text = CTOUP.MOBILE
PAN.Text = CTOUP.PAN
AADHAR.Text = CTOUP.AADHAR
BALANCE.Text = CTOUP.BALANCE
PHOTO.Picture = LoadPicture(CTOUP.PHOTO)
SIGNATURE.Picture = LoadPicture(CTOUP.SIGNATURE)

End Sub

