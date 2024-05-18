VERSION 5.00
Begin VB.Form ACCOUNTUPDATION 
   BackColor       =   &H00404000&
   Caption         =   "Review and update your account details"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   315
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11100
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   1575
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
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton UPDATEDETAILS 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update details"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton UPDATE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6960
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
      Left            =   3840
      TabIndex        =   16
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
      Left            =   3840
      TabIndex        =   7
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
      Left            =   3840
      TabIndex        =   6
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
      Left            =   3840
      TabIndex        =   5
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
      Left            =   3840
      TabIndex        =   4
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
      Left            =   3840
      TabIndex        =   3
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
      Left            =   3840
      TabIndex        =   2
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
      Left            =   3840
      TabIndex        =   1
      Top             =   5880
      Width           =   6855
   End
   Begin VB.CommandButton BACK 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1695
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
      Left            =   240
      TabIndex        =   17
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
      Left            =   240
      TabIndex        =   15
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
      Left            =   240
      TabIndex        =   14
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
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Image SIGNATURE 
      Height          =   855
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Image PHOTO 
      Height          =   2055
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2295
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
      Left            =   240
      TabIndex        =   12
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
      Left            =   240
      TabIndex        =   11
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
      Left            =   240
      TabIndex        =   10
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
      Left            =   240
      TabIndex        =   9
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
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   10
      X1              =   3000
      X2              =   1560
      Y1              =   -480
      Y2              =   -480
   End
End
Attribute VB_Name = "ACCOUNTUPDATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim ACC As Double


Private Sub CANCEL_Click()
NAME2.Text = ""
AGE.Text = ""
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
UPDATEDETAILS.Visible = False
UPDATE.Visible = True

ACCOUNTUPDATION.Hide
EMPLOYEEOPTION.Show

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

Private Sub UPDATE_Click()
NAME2.Enabled = True
AGE.Enabled = True
GENDER.Enabled = True
AADHAR.Enabled = True
PAN.Enabled = True
MOBILE.Enabled = True
UPDATE.Visible = False
UPDATEDETAILS.Visible = True
End Sub

Private Sub UPDATEDETAILS_Click()
On Error GoTo E1
ACC = ACCOUNT.Text
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ACC)
MSG = MsgBox("Are you sure you want to update this account", vbYesNoCancel + vbExclamation + vbDefaultButton3 + vbSystemModal, "C t O Bank: Account Updation")
If MSG = 6 Then
    RS.EDIT
    RS.Fields(2).Value = AGE.Text
    RS.Fields(3).Value = GENDER.Text
    RS.Fields(4).Value = AADHAR.Text
    RS.Fields(5).Value = MOBILE.Text
    RS.Fields(9).Value = PAN.Text
    RS.UPDATE
    MsgBox ("New account detail has been updated succcessfully.")
    NAME2.Text = ""
    AGE.Text = ""
    GENDER.Text = ""
    MOBILE.Text = ""
    PAN.Text = ""
    AADHAR.Text = ""
    BALANCE.Text = ""
    PHOTO.Picture = LoadPicture("")
    SIGNATURE.Picture = LoadPicture("")
    ACCOUNTUPDATION.Hide
    EMPLOYEEOPTION.Show
    
ElseIf MSG = 7 Then
    NAME2.Text = ""
    AGE.Text = ""
    GENDER.Text = ""
    MOBILE.Text = ""
    PAN.Text = ""
    AADHAR.Text = ""
    BALANCE.Text = ""
    PHOTO.Picture = LoadPicture("")
    SIGNATURE.Picture = LoadPicture("")
    ACCOUNTUPDATION.Hide
    EMPLOYEEOPTION.Show
ElseIf MSG = 2 Then
    ACCOUNTUPDATION.Show
End If
Exit Sub
E1:
    MsgBox "Blank fields are not allowed", , "C t O Bank"

End Sub


