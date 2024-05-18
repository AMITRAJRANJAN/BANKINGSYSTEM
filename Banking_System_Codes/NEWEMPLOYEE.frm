VERSION 5.00
Begin VB.Form NEWEMPLOYEE 
   BackColor       =   &H00404000&
   Caption         =   "Create new employee_id"
   ClientHeight    =   6555
   ClientLeft      =   435
   ClientTop       =   615
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5760
      Width           =   1575
   End
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
      TabIndex        =   20
      Top             =   5760
      Width           =   1575
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
      Index           =   0
      Left            =   -3360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton HIDEN 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Password"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton SHOWN 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CREATE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Create"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton HIDEE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hide Password"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton SHOWW 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton VERIFY 
      BackColor       =   &H0080C0FF&
      Caption         =   "Verify"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox EMPLOYEEPASSWORD 
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
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "."
      TabIndex        =   8
      ToolTipText     =   "Enter your employee_password"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox EMPLOYEEID 
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
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Enter your employee_id"
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox NEWEMPLOYEEPASSWORD 
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
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "."
      TabIndex        =   3
      ToolTipText     =   "Create password for the new employee_id"
      Top             =   5040
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox NEWEMPLOYEENAME 
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
      Left            =   3960
      MaxLength       =   16
      TabIndex        =   2
      ToolTipText     =   "Enter the name of new employee"
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox NEWEMPLOYEEID 
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
      Left            =   3960
      MaxLength       =   16
      TabIndex        =   1
      ToolTipText     =   "Create employee_id of the new employee"
      Top             =   4440
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ComboBox NEWEMPLOYEEDESIGNATION 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      ItemData        =   "NEWEMPLOYEE.frx":0000
      Left            =   1680
      List            =   "NEWEMPLOYEE.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select Employee Designation"
      Top             =   3840
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Create new Employee_id"
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
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Verify your authorisaton"
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
      TabIndex        =   11
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "*   Employee password"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "*   Employee id"
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
      TabIndex        =   9
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "*   Employee password"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "*  New employee name"
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
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "*   Employee id"
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
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "NEWEMPLOYEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim E_ID As Double
Dim E_P As Double
Dim CE_P As Double
Dim NE_P As Double
Dim CNE_P As Double
Dim EDCODE As Double


Private Sub BACK_Click(Index As Integer)
NEWEMPLOYEE.Hide
EMPLOYEEOPTION.Show

End Sub

Private Sub CREATE_Click()
On Error GoTo E2
NE_P = NEWEMPLOYEEPASSWORD.Text
CNE_P = CONVERT_PASSWORD(NE_P)
If NEWEMPLOYEEDESIGNATION.ItemData(0) Then
    EDCODE = 4444
ElseIf NEWEMPLOYEEDESIGNATION.ItemData(1) Then
    EDCODE = 333
ElseIf NEWEMPLOYEEDESIGNATION.ItemData(2) Then
    EDCODE = 22
End If

Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\EMPLOYEE.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM EMPLOYEEDATA")
RS.AddNew
RS.Fields(0).Value = NEWEMPLOYEEID.Text
RS.Fields(1).Value = CNE_P
RS.Fields(2).Value = NEWEMPLOYEENAME.Text
RS.Fields(3).Value = NEWEMPLOYEEDESIGNATION.Text
RS.Fields(4).Value = EDCODE
RS.UPDATE
MsgBox "New employee has been added successfully.", , "C t O Bank"
NEWEMPLOYEEID.Text = ""
NEWEMPLOYEEPASSWORD.Text = ""
NEWEMPLOYEENAME.Text = ""
NEWEMPLOYEEDESIGNATION.Text = ""
HIDEN.Visible = False
SHOWN.Visible = False
NEWEMPLOYEE.Hide
EMPLOYEEOPTION.Show

Exit Sub
E2:
    MsgBox "Please enter the details to create new employee id", , "C t O Bank"
End Sub

Private Sub CANCEL_Click()
EMPLOYEEID.Text = ""
EMPLOYEEPASSWORD.Text = ""
NEWEMPLOYEEID.Text = ""
NEWEMPLOYEEPASSWORD.Text = ""
NEWEMPLOYEEPASSWORD.PasswordChar = "."
NEWEMPLOYEENAME.Text = ""
NEWEMPLOYEEDESIGNATION = ""
Label8.Visible = False
Label2.Visible = False
Label3.Visible = False
Label5.Visible = False
NEWEMPLOYEEDESIGNATION.Visible = False
NEWEMPLOYEENAME.Visible = False
NEWEMPLOYEEID.Visible = False
NEWEMPLOYEEPASSWORD.Visible = False
SHOWN.Visible = False
HIDEN.Visible = False
CREATE.Visible = False

NEWEMPLOYEE.Hide
EMPLOYEEOPTION.Show

End Sub

Private Sub HIDEE_Click()
EMPLOYEEPASSWORD.PasswordChar = "."
SHOWW.Visible = True
HIDEE.Visible = False
End Sub

Private Sub HIDEN_Click()
NEWEMPLOYEEPASSWORD.PasswordChar = "."
SHOWN.Visible = True
HIDEN.Visible = False
End Sub

Private Sub SHOWW_Click()
EMPLOYEEPASSWORD.PasswordChar = ""
SHOWW.Visible = False
HIDEE.Visible = True
End Sub

Private Sub SHOWN_Click()
NEWEMPLOYEEPASSWORD.PasswordChar = ""
SHOWN.Visible = False
HIDEN.Visible = True
End Sub

Private Sub VERIFY_Click()
On Error GoTo E1
E_ID = EMPLOYEEID.Text
E_P = EMPLOYEEPASSWORD.Text
CE_P = CONVERT_PASSWORD(E_P)
Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\EMPLOYEE.mdb")
Set RS = DB.OpenRecordset("SELECT * FROM EMPLOYEEDATA WHERE E_DESIGNATION_CODE=" & 55555)
If RS.Fields(0).Value = E_ID And RS.Fields(1).Value = CE_P Then
    MsgBox "Create new empoyee id", , "C t O Bank"
    Label8.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label5.Visible = True
    NEWEMPLOYEEDESIGNATION.Visible = True
    NEWEMPLOYEENAME.Visible = True
    NEWEMPLOYEEID.Visible = True
    NEWEMPLOYEEPASSWORD.Visible = True
    SHOWN.Visible = True
    CREATE.Visible = True
Else
    MsgBox "You are not authorised to create new employee id", , "C t O Bank"
End If
Exit Sub
E1:
    MsgBox "Please enter a valid employee id and password", , "C t O Bank"
End Sub

Private Function CONVERT_PASSWORD(X As Double)
Dim RET As Double
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PASSWORD = RET
End Function

