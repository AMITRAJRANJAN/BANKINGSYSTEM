VERSION 5.00
Begin VB.Form EMPLOYEELOGIN 
   BackColor       =   &H00404000&
   Caption         =   "Employee Login"
   ClientHeight    =   2850
   ClientLeft      =   3405
   ClientTop       =   2610
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
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
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox EMPLOYEEPASSWORD 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "."
      TabIndex        =   2
      ToolTipText     =   "Enter your employee_password"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox EMPLOYEEID 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Enter your 9 digit employee_id"
      Top             =   240
      Width           =   3405
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "* Employee password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "* Employee id"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "EMPLOYEELOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim EID As Double
Dim OP As Double
Dim CP As Double
Dim EDCODE As Long



Private Sub BACK_Click()
EMPLOYEEPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
EMPLOYEELOGIN.Hide

End Sub

Private Sub END_Click()
EMPLOYEEID.Text = ""
EMPLOYEEPASSWORD.Text = ""
EMPLOYEEPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
EMPLOYEELOGIN.Hide

End Sub

Private Sub HIDEE_Click()
EMPLOYEEPASSWORD.PasswordChar = "."
HIDEE.Visible = False
SHOWW.Visible = True
End Sub


Private Sub SHOWW_Click()
EMPLOYEEPASSWORD.PasswordChar = ""
SHOWW.Visible = False
HIDEE.Visible = True
End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1
    EID = EMPLOYEEID.Text
    OP = EMPLOYEEPASSWORD.Text
    CP = CONVERT_PASSWORD(OP)
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\EMPLOYEE.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM EMPLOYEEDATA WHERE E_ID=" & EID)

    If RS.Fields(0).Value = EID And RS.Fields(1).Value = CP Then
    EMPLOYEEID.Text = ""
    EMPLOYEEPASSWORD.Text = ""
    CTOHOME.Hide
    EMPLOYEELOGIN.Hide
    EMPLOYEEOPTION.Show
    Else
    MsgBox "Invalid detail", , "C t O Bank"
    EMPLOYEEID.Text = ""
    EMPLOYEEPASSWORD.Text = ""

End If
Exit Sub
E1:
    MsgBox "Please enter a valid id and password."

End Sub
Private Function CONVERT_PASSWORD(X As Double)
Dim RET As Double
RET = X + 23 - 5900 - 2 + 7195 - 29 + 3 - 5 - 2091 + 4307
CONVERT_PASSWORD = RET
End Function

