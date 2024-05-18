VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PHOTOSIGNATURE 
   BackColor       =   &H00404000&
   Caption         =   "Upload photo and signature"
   ClientHeight    =   5160
   ClientLeft      =   1680
   ClientTop       =   2460
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton UPLOADPHOTO 
      BackColor       =   &H0080C0FF&
      Caption         =   "Browse Photograph"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton UPLOADSIGNATURE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Browse Signature"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton SUBBMIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Subbmit"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton RESET 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5760
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Upload Photo and Signature"
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
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Image PHOTO 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   240
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image SIGNATURE 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "PHOTOSIGNATURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHOTOADDRESS As String
Dim SIGNATUREADDRESS As String


Private Sub END_Click()
End
End Sub

Private Sub RESET_Click()
PHOTO.Picture = LoadPicture("")
SIGNATURE.Picture = LoadPicture("")
If PHOTO.DataChanged = True And SIGNATURE.DataChanged = True Then
    SUBBMIT.Visible = False
    RESET.Visible = False
End If

End Sub

Private Sub SUBBMIT_Click()
On Error GoTo E1
PHOTOADDRESS = CommonDialog1.FileName
SIGNATUREADDRESS = CommonDialog2.FileName
NEWACC.NEWPHOTO = PHOTOADDRESS
NEWACC.NEWSIGNATURE = SIGNATUREADDRESS
MsgBox "Your data has been subbmitted succesfully", , "C t O Bank"

PHOTO.Picture = LoadPicture("")
SIGNATURE.Picture = LoadPicture("")
    SUBBMIT.Visible = False
    RESET.Visible = False
PHOTOSIGNATURE.Hide
NEWACCOUNTREVIEW.Show
Exit Sub
E1:
    MsgBox "Please upload a photograph and a signature", , "C t O Bank"

End Sub

Private Sub UPLOADPHOTO_Click()
CommonDialog1.ShowOpen
CommonDialog1.DialogTitle = "Browse your photograph"
PHOTO.Picture = LoadPicture(CommonDialog1.FileName)
If PHOTO.DataChanged = True And SIGNATURE.DataChanged = True Then
    SUBBMIT.Visible = True
    RESET.Visible = True
End If

End Sub

Private Sub UPLOADSIGNATURE_Click()
CommonDialog2.ShowOpen
CommonDialog2.DialogTitle = "Browse your signature"
SIGNATURE.Picture = LoadPicture(CommonDialog2.FileName)
If PHOTO.DataChanged = True And SIGNATURE.DataChanged = True Then
    SUBBMIT.Visible = True
    RESET.Visible = True
End If

End Sub
