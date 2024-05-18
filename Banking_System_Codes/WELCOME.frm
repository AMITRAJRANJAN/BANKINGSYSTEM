VERSION 5.00
Begin VB.Form WELCOME 
   BackColor       =   &H00800080&
   Caption         =   "Welcome to C t O Bank"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   14100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton GO 
      BackColor       =   &H00FF80FF&
      Caption         =   "GO   --->"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label WELCOMEE 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "WELCOME TO C T O BANK"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   30
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Shape CD9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4440
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4200
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3960
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3720
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4440
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4200
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3960
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3720
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CD1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape C 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape CU2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR14 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9960
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR13 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9720
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR12 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9480
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9240
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR15 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3960
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL12 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3720
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL13 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL14 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   3240
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4200
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OO 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9960
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9720
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9480
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9240
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   10200
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9960
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9720
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9480
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OR1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9240
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape O 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape OL1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   7560
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   7320
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   7080
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6840
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   9000
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   8760
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   8520
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   8280
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   8040
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TR5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   7800
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   5160
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4920
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4680
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   4440
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape T 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6360
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6120
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   5880
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   5640
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TL5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape TD2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   255
      Left            =   6600
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "WELCOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer


Private Sub GO_Click()
WELCOME.Hide
CTOHOME.Show

End Sub

Private Sub Timer1_Timer()
I = I + 1

If I = 1 Then
    C.Visible = True
    CU1.Visible = True
    CD1.Visible = True
    
    TD1.Visible = True
    
    O.Visible = True
    OL1.Visible = True
    OR1.Visible = True
    
    End If
    
If I = 2 Then
    CU2.Visible = True
    CD2.Visible = True
    
    TD2.Visible = True
    
    OL2.Visible = True
    OL3.Visible = True
    OR2.Visible = True
    OR3.Visible = True
    End If
    
If I = 3 Then
    CU3.Visible = True
    CD3.Visible = True
    
    TD3.Visible = True
    
    OL4.Visible = True
    OR4.Visible = True
    
    End If
    
If I = 4 Then
    CU4.Visible = True
    CD4.Visible = True
    
    TD4.Visible = True
    
    OL5.Visible = True
    OL6.Visible = True
    OR5.Visible = True
    OR6.Visible = True

End If

If I = 5 Then
    CU5.Visible = True
    CD5.Visible = True
    
    TD5.Visible = True
    
    OL7.Visible = True
    OR7.Visible = True

    End If
    
If I = 6 Then
    CU6.Visible = True
    CD6.Visible = True
    
    TD6.Visible = True
    
    OL8.Visible = True
    OL9.Visible = True
    OR8.Visible = True
    OR9.Visible = True

        
End If
    
If I = 7 Then
    CU7.Visible = True
    CD7.Visible = True
    
    TD7.Visible = True
    
    OL10.Visible = True
    OL11.Visible = True
    OR10.Visible = True
    OR11.Visible = True

    End If
    
If I = 8 Then
    CU8.Visible = True
    CD8.Visible = True
    CD9.Visible = True
    
    TD8.Visible = True
    
    OO.Visible = True

End If

If I = 9 Then
    TL14.Visible = True
    TL13.Visible = True
    TL12.Visible = True
    TL11.Visible = True
    TL10.Visible = True
    TR15.Visible = True
    TR14.Visible = True
    TR13.Visible = True
    TR12.Visible = True
    TR11.Visible = True
    TR10.Visible = True
End If

If I = 10 Then
    TL9.Visible = True
    TL8.Visible = True
    TL7.Visible = True
    TL6.Visible = True
    TL5.Visible = True
    TR9.Visible = True
    TR8.Visible = True
    TR7.Visible = True
    TR6.Visible = True
    TR5.Visible = True
End If

If I = 11 Then
    TL4.Visible = True
    TL3.Visible = True
    TL2.Visible = True
    TL1.Visible = True
    T.Visible = True
    TR4.Visible = True
    TR3.Visible = True
    TR2.Visible = True
    TR1.Visible = True

End If
If I = 12 Then
    WELCOMEE.Visible = True
    GO.Visible = True
    End If
End Sub
