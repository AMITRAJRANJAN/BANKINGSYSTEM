VERSION 5.00
Begin VB.Form EMPLOYEEOPTION 
   BackColor       =   &H00FF00FF&
   Caption         =   "Employee option's"
   ClientHeight    =   10935
   ClientLeft      =   -30
   ClientTop       =   315
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame BLOCKACCOUNTFRAME 
      BackColor       =   &H00404000&
      Caption         =   "Block account"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   4680
      TabIndex        =   52
      Top             =   7440
      Visible         =   0   'False
      Width           =   15495
      Begin VB.TextBox BACCOUNTSTATUS 
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
         Left            =   10800
         TabIndex        =   67
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CANCELBLOCK 
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
         Height          =   615
         Left            =   13800
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton ACCOUNTBLOCK 
         BackColor       =   &H0080C0FF&
         Caption         =   "Block account"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox BMOBILE 
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
         Left            =   3720
         TabIndex        =   58
         Top             =   2760
         Width           =   6855
      End
      Begin VB.TextBox BPAN 
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
         Left            =   3720
         TabIndex        =   57
         Top             =   2280
         Width           =   6855
      End
      Begin VB.TextBox BAADHAR 
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
         Left            =   3720
         TabIndex        =   56
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox BAGE 
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
         Left            =   3720
         TabIndex        =   55
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox BNAME 
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
         Left            =   3720
         TabIndex        =   54
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox BACCOUNT 
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
         Left            =   3720
         TabIndex        =   53
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label Label19 
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
         Left            =   120
         TabIndex        =   66
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   65
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label17 
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
         Left            =   120
         TabIndex        =   64
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label16 
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
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label15 
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
         Left            =   120
         TabIndex        =   62
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label14 
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
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame ACTIVATEACCOUNTFRAME 
      BackColor       =   &H00404000&
      Caption         =   "Activate account"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   4680
      TabIndex        =   37
      Top             =   7440
      Visible         =   0   'False
      Width           =   15495
      Begin VB.TextBox AACCOUNTSTATUS 
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
         Left            =   10800
         TabIndex        =   68
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox AACCOUNT 
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
         Left            =   3720
         TabIndex        =   45
         Top             =   360
         Width           =   6855
      End
      Begin VB.TextBox ANAME 
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
         Left            =   3720
         TabIndex        =   44
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox AAGE 
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
         Left            =   3720
         TabIndex        =   43
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox AAADHAR 
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
         Left            =   3720
         TabIndex        =   42
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox APAN 
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
         Left            =   3720
         TabIndex        =   41
         Top             =   2280
         Width           =   6855
      End
      Begin VB.TextBox AMOBILE 
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
         Left            =   3720
         TabIndex        =   40
         Top             =   2760
         Width           =   6855
      End
      Begin VB.CommandButton ACCOUNTACTIVATE 
         BackColor       =   &H0080C0FF&
         Caption         =   "Activate account"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CommandButton CANCELACTIVATE 
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
         Height          =   615
         Left            =   13800
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label13 
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
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label12 
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
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label11 
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
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   48
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Activate Account"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5760
      Width           =   3975
   End
   Begin VB.CommandButton BLOCKACCOUNT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Block Account"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      Picture         =   "Form1.frx":13ED
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Frame PASSWORDRESETFRAME 
      BackColor       =   &H00404000&
      Caption         =   "Reset password"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   4680
      TabIndex        =   20
      Top             =   7440
      Visible         =   0   'False
      Width           =   15495
      Begin VB.TextBox RACCOUNTSTATUS 
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
         Left            =   10800
         TabIndex        =   69
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CANCELRESET 
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
         Height          =   615
         Left            =   13800
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton RESETPASSWORD 
         BackColor       =   &H0080C0FF&
         Caption         =   "Reset Password"
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
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox RMOBILE 
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
         Left            =   3720
         TabIndex        =   26
         Top             =   2760
         Width           =   6855
      End
      Begin VB.TextBox RPAN 
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
         Left            =   3720
         TabIndex        =   25
         Top             =   2280
         Width           =   6855
      End
      Begin VB.TextBox RAADHAR 
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
         Left            =   3720
         TabIndex        =   24
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox RAGE 
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
         Left            =   3720
         TabIndex        =   23
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox RNAME 
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
         Left            =   3720
         TabIndex        =   22
         Top             =   840
         Width           =   6855
      End
      Begin VB.TextBox RACCOUNT 
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
         Left            =   3720
         TabIndex        =   21
         Top             =   360
         Width           =   6855
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
         Left            =   120
         TabIndex        =   32
         Top             =   2760
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
         Left            =   120
         TabIndex        =   31
         Top             =   1320
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
         Left            =   120
         TabIndex        =   30
         Top             =   2280
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
         Left            =   120
         TabIndex        =   29
         Top             =   1800
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
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   3255
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame SEARCHACCOUNTFRAME 
      BackColor       =   &H00404000&
      Caption         =   "Search your account to reset password"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   11640
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   8535
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
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
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   16
         ToolTipText     =   "Enter your 9 digit account number"
         Top             =   1200
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
         Left            =   120
         TabIndex        =   15
         Top             =   480
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
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   3855
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
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   13
         ToolTipText     =   "Enter your 5 digit Mobile number"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.CommandButton CANCELSEARCH 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
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
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   3855
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
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   3855
      End
   End
   Begin VB.CommandButton PASSWORDRESET 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset an account password"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Picture         =   "Form1.frx":26C5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
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
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10080
      Width           =   1695
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
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton LINKACREDITCARD 
      BackColor       =   &H0080C0FF&
      Caption         =   "Link a credit card to an account"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15120
      Picture         =   "Form1.frx":38AF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton LINKADEBITCARD 
      BackColor       =   &H0080C0FF&
      Caption         =   "Link a debit card to an account"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9960
      Picture         =   "Form1.frx":49B1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton UPDATEACCOUNT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update details of account"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      Picture         =   "Form1.frx":5AB3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton CREATEEMPLOYEE 
      BackColor       =   &H0080C0FF&
      Caption         =   "Create a new employee_id"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   15120
      Picture         =   "Form1.frx":6C9B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton CLOSEACCOUNT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close an account"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9960
      Picture         =   "Form1.frx":7CCD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton OPENACCOUNT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Open a new account"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      Picture         =   "Form1.frx":8C8A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4215
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
      Height          =   1575
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label1 
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
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   10575
      Left            =   240
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404000&
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   4680
      Top             =   240
      Width           =   15495
   End
End
Attribute VB_Name = "EMPLOYEEOPTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset

Dim ANUMBER As Double

Dim TODAY As Variant
Dim I As Integer

'Dim DB As Database
'Dim RS As Recordset

Dim A As Double
Dim M As Double

Dim MSG As VbMsgBoxResult

Dim STATUS As Integer





Private Sub ACCOUNTACTIVATE_Click()
ANUMBER = AACCOUNT.Text
MSG = MsgBox("Are you sure to activate this account ", vbYesNoCancel, "CtO Account activation confirmation")
If MSG = 6 Then
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
    RS.EDIT
    RS.Fields(16).Value = -1
    RS.UPDATE
    MsgBox "The account has been activated", , "C t O Bank"
    AACCOUNT.Text = ""
    ANAME.Text = ""
    AMOBILE.Text = ""
    APAN.Text = ""
    AAADHAR.Text = ""
    AAGE.Text = ""
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    ACTIVATEACCOUNTFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 7 Then
    MsgBox "The account is in block state", , "C t O Bank"
    AACCOUNT.Text = ""
    ANAME.Text = ""
    AMOBILE.Text = ""
    APAN.Text = ""
    AAADHAR.Text = ""
    AAGE.Text = ""
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    ACCOUNTACTIVATEFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 2 Then
    MsgBox "The account is in block state", , "C t O Bank"
    AACCOUNT.Text = ""
    ANAME.Text = ""
    AMOBILE.Text = ""
    APAN.Text = ""
    AAADHAR.Text = ""
    AAGE.Text = ""
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    ACCOUNTACTIVATEFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False
End If
End Sub

Private Sub ACCOUNTBLOCK_Click()
ANUMBER = BACCOUNT.Text
MSG = MsgBox("Are you sure to block this account ", vbYesNoCancel, "CtO Password reset confirmation")
If MSG = 6 Then
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
    RS.EDIT
    RS.Fields(16).Value = 0
    RS.UPDATE
    MsgBox "The account has been blocked", , "C t O Bank"
    BACCOUNT.Text = ""
    BNAME.Text = ""
    BMOBILE.Text = ""
    BPAN.Text = ""
    BAADHAR.Text = ""
    BAGE.Text = ""
    SEARCHVIAACCOUNT.Value = Fase
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    BLOCKACCOUNTFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 7 Then
    MsgBox "The account is in active state", , "C t O Bank"
    BACCOUNT.Text = ""
    BNAME.Text = ""
    BMOBILE.Text = ""
    BPAN.Text = ""
    BAADHAR.Text = ""
    BAGE.Text = ""
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    BLOCKACCOUNTFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 2 Then
    MsgBox "The account is in active state", , "C t O Bank"
    BACCOUNT.Text = ""
    BNAME2.Text = ""
    BMOBILE.Text = ""
    BPAN.Text = ""
    BAADHAR.Text = ""
    BAGE.Text = ""
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    BLOCKACCOUNTFRAME.Visible = False
    SEARCHACCOUNTFRAME.Visible = False
End If

End Sub

Private Sub BACK_Click()
EMPLOYEEOPTION.Hide
CTOHOME.Show
End Sub

Private Sub BLOCKACCOUNT_Click()
SEARCHACCOUNTFRAME.Visible = True
STATUS = 1
End Sub

Private Sub CANCELACTIVATE_Click()
AACCOUNT.Text = ""
ANAME.Text = ""
AMOBILE.Text = ""
APAN.Text = ""
AAADHAR.Text = ""
AAGE.Text = ""
ACTIVATEACCOUNTFRAME.Visible = False
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False
SEARCHACCOUNTFRAME.Visible = False
End Sub

Private Sub CANCELBLOCK_Click()
BACCOUNT.Text = ""
BNAME.Text = ""
BMOBILE.Text = ""
BPAN.Text = ""
BAADHAR.Text = ""
BAGE.Text = ""
BLOCKACCOUNTFRAME.Visible = False
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False
SEARCHACCOUNTFRAME.Visible = False

End Sub

Private Sub CANCELRESET_Click()
RACCOUNT.Text = ""
RNAME.Text = ""
RMOBILE.Text = ""
RPAN.Text = ""
RAADHAR.Text = ""
RAGE.Text = ""
PASSWORDRESETFRAME.Visible = False
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False
SEARCHACCOUNTFRAME.Visible = False

End Sub

Private Sub CLOSEACCOUNT_Click()
'EMPLOYEEOPTION.Hide
ACCOUNTFORCLOSURE.Show
End Sub

Private Sub Command1_Click()
SEARCHACCOUNTFRAME.Visible = True
STATUS = 2
End Sub

Private Sub CREATEEMPLOYEE_Click()
'EMPLOYEEOPTION.Hide
NEWEMPLOYEE.Show
End Sub

Private Sub END_Click()
EMPLOYEEOPTION.Hide
CTOHOME.Show

End Sub

Private Sub LINKACREDITCARD_Click()
'EMPLOYEEOPTION.Hide
LINKCREDITCARD.Show
End Sub

Private Sub LINKADEBITCARD_Click()
'EMPLOYEEOPTION.Hide
LINKDEBITCARD.Show

End Sub

Private Sub OPENACCOUNT_Click()
'EMPLOYEEOPTION.Hide
NEWACCOUNT.Show
End Sub

Private Sub PASSWORDRESET_Click()
SEARCHACCOUNTFRAME.Visible = True
STATUS = 3
End Sub

Private Sub RESETPASSWORD_Click()
ANUMBER = RACCOUNT.Text
MSG = MsgBox("Are you sure to reset this account password", vbYesNoCancel, "CtO Password reset confirmation")
If MSG = 6 Then
    Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & ANUMBER)
    RS.EDIT
    RS.Fields(7).Value = 0
    RS.UPDATE
    MsgBox "The account password is now eligible to reset", , "C t O Bank"
    RACCOUNT.Text = ""
    RNAME.Text = ""
    RMOBILE.Text = ""
    RPAN.Text = ""
    RAADHAR.Text = ""
    RAGE.Text = ""
    PASSWORDRESETFRAME.Visible = False
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 7 Then
    MsgBox "Login using your original password", , "C t O Bank"
    RACCOUNT.Text = ""
    RNAME.Text = ""
    RMOBILE.Text = ""
    RPAN.Text = ""
    RAADHAR.Text = ""
    RAGE.Text = ""
    PASSWORDRESETFRAME.Visible = False
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    SEARCHACCOUNTFRAME.Visible = False

ElseIf MSG = 2 Then
    MsgBox "Login using your original password", , "C t O Bank"
    RACCOUNT.Text = ""
    RNAME.Text = ""
    RMOBILE.Text = ""
    RPAN.Text = ""
    RAADHAR.Text = ""
    RAGE.Text = ""
    PASSWORDRESETFRAME.Visible = False
    SEARCHVIAACCOUNT.Value = False
    SEARCHVIAMOBILE.Value = False
    ACCOUNTNUMBERT.Text = ""
    MOBILENUMBERT.Text = ""
    ACCOUNTNUMBERL.Enabled = False
    ACCOUNTNUMBERT.Enabled = False
    MOBILENUMBERL.Enabled = False
    MOBILENUMBERT.Enabled = False
    SEARCHACCOUNTFRAME.Visible = False
End If

End Sub


Private Sub Timer1_Timer()
I = I + 1
TODAY = Now()
TDATE.Caption = Format(TODAY)

End Sub

Private Sub UPDATEACCOUNT_Click()
'EMPLOYEEOPTION.Hide
ACCOUNTFORUPDATION.Show
End Sub




Private Sub CANCELSEARCH_Click()
SEARCHVIAACCOUNT.Value = False
SEARCHVIAMOBILE.Value = False
ACCOUNTNUMBERT.Text = ""
MOBILENUMBERT.Text = ""
ACCOUNTNUMBERL.Enabled = False
ACCOUNTNUMBERT.Enabled = False
MOBILENUMBERL.Enabled = False
MOBILENUMBERT.Enabled = False

    
SEARCHACCOUNTFRAME.Visible = False
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


Private Sub SUBBMIT_Click()
On Error GoTo E1

If SEARCHVIAACCOUNT.Value = True Then
    A = ACCOUNTNUMBERT.Text
    If STATUS = 3 Then
        PASSWORDRESETFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
            If RS.Fields(16).Value = -1 Then
                RACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                RACCOUNTSTATUS.Text = "BLOCKED"
            End If
        RACCOUNT.Text = RS.Fields(0).Value
        RNAME.Text = RS.Fields(1).Value
        RMOBILE.Text = RS.Fields(5).Value
        RPAN.Text = RS.Fields(9).Value
        RAADHAR.Text = RS.Fields(4).Value
        RAGE.Text = RS.Fields(2).Value
    
    ElseIf STATUS = 2 Then
        ACTIVATEACCOUNTFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
            If RS.Fields(16).Value = -1 Then
                AACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                AACCOUNTSTATUS.Text = "BLOCKED"
            End If
        AACCOUNT.Text = RS.Fields(0).Value
        ANAME.Text = RS.Fields(1).Value
        AMOBILE.Text = RS.Fields(5).Value
        APAN.Text = RS.Fields(9).Value
        AAADHAR.Text = RS.Fields(4).Value
        AAGE.Text = RS.Fields(2).Value
        
    ElseIf STATUS = 1 Then
        BLOCKACCOUNTFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_NUMBER=" & A)
        If RS.Fields(16).Value = -1 Then
                BACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                BACCOUNTSTATUS.Text = "BLOCKED"
            End If
        BACCOUNT.Text = RS.Fields(0).Value
        BNAME.Text = RS.Fields(1).Value
        BMOBILE.Text = RS.Fields(5).Value
        BPAN.Text = RS.Fields(9).Value
        BAADHAR.Text = RS.Fields(4).Value
        BAGE.Text = RS.Fields(2).Value
        
    End If
    
ElseIf SEARCHVIAMOBILE.Value = True Then
    M = MOBILENUMBERT.Text
    If STATUS = 3 Then
        PASSWORDRESETFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_MOBILE=" & M)
            If RS.Fields(16).Value = -1 Then
                RACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                RACCOUNTSTATUS.Text = "BLOCKED"
            End If
        RACCOUNT.Text = RS.Fields(0).Value
        RNAME.Text = RS.Fields(1).Value
        RMOBILE.Text = RS.Fields(5).Value
        RPAN.Text = RS.Fields(9).Value
        RAADHAR.Text = RS.Fields(4).Value
        RAGE.Text = RS.Fields(2).Value
        
    ElseIf STATUS = 2 Then
        ACTIVATEACCOUNTFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_MOBILE=" & M)
            If RS.Fields(16).Value = -1 Then
                AACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                AACCOUNTSTATUS.Text = "BLOCKED"
            End If

        AACCOUNT.Text = RS.Fields(0).Value
        ANAME.Text = RS.Fields(1).Value
        AMOBILE.Text = RS.Fields(5).Value
        APAN.Text = RS.Fields(9).Value
        AAADHAR.Text = RS.Fields(4).Value
        AAGE.Text = RS.Fields(2).Value
        
    ElseIf STATUS = 1 Then
        BLOCKACCOUNTFRAME.Visible = True
        Set DB = OpenDatabase("C:\Users\mmani\Desktop\BCA PROJECT\ACCOUNT.mdb")
        Set RS = DB.OpenRecordset("SELECT * FROM ACCOUNTDATA WHERE ACCOUNT_MOBILE=" & M)
            If RS.Fields(16).Value = -1 Then
                BACCOUNTSTATUS.Text = "ACTIVE"
            ElseIf RS.Fields(16).Value = 0 Then
                BACCOUNTSTATUS.Text = "BLOCKED"
            End If
        BACCOUNT.Text = RS.Fields(0).Value
        BNAME.Text = RS.Fields(1).Value
        BMOBILE.Text = RS.Fields(5).Value
        BPAN.Text = RS.Fields(9).Value
        BAADHAR.Text = RS.Fields(4).Value
        BAGE.Text = RS.Fields(2).Value
        
    End If
    
    
End If

Exit Sub
E1:
   MsgBox "Account does not exist.", , "C t O Bank"
End Sub



