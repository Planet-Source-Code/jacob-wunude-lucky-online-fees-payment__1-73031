VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H00000080&
   ClientHeight    =   600
   ClientLeft      =   5055
   ClientTop       =   5430
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   4905
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   360
      Top             =   840
   End
   Begin VB.Label lblAnalysis 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Analysing DataBase, Please wait.."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblUpdate 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Updating Record, Please wait..."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Searching For Record, Please Wait..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creating User Account, Please Wait..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblPin 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Generating Pin, Please Waiting..."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblPinNum 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Updating Account, Please wait..."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblSave 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saving Record, Please Wait..."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
