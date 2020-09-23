VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDisplay 
   BackColor       =   &H80000004&
   ClientHeight    =   11400
   ClientLeft      =   270
   ClientTop       =   255
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAcctBal 
      BackColor       =   &H80000009&
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   10680
      TabIndex        =   45
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label lblCardInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Details"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmDisplay.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCheckBal 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmDisplay.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraMatNo 
      BackColor       =   &H80000009&
      Caption         =   "Enter  Matric Number"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtMatNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraDisplay 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   6960
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Label lblSchool 
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   49
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   48
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MATRIC NO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SURNAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FIRSTNAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblMatNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblSname 
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   35
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblFname 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblSex 
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENT LEVEL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblCurrLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SCHOOL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblDept 
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label txtPic 
         AutoSize        =   -1  'True
         BackColor       =   &H000040C0&
         Height          =   195
         Left            =   5640
         TabIndex        =   28
         Top             =   2400
         Width           =   45
      End
   End
   Begin VB.Frame fraReceipt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Fees Receipt"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   5280
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdViewReceipt 
         Caption         =   "&View Receipt"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   26
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboReceipt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         ItemData        =   "frmDisplay.frx":0614
         Left            =   1680
         List            =   "frmDisplay.frx":0621
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdCloseReceipt 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame fraPayFees 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pay Fees"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   5280
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdclosePayfees 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox lblAcctBal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboFees 
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         ItemData        =   "frmDisplay.frx":0651
         Left            =   1680
         List            =   "frmDisplay.frx":065E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdMakePayment 
         Caption         =   "&Make Payment"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Fees To Pay"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   6600
   End
   Begin VB.Frame fraRecharge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scrach Card Serial and Pin numbers"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   4920
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtCardAmt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecharge 
         Caption         =   "&Recharge"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtPinNum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtSerialNum 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Number"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   -240
      Picture         =   "frmDisplay.frx":068E
      ScaleHeight     =   465
      ScaleWidth      =   9975
      TabIndex        =   6
      Top             =   0
      Width           =   9975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7920
      Picture         =   "frmDisplay.frx":0F53
      ScaleHeight     =   420
      ScaleWidth      =   7335
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.Image Image2 
         Height          =   330
         Left            =   6960
         Picture         =   "frmDisplay.frx":1818
         Top             =   0
         Width           =   330
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   14880
      Top             =   10200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraLinks 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   15255
      Begin VB.Label lblAcctInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| A/C Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   10680
         MouseIcon       =   "frmDisplay.frx":1E32
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| Receipts |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   12600
         MouseIcon       =   "frmDisplay.frx":213C
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblPayFees 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| Pay Fees |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   9480
         MouseIcon       =   "frmDisplay.frx":2446
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblRechargeAcct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| Recharge A/C |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   7800
         MouseIcon       =   "frmDisplay.frx":2750
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSignOut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SignOut |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   13800
         MouseIcon       =   "frmDisplay.frx":2A5A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblHome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| Home"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   6960
         MouseIcon       =   "frmDisplay.frx":2D64
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDisplay.frx":306E
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim newRec As String
Dim Lnum As String
Public strPictureName
Dim MatNum  As String
Public FN As String
Public FT As String
Private Sub cmdCancel_Click()
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub

Private Sub cmdClose_Click()
fraRecharge.Visible = False
End Sub

Private Sub cmdclosePayfees_Click()
fraPayFees.Visible = False
End Sub

Private Sub cmdCloseReceipt_Click()
fraReceipt.Visible = False
End Sub

Private Sub cmdDisplay_Click()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal
With rs
        .MoveFirst
        .Find "MatNo='" & txtMatNo & "'"
        If .EOF Then
            MsgBox "Matric Number is incorrect or does not Exist, Contact Registra", vbInformation
            txtMatNo = ""
            txtMatNo.SetFocus
        Else
            fraMatNo.Visible = False
            fraDisplay.Visible = True
            fraLinks.Visible = True
            Image3.Visible = True
            lblMatNo = .Fields(0)
            lblSname = .Fields(1)
            lblFname = .Fields(2)
            lblSex = .Fields(4)
            lblSchool = .Fields(21)
            lblDept = .Fields(22)
            lblCurrLevel = .Fields(23)
            strPictureName = .Fields(26) & ""
            Image1.Picture = LoadPicture(App.Path & "\images\" & strPictureName)
End If
    End With
End Sub

Private Sub cmdMakePayment_Click()
frmDisplay.Refresh
RepAmt = Val(lblAcctBal)
If cboFees.Text = "School Fee" And RepAmt >= 20000 And lblSchool = "Applied Sciences" Or lblSchool = "Engineering" Then
    'lblAcctBal = lblAcctBal - 20000
    Call ASchoolFees
ElseIf cboFees.Text = "School Fee" And RepAmt >= 18000 And lblSchool = "Management Sciences" Or lblSchool = "Arts" Then
    Call MSchoolFees
ElseIf cboFees.Text = "Acceptance Fee" And RepAmt >= 10000 Then
    'lblAcctBal = lblAcctBal - 10000
    Call AcceptanceFees
ElseIf cboFees.Text = "Stationary Fee" And RepAmt >= 1500 Then
    'RepAmt = Val(lblAcctBal) - 1500
    Call StationaryFees
Else
    MsgBox "Insufficient Amount in Account, check Account Balance or Select the appropriate fees", vbInformation, "Pay Fees"
    cboFees.SetFocus
    fraPayFees.Visible = False
End If
End Sub

Private Sub ASchoolFees()
prcNewRepNo
With rsReceipts
            .MoveFirst
            .Find "MatNo='" & lblMatNo & "'"
            .Find "ReceiptType='" & cboFees & "'"
            .Find "Level='" & lblCurrLevel & "'"
            
        If .EOF Then
            .AddNew
            .Fields(0) = lblMatNo
            .Fields(1) = lblSname
            .Fields(2) = lblFname
            .Fields(3) = lblCurrLevel
            .Fields(4) = lblDept
            .Fields(5) = NewRepNo
            .Fields(6) = 20000
            .Fields(7) = cboFees
            .Fields(8) = Date
            .Fields(9) = lblSchool
            .Update
            RepAmt = RepAmt - 20000
            UpdateUsedcards
            frmWait.lblPinNum.Visible = True
            frmWait.Show vbModal
            MsgBox "Fees Successfully Paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        Else
            MsgBox "Fees has already been paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        End If
    End With
End Sub

Private Sub MSchoolFees()
prcNewRepNo
With rsReceipts
            .MoveFirst
            .Find "MatNo='" & lblMatNo & "'"
            .Find "ReceiptType='" & cboFees & "'"
            .Find "Level='" & lblCurrLevel & "'"
            
        If .EOF Then
            .AddNew
            .Fields(0) = lblMatNo
            .Fields(1) = lblSname
            .Fields(2) = lblFname
            .Fields(3) = lblCurrLevel
            .Fields(4) = lblDept
            .Fields(5) = NewRepNo
            .Fields(6) = 18000
            .Fields(7) = cboFees
            .Fields(8) = Date
            .Fields(9) = lblSchool
            .Update
            RepAmt = RepAmt - 18000
            UpdateUsedcards
            frmWait.lblPinNum.Visible = True
            frmWait.Show vbModal
            MsgBox "Fees Successfully Paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        Else
            MsgBox "Fees has already been paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        End If
    End With
End Sub

Private Sub AcceptanceFees()
prcNewRepNo
With rsReceipts
            .MoveFirst
            .Find "MatNo='" & lblMatNo & "'"
            .Find "ReceiptType='" & cboFees & "'"
            .Find "Level='" & lblCurrLevel & "'"
            
        If .EOF Then
            .AddNew
            .Fields(0) = lblMatNo
            .Fields(1) = lblSname
            .Fields(2) = lblFname
            .Fields(3) = lblCurrLevel
            .Fields(4) = lblDept
            .Fields(5) = NewRepNo
            .Fields(6) = 10000
            .Fields(7) = cboFees
            .Fields(8) = Date
            .Fields(9) = lblSchool
            .Update
            RepAmt = RepAmt - 10000
            UpdateUsedcards
            frmWait.lblPinNum.Visible = True
            frmWait.Show vbModal
            MsgBox "Fees Successfully Paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        Else
            MsgBox "Fees has already been paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        End If
    End With
End Sub

Private Sub StationaryFees()
prcNewRepNo
With rsReceipts
            .MoveFirst
            .Find "MatNo='" & lblMatNo & "'"
            .Find "ReceiptType='" & cboFees & "'"
            .Find "Level='" & lblCurrLevel & "'"
            
        If .EOF Then
            .AddNew
            .Fields(0) = lblMatNo
            .Fields(1) = lblSname
            .Fields(2) = lblFname
            .Fields(3) = lblCurrLevel
            .Fields(4) = lblDept
            .Fields(5) = NewRepNo
            .Fields(6) = 1500
            .Fields(7) = cboFees
            .Fields(8) = Date
            .Fields(9) = lblSchool
            .Update
            RepAmt = RepAmt - 1500
            UpdateUsedcards
            frmWait.lblPinNum.Visible = True
            frmWait.Show vbModal
            MsgBox "Fees Successfully Paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        Else
            MsgBox "Fees has already been paid", vbInformation, "Pay Fees"
            fraPayFees.Visible = False
        End If
    End With
End Sub

Private Sub UpdateUsedcards()
With rsUsedcards
                .MoveFirst
                .Find "MatNo='" & lblMatNo & "'"
                
            If .EOF Then
                MsgBox "Please Matric Number is invalid, Contact registra", vbInformation, "Pay Fees"
                fraPayFees.Visible = False
            Else
                .Fields(6) = RepAmt
                .Update
            End If
End With
End Sub
Private Sub cmdRecharge_Click()
With rsUnUsedCards
                .MoveFirst
                .Find "SerialNumber ='" & txtSerialNum & "'"
                .Find "PinNumber ='" & txtPinNum & "'"
            
            If .EOF Then
                MsgBox "Invalid Pin or Serial Number entered", vbInformation, "Card Error"
                txtSerialNum.Text = ""
                txtSerialNum.SetFocus
               ' fraRecharge.Visible = False
            Else
                CardAmt = .Fields(3)
            End If
End With
'Call DownloadCardDetials
With rsCtrans
            .MoveFirst
            .Find "SerialNum ='" & txtSerialNum & "'"
            .Find "PinNum='" & txtPinNum & "'"
            
         If .EOF Then
        .AddNew
        .Fields(0) = lblMatNo
        .Fields(1) = lblSname
        .Fields(2) = lblFname
        .Fields(3) = lblSex
        .Fields(4) = lblCurrLevel
        .Fields(5) = lblDept
        .Fields(6) = Date
        .Fields(7) = txtSerialNum
        .Fields(8) = txtPinNum
        .Fields(9) = CardAmt
        .Fields(10) = strPictureName
        .Update
        Call UploadCard
    Else
        MsgBox "This Card Has been used OR Serial or Pin Number does not exist", vbInformation, "Recharge Acct"
        txtSerialNum = ""
        txtPinNum = ""
        txtSerialNum.SetFocus
    End If
    fraRecharge.Visible = False
End With
End Sub

Private Sub DownloadCardDetials()
With rsUnUsedCards
                .MoveFirst
                .Find "SerialNumber ='" & txtSerialNum & "'"
                .Find "PinNumber ='" & txtPinNum & "'"
            
            If .EOF Then
                MsgBox "Invalid Pin or Serial Number entered", vbInformation, "Card Error"
                txtSerialNum.Text = ""
                txtSerialNum.SetFocus
                fraRecharge.Visible = False
            Else
                CardAmt = .Fields(3)
            End If
End With
End Sub

Private Sub UploadCard()

With rsUsedcards
                .MoveFirst
                .Find "MatNo ='" & lblMatNo & "'"
        
            If .EOF Then
                 MsgBox "Invalid Matric Number Diplayed...", vbInformation
                txtSerialNum.SetFocus
                fraRecharge.Visible = False
            Else
                .Fields(6) = CardAmt + Val(txtCardAmt)
                .Update
                frmWait.lblPinNum.Visible = True
                frmWait.Show vbModal
                MsgBox "Your Account has been Credited with: N" & CardAmt & "  ", vbInformation, "Acct Status"
           End If
End With
End Sub


Private Sub cmdViewReceipt_Click()
If cboReceipt = "School Fee" And lblSchool = "Applied Sciences" Or lblSchool = "Engineering" Then
    Call ReportSReceipt
ElseIf cboReceipt = "Acceptance Fee" Or cboReceipt = "Stationary Fee" Then
    Call ReportOReceipts
ElseIf cboReceipt = "School Fee" And lblSchool = "Management Sciences" Or lblSchool = "Arts" Then
    Call ReportMReceipt
        Else
            MsgBox "Please Select Receipt Type", vbExclamation, "Query"
            cboReceipt.SetFocus
End If
End Sub


Private Sub ReportSReceipt()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal
            rsReceipts.MoveFirst
            rsReceipts.Find "MatNo='" & lblMatNo & "'"
          
        If rsReceipts.EOF Then
            MsgBox "Record Not Found", vbInformation, "Receipt Error"
    Else
        With deOnlineFees
             .conFees.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\OnlineFees.mdb;"
            .rsReceipts.Open "Select * From Receipts Where Matno='" & lblMatNo.Caption & "' and ReceiptType ='" & cboReceipt.Text & "'", deOnlineFees.conFees, adOpenDynamic, adLockOptimistic
            rptReceipts.Show vbModal
           .conFees.Close
    
End With
End If
End Sub

Private Sub ReportMReceipt()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal
            rsReceipts.MoveFirst
            rsReceipts.Find "MatNo='" & lblMatNo & "'"
          
        If rsReceipts.EOF Then
            MsgBox "Record Not Found", vbInformation, "Receipt Error"
    Else
        With deOnlineFees
             .conFees.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\OnlineFees.mdb;"
            .rsReceipts.Open "Select * From Receipts Where Matno='" & lblMatNo.Caption & "' and ReceiptType ='" & cboReceipt.Text & "'", deOnlineFees.conFees, adOpenDynamic, adLockOptimistic
            rptMReceipts.Show vbModal
           .conFees.Close
    
End With
End If
End Sub


Private Sub ReportOReceipts()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal

            rsReceipts.MoveFirst
            rsReceipts.Find "MatNo='" & lblMatNo & "'"
          
        If rsReceipts.EOF Then
            MsgBox "Record Not Found", vbInformation, "Receipt Error"
     Else
        With deOnlineFees
             .conFees.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\OnlineFees.mdb;"
            .rsReceipts.Open "Select * From Receipts Where Matno='" & lblMatNo.Caption & "' and ReceiptType ='" & cboReceipt.Text & "'", deOnlineFees.conFees, adOpenDynamic, adLockOptimistic
            rptOtherReceipts.Show vbModal
           .conFees.Close
        End With
End If

End Sub
Private Sub Form_Load()
Call OpenDb
Call UsedCards
Call UnUsedCards
Call Ctrans
Call Admin
Call Receipts
Timer1.Enabled = False
End Sub


Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Image2_Click()
'frmWaitTest.Show vbModal
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub


Private Sub lblPrint_Click()
On Error GoTo check_pr
   
    CD1.ShowPrinter
    
check_pr:
    If Err.Number = 32755 Then
    
    Else
        MsgBox "Error Occured: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Sub Label10_Click()
fraAcctBal.Visible = False
fraReceipt.Visible = True
cmdViewReceipt.Default = True
cboReceipt.SetFocus
End Sub

Private Sub lblAcctInfo_Click()
fraAcctBal.Visible = True
fraRecharge.Visible = False
fraPayFees.Visible = False
fraReceipt.Visible = False
End Sub

Private Sub lblCardInfo_Click()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal
With deOnlineFees
    .conFees.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\OnlineFees.mdb;"
    .rsCardInfo.Open "Select * From CardTransactions where [MatNo]='" & frmDisplay.lblMatNo & "'", deOnlineFees.conFees, adOpenDynamic, adLockOptimistic
    rptCardInfo.Show vbModal
    .conFees.Close
End With
fraAcctBal.Visible = False
End Sub

Private Sub lblCheckBal_Click()
frmWait.lblAnalysis.Visible = True
frmWait.Show vbModal
With rsUsedcards
                .MoveFirst
                .Find "MatNo ='" & lblMatNo & "'"
            If .EOF Then
                MsgBox "Matric Number Does not Exist! please Contact Registra", vbInformation, "Matric Number Error"
                fraPayFees.Visible = False
                'Unload Me
                'Me.Hide
            Else
                MsgBox "Your Student Account Balance is: N" & .Fields(6) & " Only", vbInformation, "A/C Details"
                fraAcctBal.Visible = False
            End If
    End With

End Sub

Private Sub lblHome_Click()
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub

Private Sub lblPayFees_Click()
fraAcctBal.Visible = False
With rsUsedcards
                .MoveFirst
                .Find "MatNo ='" & lblMatNo & "'"
            If .EOF Then
                MsgBox "Please Recharge your Student Account", vbInformation, "Acct Bal"
                fraPayFees.Visible = False
                'Unload Me
                'Me.Hide
                'Load frmWel
                'frmWel.Show
            Else
                fraPayFees.Visible = True
                cmdMakePayment.Default = True
                lblAcctBal = .Fields(6)
            End If
    End With
End Sub

Private Sub lblRechargeAcct_Click()
fraAcctBal.Visible = False
'Timer1.Enabled = True
With rsUsedcards
                .MoveFirst
                .Find "MatNo='" & lblMatNo & "'"
            If .EOF Then
                MsgBox "Matric Number Does not Exist! please Contact Registra", vbInformation, "Matric Number Error"
                fraRecharge.Visible = False
                'Unload Me
                'Me.Hide
            Else
                txtCardAmt.Text = .Fields(6)
                fraRecharge.Visible = True
                cmdRecharge.Default = True
                txtSerialNum = ""
                txtPinNum = ""
                txtSerialNum.SetFocus
            End If
    End With
End Sub

Private Sub lblSignOut_Click()
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub

Public Sub prcNewRepNo()
Dim C As Integer
Dim RepNo As String

NewRepNo = ""
C = 0
'rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
Do While Not rsReceipts.EOF
C = C + 1
rsReceipts.MoveNext
Loop
'rs.Close

If C = 0 Then
NewRepNo = "0000001"
Exit Sub
End If
'rs.Open "Select * from empmaster", con, adOpenDynamic, adLockOptimistic
rsReceipts.MoveLast
RepNo = rsReceipts("Receiptno")
RepNo = Mid(RepNo, 1, 7)
NewRepNo = RepNo + 1
If NewRepNo >= 1 And NewRepNo <= 9 Then
NewRepNo = "000000" + NewRepNo
ElseIf NewRepNo >= 10 And NewRepNo <= 99 Then
NewRepNo = "00000" + NewRepNo
ElseIf NewRepNo >= 100 And NewRepNo <= 999 Then
NewRepNo = "0000" + NewRepNo
ElseIf NewRepNo >= 1000 And NewRepNo <= 9999 Then
NewRepNo = "000" + NewRepNo
ElseIf NewRepNo >= 10000 And NewRepNo <= 99999 Then
NewRepNo = "00" + NewRepNo
ElseIf NewRepNo >= 100000 And NewRepNo <= 999999 Then
NewRepNo = "0" + NewRepNo
End If
'rs.Close
End Sub

