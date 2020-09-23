VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudPage 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   11400
   ClientLeft      =   270
   ClientTop       =   255
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmStudPage.frx":0000
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4800
      Picture         =   "frmStudPage.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   10575
      TabIndex        =   62
      Top             =   0
      Width           =   10575
      Begin VB.Image Image2 
         Height          =   330
         Left            =   10080
         Picture         =   "frmStudPage.frx":0BCF
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      Picture         =   "frmStudPage.frx":11E9
      ScaleHeight     =   435
      ScaleWidth      =   10815
      TabIndex        =   63
      Top             =   0
      Width           =   10815
   End
   Begin VB.Frame fraDeptRegistration 
      BackColor       =   &H0080C0FF&
      Height          =   8055
      Left            =   0
      TabIndex        =   31
      Top             =   2880
      Visible         =   0   'False
      Width           =   15255
      Begin VB.ComboBox cboCLevel 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":1AAE
         Left            =   9000
         List            =   "frmStudPage.frx":1ABE
         TabIndex        =   61
         Top             =   4080
         Width           =   3495
      End
      Begin MSComDlg.CommonDialog dlgs 
         Left            =   5160
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDoA 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   23
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox txtLastSchool 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   22
         Top             =   3120
         Width           =   4575
      End
      Begin VB.ComboBox cboQualify 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":1AD6
         Left            =   9000
         List            =   "frmStudPage.frx":1AE6
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtSponsorAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         MaxLength       =   120
         ScrollBars      =   1  'Horizontal
         TabIndex        =   20
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox txtSponsorName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         MaxLength       =   120
         ScrollBars      =   1  'Horizontal
         TabIndex        =   19
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtMOccupation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   18
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtMotherAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtMotherName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton cmdUploadPics 
         Caption         =   "&Upload"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   25
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   24
         Top             =   5640
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   9720
         TabIndex        =   58
         Top             =   7200
         Width           =   3855
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   64
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   27
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdSubmit 
            Caption         =   "&Submit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   26
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.TextBox txtOccupation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   7560
         Width           =   3855
      End
      Begin VB.TextBox txtFatherAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   7080
         Width           =   3855
      End
      Begin VB.TextBox txtFatherName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   6600
         Width           =   2895
      End
      Begin VB.ComboBox cboMStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":1AFF
         Left            =   1920
         List            =   "frmStudPage.frx":1B09
         TabIndex        =   12
         Top             =   6000
         Width           =   3015
      End
      Begin VB.ComboBox cboReligion 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":1B1E
         Left            =   1920
         List            =   "frmStudPage.frx":1B2B
         TabIndex        =   11
         Top             =   5520
         Width           =   3015
      End
      Begin VB.ComboBox cboLgA 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":1B4E
         Left            =   1920
         List            =   "frmStudPage.frx":1D11
         TabIndex        =   10
         Top             =   5040
         Width           =   3015
      End
      Begin VB.ComboBox cboState 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":2362
         Left            =   1920
         List            =   "frmStudPage.frx":23CF
         TabIndex        =   9
         Top             =   4560
         Width           =   3015
      End
      Begin VB.ComboBox cboNation 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":2506
         Left            =   1920
         List            =   "frmStudPage.frx":2510
         TabIndex        =   8
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox txtDoB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   3600
         Width           =   3495
      End
      Begin VB.TextBox txtPoB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtHomeAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   120
         ScrollBars      =   1  'Horizontal
         TabIndex        =   5
         Top             =   2640
         Width           =   3735
      End
      Begin VB.ComboBox cboSex 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmStudPage.frx":2526
         Left            =   1920
         List            =   "frmStudPage.frx":2530
         TabIndex        =   4
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtMname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtFname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtSname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtMatNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Level"
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
         Left            =   6480
         TabIndex        =   60
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   13680
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   13680
         X2              =   13680
         Y1              =   4680
         Y2              =   7920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   6240
         X2              =   13680
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label txtPics 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Height          =   195
         Left            =   7560
         TabIndex        =   59
         Top             =   4800
         Width           =   45
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   7560
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Passport"
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
         Left            =   8160
         TabIndex        =   57
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Admission"
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
         Left            =   6480
         TabIndex        =   56
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Last School Attended"
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
         Left            =   6480
         TabIndex        =   55
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Qalification"
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
         Left            =   6480
         TabIndex        =   54
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Address Of Sponsor"
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
         Left            =   6480
         TabIndex        =   53
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of Sponsor"
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
         Left            =   6480
         TabIndex        =   52
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation "
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
         Left            =   6480
         TabIndex        =   51
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Address Of Mother"
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
         Left            =   6480
         TabIndex        =   50
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of Mother"
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
         Left            =   6480
         TabIndex        =   49
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   120
         Y2              =   7920
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
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
         Left            =   120
         TabIndex        =   48
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Address Of Father"
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
         Left            =   120
         TabIndex        =   47
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of Father"
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
         Left            =   120
         TabIndex        =   46
         Top             =   6600
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status"
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
         Left            =   120
         TabIndex        =   45
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
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
         Left            =   120
         TabIndex        =   44
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "LGA"
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
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   4215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "State Of Origin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
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
         Left            =   120
         TabIndex        =   41
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
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
         Left            =   120
         TabIndex        =   40
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Place of Birth"
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
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "MiddleName"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "FirstName"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
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
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mat No:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   840
      Width           =   15255
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome! "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label lblHome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "| Home |"
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
         Left            =   6840
         MouseIcon       =   "frmStudPage.frx":2542
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblRegistration 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registration |"
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
         MouseIcon       =   "frmStudPage.frx":284C
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmStudPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim newApp As String
Dim Answer  As Integer
Public strPictureName
Public strCvName
Public FN As String
Public FT As String


Private Sub cmdBrowse_Click()
Dim retvalue As Long
Dim fileName As String
Dim dest As String
With dlgs
        .Filter = "(*.bmp;*.jpg;*.gif;*.dat;*.pcx)| *.bmp;*.jpg;*.gif;*.dat;*.pcx|(*.psd)|*.psd|(*.All files)|*.*"
        .ShowOpen
        If .fileName <> "" Then fileName = .fileName
        If .fileName = "" Then Exit Sub
        FN = .fileName
        FT = .FileTitle
        dest = (App.Path & "\images\")
        MsgBox FN, vbInformation, "Picture Loading..."
        txtPics = FN
        cmdUploadPics.Enabled = True
    End With
End Sub

Private Sub cmdSubmit_Click()
If txtMatNo <> "" And txtSname <> "" And txtFname <> "" And txtMname <> "" And cboSex <> "" And txtHomeAdd <> "" And _
    txtPoB <> "" And txtDoB <> "" And cboNation <> "" And cboState <> "" And cboLgA <> "" And cboReligion <> "" And _
    cboMStatus <> "" And txtFatherName <> "" And txtFatherAdd <> "" And txtOccupation <> "" And txtMotherName <> "" And _
    txtMotherAdd <> "" And txtMOccupation <> "" And txtSponsorName <> "" And txtSponsorAdd <> "" And cboQualify <> "" And _
    txtLastSchool <> "" And txtDoA <> "" And txtPics <> "" Then

Set rsR = New ADODB.Recordset
    With rsR
            .ActiveConnection = con
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
    End With
    
rsR.Open "Select *from studrec", con, adOpenDynamic, adLockOptimistic
    With rsR
            .MoveFirst
            .Find "MatNo='" & txtMatNo.Text & "'"
            .Find "Level='" & cboCLevel.Text & "'"
            
        If .EOF Then
            .AddNew
            .Fields(1) = txtMatNo
            .Fields(2) = txtSname
            .Fields(3) = txtFname
            .Fields(4) = txtMname
            .Fields(5) = cboSex
            .Fields(6) = txtHomeAdd
            .Fields(7) = txtPoB
            .Fields(8) = txtDoB
            .Fields(9) = cboNation
            .Fields(10) = cboState
            .Fields(11) = cboLgA
            .Fields(12) = cboReligion
            .Fields(13) = cboMStatus
            .Fields(14) = txtFatherName
            .Fields(15) = txtFatherAdd
            .Fields(16) = txtOccupation
            .Fields(17) = txtMotherName
            .Fields(18) = txtMotherAdd
            .Fields(19) = txtMOccupation
            .Fields(20) = txtSponsorName
            .Fields(21) = txtSponsorAdd
            .Fields(22) = cboQualify
            .Fields(23) = txtLastSchool
            .Fields(24) = txtDoA
            .Fields(25) = FT
            .Fields(26) = cboCLevel
            .Update
            frmWait.lblSave.Visible = True
            frmWait.Show vbModal
            MsgBox "Registration Successful", vbInformation, "Registration"
            fraDeptRegistration.Visible = False
    
        Else
                MsgBox "You have registered with the Department", vbInformation, "Registration"
                fraDeptRegistration.Visible = False
        End If
 .Close
End With
        
Else
        MsgBox "Please supply All Necessary Information", vbInformation, " Intranet"
        txtMatNo.SetFocus
        cmdSubmit.Enabled = False
End If
End Sub




Private Sub cmdUploadPics_Click()
   frmWaitUpdate.lblPics.Visible = True
        frmWaitUpdate.Show vbModal
Image1.Picture = LoadPicture(FN)
End Sub

Private Sub cmdView_Click()
Set rsA = New ADODB.Recordset
    With rsA
            .ActiveConnection = con
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
    End With
    
rsA.Open "Select *from Assignment", con, adOpenDynamic, adLockOptimistic

    With rsA
            .Find "Level='" & cboLevel.Text & "'"
            .Find "Coursecode='" & cboCCode.Text & "'"
        If .EOF Then
            MsgBox "No Assignment yet for this Level", vbInformation, "Assignment"
            fraLevelAssignments.Visible = False
        Else
            frmWaitTest.Show vbModal
            fraLevelAssignments.Visible = True
            txtViewAssignment.Text = .Fields(3)
        End If
    .Close
    End With
End Sub

Private Sub Form_Load()
Label3.Caption = "Welcome!  " & ActUser
fraRegistration.Visible = False
Call OpenDb
End Sub

Private Sub Image2_Click()
frmWaitTest.Show vbModal
Unload Me
Me.Hide
frmWel.Show
End Sub

Private Sub Image3_Click()

End Sub

Private Sub lblAboutUs_Click()
frmWaitTest.Show vbModal
frmContact_About.lblContactUs.Visible = True
frmContact_About.lblContactUs.Caption = "About Us"
frmContact_About.lblContact.Visible = False
frmContact_About.fraContact.Visible = False
frmContact_About.fraAbout.Visible = True
Unload Me
Me.Hide
frmContact_About.Show
End Sub

Private Sub lblAssignments_Click()
frmWaitTest.Show vbModal
fraDeptRegistration.Visible = False
fraLevelAssignments.Visible = True
End Sub

Private Sub lblContacUs_Click()
frmWaitTest.Show vbModal
frmContact_About.lblContactUs.Visible = True
frmContact_About.lblContact.Visible = True
frmContact_About.fraContact.Visible = True
frmContact_About.fraAbout.Visible = False
Unload Me
Me.Hide
frmContact_About.Show
End Sub

Private Sub lblCourseRegistration_Click()
frmWaitTest.Show vbModal
frmCourseReg.picSearch.Visible = True
frmCourseReg.Show
fraRegistration.Visible = False
End Sub

Private Sub lblDeptRegistration_Click()
frmWaitTest.Show vbModal
fraLevelAssignments.Visible = False
fraDeptRegistration.Visible = True
fraAssignment.Visible = False
fraRegistration.Visible = False
End Sub

Private Sub lblFAQs_Click()
frmWaitTest.Show vbModal
frmDisplay.fraFAQs.Visible = True
frmDisplay.fraDeptNews.Visible = False
frmDisplay.fraNoticeBoard.Visible = False
frmDisplay.fraRivpolySite.Visible = False
'frmDisplay.Caption = "Computer Science Department Intranet: Frequently Ask Questions"
frmDisplay.Show
End Sub

Private Sub lblHome_Click()
frmWaitTest.Show vbModal
Unload Me
Me.Hide
frmWel.Show
End Sub

Private Sub lblRegistration_Click()
fraRegistration.Visible = True
End Sub

