VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdmin 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   15690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11880
   ScaleWidth      =   15690
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8040
      Picture         =   "frmAdmin.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   7455
      TabIndex        =   67
      Top             =   0
      Width           =   7455
      Begin VB.Image Image2 
         Height          =   330
         Left            =   7080
         Picture         =   "frmAdmin.frx":08C5
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "ppppppp"
      Height          =   495
      Left            =   0
      TabIndex        =   63
      Top             =   3000
      Width           =   15480
      Begin VB.Label lblRegisteredStudents 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reistered Students|"
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
         Left            =   10440
         MouseIcon       =   "frmAdmin.frx":0EDF
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblNoticeBoard 
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
         Left            =   7080
         MouseIcon       =   "frmAdmin.frx":11E9
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAboutUS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "About Us |"
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
         MouseIcon       =   "frmAdmin.frx":14F3
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblGenPin_SNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Pin Number|"
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
         Left            =   8160
         MouseIcon       =   "frmAdmin.frx":17FD
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblContactUs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Us |"
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
         Left            =   12480
         MouseIcon       =   "frmAdmin.frx":1B07
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Matriculation Number"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      TabIndex        =   60
      Top             =   3600
      Width           =   6735
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5160
         Picture         =   "frmAdmin.frx":1E11
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAdmin.frx":27FB
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMat 
         Appearance      =   0  'Flat
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
         Left            =   840
         TabIndex        =   61
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog dlgs 
      Left            =   600
      Top             =   10080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   3240
      TabIndex        =   55
      Top             =   10080
      Width           =   8175
      Begin VB.CommandButton cmdANew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   240
         Picture         =   "frmAdmin.frx":343D
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdASave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1080
         Picture         =   "frmAdmin.frx":3E27
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2760
         Picture         =   "frmAdmin.frx":4811
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdADelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1920
         Picture         =   "frmAdmin.frx":4D93
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAFirst 
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   3600
         Picture         =   "frmAdmin.frx":577D
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAPrevious 
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   4680
         Picture         =   "frmAdmin.frx":63C0
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdANext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   5640
         Picture         =   "frmAdmin.frx":7003
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdALast 
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   6720
         Picture         =   "frmAdmin.frx":7C46
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   1320
      TabIndex        =   48
      Top             =   6840
      Width           =   13815
      Begin VB.TextBox txtNsponsor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtPhoneSponsor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11760
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtAsponsor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtPOccupation 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11760
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   13
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtPname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   54
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Address of Sponsor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   53
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Sponsor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10560
         TabIndex        =   51
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Address of Parent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   50
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Parents"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACADEMIC PROFILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1320
      TabIndex        =   39
      Top             =   7920
      Width           =   13815
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   57
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdUploadPics 
         Caption         =   "&Upload"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   56
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtLsAttended 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         TabIndex        =   23
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtDoAddmission 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cboBaddmission 
         Height          =   315
         ItemData        =   "frmAdmin.frx":8889
         Left            =   2160
         List            =   "frmAdmin.frx":889C
         Sorted          =   -1  'True
         TabIndex        =   21
         Text            =   "Select..."
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cboFaculty 
         Height          =   315
         ItemData        =   "frmAdmin.frx":88BB
         Left            =   6120
         List            =   "frmAdmin.frx":88CB
         TabIndex        =   19
         Text            =   "Select..."
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         ItemData        =   "frmAdmin.frx":8909
         Left            =   6120
         List            =   "frmAdmin.frx":8A00
         Sorted          =   -1  'True
         TabIndex        =   20
         Text            =   "Select..."
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         ItemData        =   "frmAdmin.frx":912D
         Left            =   2160
         List            =   "frmAdmin.frx":9140
         TabIndex        =   24
         Text            =   "Select..."
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtMatNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cboSession 
         Height          =   315
         ItemData        =   "frmAdmin.frx":915D
         Left            =   2160
         List            =   "frmAdmin.frx":9173
         TabIndex        =   25
         Text            =   "Select..."
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label txtPic 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9960
         TabIndex        =   58
         Top             =   1920
         Width           =   45
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   9960
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Last School Attended"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   47
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Addmission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   46
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic For Addmission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Matriculation Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "School"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   42
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Session"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Attach Postport"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   59
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PERSONAL PROFILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1320
      TabIndex        =   26
      Top             =   4680
      Width           =   13815
      Begin VB.ComboBox cboMstatus 
         Height          =   315
         ItemData        =   "frmAdmin.frx":91B9
         Left            =   10320
         List            =   "frmAdmin.frx":91C3
         TabIndex        =   11
         Text            =   "Select..."
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboReligion 
         Height          =   315
         ItemData        =   "frmAdmin.frx":91D8
         Left            =   5520
         List            =   "frmAdmin.frx":91E8
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "Select..."
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtParmAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10320
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtPoB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtDoB 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtMname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtSname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtFname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "frmAdmin.frx":9218
         Left            =   1440
         List            =   "frmAdmin.frx":9222
         TabIndex        =   3
         Text            =   "Select..."
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboLga 
         Height          =   315
         ItemData        =   "frmAdmin.frx":9234
         Left            =   5520
         List            =   "frmAdmin.frx":9B38
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Select..."
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         ItemData        =   "frmAdmin.frx":B801
         Left            =   1680
         List            =   "frmAdmin.frx":B874
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "Select..."
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cboNation 
         Height          =   315
         ItemData        =   "frmAdmin.frx":B9A2
         Left            =   1440
         List            =   "frmAdmin.frx":B9AC
         TabIndex        =   9
         Text            =   "Select..."
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   38
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   37
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Parmanent Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   36
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Place Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   35
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Middlename"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "LGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "State of Origin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2655
      Left            =   0
      TabIndex        =   62
      Top             =   360
      Width           =   15480
      _cx             =   27305
      _cy             =   4683
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      Picture         =   "frmAdmin.frx":B9C2
      ScaleHeight     =   435
      ScaleWidth      =   10935
      TabIndex        =   68
      Top             =   0
      Width           =   10935
   End
   Begin VB.Image Image3 
      Height          =   11580
      Left            =   0
      Picture         =   "frmAdmin.frx":C287
      Stretch         =   -1  'True
      Top             =   840
      Width           =   15690
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPictureName
Public FN As String
Public FT As String


Private Sub cmdAClose_Click()
frmWaitTest.Show vbModal
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub

Private Sub cmdADelete_Click()
Dim MsgR As String
    MsgR = MsgBox("Are you sure you want to Delete this Record! ", vbYesNo, "Deletion Confirmation")

    If MsgR = vbYes Then
        
        With rs
            .Delete
            .Update
        End With
        MsgBox "Record Deleted!", vbInformation, "Academic System"
        cmdANew_Click
    End If

End Sub

Private Sub cmdAFirst_Click()
rs.MoveFirst
UpdateA
cmdBrowse.Enabled = True
End Sub

Private Sub cmdALast_Click()
rs.MoveLast
UpdateA
cmdBrowse.Enabled = True
End Sub

Private Sub cmdANew_Click()
    prcEnableTex
    txtSname = ""
    txtFname = ""
    txtMname = ""
    cboSex = ""
    txtDoB = ""
    cboLga = ""
    cboState = ""
    cboNation = ""
    txtMatNo = ""
    cboFaculty = ""
    cboDept = ""
    cboBaddmission = ""
    cboLevel = ""
    cboSession = ""
    txtNsponsor = ""
    txtLsAttended = ""
    txtPOccupation = ""
    txtPAddress = ""
    txtParmAddress = ""
    txtPname = ""
    cboMstatus = ""
    txtAsponsor = ""
    txtDoAddmission = ""
    cboReligion = ""
    txtPAddress = ""
    txtPoB = ""
    txtPhoneSponsor = ""
    Image1.Picture = LoadPicture("")
    txtPic = ""
    txtSname.SetFocus
    cmdASave.Enabled = True
    prcDisable
    cmdBrowse.Enabled = True
End Sub

Private Sub UpdateA()
    txtMatNo = rs.Fields(0)
    txtSname = rs.Fields(1)
    txtFname = rs.Fields(2)
    txtMname = rs.Fields(3)
    cboSex = rs.Fields(4)
    txtParmAddress = rs.Fields(5)
    txtPoB = rs.Fields(6)
    txtDoB = rs.Fields(7)
    cboNation = rs.Fields(8)
    cboState = rs.Fields(9)
    cboLga = rs.Fields(10)
    cboReligion = rs.Fields(11)
    cboMstatus = rs.Fields(12)
    txtPname = rs.Fields(13)
    txtPAddress = rs.Fields(14)
    txtPOccupation = rs.Fields(15)
    txtNsponsor = rs.Fields(16)
    txtAsponsor = rs.Fields(17)
    cboBaddmission = rs.Fields(18)
    txtDoAddmission = rs.Fields(19)
    txtLsAttended = rs.Fields(20)
    cboFaculty = rs.Fields(21)
    cboDept = rs.Fields(22)
    cboLevel = rs.Fields(23)
    cboSession = rs.Fields(24)
    txtPhoneSponsor = rs.Fields(25)
    strPictureName = rs.Fields(26) & ""
    Image1.Picture = LoadPicture(App.Path & "\images\" & strPictureName)
End Sub
Private Sub ClearA()
    txtSname = ""
    txtFname = ""
    txtMname = ""
    cboSex = ""
    txtDoB = ""
    cboLga = ""
    cboState = ""
    cboNation = ""
    txtMatNo = ""
    cboFaculty = ""
    cboDept = ""
    cboBaddmission = ""
    cboLevel = ""
    cboSession = ""
    txtNsponsor = ""
    txtLsAttended = ""
    txtPOccupation = ""
    txtPAddress = ""
    txtParmAddress = ""
    txtPname = ""
    cboMstatus = ""
    txtAsponsor = ""
    txtDoAddmission = ""
    cboReligion = ""
    txtPAddress = ""
    txtPoB = ""
    txtPhoneSponsor = ""
    txtSname.SetFocus
End Sub


Private Sub cmdANext_Click()
cmdBrowse.Enabled = True
If Not rs.EOF Then
    rs.MoveNext
    If Not rs.EOF Then
        UpdateA
    Else
        rs.MoveLast
        UpdateA
    End If
End If
End Sub

Private Sub cmdAPrevious_Click()
cmdBrowse.Enabled = True
If Not rs.BOF Then
rs.MovePrevious
    If Not rs.BOF Then
        UpdateA
    Else
        rs.MoveFirst
        UpdateA
    End If
End If
End Sub

Private Sub cmdASave_Click()
If txtSname = "" Or txtFname = "" Or cboSex = "" Or txtDoB = "" Or _
         cboLga.Text = "" Or cboState.Text = "" Or cboNation = "" Or txtMatNo = "" _
         Or cboFaculty = "" Or cboDept = "" Or _
         cboLevel = "" Or cboSession = "" Then
        
        MsgBox "Supply All necessary Information", vbInformation
        Exit Sub
Else
      
     With rs
        '.MoveFirst
        .Find "MatNo ='" & txtMatNo & "'"
        .Find "Level='" & cboLevel.Text & "'"
        If .EOF Then
            .AddNew
            .Fields(0) = txtMatNo
            .Fields(1) = txtSname
            .Fields(2) = txtFname
            .Fields(3) = txtMname
            .Fields(4) = cboSex
            .Fields(5) = txtParmAddress
            .Fields(6) = txtPoB
            .Fields(7) = txtDoB
            .Fields(8) = cboNation
            .Fields(9) = cboState
            .Fields(10) = cboLga
            .Fields(11) = cboReligion
            .Fields(12) = cboMstatus
            .Fields(13) = txtPname
            .Fields(14) = txtPAddress
            .Fields(15) = txtPOccupation
            .Fields(16) = txtNsponsor
            .Fields(17) = txtAsponsor
            .Fields(18) = cboBaddmission
            .Fields(19) = txtDoAddmission
            .Fields(20) = txtLsAttended
            .Fields(21) = cboFaculty
            .Fields(22) = cboDept
            .Fields(23) = cboLevel
            .Fields(24) = cboSession
            .Fields(25) = txtPhoneSponsor
            .Fields(26) = FT
            .Update
            UploadCardTransaction
            UploadUsedCards
            frmWait.lblSave.Visible = True
            frmWait.Show vbModal
            MsgBox "   New Record Saved    ", vbInformation
        Else
           MsgBox "Matric Number ! Or Level Already Exist!", vbInformation
           cmdASave.Enabled = True
       End If
       
    End With
    cmdANew.Enabled = True
    'cmdAUpdate.Enabled = True
    cmdASave.Enabled = False
    prcEnable
End If
End Sub

Private Sub UploadCardTransaction()
With rsCtrans
            .AddNew
            .Fields(0) = txtMatNo
            .Fields(1) = txtSname
            .Fields(2) = txtFname
            .Fields(3) = cboSex
            .Fields(4) = cboLevel
            .Fields(5) = cboDept
            .Fields(6) = Date
            .Fields(9) = 0
            .Fields(10) = FT
            .Update
End With
End Sub

Private Sub UploadUsedCards()
With rsUsedcards
            .AddNew
            .Fields(0) = txtMatNo
            .Fields(1) = txtSname
            .Fields(2) = txtFname
            .Fields(3) = cboSex
            .Fields(4) = cboLevel
            .Fields(5) = cboDept
            .Fields(6) = 0
            .Fields(7) = FT
            .Update
End With
End Sub

Public Sub UPDATEREC()

If txtSname = "" Or txtFname = "" Or cboSex = "" Or txtDoB = "" Or _
         cboLga.Text = "" Or cboState.Text = "" Or cboNation = "" Or txtMatNo = "" _
         Or cboFaculty = "" Or cboDept = "" Or _
         cboLevel = "" Or cboSession = "" Then
        
        MsgBox "Supply All necessary Information", vbInformation
        Exit Sub
Else
        
     With rs
        .MoveFirst
        .Find "MatNo ='" & txtMatNo & "'"
        If .EOF Then
            MsgBox "Matric Number does not Exist!", vbInformation
            Exit Sub
        Else
            .Fields(0) = txtMatNo
            .Fields(1) = txtSname
            .Fields(2) = txtFname
            .Fields(3) = txtMname
            .Fields(4) = cboSex
            .Fields(5) = txtParmAddress
            .Fields(6) = txtPoB
            .Fields(7) = txtDoB
            .Fields(8) = cboNation
            .Fields(9) = cboState
            .Fields(10) = cboLga
            .Fields(11) = cboReligion
            .Fields(12) = cboMstatus
            .Fields(13) = txtPname
            .Fields(14) = txtPAddress
            .Fields(15) = txtPOccupation
            .Fields(16) = txtNsponsor
            .Fields(17) = txtAsponsor
            .Fields(18) = cboBaddmission
            .Fields(19) = txtDoAddmission
            .Fields(20) = txtLsAttended
            .Fields(21) = cboFaculty
            .Fields(22) = cboDept
            .Fields(23) = cboLevel
            .Fields(24) = cboSession
            .Fields(25) = txtPhoneSponsor
            .Fields(26) = FT
            .Update
            frmWait.lblUpdate.Visible = True
            frmWait.Show vbModal
            MsgBox "   New Record Saved    ", vbInformation
       End If
    
    End With
    cmdANew.Enabled = True
    'cmdAUpdate.Enabled = True
    cmdASave.Enabled = False
    prcEnable
End If
End Sub

Public Sub prcEnable()
cmdANext.Enabled = True
cmdADelete.Enabled = True
cmdALast.Enabled = True
cmdAPrevious.Enabled = True
cmdAFirst.Enabled = True
End Sub

Public Sub prcDisable()
cmdANext.Enabled = False
cmdADelete.Enabled = False
cmdALast.Enabled = False
cmdAPrevious.Enabled = False
cmdAFirst.Enabled = False
cmdUploadPics.Enabled = False
End Sub

Public Sub prcEnableTex()
    txtSname.Enabled = True
    txtFname.Enabled = True
    txtMname.Enabled = True
    cboSex.Enabled = True
    txtDoB.Enabled = True
    cboLga.Enabled = True
    cboState.Enabled = True
    cboNation.Enabled = True
    txtMatNo.Enabled = True
    cboFaculty.Enabled = True
    cboDept.Enabled = True
    cboBaddmission.Enabled = True
    cboLevel.Enabled = True
    cboSession.Enabled = True
    txtNsponsor.Enabled = True
    txtLsAttended.Enabled = True
    txtPOccupation.Enabled = True
    txtPAddress.Enabled = True
    txtParmAddress.Enabled = True
    txtPname.Enabled = True
    cboMstatus.Enabled = True
    txtAsponsor.Enabled = True
    txtDoAddmission.Enabled = True
    cboReligion.Enabled = True
    txtPAddress.Enabled = True
    txtPoB.Enabled = True
    txtPhoneSponsor.Enabled = True
End Sub

Public Sub prcDisableText()
    txtSname.Enabled = False
    txtFname.Enabled = False
    txtMname.Enabled = False
    cboSex.Enabled = False
    txtDoB.Enabled = False
    cboLga.Enabled = False
    cboState.Enabled = False
    cboNation.Enabled = False
    txtMatNo.Enabled = False
    cboFaculty.Enabled = False
    cboDept.Enabled = False
    cboBaddmission.Enabled = False
    cboLevel.Enabled = False
    cboSession.Enabled = False
    txtNsponsor.Enabled = False
    txtLsAttended.Enabled = False
    txtPOccupation.Enabled = False
    txtPAddress.Enabled = False
    txtParmAddress.Enabled = False
    txtPname.Enabled = False
    cboMstatus.Enabled = False
    txtAsponsor.Enabled = False
    txtDoAddmission.Enabled = False
    cboReligion.Enabled = False
    txtPAddress.Enabled = False
    txtPoB.Enabled = False
    txtPhoneSponsor.Enabled = False
End Sub


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
        txtPic = FN
        cmdUploadPics.Enabled = True
        'cmdBrowse.Enabled = False
    End With
End Sub

Private Sub cmdSearch_Click()
   
With rs
        .MoveFirst
        .Find "MatNo ='" & txtMat & "'"
    If .EOF Then
        MsgBox "Matric Number does not Exist: Please contact Administrator", vbInformation, "Record Error"
        'fraMatNo.Visible = True
        txtMat = ""
        txtMat.SetFocus
        
    Else
        prcEnableTex
        frmWait.lblSearch.Visible = True
        frmWait.Show vbModal
        cmdUpdate.Enabled = True
        cmdUpdate.Default = True
        cmdASave.Enabled = False
        'cmdAUpdate.Enabled = False
        txtMatNo = rs.Fields(0)
        txtSname = rs.Fields(1)
        txtFname = rs.Fields(2)
        txtMname = rs.Fields(3)
        cboSex = rs.Fields(4)
        txtParmAddress = rs.Fields(5)
        txtPoB = rs.Fields(6)
        txtDoB = rs.Fields(7)
        cboNation = rs.Fields(8)
        cboState = rs.Fields(9)
        cboLga = rs.Fields(10)
        cboReligion = rs.Fields(11)
        cboMstatus = rs.Fields(12)
        txtPname = rs.Fields(13)
        txtPAddress = rs.Fields(14)
        txtPOccupation = rs.Fields(15)
        txtNsponsor = rs.Fields(16)
        txtAsponsor = rs.Fields(17)
        cboBaddmission = rs.Fields(18)
        txtDoAddmission = rs.Fields(19)
        txtLsAttended = rs.Fields(20)
        cboFaculty = rs.Fields(21)
        cboDept = rs.Fields(22)
        cboLevel = rs.Fields(23)
        cboSession = rs.Fields(24)
        txtPhoneSponsor = rs.Fields(25)
        strPictureName = .Fields(26) & " "
        Image1.Picture = LoadPicture(App.Path & "\images\" & strPictureName)
    End If

End With
End Sub


Private Sub cmdUpdate_Click()
cmdBrowse_Click
cmdUploadPics_Click
UPDATEREC
cmdASave.Enabled = True
'cmdAUpdate.Enabled = True
cmdANew.Enabled = True
cmdANew.Default = True
End Sub

Private Sub cmdUploadPics_Click()
        frmWaitUpdate.lblPics.Visible = True
        frmWaitUpdate.Show vbModal
Image1.Picture = LoadPicture(FN)
'cmdUploadPics.Enabled = False
End Sub


Private Sub Form_Load()
ShockwaveFlash1.Movie = App.Path & "\" & "Headerfees.swf"
   ShockwaveFlash1.Play
   ShockwaveFlash1.Loop = True
Call OpenDb
Call UsedCards
Call UnUsedCards
Call Ctrans
Call Admin
prcDisable
prcDisableText
cmdBrowse.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmWel.WindowState = 2
End Sub

Private Sub Image2_Click()
'frmWaitTest.Show vbModal
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub


Private Sub lblGenPin_SNum_Click()
Me.Enabled = False
frmWaitTest.Show vbModal
Load frmPinNum
frmPinNum.Show
End Sub

Private Sub lblHome_Click()
frmWaitTest.Show vbModal
Unload Me
Me.Hide
frmWel.Show
End Sub

Private Sub lblNoticeBoard_Click()
Unload Me
Me.Hide
Load frmWel
frmWel.Show
End Sub

Private Sub lblRegisteredStudents_Click()
frmWait.lblSearch.Visible = True
frmWait.Show vbModal
With deOnlineFees
    rptAdmin.Show vbModal
End With
End Sub
