VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form frmWel 
   Caption         =   "Welcome to Department of Computer Science Rivers State Polytechnic Bori... Fees Payment Software"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   15240
   Icon            =   "frmWel.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLoginID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   3720
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtLoginPword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9720
      PasswordChar    =   "+"
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15255
      _cx             =   26908
      _cy             =   4683
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      ToolTipText     =   "Log-In to Registration  Courses"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "cyberjakes@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000D&
      Caption         =   "News Flash>>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      ToolTipText     =   "Log-In to Registration  Courses"
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      ToolTipText     =   "Log-In to Registration  Courses"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ">> Not Registered? Click Here to "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label lblSignUp 
      BackStyle       =   0  'Transparent
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7920
      MouseIcon       =   "frmWel.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   7920
      X2              =   8640
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   $"frmWel.frx":074C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   -8640
      TabIndex        =   7
      Top             =   2640
      Width           =   31335
   End
   Begin VB.Image Image1 
      Height          =   11655
      Left            =   -840
      Picture         =   "frmWel.frx":08B5
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   16095
   End
End
Attribute VB_Name = "frmWel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    If (txtLoginID.Text = "staff" And txtLoginPword.Text = "compscience") Or _
        (txtLoginID.Text = "STAFF" And txtLoginPword.Text = "COMPSCIENCE") Then
        frmWaitTest.Show vbModal
        Me.Hide
        frmAdmin.Show
        txtLoginPword = ""
        txtLoginID = ""
        
    ElseIf (txtLoginID.Text = "hodcompscience" And txtLoginPword.Text = "administrator") Or _
           (txtLoginID.Text = "HODCOMPSCIENCE" And txtLoginPword.Text = "ADMINISTRATOR") Then
        frmWaitTest.Show vbModal
        Me.Hide
        frmAdmin.Show
        txtLoginPword = ""
        txtLoginID = ""
  Else
    
    With rs
        .MoveFirst
        .Find "UserId='" & txtLoginID & "'"
        .Find "Password ='" & txtLoginPword & "'"
        If .EOF Then
            MsgBox "Username, Password is incorrect or does not Exist", vbInformation
            txtMatNo = ""
            txtLoginPword = ""
            txtLoginID.SetFocus
        Else
            frmWaitTest.Show vbModal
            Me.Hide
            frmDisplay.fraMatNo.Visible = True
            frmDisplay.cmdDisplay.Default = True
            frmDisplay.Show
            txtLoginPword = ""
            txtLoginID = ""
     
End If
    End With
    End If
End Sub


Private Sub cmdSearch_Click()

With rsR
        .MoveFirst
        .Find "UserId ='" & txtMatNo.Text & "'"
        .Find "Level='" & cboLevel.Text & "'"
        
    If .EOF Then
        MsgBox "This Student is not a Registered Student of the Department", vbInformation, "Intranet"
        fraSearch.Visible = False
Else
        
        frmWaitTest.Show vbModal
        fraSearch.Visible = False
        fraSearchResult.Visible = True
        lblSn = .Fields(0)
        lblMatNo = .Fields(1)
        lblSname = .Fields(2)
        lblFname = .Fields(3)
        lblSex = .Fields(5)
        lblPermAdd = .Fields(6)
        lblPoB = .Fields(7)
        lblDoB = .Fields(8)
        lblState = .Fields(10)
        lblLga = .Fields(11)
        lblQual = .Fields(22)
        lblDoA = .Fields(24)
        lblLastSch = .Fields(23)
        strPictureName = .Fields(25) & ""
        Image1.Picture = LoadPicture(App.Path & "\images\" & strPictureName)
 
 End If
 End With
End Sub

Private Sub Form_Initialize()
   ShockwaveFlash1.Movie = App.Path & "\" & "Headerfees.swf"
   ShockwaveFlash1.Play
   ShockwaveFlash1.Loop = True
    
End Sub

Private Sub Form_Load()
Call OpenDb
Call Admin
Call UserLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
Load frmAbout
frmAbout.Show
End Sub


Private Sub imgHome_Click()
Me.Show
End Sub


Private Sub lblSignUp_Click()
frmSignup.Show
End Sub


Private Sub lblSignUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSignUp.ForeColor = vbGreen
End Sub

Private Sub lblSignUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSignUp.ForeColor = vbRed
End Sub


Private Sub mnuExit_Click()
    End
End Sub


Private Sub Text2_Change()
Text2.Enabled = False
End Sub

Private Sub Timer1_Timer()
  Label1.Caption = Right$(Label1.Caption, Len(Label1.Caption) - 1) & Left(Label1.Caption, 1)
  Label6.ForeColor = vbGreen
End Sub


