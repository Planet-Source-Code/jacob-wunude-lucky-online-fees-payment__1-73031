VERSION 5.00
Begin VB.Form frmSignup 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4530
   ClientLeft      =   4365
   ClientTop       =   4740
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6105
   Begin VB.CommandButton cmdSignup 
      Caption         =   "&Sign-up"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   -1200
      Picture         =   "frmSignup.frx":0000
      ScaleHeight     =   420
      ScaleWidth      =   7335
      TabIndex        =   14
      Top             =   0
      Width           =   7335
      Begin VB.Image imgclose 
         Height          =   330
         Left            =   6960
         Picture         =   "frmSignup.frx":08C5
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.TextBox txtConfPassWd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "ConfPassWd"
      DataSource      =   "adoInfo"
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "+"
      TabIndex        =   4
      ToolTipText     =   "Re-Enter Password to Comfirm"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtPassWd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "PassWd"
      DataSource      =   "adoInfo"
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "+"
      TabIndex        =   3
      ToolTipText     =   "Enter Password"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtUserId 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "UserId"
      DataSource      =   "adoInfo"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Enter Username"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "E_Mail"
      DataSource      =   "adoInfo"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Enter your E-mail Address"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox txtMat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Enter Matric Number"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E3C5AC&
      Caption         =   "Let Me drag the Whole Form!!!!"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "frmSignup.frx":0EDF
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Preffered UserID"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mat/No."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF80FF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   4260
      Left            =   0
      Picture         =   "frmSignup.frx":1DA9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E3C5AC&
      BorderWidth     =   4
      Height          =   5535
      Left            =   1080
      Top             =   -240
      Width           =   7335
   End
End
Attribute VB_Name = "frmSignup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Msg     As String
Dim PassWd  As Integer
Public candragform As Boolean

Private Declare Function SendMessage Lib "User32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "User32" ()

'FOR OPENING BROWSER

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long
                

Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2


Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
    
     candragform = True
      
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If candragform = True Then
    
    Dim lngReturnValue As Long
    
        If Button = 1 Then
        
        Call ReleaseCapture
        
        'To move form I have used ME, instead you can write the control name
        lngReturnValue = SendMessage(frmSignup.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, &O1)
        
        End If
        
    End If
    
End Sub

Private Sub imgclose_Click()
Unload Me
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    
    If Button = 1 Then
    
    Call ReleaseCapture
    
    'To move form I have used ME, instead you can write the control name
    lngReturnValue = SendMessage(frmSignup.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, &O1)
    
    End If
End Sub

Public Sub OpenURL(urlADD As String, sourceHWND As Long)

     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
     
End Sub

Private Sub cmdSignup_Click()
If txtUserId = "" Or txtPassWd = "" Or txtConfPassWd = "" Or _
         txtPassWd.Text <> txtConfPassWd.Text Or txtMat.Text = "" Or _
         txtEmail.Text = "" Then
        MsgBox "Password and Comfirm password does not match", vbInformation
        Exit Sub
Else
       
        rs.MoveFirst
        rs.Find "MatNo ='" & txtMat.Text & "'"
        If rs.EOF Then
            MsgBox "Matric number not Registered! Please contact Your Administrator", vbExclamation, "Mat Error"
            txtMat.Text = ""
            txtMat.SetFocus
       Else
        
    With rsUserLog
        '.MoveFirst
        .Find "Username ='" & txtUserId.Text & "'"
        .Find "Password='" & txtPassWd.Text & "'"
        If .EOF Then
            .AddNew
            .Fields(0) = txtUserId
            .Fields(1) = txtPassWd
            .Fields(2) = txtConfPassWd
            .Fields(3) = txtEmail
            .Fields(4) = txtMat
            .Update
    With rs
            .Fields(27) = txtUserId
            .Fields(28) = txtPassWd
            .Update
             frmWait.lblAccount.Visible = True
             frmWait.Show vbModal
             MsgBox "Sign Up Successful!!!", vbInformation, "Account"
             MsgBox "Your username is  " & txtUserId.Text & "  Password not Shown", vbInformation, "Sign Up"
             Unload Me
             Me.Hide
    End With
        Else
           MsgBox "Username ! and  Password Already Exist!", vbInformation
           txtUserId = ""
           txtPassWd = ""
           txtUserId.SetFocus
       End If
    End With
End If
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
Me.Hide
End Sub

Private Sub Form_Load()
Call OpenDb
Call UserLog
Call Admin
Check1.Value = vbChecked
End Sub
