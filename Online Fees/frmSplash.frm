VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   210
   ClientLeft      =   3600
   ClientTop       =   6330
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerPb 
      Interval        =   100
      Left            =   720
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   4200
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   160
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Timer1_Timer()
'ProgressBar1.Value = ProgressBar1.Value + 5
'If ProgressBar1.Value = 100 Then
'Unload Me
'frmWel.Show
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = vbDefault
End Sub

Private Sub Image1_Click()

End Sub

Private Sub TimerPb_Timer()
 ProgressBar1.Value = ProgressBar1.Value + 1
    Select Case ProgressBar1.Value

    Case 20
    Label1.Caption = "Loading Please Wait."
    
    Case 40
    Label1.Caption = "Loading Please Wait.."
    
    Case 60
    Label1.Caption = "Loading Please Wait..."
    
    Case 80
    Label1.Caption = "Loading Please Wait...."
    
    Case 100
    Label1.Caption = "Loading Please Wait....."
    
    Case 120
     Label1.Caption = "Loading Please Wait......"
     
    Case 140
     Label1.Caption = "Loading Please Wait......."
      
    Case 160
     Label1.Caption = "Loading Please Wait........"
      
    End Select
    
    If ProgressBar1.Value = ProgressBar1.Max Then
    Unload Me
    frmWel.Show
    End If
End Sub
