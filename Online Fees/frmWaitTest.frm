VERSION 5.00
Begin VB.Form frmWaitTest 
   BackColor       =   &H80000007&
   ClientHeight    =   255
   ClientLeft      =   6000
   ClientTop       =   5715
   ClientWidth     =   2745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleWidth      =   2745
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   2880
      Top             =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Loading page Please Wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmWaitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
