VERSION 5.00
Begin VB.Form frmAbout 
   ClientHeight    =   3000
   ClientLeft      =   5205
   ClientTop       =   4170
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4695
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4695
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   2640
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    Unload Me
    End
End Sub

Private Sub Form_Load()

    Text1.Text = "(CjS)Online Fees Payment, Pin and Serial Number Generator can be use for commercial purpose. "
    Text1.Text = Text1.Text & "It may not be sold or redistributed for any profit or gratification. "
    Text1.Text = Text1.Text & "If you wish to allow other's to use it, any and all copyrights must remain intact. "
    Text1.Text = Text1.Text & "[CjS], nor myself, cant be held responsible for any damages of any kind. This program "
    Text1.Text = Text1.Text & "is available on CD for Purchase." & vbCrLf & vbCrLf
    Text1.Text = Text1.Text & "This program was only tested on Windows XP, and may work on other operating systems. "
    Text1.Text = Text1.Text & "If the program does not work, you can email me and I will see what I can do, but my time is limited, "
    Text1.Text = Text1.Text & "and this is a free program ;)"
    Text1.Text = Text1.Text & vbCrLf & vbCrLf & "Email: cyberjakes@hotmail.com" & vcrlf & vbCrLf
    Text1.Text = Text1.Text & "Mobile No: 07035339853, 08088566288"
    
End Sub

Private Sub Text1_Click()
Text1.Locked = True
End Sub
