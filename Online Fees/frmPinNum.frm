VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPinNum 
   ClientHeight    =   6000
   ClientLeft      =   2565
   ClientTop       =   3225
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmPinNum.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   9615
   Begin VB.Frame frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   9615
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
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
         Left            =   3720
         Picture         =   "frmPinNum.frx":0131
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
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
         Left            =   4560
         Picture         =   "frmPinNum.frx":0B1B
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3960
         Width           =   735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Enter Card Amount"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6960
         TabIndex        =   24
         Top             =   2760
         Width           =   2535
         Begin VB.TextBox txtCardAmt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Character Configuration For Serial Number:"
         Height          =   2055
         Left            =   3480
         TabIndex        =   15
         Top             =   120
         Width           =   3375
         Begin VB.CheckBox chkAllS 
            Caption         =   "Check All"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chkOptS 
            Caption         =   "[0-9]"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkOptS 
            Caption         =   "[a-z]"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkOptS 
            Caption         =   "[A-Z]"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CheckBox chkOptS 
            Caption         =   "[misc. (!@#%) etc.]"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtCharsS 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtMiscS 
            Height          =   285
            Left            =   400
            TabIndex        =   16
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "8-25 Characters"
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   3000
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H8000000D&
         Caption         =   "Clear"
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdGen 
         BackColor       =   &H8000000D&
         Caption         =   "Generate"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Character Configuration For Pin Number:"
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtMiscP 
            Height          =   285
            Left            =   400
            TabIndex        =   13
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox chkAllP 
            Caption         =   "Check All"
            Height          =   255
            Left            =   1920
            TabIndex        =   6
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCharsP 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chkOptP 
            Caption         =   "[misc. (!@#%) etc.]"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CheckBox chkOptP 
            Caption         =   "[A-Z]"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CheckBox chkOptP 
            Caption         =   "[a-z]"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkOptP 
            Caption         =   "[0-9]"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   2535
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3000
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label1 
            Caption         =   "10-25 Characters"
            Height          =   255
            Left            =   600
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pin Number                                    Serial Number"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   6735
         Begin VB.TextBox txtSNum 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Century Schoolbook"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtPNum 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Century Schoolbook"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Image Image2 
         Height          =   2100
         Left            =   7200
         Picture         =   "frmPinNum.frx":109D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2265
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   975
         Left            =   3480
         Shape           =   4  'Rounded Rectangle
         Top             =   3840
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin  and Serial Number Generator Interface"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   480
      TabIndex        =   26
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmPinNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAllP_Click()
S = 0
    If chkAllP.Value = "1" Then
        For S = 0 To 3
            chkOptP(S).Value = "1"
        Next
    ElseIf chkAllP.Value = "0" Then
        For S = 0 To 3
            chkOptP(S).Value = "0"
        Next
    End If

End Sub

Private Sub chkAllS_Click()
t = 0
    If chkAllS.Value = "1" Then
        For t = 0 To 3
            chkOptS(t).Value = "1"
        Next
    ElseIf chkAllS.Value = "0" Then
        For t = 0 To 3
            chkOptS(t).Value = "0"
        Next
    End If

End Sub

Private Sub cmdAbout_Click()

    frmAbout.Show

End Sub

Private Sub cmdClear_Click()
    txtCardAmt.Text = ""
    txtPNum.Text = ""
    txtSNum = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Me.Hide
    frmAdmin.Enabled = True
    frmAdmin.Show
End Sub

Private Sub SerialNumGen()
iSelectS = ""
strCharsetS = ""
strCustomizeS = txtMiscS.Text
Charset7 = strCustomizeS

S = 0

    If txtCharsS.Text = "" Then
        MsgBox ("You Must Enter a Numeric Value.")
        Exit Sub
    ElseIf chkOptS(0).Value = "0" And chkOptS(1).Value = "0" And chkOptS(2).Value = "0" And chkOptS(3).Value = "0" Then
        MsgBox ("You Must Choose an option.")
        Exit Sub
    Else
    
        If chkOptS(0).Value > 0 Then
            strCharsetS = strCharsetS & Charset4
        End If
        
        If chkOptS(1).Value > 0 Then
            strCharsetS = strCharsetS & Charset5
        End If
        
        If chkOptS(2).Value > 0 Then
            strCharsetS = strCharsetS & Charset6
        End If
        
        If chkOptS(3).Value > 0 Then
            strCharsetS = strCharsetS & Charset7
        End If
            
        RandomSerialNum (iLenS)
        'txtSNum.Text = ""
    End If
End Sub
Private Sub cmdGen_Click()
frmWait.lblPin.Visible = True
frmWait.Show vbModal
txtCardAmt.Text = ""
iSelect = ""
strCharset = ""
strCustomize = txtMiscP.Text
Charset3 = strCustomize
X = 0


    If txtCharsP.Text = "" Then
        MsgBox ("You Must Enter a Numeric Value.")
        Exit Sub
    ElseIf chkOptP(0).Value = "0" And chkOptP(1).Value = "0" And chkOptP(2).Value = "0" And chkOptP(3).Value = "0" Then
        MsgBox ("You Must Choose an option.")
        Exit Sub
    Else
    
        If chkOptP(0).Value > 0 Then
            strCharset = strCharset & Charset0
        End If
        
        If chkOptP(1).Value > 0 Then
            strCharset = strCharset & Charset1
        End If
        
        If chkOptP(2).Value > 0 Then
            strCharset = strCharset & Charset2
        End If
        
        If chkOptP(3).Value > 0 Then
            strCharset = strCharset & Charset3
        End If
            
        RandomPassword (iLen)
        'txtPNum.Text = ""
        cmdSave.Enabled = True
        txtCardAmt.SetFocus
    End If
Call SerialNumGen
End Sub

Private Sub cmdSave_Click()
    Dim sFile As String
    Dim iFile As Integer
    iFile = FreeFile
    
    If txtPNum.Text = "" And txtSNum.Text = "" And txtCardAmt.Text = "" Then
        MsgBox ("Nothing to Save. Please create a Pin, Serial Number and Enter the Card Amount before saving.")
        txtCharsP.SetFocus
    Else
        'MsgBox ("Save Code Here.")
       ' dlgCommon.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
       ' dlgCommon.ShowSave
        'sFile = dlgCommon.fileName
        
         '   Open sFile For Append As #iFile
          '  Write #iFile, txtPword.Text
           ' MsgBox ("File " & sFile & " saved!")
        'Close #iFile

With rsUnUsedCards
                .MoveFirst
                .Find "SerialNumber='" & txtSNum & "'"
                .Find "PinNumber='" & txtPNum & "'"
                If .EOF Then
                    .AddNew
                    .Fields(1) = txtSNum
                    .Fields(2) = txtPNum
                    .Fields(3) = txtCardAmt
                    .Update
                    frmWait.lblPinNum.Visible = True
                    frmWait.Show vbModal
                    MsgBox "Pin and Serial Number Generated and saved successfuly", vbInformation, "Pin Number"
                Else
                    MsgBox "Error Saving Pin and Serial Number", vbInformation, "Pin Error"
                    MsgBox "Duplicate Pin and Serial Numbers are not Allowed: Click on 'Generate Button' again...", vbInformation, "Pin Error"
                    txtCardAmt.SetFocus
                    cmdSave.Enabled = True
                End If

        End With
        cmdSave.Enabled = False
End If
End Sub

Private Sub Command1_Click()
 RandomPassword (iLen)
 RandomPassword (iLenS)
End Sub

Private Sub Form_Load()
Call OpenDb
Call UnUsedCards
    bLoaded = "0"
    frmPinNum.Caption = "[CjS] Pin and Serial Number Generator - v" & App.Major & "." & App.Minor
    txtCharsP.Text = "10"
    chkOptP(2).Value = "1"
    chkOptP(0).Value = "1"
    txtCharsS.Text = "8"
    chkOptS(0).Value = "1"
    iLenS = txtCharsS.Text
    txtMiscS.Text = "!@#$%^&*-_="
    iLen = txtCharsP.Text
    txtMiscP.Text = "!@#$%^&*-_="
End Sub

Private Sub txtCharsP_LostFocus()
'dont let them enter bad data

 Select Case txtCharsP.Text
        Case "10"
            iLen = txtCharsP.Text
        Case "11"
            iLen = txtCharsP.Text
        Case "12"
            iLen = txtCharsP.Text
        Case "13"
            iLen = txtCharsP.Text
        Case "14"
            iLen = txtCharsP.Text
        Case "15"
            iLen = txtCharsP.Text
        Case "16"
            iLen = txtCharsP.Text
        Case "17"
            iLen = txtCharsP.Text
        Case "18"
            iLen = txtCharsP.Text
        Case "19"
            iLen = txtCharsP.Text
        Case "20"
            iLen = txtCharsP.Text
        Case "21"
            iLen = txtCharsP.Text
        Case "22"
            iLen = txtCharsP.Text
        Case "23"
            iLen = txtCharsP.Text
        Case "24"
            iLen = txtCharsP.Text
        Case "25"
            iLen = txtCharsP.Text
        Case Else
            Beep
            txtCharsP.Text = "10"
            MsgBox ("Must Enter a Numeric Value between 10 and 25.")
            txtCharsP.SetFocus
    End Select
    
End Sub


Private Sub txtCharsS_LostFocus()
'dont let them enter bad data

 Select Case txtCharsS.Text
        Case "8"
            iLenS = txtCharsS.Text
        Case "9"
            iLenS = txtCharsS.Text
        Case "10"
            iLenS = txtCharsS.Text
        Case "11"
            iLenS = txtCharsS.Text
        Case "12"
            iLenS = txtCharsS.Text
        Case "13"
            iLenS = txtCharsS.Text
        Case "14"
            iLenS = txtCharsS.Text
        Case "15"
            iLenS = txtCharsS.Text
        Case "16"
            iLenS = txtCharsS.Text
        Case "17"
            iLenS = txtCharsS.Text
        Case "18"
            iLenS = txtCharsS.Text
        Case "19"
            iLenS = txtCharsS.Text
        Case "20"
            iLenS = txtCharsS.Text
        Case "21"
            iLenS = txtCharsS.Text
        Case "22"
            iLenS = txtCharsS.Text
        Case "23"
            iLenS = txtCharsS.Text
        Case "24"
            iLenS = txtCharsS.Text
        Case "25"
            iLenS = txtCharsS.Text
        Case Else
            Beep
            txtCharsS.Text = "8"
            MsgBox ("Must Enter a Numeric Value between 8 and 25.")
            txtCharsS.SetFocus
    End Select

End Sub
