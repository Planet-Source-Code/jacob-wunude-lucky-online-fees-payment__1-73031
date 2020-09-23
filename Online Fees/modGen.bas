Attribute VB_Name = "modGen"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROGRAM:   (CjS)Pin and Serial Number Generator                         '
'AUTHOR:    Jacob Wunude                                       '
'DATE:      19/06/2009                                      '
'COPYRIGHT: Copyright 2009 All Rights Reserved              '
'                                                           '
'COMMENTS:  This program is a simple password generator.    '
'           Use at your own risk. [CjS] can NOT be held      '
'           responsible for any damages done.               '
'           If you need support, you can email me at        '
'           cyberjakes@hotmail.com. Although, I have     '
'           limited time, I will answer your emails as      '
'           soon as possible.                               '
'                                                           '
'THANKS:    To Clint for the better code. :) Appreciate it  '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOTE:      *** THE ABOVE MUST REMAIN HERE. REMOVAL OF THIS IS A VIOLATION OF THE COPYRIGHT ***'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const Charset0 As String = "0123456789"                  'numbers
Public Const Charset1 As String = "abcdefghijklmnopqrstuvwxyz"  'lower case
Public Const Charset2 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 'upper case
Public Charset3 As String                                       'symbols.. customize on form
Public Const Charset4 As String = "0123456789"                  'numbers
Public Const Charset5 As String = "abcdefghijklmnopqrstuvwxyz"  'lower case
Public Const Charset6 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 'upper case
Public Charset7 As String                                       'symbols.. customize on form

Public strCustomize As String
Public strCharset As String
Public strCustomizeS As String
Public strCharsetS As String
Public iLenS As String
Public sPwordS As String
Public iLen As String
Public sPword As String
Public S As Integer
Public i As Integer
Public X As Integer

Public Function RandomPassword(Optional PasswordLen As Long = 25) As String
Randomize

   Do Until Len(RandomPassword) >= PasswordLen
      RandomPassword = RandomPassword & Mid(strCharset, Int((Len(strCharset) * Rnd) + 1), 1)
   Loop
   frmPinNum.txtPNum.Text = RandomPassword
   frmPinNum.txtSNum.Text = RandomPassword
End Function

Public Function RandomSerialNum(Optional SerialLen As Long = 25) As String
Randomize

   Do Until Len(RandomSerialNum) >= SerialLen
      RandomSerialNum = RandomSerialNum & Mid(strCharsetS, Int((Len(strCharsetS) * Rnd) + 1), 1)
   Loop
   frmPinNum.txtSNum.Text = RandomSerialNum
End Function

