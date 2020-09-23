Attribute VB_Name = "modReg"
Public con As ADODB.Connection
'Public  As ADODB.Recordset
'Public  As ADODB.Recordset
Public rsUsedcards, rsCtrans, rsUnUsedCards, rs, rsUserLog, rsReceipts As Recordset
'Public  As ADODB.Recordset
'Public  As ADODB.Recordset
Public ActUser As String
Public CardAmt  As Integer
Public Amount  As Double
Public RepAmt As Double
Public strPictureName
Public FN As String
Public FT As String
Public NewRepNo As String

Public Sub OpenDb()
Set con = New ADODB.Connection
con.Provider = "Microsoft.Jet.OLEDB.4.0"
con.Open App.Path & "./OnlineFees.mdb"
End Sub

Public Sub UserLog()
Set rsUserLog = New ADODB.Recordset
    With rsUserLog
                .ActiveConnection = con
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open "Login"
    End With
End Sub

Public Sub Receipts()
Set rsReceipts = New ADODB.Recordset
    With rsReceipts
                    .ActiveConnection = con
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open "Receipts"
    End With
End Sub

Public Sub Receipt()
Set rsReceipts = New ADODB.Recordset
    With rsReceipts
                    .ActiveConnection = con
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    '.Open "Receipts"
    End With
End Sub

Public Sub UsedCards()
Set rsUsedcards = New ADODB.Recordset
    With rsUsedcards
                    .ActiveConnection = con
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open "Usedcards"
    End With
End Sub

Public Sub Ctrans()
Set rsCtrans = New ADODB.Recordset
    With rsCtrans
                .ActiveConnection = con
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockOptimistic
                .Open "CardTransactions"
    End With
End Sub

Public Sub UnUsedCards()
Set rsUnUsedCards = New ADODB.Recordset
    With rsUnUsedCards
                    .ActiveConnection = con
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Open "UnUsedCards"
    End With
End Sub

Public Sub Admin()
Set rs = New ADODB.Recordset
    With rs
            .ActiveConnection = con
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "Admin"
    End With
End Sub
