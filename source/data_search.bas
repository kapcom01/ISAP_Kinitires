Attribute VB_Name = "data_search"
Public rs As ADODB.Recordset
Public conn As ADODB.Connection

Public Sub db_connect()
On Error GoTo errorhandler
Set conn = New ADODB.Connection

conn.CursorLocation = adUseClient
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & core_path & ";Mode=ReadWrite;Persist Security Info=False"

'If Not rs.EOF = True Then
'rs.MoveLast
'rs.MoveFirst
'End If

Exit Sub
errorhandler:
    MsgBox Err.Description
End Sub

Public Sub db_query(SQL_query As String)
On Error GoTo errorhandler
Set rs = New ADODB.Recordset
rs.Open SQL_query, conn, adOpenStatic, adLockOptimistic, adCmdText
Exit Sub
errorhandler:
    MsgBox Err.Description, , "Σφάλμα στο db_query (data_search.mod)"
End Sub

Public Sub db_close()
    conn.Close
    Set conn = Nothing
    Set rs = Nothing
End Sub

Public Sub rs_close()
On Error GoTo errorhandler
    rs.Close
    Set rs = Nothing
Exit Sub
errorhandler:
MsgBox Err.Description, , "Error: " & Err.Number & " στο rs_close (data_search.mod)"
End Sub
