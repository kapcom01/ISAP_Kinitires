Attribute VB_Name = "copyfiles"
Public backup_path As String
Public core_path As String
Public Function CopyFile(Source As String, Destiny As String, Optional BlockSize As Long = 32765) As Boolean
    On Error GoTo CopyFile_Err
    Dim Pos As Long
    Dim posicao As Long
    Dim pbyte As String
    Dim buffer As Long
    Dim Exist As String
    Dim LenSource As Long
    Dim FFSource As Integer, FFDestiny As Integer
    
100     buffer = BlockSize
102     posicao = 1
104     Exist = ""
106     Exist = Dir$(Destiny)
108     If Exist <> "" Then Kill Destiny
110     FFSource = FreeFile
112     Open Source For Binary As #FFSource
114     FFDestiny = FreeFile
116     Open Destiny For Binary As #FFDestiny
118     LenSource = LOF(FFSource)
120     For Pos = 1 To LenSource Step buffer
122     If Pos + buffer > LenSource Then buffer = (LenSource - Pos) + 1
124     pbyte = Space$(buffer)
126     Get #FFSource, Pos, pbyte
128     Put #FFDestiny, posicao, pbyte
130     posicao = posicao + buffer

132 frm_unloading.ProgressBar2.Value = (Round((((Pos / 100) * 100) / (LenSource / 100)), 2))
134 DoEvents

Next
136 Close #FFSource
138 Close #FFDestiny
140 End
Exit Function
CopyFile_Err:
MsgBox "Σφάλμα κατά τη διμιουργία αντίγραφου ασφαλείας!" & vbCrLf & _
"Αριθμός: " & Err.Number & vbCrLf & _
"Γραμμή: " & Erl & vbCrLf & vbCrLf & _
"Περιγραφή: " & Err.Description & vbCrLf & vbCrLf & _
"Η διαδικασία ακυρώθηκε!", vbCritical, "Παρουσιάστηκε σφάλμα!"
Screen.MousePointer = vbDefault
End
End Function

