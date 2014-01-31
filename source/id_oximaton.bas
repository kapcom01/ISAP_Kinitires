Attribute VB_Name = "id_oximaton"
Public Function typos_oximatos(id_no As String) As String
If Len(id_no) = 3 Or 4 Then
    If Mid(Right(id_no, 3), 1, 1) = 1 Then
        typos_oximatos = "���"
    ElseIf Mid(Right(id_no, 3), 1, 1) = 2 Then
        typos_oximatos = "��"
    ElseIf Mid(Right(id_no, 3), 1, 1) = 3 Then
        typos_oximatos = "��"
    End If
End If
End Function

Public Function paralabi_oximatos(id_no As String) As String
If Len(id_no) = 3 Then
    If Mid(id_no, 1, 1) = 1 Then
        If id_no <= 145 Then
            paralabi_oximatos = "8��"
        Else
            paralabi_oximatos = "10��"
        End If
    ElseIf Mid(id_no, 1, 1) = 2 Then
        If id_no <= 215 Then
            paralabi_oximatos = "8��"
        Else
            paralabi_oximatos = "10��"
        End If
    ElseIf Mid(id_no, 1, 1) = 3 Then
        If id_no <= 315 Then
            paralabi_oximatos = "8��"
        Else
            paralabi_oximatos = "10��"
        End If
    End If
ElseIf Len(id_no) = 4 Then
    paralabi_oximatos = "11��"
End If
End Function
