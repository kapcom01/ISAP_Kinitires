Attribute VB_Name = "refresher"
Public search_frame, toolbar_frame, edit_frame, ergasia_date, ergasia_str, edit_oxima As String
Public history_type, history_id As String
Public kio_oxima(1) As String
Public kio_i As Integer
Public antikatastasi_me As String

Public Sub refresh_main_contents()

Dim i, Y, q As Integer
Y = 0
repositioner.reposition_main_contents
If edit_frame = "off" Then


'ανανεωση των lbl εργασιων σε εξελειξη
db_query "SELECT * FROM ergasies"
If rs.RecordCount <> 0 Then
    frm_main.lbl_no_work.Visible = False
            'unload ολων εκτος του 0
            On Error Resume Next
            For q = 1 To frm_main.lbl_work_progress.Count - 1
                Unload frm_main.lbl_work_progress(q)
            Next q
    If rs.RecordCount > 1 Then
        rs.MoveFirst
        frm_main.lbl_work_progress(0).Caption = rs!lbl_caption & "(από " & rs!Date & ")"
        frm_main.lbl_work_progress(0).Visible = True
        rs.MoveNext
            
            'load οσων χρειαζεται επιπλεον απ'το 0
            For q = 1 To rs.RecordCount - 1
                Load frm_main.lbl_work_progress(q)
                frm_main.lbl_work_progress(q).Top = frm_main.lbl_work_progress(q - 1).Top + 450
                frm_main.lbl_work_progress(q).Caption = rs!lbl_caption & "(από " & rs!Date & ")"
                frm_main.lbl_work_progress(q).Visible = True
                rs.MoveNext
            Next q
    
    Else
        rs.MoveFirst
        frm_main.lbl_work_progress(0).Caption = rs!lbl_caption & "(από " & rs!Date & ")"
        frm_main.lbl_work_progress(0).Visible = True
    End If
Else
    frm_main.lbl_no_work.Visible = True
    frm_main.lbl_work_progress(0).Caption = ""
    frm_main.lbl_work_progress(0).Visible = False
    For q = 1 To frm_main.lbl_work_progress.Count - 1
        Unload frm_main.lbl_work_progress(q)
    Next q
End If
rs_close
'τελος ανανεωσης των lbl εργασιων σε εξελειξη
    
    
    frm_main.lbl_oxima_KIO.Visible = False
    frm_main.lbl_oxima_IO.Visible = False
    frm_main.lbl_oxima_RO.Visible = False
    frm_main.frm_edit.Visible = False
    frm_main.lbl_title.Caption = Date
    For i = 0 To frm_main.lbl_work_progress.Count - 1
        If frm_main.lbl_work_progress.Item(i).Caption <> "" Then
        Y = Y + 1
        End If
    Next i
    frm_main.lbl_status.Caption = Y & " εργασίες σε εξέλιξη"
ElseIf edit_frame <> "off" Then
    refresh_edit_contents "1"
    frm_main.frm_edit.Visible = True
End If

refresh_history_contents

toolbar_frame = "on"
frm_main.Timer2.Enabled = True

End Sub

Public Sub refresh_buttons(Button As String)
If Button = "oxima" Then
    frm_main.shape_oxima.BackStyle = 1
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "kinitiras" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 1
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "hz" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 1
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "sympiestis" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 1
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "kio" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &HFFFF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "io" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &HFFFF&
    frm_main.lbl_oxima_RO.ForeColor = &H80FF&
ElseIf Button = "ro" Then
    frm_main.shape_oxima.BackStyle = 0
    frm_main.shape_kinitiras.BackStyle = 0
    frm_main.shape_hz.BackStyle = 0
    frm_main.shape_sympiestis.BackStyle = 0
    frm_main.lbl_oxima_KIO.ForeColor = &H80FF&
    frm_main.lbl_oxima_IO.ForeColor = &H80FF&
    frm_main.lbl_oxima_RO.ForeColor = &HFFFF&
End If
End Sub

Public Sub refresh_edit_contents(level As String)
If level = "2" Then GoTo level2

level1:
If edit_frame = "blabi" Then
    frm_main.frm_edit.Caption = "Σημειώσεις Βλάβης"
ElseIf edit_frame = "xil4" Then
    frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 4"
ElseIf edit_frame = "xil1" Then
    frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 1"
End If

db_query "SELECT * FROM stoixeia WHERE oxima='" & edit_oxima & "'"

frm_main.lbl_oxima_KIO.Caption = typos_oximatos(edit_oxima) & ": " & edit_oxima
frm_main.lbl_oxima_KIO.Visible = True
frm_main.lbl_oxima_IO.Caption = typos_oximatos(rs.Fields("oxima2")) & ": " & rs.Fields("oxima2")
frm_main.lbl_oxima_IO.Visible = True
If Not rs.Fields("oxima3") = "" Then
frm_main.lbl_oxima_RO.Caption = typos_oximatos(rs.Fields("oxima3")) & ": " & rs.Fields("oxima3")
frm_main.lbl_oxima_RO.Visible = True
Else
frm_main.lbl_oxima_RO.Visible = False
End If

   

'ευρεση μικρότερου ΚΙΟ για (0) και μεγαλύτερου για (1)
Dim kio_no1, kio_no2, kio_no3 As Integer
kio_no1 = edit_oxima
If IsNull(rs!oxima2) = True Then
    kio_no2 = 0
Else
    kio_no2 = rs!oxima2
End If
If IsNull(rs!oxima3) = True Then
    kio_no3 = 0
Else
    kio_no3 = rs!oxima3
End If

If typos_oximatos(edit_oxima) = "ΚΙΟ" Then
    If typos_oximatos(rs!oxima2) = "ΚΙΟ" Then
        If kio_no1 > kio_no2 Then
            kio_oxima(0) = rs!oxima2
            kio_oxima(1) = edit_oxima
        Else
            kio_oxima(0) = edit_oxima
            kio_oxima(1) = rs!oxima2
        End If
    ElseIf typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
        If kio_no1 > kio_no3 Then
            kio_oxima(0) = rs!oxima3
            kio_oxima(1) = edit_oxima
        Else
            kio_oxima(0) = edit_oxima
            kio_oxima(1) = rs!oxima3
        End If
    Else
        kio_oxima(0) = edit_oxima
    End If
ElseIf typos_oximatos(rs!oxima2) = "ΚΙΟ" Then
    If typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
        If kio_no2 > kio_no3 Then
            kio_oxima(0) = rs!oxima3
            kio_oxima(1) = rs!oxima2
        Else
            kio_oxima(0) = rs!oxima2
            kio_oxima(1) = rs!oxima3
        End If
    Else
        kio_oxima(0) = rs!oxima2
    End If
ElseIf typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
    kio_oxima(0) = rs!oxima3
End If
'η ευρεση έγινε

rs_close

level2:
If edit_frame = "blabi" Then
    kio_i = 0
Else
    For kio_i = 0 To 1
        If kio_oxima(kio_i) = edit_oxima Then
            Exit For
        End If
    Next
End If

'On Error GoTo errorhandler
db_query "SELECT * FROM stoixeia WHERE oxima='" & edit_oxima & "'"
If typos_oximatos(edit_oxima) = "ΚΙΟ" Then
    
    frm_main.kin1(kio_i) = "Κινητήρας: " & rs!kinitiras1
    frm_main.kin2(kio_i) = "Κινητήρας: " & rs!kinitiras2
    frm_main.kin3(kio_i) = "Κινητήρας: " & rs!kinitiras3
    frm_main.kin4(kio_i) = "Κινητήρας: " & rs!kinitiras4
    rs_close
    
        'Ρουφιγμα όλων των πληροφοριών για το ΚΙΟ
    db_query "SELECT * FROM ergasies WHERE kio" & kio_i & "='" & edit_oxima & "'"
    frm_main.over_kin1(kio_i) = rs.Fields("over_kin1_" & kio_i).Value & ""
    frm_main.txt_kin1(kio_i) = rs.Fields("kin1_" & kio_i).Value & ""
    frm_main.over_kin2(kio_i) = rs.Fields("over_kin2_" & kio_i).Value & ""
    frm_main.txt_kin2(kio_i) = rs.Fields("kin2_" & kio_i).Value & ""
    frm_main.over_kin3(kio_i) = rs.Fields("over_kin3_" & kio_i).Value & ""
    frm_main.txt_kin3(kio_i) = rs.Fields("kin3_" & kio_i).Value & ""
    frm_main.over_kin4(kio_i) = rs.Fields("over_kin4_" & kio_i).Value & ""
    frm_main.txt_kin4(kio_i) = rs.Fields("kin4_" & kio_i).Value & ""
    frm_main.txt_kio(kio_i) = rs.Fields("txt_kio" & kio_i).Value & ""
    rs_close
        'Τέλος ρουφίγματος του ΚΙΟ
    
    frm_main.FRM_KIO(0).Visible = False
    frm_main.FRM_KIO(1).Visible = False
    frm_main.FRM_KIO(kio_i).Caption = "Σημειώσεις στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        If edit_frame = "blabi" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Βλάβης στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil4" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 4 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil1" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 1 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        End If
    frm_main.FRM_KIO(kio_i).Visible = True
    frm_main.frm_IO.Visible = False
    frm_main.FRM_RO.Visible = False

ElseIf typos_oximatos(edit_oxima) = "ΙΟ" Then
  
    frm_main.sym_io = "Συμπιεστής: " & rs!sympiestis
    frm_main.hz_io = "Η/Ζ: " & rs!hz
    rs_close
    
        'Ρουφιγμα όλων των πληροφοριών για το ΙΟ
    db_query "SELECT * FROM ergasies WHERE io='" & edit_oxima & "'"
    frm_main.over_sym_io = rs.Fields("over_sym_io").Value & ""
    frm_main.txt_sym_io = rs.Fields("sym_io").Value & ""
    frm_main.over_hz_io = rs.Fields("over_hz_io").Value & ""
    frm_main.txt_hz_io = rs.Fields("hz_io").Value & ""
    frm_main.txt_io = rs.Fields("io_txt").Value & ""
    rs_close
        'Τέλος ρουφίγματος του ΙΟ

    
    frm_main.frm_IO.Caption = "Σημειώσεις στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        If edit_frame = "blabi" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Βλάβης στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil4" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 4 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil1" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 1 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        End If
    frm_main.FRM_KIO(0).Visible = False
    frm_main.FRM_KIO(1).Visible = False
    frm_main.frm_IO.Visible = True
    frm_main.FRM_RO.Visible = False
ElseIf typos_oximatos(edit_oxima) = "ΡΟ" Then
    
    frm_main.sym_ro = "Συμπιεστής: " & rs!sympiestis
    frm_main.hz_ro = "Η/Ζ: " & rs!hz
    rs_close
    
        'Ρουφιγμα όλων των πληροφοριών για το ΡΟ
    db_query "SELECT * FROM ergasies WHERE ro='" & edit_oxima & "'"
    frm_main.over_sym_ro = rs.Fields("over_sym_ro").Value & ""
    frm_main.txt_sym_ro = rs.Fields("sym_ro").Value & ""
    frm_main.over_hz_ro = rs.Fields("over_hz_ro").Value & ""
    frm_main.txt_hz_ro = rs.Fields("hz_ro").Value & ""
    frm_main.txt_ro = rs.Fields("ro_txt").Value & ""
    rs_close
        'Τέλος ρουφίγματος του ΡΟ

    
    frm_main.FRM_RO.Caption = "Σημειώσεις στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        If edit_frame = "blabi" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Βλάβης στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil4" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 4 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        ElseIf edit_frame = "xil1" Then
            frm_main.frm_edit.Caption = "Σημειώσεις Χιλιομετρικής 1 στο " & typos_oximatos(edit_oxima) & " " & edit_oxima
        End If
    frm_main.FRM_KIO(0).Visible = False
    frm_main.FRM_KIO(1).Visible = False
    frm_main.frm_IO.Visible = False
    frm_main.FRM_RO.Visible = True
End If

Exit Sub

errorhandler:
    MsgBox "Το όχημα " & edit_oxima & " δεν έχει καταχωρημένους κινητήρες. Η εισαγωγή ακυρώνεται. Η εισαγγή θα γίνει δεκτή μόνο αν καταχρηθούν κινητήρες.", , Err.Description
End Sub


Public Sub update_eisagogis()
Dim temp As String
    db_query "SELECT * FROM ergasies WHERE lbl_caption='" & ergasia_str & "'"

If typos_oximatos(edit_oxima) = "ΚΙΟ" Then
    temp = frm_main.txt_kin1(kio_i)
    rs.Fields("over_kin1_" & kio_i) = frm_main.over_kin1(kio_i)
    rs.Fields("kin1_" & kio_i).Value = temp
    rs.Fields("over_kin2_" & kio_i).Value = frm_main.over_kin2(kio_i)
    rs.Fields("kin2_" & kio_i).Value = frm_main.txt_kin2(kio_i)
    rs.Fields("over_kin3_" & kio_i).Value = frm_main.over_kin3(kio_i)
    rs.Fields("kin3_" & kio_i).Value = frm_main.txt_kin3(kio_i)
    rs.Fields("over_kin4_" & kio_i).Value = frm_main.over_kin4(kio_i)
    rs.Fields("kin4_" & kio_i).Value = frm_main.txt_kin4(kio_i)
    rs.Fields("txt_kio" & kio_i).Value = frm_main.txt_kio(kio_i)

ElseIf typos_oximatos(edit_oxima) = "ΡΟ" Then
    rs.Fields("over_sym_ro").Value = frm_main.over_sym_ro
    rs.Fields("sym_ro").Value = frm_main.txt_sym_ro
    rs.Fields("over_hz_ro").Value = frm_main.over_hz_ro
    rs.Fields("hz_ro").Value = frm_main.txt_hz_ro
    rs.Fields("ro_txt").Value = frm_main.txt_ro

ElseIf typos_oximatos(edit_oxima) = "ΙΟ" Then
    rs.Fields("over_sym_io").Value = frm_main.over_sym_io
    rs.Fields("sym_io").Value = frm_main.txt_sym_io
    rs.Fields("over_hz_io").Value = frm_main.over_hz_io
    rs.Fields("hz_io").Value = frm_main.txt_hz_io
    rs.Fields("io_txt").Value = frm_main.txt_io

End If
    rs.Update
    rs_close
End Sub


Public Sub refresh_history_contents()
Dim q As Integer
    reposition_history_contents
If history_type <> "off" Then
    If history_type = "oxima" Then
        frm_main.frm_history.Caption = "Ιστορικό του Οχήματος " & typos_oximatos(history_id) & " " & history_id
        frm_main.frm_history.Visible = True
    ElseIf history_type = "hz" Then
        frm_main.frm_history.Caption = "Ιστορικό του Η/Ζ " & history_id
        frm_main.frm_history.Visible = True
    ElseIf history_type = "sympiestis" Then
        frm_main.frm_history.Caption = "Ιστορικό του Συμπιεστή " & history_id
        frm_main.frm_history.Visible = True
    ElseIf history_type = "kinitiras" Then
        frm_main.frm_history.Caption = "Ιστορικό του Κινητήρα " & history_id
        frm_main.frm_history.Visible = True
    End If
    frm_main.list_blabi.ListItems.Clear
    frm_main.list_xil1.ListItems.Clear
    frm_main.list_xil4.ListItems.Clear
    frm_main.tab_stoixeia_ioro.Visible = False
    frm_main.tab_stoixeia_kio.Visible = False
    frm_main.tab_stoixeia_organo.Visible = False
    frm_main.t_sim_label.Caption = ""
    frm_main.t_txt_sim = ""

If history_type = "oxima" Then
    If typos_oximatos(history_id) = "ΚΙΟ" Then
        db_query "SELECT * FROM stoixeia WHERE oxima='" & history_id & "'"
        If rs.RecordCount = 0 Then
            frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στους διαθέσιμους"
            GoTo next_job
        End If
        frm_main.t_kin1 = rs!kinitiras1 & ""
        frm_main.t_kin2 = rs!kinitiras2 & ""
        frm_main.t_kin3 = rs!kinitiras3 & ""
        frm_main.t_kin4 = rs!kinitiras4 & ""
        rs_close
    Else
        db_query "SELECT * FROM stoixeia WHERE oxima='" & history_id & "'"
        If rs.RecordCount = 0 Then
            frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στους διαθέσιμους"
            GoTo next_job
        End If
        frm_main.t_hz = rs!hz & ""
        frm_main.t_sym = rs!sympiestis & ""
        rs_close
    End If
ElseIf history_type = "kinitiras" Then
    db_query "SELECT * FROM stoixeia"
        If rs.RecordCount = 0 Then
            frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στους διαθέσιμους"
            GoTo next_job
        End If
    For q = 1 To rs.RecordCount
        If rs!kinitiras1 & "" = history_id Or rs!kinitiras2 & "" = history_id Or rs!kinitiras3 & "" = history_id And rs!kinitiras4 & "" = history_id Then
            frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στο " & typos_oximatos(rs!oxima) & " " & rs!oxima
            Exit For
        End If
    Next q
    rs_close
Else
    db_query "SELECT * FROM stoixeia WHERE " & history_type & "='" & history_id & "'"
        If rs.RecordCount = 0 Then
            frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στους διαθέσιμους"
            GoTo next_job
        End If
    frm_main.t_lbl_stoixeia_organou = "Βρίσκεται στο " & typos_oximatos(rs!oxima) & " " & rs!oxima
    rs_close
End If
next_job:
db_query "SELECT * FROM istoriko WHERE history_type='" & history_type & "' AND history_id='" & history_id & "'"
    frm_main.history_na.Visible = False
    frm_main.tabstrip_istoriko.Visible = True
If rs.RecordCount > 1 Then
    rs.MoveFirst
    For q = 1 To rs.RecordCount
        If rs!ergasia = "xil1" Then
            frm_main.list_xil1.ListItems.Add
            frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
            frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
            frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
            frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(4) = rs!simeioseis & ""
        ElseIf rs!ergasia = "xil4" Then
            frm_main.list_xil4.ListItems.Add
            frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
            frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
            frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
            frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(4) = rs!simeioseis & ""
        ElseIf rs!ergasia = "blabi" Then
            frm_main.list_blabi.ListItems.Add
            frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
            frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
            frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
            frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(4) = rs!simeioseis & ""
        End If
        rs.MoveNext
    Next q
ElseIf rs.RecordCount = 1 Then
    rs.MoveFirst
    If rs!ergasia = "xil1" Then
        frm_main.list_xil1.ListItems.Add
        frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
        frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
        frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
        frm_main.list_xil1.ListItems(frm_main.list_xil1.ListItems.Count).SubItems(4) = rs!simeioseis & ""
    ElseIf rs!ergasia = "xil4" Then
        frm_main.list_xil4.ListItems.Add
        frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
        frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
        frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
        frm_main.list_xil4.ListItems(frm_main.list_xil4.ListItems.Count).SubItems(4) = rs!simeioseis & ""
    ElseIf rs!ergasia = "blabi" Then
        frm_main.list_blabi.ListItems.Add
        frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(1) = rs!date_in & " - " & rs!date_out
        frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(2) = rs.Fields("8esi_prin") & " > " & rs.Fields("8esi_meta") & ""
        frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(3) = rs!antikatastasi_me & ""
        frm_main.list_blabi.ListItems(frm_main.list_blabi.ListItems.Count).SubItems(4) = rs!simeioseis & ""
    End If
End If
rs_close

Else
frm_main.frm_history.Visible = False
End If

End Sub

