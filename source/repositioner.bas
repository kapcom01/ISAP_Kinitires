Attribute VB_Name = "repositioner"
Public Sub reposition_search_buttons()
frm_main.shape_oxima.Left = (frm_main.frm_search.Width / 4) - 1560
frm_main.shape_kinitiras.Left = ((frm_main.frm_search.Width / 4) * 2) - 1700
frm_main.shape_hz.Left = ((frm_main.frm_search.Width / 4) * 3) - 1450
frm_main.shape_sympiestis.Left = frm_main.frm_search.Width - 2080
frm_main.lbl_oxima.Left = (frm_main.frm_search.Width / 4) - 1400
frm_main.lbl_kinitiras.Left = ((frm_main.frm_search.Width / 4) * 2) - 1600
frm_main.lbl_hz.Left = ((frm_main.frm_search.Width / 4) * 3) - 1300
frm_main.lbl_sympiestis.Left = frm_main.frm_search.Width - 1980
End Sub

Public Sub reposition_history_contents()
frm_main.tabstrip_istoriko.Width = frm_main.Width - 600
frm_main.tabstrip_istoriko.Height = frm_main.Height - 3100
frm_main.tab_blabi.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_blabi.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_sim.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_sim.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_stoixeia_ioro.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_stoixeia_ioro.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_stoixeia_kio.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_stoixeia_kio.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_stoixeia_organo.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_stoixeia_organo.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_xil1.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_xil1.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.tab_xil4.Width = frm_main.tabstrip_istoriko.Width - 200
frm_main.tab_xil4.Height = frm_main.tabstrip_istoriko.Height - 700
frm_main.list_blabi.Width = frm_main.tab_blabi.Width
frm_main.list_blabi.Height = frm_main.tab_blabi.Height
frm_main.list_blabi.ColumnHeaders(5).Width = frm_main.tab_blabi.Width - 5000
frm_main.list_xil1.Width = frm_main.tab_xil1.Width
frm_main.list_xil1.Height = frm_main.tab_xil1.Height
frm_main.list_xil1.ColumnHeaders(5).Width = frm_main.tab_xil1.Width - 5000
frm_main.list_xil4.Width = frm_main.tab_xil4.Width
frm_main.list_xil4.Height = frm_main.tab_xil4.Height
frm_main.list_xil4.ColumnHeaders(5).Width = frm_main.tab_xil4.Width - 5000
frm_main.t_sim_label.Width = frm_main.tab_sim.Width - 300
frm_main.t_txt_sim.Width = frm_main.tab_sim.Width - 200
frm_main.t_txt_sim.Height = frm_main.tab_sim.Height - 1300
frm_main.t_lbl_stoixeia_organou.Width = frm_main.tab_stoixeia_organo.Width - 300
End Sub

Public Sub reposition_main_contents()
frm_main.frm_search.Left = frm_main.Width - 600
frm_main.frm_search.Width = 200
frm_main.lbl_no_work.Left = (frm_main.Width / 2) - 2500
frm_main.lbl_no_work.Top = (frm_main.Height / 2) - 250
frm_main.lbl_title.Width = frm_main.Width - 3000
frm_main.frm_toolbar.Width = frm_main.Width
frm_main.frm_toolbar.Top = frm_main.Height - frm_main.frm_toolbar.Height - 410

If search_frame = "on" Then
    frm_main.frm_search.Width = frm_main.frm_search.Left - 1960
    frm_main.frm_search.Left = 2160
    repositioner.reposition_search_buttons
End If

If edit_frame <> "off" Then
    frm_main.frm_edit.Width = frm_main.Width - 350
    frm_main.frm_edit.Height = frm_main.Height - 3100
    frm_main.frm_history.Top = 2160
    frm_main.frm_history.Height = frm_main.Height - 3100
    frm_main.frm_history.Width = frm_main.Width - 350
    'repositioner.reposition_history_contents
    repositioner.reposition_edit_contents
ElseIf edit_frame = "off" Then
    frm_main.frm_history.Top = 1700
    frm_main.frm_history.Height = frm_main.Height - 2640
    frm_main.frm_history.Width = frm_main.Width - 350
    'repositioner.reposition_history_contents
End If

frm_main.lbl_edit.Left = frm_main.frm_toolbar.Width - 1800
frm_main.lbl_history.Left = frm_main.frm_toolbar.Width - frm_main.lbl_history.Width - 1800
frm_main.Shape2.Height = frm_main.frm_toolbar.Height
frm_main.Shape2.Width = frm_main.frm_toolbar.Width - 120
frm_main.Line1.X2 = frm_main.frm_toolbar.Width

End Sub

Public Sub reposition_edit_contents()
Dim text_height As Integer

'Τα 3 frame kio,io,ro
frm_main.lbl_oxima_KIO.Left = (frm_main.Width / 3) - 2500
frm_main.lbl_oxima_IO.Left = ((frm_main.Width / 3) * 2) - 2500
frm_main.lbl_oxima_RO.Left = frm_main.Width - 2500
frm_main.FRM_KIO(0).Width = frm_main.frm_edit.Width - 1800
frm_main.FRM_KIO(0).Height = frm_main.frm_edit.Height - 600
frm_main.FRM_KIO(1).Width = frm_main.frm_edit.Width - 1800
frm_main.FRM_KIO(1).Height = frm_main.frm_edit.Height - 600
frm_main.frm_IO.Width = frm_main.frm_edit.Width - 1800
frm_main.frm_IO.Height = frm_main.frm_edit.Height - 600
frm_main.FRM_RO.Width = frm_main.frm_edit.Width - 1800
frm_main.FRM_RO.Height = frm_main.frm_edit.Height - 600
frm_main.cmd_save.Left = frm_main.frm_edit.Width - 1350
frm_main.cmd_edit.Left = frm_main.frm_edit.Width - 1350
frm_main.cmd_exit.Left = frm_main.frm_edit.Width - 1350
frm_main.cmd_save.Top = (frm_main.frm_edit.Height / 3) - 800
frm_main.cmd_edit.Top = ((frm_main.frm_edit.Height / 3) * 2) - 800
frm_main.cmd_exit.Top = (frm_main.frm_edit.Height - 800)

'----------------------------------------------------------------------
'ΣΗΜΑΝΤΙΚΗ ΠΛΗΡΟΦΟΡΙΑ:
'---------------------
'Αν είναι να διαβάσεις τα παρακάτω για να βρεις κατι
'ξεχνα το! Οτι ακολουθεί λειτουργεί μόνο και μόνο από θαύμα!
'Ουτε εγώ καταλαβαινα τι έγραφα. Αλλαζα αριθμούς και οταν
'φαινόταν οκ το αφηνα.
'----------------------------------------------------------------------

'Τα περιεχόμενα του kio
Dim i As Integer
For i = 0 To 1
    text_height = ((frm_main.FRM_KIO(i).Height - 190) - (255 * 5) - 500) / 5
    
    frm_main.txt_kin1(i).Height = text_height
    frm_main.txt_kin1(i).Width = frm_main.FRM_KIO(i).Width - 300
    
    
    frm_main.kin2(i).Top = text_height + 255 + 100 + 90
    frm_main.over_kin2(i).Top = text_height + 255 + 100 + 90
    frm_main.cmd_replace_kin2(i).Top = text_height + 255 + 100 + 90
    frm_main.txt_kin2(i).Top = text_height + (255 * 2) + 100 + 140
    frm_main.txt_kin2(i).Height = text_height
    frm_main.txt_kin2(i).Width = frm_main.FRM_KIO(i).Width - 300
    
    frm_main.kin3(i).Top = (text_height * 2) + (255 * 2) + (100 * 2) + 90
    frm_main.over_kin3(i).Top = (text_height * 2) + (255 * 2) + (100 * 2) + 90
    frm_main.cmd_replace_kin3(i).Top = (text_height * 2) + (255 * 2) + (100 * 2) + 90
    frm_main.txt_kin3(i).Top = (text_height * 2) + (255 * 3) + (100 * 2) + 140
    frm_main.txt_kin3(i).Height = text_height
    frm_main.txt_kin3(i).Width = frm_main.FRM_KIO(i).Width - 300
    
    frm_main.kin4(i).Top = (text_height * 3) + (255 * 3) + (100 * 3) + 90
    frm_main.over_kin4(i).Top = (text_height * 3) + (255 * 3) + (100 * 3) + 90
    frm_main.cmd_replace_kin4(i).Top = (text_height * 3) + (255 * 3) + (100 * 3) + 90
    frm_main.txt_kin4(i).Top = (text_height * 3) + (255 * 4) + (100 * 3) + 140
    frm_main.txt_kin4(i).Height = text_height
    frm_main.txt_kin4(i).Width = frm_main.FRM_KIO(i).Width - 300
    
    frm_main.Label_kio(i).Top = (text_height * 4) + (255 * 4) + (100 * 4) + 90
    frm_main.txt_kio(i).Top = (text_height * 4) + (255 * 5) + (100 * 4) + 140
    frm_main.txt_kio(i).Height = text_height
    frm_main.txt_kio(i).Width = frm_main.FRM_KIO(i).Width - 300
Next i

'Τα περιεχόμενα του io
    text_height = ((frm_main.frm_IO.Height - 190) - (255 * 2) - 488) / 3
    
    frm_main.txt_sym_io.Height = text_height
    frm_main.txt_sym_io.Width = frm_main.frm_IO.Width - 300
    
    
    frm_main.hz_io.Top = text_height + 255 + 100 + 90
    frm_main.over_hz_io.Top = text_height + 255 + 100 + 90
    frm_main.cmd_replace_hz_io.Top = text_height + 255 + 100 + 90
    frm_main.txt_hz_io.Top = text_height + (255 * 2) + 100 + 140
    frm_main.txt_hz_io.Height = text_height
    frm_main.txt_hz_io.Width = frm_main.frm_IO.Width - 300
    
    frm_main.label_io.Top = (text_height * 2) + (255 * 2) + (100 * 2) + 90
    frm_main.txt_io.Top = (text_height * 2) + (255 * 3) + (100 * 2) + 140
    frm_main.txt_io.Height = text_height
    frm_main.txt_io.Width = frm_main.frm_IO.Width - 300

'Τα περιεχόμενα του ro
    text_height = ((frm_main.FRM_RO.Height - 190) - (255 * 2) - 488) / 3
    
    frm_main.txt_sym_ro.Height = text_height
    frm_main.txt_sym_ro.Width = frm_main.FRM_RO.Width - 300
    
    
    frm_main.hz_ro.Top = text_height + 255 + 100 + 90
    frm_main.over_hz_ro.Top = text_height + 255 + 100 + 90
    frm_main.cmd_replace_hz_ro.Top = text_height + 255 + 100 + 90
    frm_main.txt_hz_ro.Top = text_height + (255 * 2) + 100 + 140
    frm_main.txt_hz_ro.Height = text_height
    frm_main.txt_hz_ro.Width = frm_main.FRM_RO.Width - 300
    
    frm_main.Label_ro.Top = (text_height * 2) + (255 * 2) + (100 * 2) + 90
    frm_main.txt_ro.Top = (text_height * 2) + (255 * 3) + (100 * 2) + 140
    frm_main.txt_ro.Height = text_height
    frm_main.txt_ro.Width = frm_main.FRM_RO.Width - 300

End Sub

