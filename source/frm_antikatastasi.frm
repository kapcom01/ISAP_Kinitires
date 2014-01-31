VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_antikatastasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Διαθέσιμοι "
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frm_antikatastasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frm_antikatastasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public label_over As String

Private Sub form_load()
ListView1.ListItems.Clear
db_query "SELECT * FROM dia8esimoi_" & antikatastasi_me & " WHERE paralabi='" & paralabi_oximatos(edit_oxima) & "'"
Me.Caption = Me.Caption & antikatastasi_me & " " & paralabi_oximatos(edit_oxima)
If Not rs.RecordCount = 0 Then
    For m = 1 To rs.RecordCount
        ListView1.ListItems.Add m, , rs.Fields("id_" & antikatastasi_me)
        rs.MoveNext
    Next
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frm_main.Enabled = True
End Sub

Private Sub ListView1_DblClick()
If label_over = "kin10" Then
    frm_main.over_kin1(0) = ListView1.SelectedItem
ElseIf label_over = "kin20" Then
    frm_main.over_kin2(0) = ListView1.SelectedItem
ElseIf label_over = "kin30" Then
    frm_main.over_kin3(0) = ListView1.SelectedItem
ElseIf label_over = "kin40" Then
    frm_main.over_kin4(0) = ListView1.SelectedItem
ElseIf label_over = "kin11" Then
    frm_main.over_kin1(1) = ListView1.SelectedItem
ElseIf label_over = "kin21" Then
    frm_main.over_kin2(1) = ListView1.SelectedItem
ElseIf label_over = "kin31" Then
    frm_main.over_kin3(1) = ListView1.SelectedItem
ElseIf label_over = "kin41" Then
    frm_main.over_kin4(1) = ListView1.SelectedItem
ElseIf label_over = "sym_io" Then
    frm_main.over_sym_io = ListView1.SelectedItem
ElseIf label_over = "hz_io" Then
    frm_main.over_hz_io = ListView1.SelectedItem
ElseIf label_over = "sym_ro" Then
    frm_main.over_sym_ro = ListView1.SelectedItem
ElseIf label_over = "hz_ro" Then
    frm_main.over_hz_ro = ListView1.SelectedItem
End If
Unload Me
frm_main.Enabled = True
End Sub
