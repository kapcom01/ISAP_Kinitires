VERSION 5.00
Begin VB.Form frm_nea_eisagogi 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ΝΕΑ ΕΙΣΑΓΩΓΗ"
   ClientHeight    =   2355
   ClientLeft      =   4545
   ClientTop       =   2985
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_nea_eisagogi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6450
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ΟΚ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "’κυρο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txt_oxima_no 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frm_nea_eisagogi.frx":5812
      Left            =   3240
      List            =   "frm_nea_eisagogi.frx":581F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lbl_paralabi 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός οχήματος:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Νέα Εισαγωγή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Εισαγωγή οχημάτων για:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frm_nea_eisagogi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kio0, kio1, io, ro As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim i As Integer
If Combo1.ListIndex = 0 Or Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Then
    If Len(txt_oxima_no) >= 3 Then
        db_query "SELECT * FROM ergasies WHERE 1=0"
        rs.AddNew
        If Combo1.ListIndex = 0 Then
            rs!ergasia = "xil4"
            rs!lbl_caption = "Επιθεώρηση Χιλιομετρικής 4 του συρμού " & Label4.Caption & " σε εξέλιξη..."
            rs!oxima = txt_oxima_no.Text
            symplirosi_xiliometrikis "nea_eisagogi"
        ElseIf Combo1.ListIndex = 1 Then
            rs!ergasia = "xil1"
            rs!lbl_caption = "Επιθεώρηση Χιλιομετρικής 1 του συρμού " & Label4.Caption & " σε εξέλιξη..."
            rs!oxima = txt_oxima_no.Text
            symplirosi_xiliometrikis "nea_eisagogi"
        ElseIf Combo1.ListIndex = 2 Then
            rs!ergasia = "blabi"
            rs!lbl_caption = "Επισκευή βλάβης του οχήματος " & typos_oximatos(txt_oxima_no.Text) & " " & txt_oxima_no.Text & " σε εξέλιξη..."
            rs!oxima = txt_oxima_no.Text
            If typos_oximatos(txt_oxima_no) = "ΚΙΟ" Then
                rs!kio0 = txt_oxima_no
            ElseIf typos_oximatos(txt_oxima_no) = "ΙΟ" Then
                rs!io = txt_oxima_no
            ElseIf typos_oximatos(txt_oxima_no) = "ΡΟ" Then
                rs!ro = txt_oxima_no
            End If
        End If
        rs!Date = Date
        rs.Update
        rs_close
    refresh_main_contents
    Unload Me
    Else
        MsgBox "Πρέπει να συμπληρωθεί ο αριθμός του οχήματος", , "Αριθμός οχήματος;"
    End If
Else
    MsgBox "Πρέπει να επιλέξετε το λόγο εισαγωγής του οχήματος", , "Εισαγωγή για;"
End If
End Sub

Private Sub txt_oxima_no_Change()
On Error GoTo errorhandler
Label4.Caption = ""
lbl_paralabi = ""
If Len(txt_oxima_no) >= 3 Then
    db_query ("SELECT oxima, oxima2, oxima3 FROM stoixeia WHERE oxima='" & txt_oxima_no.Text & "'")
    lbl_paralabi = paralabi_oximatos(txt_oxima_no.Text) & " παραλαβής"
    If rs.Fields("oxima3") & "" = "" Then
        Label4.Caption = rs!oxima & "-" & rs!oxima2
    Else
        Label4.Caption = rs!oxima & "-" & rs!oxima2 & "-" & rs!oxima3
    End If
    symplirosi_xiliometrikis
    rs_close
End If
Exit Sub

errorhandler:
Dim click As Integer
    If Err.Number = 3021 Then
        click = MsgBox("Ο αριθμός του οχήματος που πληκτρολογήσατε δεν υπάρχει. Θέλετε να κάνετε Νέα Καταχώρηση;", vbYesNo, "Παρουσιάστηκε σφάλμα")
        If click = vbYes Then
            Unload Me
            frm_nea_kataxorisi.Show
        Else
            txt_oxima_no.Text = ""
        End If
    Else
        MsgBox "’γνωστο σφάλμα", , Err.Number
    End If
End Sub


Private Sub symplirosi_xiliometrikis(Optional job As String)
On Error GoTo errorhandler

If job = "nea_eisagogi" Then GoTo nea_eisagogi

Dim kio_no1, kio_no2, kio_no3 As Integer
kio0 = ""
kio1 = ""
io = ""
ro = ""

    
    'διαχωρισμος βαγονιων στις μεταβλητες
If typos_oximatos(rs!oxima) = "ΙΟ" Then
    io = rs!oxima
ElseIf typos_oximatos(rs!oxima) = "ΡΟ" Then
    ro = rs!oxima
End If
If typos_oximatos(rs!oxima2) = "ΙΟ" Then
    io = rs!oxima2
ElseIf typos_oximatos(rs!oxima2) = "ΡΟ" Then
    ro = rs!oxima2
End If
If typos_oximatos(rs!oxima3 & "") = "ΙΟ" Then
    io = rs!oxima3
ElseIf typos_oximatos(rs!oxima3 & "") = "ΡΟ" Then
    ro = rs!oxima3
End If

kio_no1 = rs!oxima
kio_no2 = rs!oxima2
If IsNull(rs!oxima3) = False Then
kio_no3 = rs!oxima3
End If

If typos_oximatos(rs!oxima) = "ΚΙΟ" Then
    If typos_oximatos(rs!oxima2) = "ΚΙΟ" Then
        If kio_no1 > kio_no2 Then
            kio0 = rs!oxima2
            kio1 = rs!oxima
        Else
            kio0 = rs!oxima
            kio1 = rs!oxima2
        End If
    ElseIf typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
        If kio_no1 > kio_no3 Then
            kio0 = rs!oxima3
            kio1 = rs!oxima
        Else
            kio0 = rs!oxima
            kio1 = rs!oxima3
        End If
    Else
        kio0 = rs!oxima
    End If
ElseIf typos_oximatos(rs!oxima2) = "ΚΙΟ" Then
    If typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
        If kio_no2 > kio_no3 Then
            kio0 = rs!oxima3
            kio1 = rs!oxima2
        Else
            kio0 = rs!oxima2
            kio1 = rs!oxima3
        End If
    Else
        kio0 = rs!oxima2
    End If
ElseIf typos_oximatos(rs!oxima3 & "") = "ΚΙΟ" Then
    kio0 = rs!oxima3
End If
Exit Sub
    'αφου έχει γίνει ο διαχωρισμό, γίνετε η νεα εγγραφη στις εργασιες
nea_eisagogi:
rs!kio0 = kio0
rs!kio1 = kio1
rs!io = io
rs!ro = ro
Exit Sub

errorhandler:
    MsgBox Err.Description, , "Σφάλμα κατά την: symplirosi_xiliometrikis"
End Sub

