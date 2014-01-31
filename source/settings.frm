VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_settings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ρυθμίσεις"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   25
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ενεργοποίηση κατάστασης στη γραμμή εργαλείων"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Πηρύνας προγράμματος (μην το πειράζεις άμα δεν ξέρεις)"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   0
      TabIndex        =   17
      Top             =   2880
      Width           =   8535
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "στη διεύθυνση του προγράμματος"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   600
         Width           =   4575
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "σε Δισκέτα"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "στο φάκελο:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Text            =   "C:"
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Αναζήτηση"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog cd2 
         Left            =   7920
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Θέση αποθήκευσης της βάσης δεδομένων:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.TextBox txt_sec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      MaxLength       =   5
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Αναζήτηση"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   4800
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6960
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt_path 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Text            =   "C:"
      Top             =   5160
      Width           =   4695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "στο φάκελο:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "σε Δισκέτα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Χρήση κωδικού κατά την : αποθήκευση και επεξεργασία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "’κυρο"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Αποθήκευση + ΟΚ"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   8760
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line4 
      X1              =   -120
      X2              =   8640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(σε ms, 1000ms=1sec)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Χρόνος αναμονής της μπάρας ""Εργαλεία"":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Θέση αποθήκευσης των αντιγραφων ασφαλείας:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ρυθμίσεις:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   8640
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Επιβεβαίωση:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4365
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Νέος:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Παλιός:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frm_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
If Check2.Value = 1 Then
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
Else
    Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Check2.Value = 1 Then
    WriteINI "HSAP_1_0", "protection", "on", App.Path & "\Settings.ini"
    frm_main.protection = "on"
Else
    WriteINI "HSAP_1_0", "protection", "off", App.Path & "\Settings.ini"
    frm_main.protection = "off"
End If

If Text1 <> "" Then
    If Text1 = ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then
        If Text2.Text = Text3.Text Then
            WriteINI "HSAP_1_0", "password", Text2, App.Path & "\Settings.ini"
            MsgBox "Ο κωδικός άλλαξε με επιτυχία.", , "Νέος κωδικός"
        Else
            MsgBox "Ο νέος κωδικός δεν είναι ίδιος με αυτόν που δώθηκε κατά την επιβεβαίωση. Η αλλαγή δεν έγινε.", vbCritical, "Αποτυχία επιβεβαίωσης νέου κωδικού"
        End If
    Else
        MsgBox "Ο κωδικός δεν είναι σωστός. Η αλλαγή δεν έγινε.", vbCritical, "Η πρόσβαση δεν επιτράπηκε"
    End If
End If
If Option1.Value = True Then
    backup_path = "A:\data1.mdb"
    WriteINI "HSAP_1_0", "backup_path", backup_path, App.Path & "\Settings.ini"
ElseIf Option2.Value = True Then
    backup_path = txt_path
    WriteINI "HSAP_1_0", "backup_path", backup_path, App.Path & "\Settings.ini"
End If
If Option4.Value = True Then
    WriteINI "HSAP_1_0", "core_path", "A:\data1.mdb", App.Path & "\Settings.ini"
ElseIf Option5.Value = True Then
    WriteINI "HSAP_1_0", "core_path", "app.path", App.Path & "\Settings.ini"
ElseIf Option3.Value = True Then
    WriteINI "HSAP_1_0", "core_path", Text4, App.Path & "\Settings.ini"
End If
On Error GoTo timererror
If txt_sec > 10000 Or txt_sec < 500 Then
    MsgBox "Ο χρόνος πρέπει να είναι από μισό(500ms) εως 10 δευτερόλεπτα(10000ms)."
    txt_sec = frm_main.Timer3
Else
    frm_main.Timer3.Interval = txt_sec
    WriteINI "HSAP_1_0", "tools_timer", txt_sec, App.Path & "\Settings.ini"
End If

If Check1.Value = 1 Then
        WriteINI "HSAP_1_0", "status", "on", App.Path & "\Settings.ini"
        frm_main.Timer4.Enabled = True
        frm_main.lbl_edit.Visible = True
        frm_main.lbl_history.Visible = True
Else
    WriteINI "HSAP_1_0", "status", "off", App.Path & "\Settings.ini"
        frm_main.Timer4.Enabled = False
        frm_main.lbl_edit.Visible = False
        frm_main.lbl_history.Visible = False
End If

Unload Me
Exit Sub
timererror:
    MsgBox "Μονο αριθμοί επιτρέπονται!"
    txt_sec = frm_main.Timer3
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
cd.FileName = "data1.mdb"
cd.Filter = "Αρχείο βάσης δεδομένων | *.mdb"
cd.ShowSave
txt_path = cd.FileName
End Sub

Private Sub Command4_Click()
cd2.FileName = "data1.mdb"
cd2.Filter = "Αρχείο βάσης δεδομένων | *.mdb"
cd2.ShowSave
Text4 = cd.FileName
End Sub

Private Sub form_load()
If frm_main.protection = "on" Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If

If Mid(ReadINI("HSAP_1_0", "core_path", App.Path & "\Settings.ini"), 1, 3) = "A:\" Then
    Option4.Value = True
ElseIf ReadINI("HSAP_1_0", "core_path", App.Path & "\Settings.ini") = "app.path" Then
    Option5.Value = True
Else
    Option3.Value = True
End If
txt_sec = frm_main.Timer3.Interval
If Mid(backup_path, 1, 3) = "A:\" Then
    Option1.Value = True
Else
    Option2.Value = True
    txt_path.Text = backup_path
End If
If ReadINI("HSAP_1_0", "protection", App.Path & "\Settings.ini") = "off" Then
    Check2.Value = 0
ElseIf ReadINI("HSAP_1_0", "protection", App.Path & "\Settings.ini") = "on" Then
    Check2.Value = 1
End If

If ReadINI("HSAP_1_0", "status", App.Path & "\Settings.ini") = "on" Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If


End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    txt_path.Enabled = False
    Command3.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    txt_path.Enabled = True
    Command3.Enabled = True
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
    Text4.Enabled = True
    Command4.Enabled = True
Else
    Text4.Enabled = False
    Command4.Enabled = False
End If
End Sub

Private Sub Option4_Click()
If Option3.Value = True Then
    Text4.Enabled = True
    Command4.Enabled = True
Else
    Text4.Enabled = False
    Command4.Enabled = False
End If
End Sub

Private Sub Option5_Click()
If Option3.Value = True Then
    Text4.Enabled = True
    Command4.Enabled = True
Else
    Text4.Enabled = False
    Command4.Enabled = False
End If
End Sub
