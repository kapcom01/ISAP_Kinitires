VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_main 
   BackColor       =   &H00008000&
   Caption         =   "Η.Σ.Α.Π. - Τμήμα κινητήρων   -   Έκδοση 1.1  -  by Kapcom A.E."
   ClientHeight    =   6600
   ClientLeft      =   1230
   ClientTop       =   1290
   ClientWidth     =   9060
   ForeColor       =   &H00FFFFFF&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin VB.Frame frm_toolbar 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Εργαλεία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1005
      Left            =   0
      MouseIcon       =   "main.frx":5812
      TabIndex        =   8
      Top             =   5520
      Width           =   12495
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ημερολόγιο"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10080
         MouseIcon       =   "main.frx":5B1C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   400
         Width           =   1455
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   5280
         Top             =   240
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1920
         Top             =   240
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2280
         Top             =   240
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ρυθμίσεις"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7080
         MouseIcon       =   "main.frx":5E26
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Νέα εισαγωγή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   600
         MouseIcon       =   "main.frx":6130
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   400
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Νέα καταχώρηση"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3000
         MouseIcon       =   "main.frx":643A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label lbl_history 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   84
         Top             =   15
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label lbl_edit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   83
         Top             =   15
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1320
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9220
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   9135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Εργαλεία"
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
         Left            =   120
         MouseIcon       =   "main.frx":6744
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frm_history 
      BackColor       =   &H00008000&
      Caption         =   "Ιστορικό οχήματος"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3840
      Left            =   120
      TabIndex        =   80
      Top             =   1700
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame tab_stoixeia_kio 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   97
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin VB.Label t_kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":6A4E
            MousePointer    =   99  'Custom
            TabIndex        =   105
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lt_kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 1:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   104
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lt_kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 2:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   103
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lt_kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 3:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   102
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lt_kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 4:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   101
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label t_kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":6D58
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label t_kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":7062
            MousePointer    =   99  'Custom
            TabIndex        =   99
            Top             =   2640
            Width           =   2895
         End
         Begin VB.Label t_kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":736C
            MousePointer    =   99  'Custom
            TabIndex        =   98
            Top             =   3240
            Width           =   2895
         End
      End
      Begin VB.Frame tab_stoixeia_ioro 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   92
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin VB.Label t_hz 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":7676
            MousePointer    =   99  'Custom
            TabIndex        =   96
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Η/Ζ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   95
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Συμπιεστής :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   94
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label t_sym 
            BackStyle       =   0  'Transparent
            Caption         =   "0000000000000000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   2160
            MouseIcon       =   "main.frx":7980
            MousePointer    =   99  'Custom
            TabIndex        =   93
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.Frame tab_sim 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   89
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin VB.TextBox t_txt_sim 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   120
            TabIndex        =   90
            Text            =   "Text1"
            Top             =   1200
            Width           =   6615
         End
         Begin VB.Label t_sim_label 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   6735
         End
      End
      Begin VB.Frame tab_blabi 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   87
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin MSComctlLib.ListView list_blabi 
            Height          =   1935
            Left            =   0
            TabIndex        =   88
            Top             =   0
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ημερομηνία"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Από > Σε"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Αντικατάσταση με"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Σημειώσεις"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame tab_xil1 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   108
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin MSComctlLib.ListView list_xil1 
            Height          =   1935
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ημερομηνία"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Από > Σε"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Αντικατάσταση με"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Σημειώσεις"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame tab_stoixeia_organo 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   111
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin VB.Label t_lbl_stoixeia_organou 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   240
            TabIndex        =   112
            Top             =   480
            Width           =   6375
         End
      End
      Begin VB.Frame tab_xil4 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   240
         TabIndex        =   106
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         Begin MSComctlLib.ListView list_xil4 
            Height          =   1095
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   1931
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ημερομηνία"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Από > Σε"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Αντικατάσταση με"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Σημειώσεις"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSComctlLib.TabStrip tabstrip_istoriko 
         Height          =   5415
         Left            =   120
         TabIndex        =   110
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   9551
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Στοιχεία"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Χιλιομετρική 1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Χιλιομετρική 4"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Βλάβη"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Σημειώσεις"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label history_label 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "main.frx":7C8A
         MousePointer    =   99  'Custom
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   15015
      End
      Begin VB.Label history_na 
         BackStyle       =   0  'Transparent
         Caption         =   "Το ιστορικό δεν είναι διαθέσιμο"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   840
         TabIndex        =   81
         Top             =   2520
         Width           =   7335
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14040
      Top             =   360
   End
   Begin VB.Frame frm_search 
      BackColor       =   &H00008000&
      Caption         =   "Αναζήτηση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   9360
      MouseIcon       =   "main.frx":7F94
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Αναζήτηση"
      Top             =   360
      Width           =   5295
      Begin VB.Label lbl_sympiestis 
         BackStyle       =   0  'Transparent
         Caption         =   "Συμπιεστής"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5040
         MouseIcon       =   "main.frx":829E
         TabIndex        =   4
         Top             =   555
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Shape shape_sympiestis 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   735
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lbl_hz 
         BackStyle       =   0  'Transparent
         Caption         =   "Η/Ζ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         MouseIcon       =   "main.frx":85A8
         TabIndex        =   3
         Top             =   555
         Width           =   615
      End
      Begin VB.Shape shape_hz 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   735
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_kinitiras 
         BackStyle       =   0  'Transparent
         Caption         =   "Κινητήρας"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         MouseIcon       =   "main.frx":88B2
         TabIndex        =   2
         Top             =   555
         Width           =   1575
      End
      Begin VB.Shape shape_kinitiras 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   735
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_oxima 
         BackStyle       =   0  'Transparent
         Caption         =   "Όχημα"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         MouseIcon       =   "main.frx":8BBC
         TabIndex        =   1
         Top             =   555
         Width           =   975
      End
      Begin VB.Shape shape_oxima 
         BackColor       =   &H0000C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frm_edit 
      BackColor       =   &H00008000&
      Caption         =   "Σημειώσεις εισαγωγής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3060
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   13215
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H0000FFFF&
         Caption         =   "Εξαγωγή"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11880
         MouseIcon       =   "main.frx":8EC6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ακύρωση εισαγωγής"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11880
         MouseIcon       =   "main.frx":91D0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ιστορικό οχήματος"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11880
         MouseIcon       =   "main.frx":94DA
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame FRM_KIO 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "kio"
         ForeColor       =   &H000080FF&
         Height          =   2535
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   11175
         Begin VB.CommandButton cmd_replace_kin4 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin2 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   300
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   90
            Width           =   2295
         End
         Begin VB.TextBox txt_kin4 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   1920
            Width           =   6855
         End
         Begin VB.TextBox txt_kin3 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   1440
            Width           =   6855
         End
         Begin VB.TextBox txt_kin2 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   960
            Width           =   6855
         End
         Begin VB.TextBox txt_kin1 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   400
            Width           =   6855
         End
         Begin VB.TextBox txt_kio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2280
            Width           =   6855
         End
         Begin VB.Label Label_kio 
            BackStyle       =   0  'Transparent
            Caption         =   "Γενικές σημειώσης οχήματος"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   1680
            Width           =   3975
         End
         Begin VB.Label over_kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   58
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   3975
         End
         Begin VB.Label over_kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   55
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label over_kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   52
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   90
            Width           =   3975
         End
         Begin VB.Label over_kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   49
            Top             =   90
            Width           =   4215
         End
      End
      Begin VB.Frame frm_IO 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "io"
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   11295
         Begin VB.CommandButton cmd_replace_hz_io 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   960
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_sym_io 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   300
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   90
            Width           =   2295
         End
         Begin VB.TextBox txt_hz_io 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   1200
            Width           =   7335
         End
         Begin VB.TextBox txt_sym_io 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   400
            Width           =   7335
         End
         Begin VB.TextBox txt_io 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1920
            Width           =   7335
         End
         Begin VB.Label over_hz_io 
            BackStyle       =   0  'Transparent
            Caption         =   "H/Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   6600
            TabIndex        =   31
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label over_sym_io 
            BackStyle       =   0  'Transparent
            Caption         =   "Συμπιεστής"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   6600
            TabIndex        =   29
            Top             =   90
            Width           =   3375
         End
         Begin VB.Label hz_io 
            BackStyle       =   0  'Transparent
            Caption         =   "Η/Ζ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label sym_io 
            BackStyle       =   0  'Transparent
            Caption         =   "Συμπιεστής"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   90
            Width           =   3975
         End
         Begin VB.Label label_io 
            BackStyle       =   0  'Transparent
            Caption         =   "Γενικές σημειώσης οχήματος"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   3255
         End
      End
      Begin VB.Frame FRM_RO 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "ro"
         ForeColor       =   &H000080FF&
         Height          =   2535
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txt_ro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   2040
            Width           =   7335
         End
         Begin VB.TextBox txt_sym_ro 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   400
            Width           =   7335
         End
         Begin VB.TextBox txt_hz_ro 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1320
            Width           =   7335
         End
         Begin VB.CommandButton cmd_replace_sym_ro 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   90
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_hz_ro 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από την:"
            Height          =   315
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1050
            Width           =   2295
         End
         Begin VB.Label Label_ro 
            BackStyle       =   0  'Transparent
            Caption         =   "Γενικές σημειώσης οχήματος"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label sym_ro 
            BackStyle       =   0  'Transparent
            Caption         =   "Συμπιεστής"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   90
            Width           =   2415
         End
         Begin VB.Label hz_ro 
            BackStyle       =   0  'Transparent
            Caption         =   "Η/Ζ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label over_sym_ro 
            BackStyle       =   0  'Transparent
            Caption         =   "Συμπιεστής"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   6600
            TabIndex        =   39
            Top             =   90
            Width           =   4455
         End
         Begin VB.Label over_hz_ro 
            BackStyle       =   0  'Transparent
            Caption         =   "H/Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Left            =   6600
            TabIndex        =   38
            Top             =   1050
            Width           =   4215
         End
      End
      Begin VB.Frame FRM_KIO 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "kio"
         ForeColor       =   &H000080FF&
         Height          =   2535
         Index           =   1
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   11295
         Begin VB.TextBox txt_kio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   2280
            Width           =   6855
         End
         Begin VB.TextBox txt_kin1 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Top             =   400
            Width           =   6855
         End
         Begin VB.TextBox txt_kin2 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            Top             =   960
            Width           =   6855
         End
         Begin VB.TextBox txt_kin3 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   1440
            Width           =   6855
         End
         Begin VB.TextBox txt_kin4 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
            Top             =   1920
            Width           =   6855
         End
         Begin VB.CommandButton cmd_replace_kin1 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   300
            Index           =   1
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   90
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   1
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   690
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin3 
            BackColor       =   &H0080FF80&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   1
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   1170
            Width           =   2295
         End
         Begin VB.CommandButton cmd_replace_kin4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Θα αντικατασταθεί από τον:"
            Height          =   315
            Index           =   1
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1650
            Width           =   2295
         End
         Begin VB.Label over_kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   79
            Top             =   90
            Width           =   4095
         End
         Begin VB.Label kin1 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   78
            Top             =   90
            Width           =   3855
         End
         Begin VB.Label over_kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   77
            Top             =   690
            Width           =   4095
         End
         Begin VB.Label kin2 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label over_kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   75
            Top             =   1170
            Width           =   4215
         End
         Begin VB.Label kin3 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   3855
         End
         Begin VB.Label over_kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητηρας"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   73
            Top             =   1650
            Width           =   4215
         End
         Begin VB.Label kin4 
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητήρας 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Label Label_kio 
            BackStyle       =   0  'Transparent
            Caption         =   "Γενικές σημειώσης οχήματος"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   2040
            Width           =   3255
         End
      End
   End
   Begin VB.Label lbl_oxima_RO 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ΡΟ: 0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   6120
      MouseIcon       =   "main.frx":97E4
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl_oxima_IO 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ΙΟ: 0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3480
      MouseIcon       =   "main.frx":9AEE
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl_oxima_KIO 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "ΚΙΟ: 0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   960
      MouseIcon       =   "main.frx":9DF8
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl_status 
      BackColor       =   &H00008000&
      Caption         =   "0 εργασίες σε εξέλιξη εδώ και 0 μέρες."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label lbl_title 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Παρασκευή 00-00-2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   10005
   End
   Begin VB.Label lbl_work_progress 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Index           =   0
      Left            =   120
      MouseIcon       =   "main.frx":A102
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   15015
   End
   Begin VB.Label lbl_no_work 
      BackColor       =   &H00008000&
      Caption         =   "Καμία εργασία σε εξέλιξη."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Image img_logo 
      Height          =   1215
      Left            =   120
      MouseIcon       =   "main.frx":A40C
      MousePointer    =   99  'Custom
      Picture         =   "main.frx":A716
      ToolTipText     =   "Μετάβαση στην αρχική σελίδα"
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public protection As String

Private Sub cmd_edit_Click()
Dim code As String
If protection = "on" Then
    code = InputBox("Η ενέργεια αυτή προστατεύεται από κωδικό. Για να προχωρήστε πληκτρολογήστε τον κωδικό.", "Προστατευμένη ενέργεια!")
    If code <> ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then Exit Sub
End If
Dim sure As VbMsgBoxResult
sure = MsgBox("Όλες οι σημειώσεις της συγκεκριμένης εισαγώσής θα χαθούν. Είστε σίγουροι οτι θέλετε να γίνει ακύρωση της εισαγωγής;", vbYesNo, "Ακύρωση εισαγωγής;")
If sure = vbYes Then
    db_query "SELECT * FROM ergasies WHERE lbl_caption='" & ergasia_str & "'"
    rs.Delete
    rs_close
    edit_frame = "off"
    img_logo_Click
End If
End Sub

Private Sub cmd_exit_Click()
Dim code As String
Dim sure As VbMsgBoxResult

If protection = "on" Then
    code = InputBox("Η ενέργεια αυτή προστατεύεται από κωδικό. Για να προχωρήστε πληκτρολογήστε τον κωδικό.", "Προστατευμένη ενέργεια!")
    If code <> ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then Exit Sub
End If

sure = MsgBox("Πρόκειτε να γίνει εξαγωγή και αποθήκευση όλων των σημειώσεων στο ιστορικό. Θέλετε να συνεχίσετε;", vbYesNo, "Εξαγωγή;")

If sure <> vbYes Then Exit Sub
If edit_frame = "blabi" Then GoTo click_kio
If lbl_oxima_RO.Visible = False Then GoTo click_io

click_ro:
    lbl_oxima_RO_Click
    DoEvents
    add_new_istoriko
click_io:
    lbl_oxima_IO_Click
    DoEvents
    add_new_istoriko
click_kio:
    lbl_oxima_KIO_Click
    DoEvents
    add_new_istoriko
db_query "SELECT * FROM ergasies WHERE lbl_caption='" & ergasia_str & "'"
rs.Delete
rs_close
edit_frame = "off"
img_logo_Click
End Sub

Private Sub cmd_replace_hz_io_Click()
frm_antikatastasi.label_over = "hz_io"
antikatastasi_me = "hz"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_hz_ro_Click()
frm_antikatastasi.label_over = "hz_ro"
antikatastasi_me = "hz"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_kin1_Click(Index As Integer)
frm_antikatastasi.label_over = "kin1" & Index
antikatastasi_me = "kinitires"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_kin2_Click(Index As Integer)
frm_antikatastasi.label_over = "kin2" & Index
antikatastasi_me = "kinitires"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_kin3_Click(Index As Integer)
frm_antikatastasi.label_over = "kin3" & Index
antikatastasi_me = "kinitires"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_kin4_Click(Index As Integer)
frm_antikatastasi.label_over = "kin4" & Index
antikatastasi_me = "kinitires"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_sym_io_Click()
frm_antikatastasi.label_over = "sym_io"
antikatastasi_me = "sympiestes"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_replace_sym_ro_Click()
frm_antikatastasi.label_over = "sym_ro"
antikatastasi_me = "sympiestes"
frm_antikatastasi.Show
Me.Enabled = False
End Sub

Private Sub cmd_save_Click()
history_type = "oxima"
history_id = edit_oxima
refresh_history_contents
End Sub

Private Sub Command1_Click()
Dim code As String
If protection = "on" Then
    code = InputBox("Η ενέργεια αυτή προστατεύεται από κωδικό. Για να προχωρήστε πληκτρολογήστε τον κωδικό.", "Προστατευμένη ενέργεια!")
    If code <> ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then Exit Sub
End If
frm_nea_eisagogi.Show
End Sub

Private Sub Command2_Click()
Dim code As String
If protection = "on" Then
    code = InputBox("Η ενέργεια αυτή προστατεύεται από κωδικό. Για να προχωρήστε πληκτρολογήστε τον κωδικό.", "Προστατευμένη ενέργεια!")
    If code <> ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then Exit Sub
End If
frm_nea_kataxorisi.Show
End Sub

Private Sub Command4_Click()
Dim code As String
If protection = "on" Then
    code = InputBox("Η ενέργεια αυτή προστατεύεται από κωδικό. Για να προχωρήστε πληκτρολογήστε τον κωδικό.", "Προστατευμένη ενέργεια!")
    If code <> ReadINI("HSAP_1_0", "password", App.Path & "\Settings.ini") Then Exit Sub
End If
frm_settings.Show
End Sub

Private Sub form_load()
If ReadINI("HSAP_1_0", "protection", App.Path & "\Settings.ini") = "on" Then
    protection = "on"
Else
    protection = "off"
End If

If ReadINI("HSAP_1_0", "core_path", App.Path & "\Settings.ini") = "app.path" Then
    core_path = App.Path & "\data1.mdb"
Else
    core_path = ReadINI("HSAP_1_0", "core_path", App.Path & "\Settings.ini")
End If

If ReadINI("HSAP_1_0", "status", App.Path & "\Settings.ini") = "on" Then
    Timer4.Enabled = True
    lbl_edit.Visible = True
    lbl_history.Visible = True
End If

Load lbl_work_progress(1)
lbl_work_progress(1).Top = lbl_work_progress(0).Top + 450

db_connect
search_frame = "off"
toolbar_frame = "off"
history_type = "off"
edit_frame = "off"
backup_path = ReadINI("HSAP_1_0", "backup_path", App.Path & "\Settings.ini")
Timer3.Interval = ReadINI("HSAP_1_0", "tools_timer", App.Path & "\Settings.ini")
refresh_main_contents
End Sub

Private Sub Form_Initialize()
If App.PrevInstance Then
        MsgBox "Το πρόγραμμα είναι ήδη ανοιχτό.", vbCritical + vbOKOnly, "Η.Σ.Α.Π."
        End
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frm_unloading.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 9180 Then
    Me.Width = 9180
End If
If Me.Height < 7000 Then
    Me.Height = 7000
End If
reposition_main_contents
If history_frame <> "off" Then
    reposition_history_contents
End If
End Sub

Private Sub frm_edit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons ""
End Sub

Private Sub frm_search_Click()
frm_search.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub frm_search_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons ""
End Sub


Private Sub history_label_Click(Index As Integer)
MsgBox history_label(Index).ToolTipText, , history_label(Index).Caption
End Sub

Public Sub img_logo_Click()
If edit_frame <> "off" Then
    update_eisagogis
End If
search_frame = "off"
edit_frame = "off"
history_type = "off"
refresh_main_contents
End Sub

Private Sub Label1_Click()
If Not toolbar_frame = "on" Then
    Timer2.Enabled = True
    frm_toolbar.Enabled = False
End If
End Sub

Private Sub lbl_hz_Click()
history_type = "hz"
history_id = InputBox("Πληκτρολογήστε τον αριθμό του Η/Ζ.", "Αναζήτηση Η/Ζ")
If Len(history_id) < 3 Then Exit Sub
refresh_history_contents
End Sub

Private Sub lbl_hz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "hz"
End Sub

Private Sub lbl_kinitiras_Click()
history_type = "kinitiras"
history_id = InputBox("Πληκτρολογήστε τον αριθμό του Κινητήρα.", "Αναζήτηση Κινητήρα")
If Len(history_id) < 3 Then Exit Sub
refresh_history_contents
End Sub

Private Sub lbl_kinitiras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "kinitiras"
End Sub

Private Sub lbl_oxima_Click()
history_type = "oxima"
history_id = InputBox("Πληκτρολογήστε τον αριθμό του Οχήματος.", "Αναζήτηση Οχήματος")
If Len(history_id) < 3 Then Exit Sub
refresh_history_contents
End Sub

Public Sub lbl_oxima_IO_Click()
If edit_frame <> "blabi" Then
    lbl_oxima_IO.FontUnderline = True
    lbl_oxima_KIO.FontUnderline = False
    lbl_oxima_RO.FontUnderline = False
End If
    update_eisagogis
    DoEvents
Dim a As Integer
For a = 1 To Len(lbl_oxima_IO)
    If Mid(lbl_oxima_IO, a, 1) = ":" Then
        Exit For
    End If
Next a
If edit_frame = "blabi" Then
    history_type = "oxima"
    history_id = Mid(lbl_oxima_IO, a + 2, Len(lbl_oxima_IO))
    refresh_history_contents
Else
    edit_oxima = Mid(lbl_oxima_IO, a + 2, Len(lbl_oxima_IO))
    frm_edit.Visible = True
    frm_history.Visible = False
    refresh_edit_contents "2"
End If
End Sub

Private Sub lbl_oxima_IO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "io"
End Sub

Public Sub lbl_oxima_KIO_Click()
If edit_frame <> "blabi" Then
    lbl_oxima_KIO.FontUnderline = True
    lbl_oxima_IO.FontUnderline = False
    lbl_oxima_RO.FontUnderline = False
End If
    update_eisagogis
    DoEvents
Dim a As Integer
For a = 1 To Len(lbl_oxima_KIO)
    If Mid(lbl_oxima_KIO, a, 1) = ":" Then
        Exit For
    End If
Next a
edit_oxima = Mid(lbl_oxima_KIO, a + 2, Len(lbl_oxima_KIO))
frm_edit.Visible = True
frm_history.Visible = False
refresh_edit_contents "2"
End Sub

Private Sub lbl_oxima_KIO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "kio"
End Sub

Private Sub lbl_oxima_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "oxima"
End Sub

Public Sub lbl_oxima_RO_Click()
If edit_frame <> "blabi" Then
    lbl_oxima_RO.FontUnderline = True
    lbl_oxima_KIO.FontUnderline = False
    lbl_oxima_IO.FontUnderline = False
End If
    update_eisagogis
    DoEvents
Dim a As Integer
For a = 1 To Len(lbl_oxima_RO)
    If Mid(lbl_oxima_RO, a, 1) = ":" Then
        Exit For
    End If
Next a
If edit_frame = "blabi" Then
    history_type = "oxima"
    history_id = Mid(lbl_oxima_RO, a + 2, Len(lbl_oxima_RO))
    refresh_history_contents
Else
    edit_oxima = Mid(lbl_oxima_RO, a + 2, Len(lbl_oxima_RO))
    frm_edit.Visible = True
    frm_history.Visible = False
    refresh_edit_contents "2"

End If
End Sub

Private Sub lbl_oxima_RO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "ro"
End Sub

Private Sub lbl_sympiestis_Click()
history_type = "sympiestis"
history_id = InputBox("Πληκτρολογήστε τον αριθμό του Συμπιεστή.", "Αναζήτηση Συμπιεστή")
If Len(history_id) < 3 Then Exit Sub
refresh_history_contents
End Sub

Private Sub lbl_sympiestis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
refresh_buttons "sympiestis"
End Sub

Private Sub lbl_work_progress_Click(Index As Integer)
Dim q As Integer
    lbl_oxima_KIO.FontUnderline = True
For q = 1 To Len(lbl_work_progress(Index).Caption)
    If Mid(lbl_work_progress(Index).Caption, q, 3) = "..." Then
        Exit For
    End If
Next q
db_query "SELECT * FROM ergasies WHERE lbl_caption='" & Mid(lbl_work_progress(Index).Caption, 1, q + 2) & "'"
edit_frame = rs!ergasia
edit_oxima = rs!oxima
ergasia_str = rs!lbl_caption
ergasia_date = rs!Date
rs_close
refresh_main_contents
End Sub

Private Sub list_xil1_click()
On Error Resume Next
If Len(list_xil1.SelectedItem.SubItems(2) & "") < 3 Then
    t_sim_label = "(" & list_xil1.SelectedItem.SubItems(1) & "):"
Else
    t_sim_label = "(" & list_xil1.SelectedItem.SubItems(1) & "): " & list_xil1.SelectedItem.SubItems(2) & ", στη θέση του μπήκε ο " & list_xil1.SelectedItem.SubItems(3)
End If
t_txt_sim = list_xil1.SelectedItem.SubItems(4)
End Sub

Private Sub list_xil4_click()
On Error Resume Next
If Len(list_xil4.SelectedItem.SubItems(2) & "") < 3 Then
    t_sim_label = "(" & list_xil4.SelectedItem.SubItems(1) & "):"
Else
    t_sim_label = "(" & list_xil4.SelectedItem.SubItems(1) & "): " & list_xil4.SelectedItem.SubItems(2) & ", στη θέση του μπήκε ο " & list_xil4.SelectedItem.SubItems(3)
End If
t_txt_sim = list_xil4.SelectedItem.SubItems(4)
End Sub

Private Sub list_blabi_click()
On Error Resume Next
If Len(list_blabi.SelectedItem.SubItems(2) & "") < 3 Then
    t_sim_label = "(" & list_blabi.SelectedItem.SubItems(1) & "):"
Else
    t_sim_label = "(" & list_blabi.SelectedItem.SubItems(1) & "): " & list_blabi.SelectedItem.SubItems(2) & ", στη θέση του μπήκε ο " & list_blabi.SelectedItem.SubItems(3)
End If
t_txt_sim = list_blabi.SelectedItem.SubItems(4)
End Sub

Private Sub t_hz_Click()
history_type = "hz"
history_id = hz.Caption
refresh_history_contents
End Sub

Private Sub t_kin1_Click()
history_type = "kinitiras"
history_id = t_kin1.Caption
refresh_history_contents
End Sub

Private Sub t_kin2_Click()
history_type = "kinitiras"
history_id = t_kin2.Caption
refresh_history_contents
End Sub

Private Sub t_kin3_Click()
history_type = "kinitiras"
history_id = t_kin3.Caption
refresh_history_contents
End Sub

Private Sub t_kin4_Click()
history_type = "kinitiras"
history_id = t_kin4.Caption
refresh_history_contents
End Sub

Private Sub t_sym_Click()
history_type = "sympiestis"
history_id = t_sym.Caption
refresh_history_contents
End Sub

Private Sub tabstrip_istoriko_Click()
If tabstrip_istoriko.SelectedItem.Index = 1 Then
    If history_type = "oxima" Then
        If typos_oximatos(history_id) = "ΚΙΟ" Then
            tab_stoixeia_kio.Visible = True
            tab_stoixeia_ioro.Visible = False
            tab_stoixeia_organo.Visible = False
        Else
            tab_stoixeia_kio.Visible = False
            tab_stoixeia_ioro.Visible = True
            tab_stoixeia_organo.Visible = False
        End If
    Else
        tab_stoixeia_kio.Visible = False
        tab_stoixeia_ioro.Visible = False
        tab_stoixeia_organo.Visible = True
    End If

    tab_xil1.Visible = False
    tab_xil4.Visible = False
    tab_blabi.Visible = False
    tab_sim.Visible = False
ElseIf tabstrip_istoriko.SelectedItem.Index = 2 Then
    tab_stoixeia_kio.Visible = False
    tab_stoixeia_ioro.Visible = False
    tab_stoixeia_organo.Visible = False
    tab_xil1.Visible = True
    tab_xil4.Visible = False
    tab_blabi.Visible = False
    tab_sim.Visible = False
ElseIf tabstrip_istoriko.SelectedItem.Index = 3 Then
    tab_stoixeia_kio.Visible = False
    tab_stoixeia_ioro.Visible = False
    tab_stoixeia_organo.Visible = False
    tab_xil1.Visible = False
    tab_xil4.Visible = True
    tab_blabi.Visible = False
    tab_sim.Visible = False
ElseIf tabstrip_istoriko.SelectedItem.Index = 4 Then
    tab_stoixeia_kio.Visible = False
    tab_stoixeia_ioro.Visible = False
    tab_stoixeia_organo.Visible = False
    tab_xil1.Visible = False
    tab_xil4.Visible = False
    tab_blabi.Visible = True
    tab_sim.Visible = False
ElseIf tabstrip_istoriko.SelectedItem.Index = 5 Then
    tab_stoixeia_kio.Visible = False
    tab_stoixeia_ioro.Visible = False
    tab_stoixeia_organo.Visible = False
    tab_xil1.Visible = False
    tab_xil4.Visible = False
    tab_blabi.Visible = False
    tab_sim.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
If search_frame = "on" Then
        If frm_search.Width < 800 Then
            frm_search.Left = Me.Width - 600
            frm_search.Width = 200
            search_frame = "off"
            frm_search.Enabled = True
            Timer1.Enabled = False
            Exit Sub
        End If
        frm_search.Left = frm_search.Left + 800
        frm_search.Width = frm_search.Width - 800
        reposition_search_buttons
ElseIf search_frame = "off" Then
        If frm_search.Left < 2960 Then
            frm_search.Left = 2160
            frm_search.Width = (Me.Width - 2160) - 400
            search_frame = "on"
            frm_search.Enabled = True
            Timer1.Enabled = False
            Exit Sub
        End If
        frm_search.Left = frm_search.Left - 800
        frm_search.Width = frm_search.Width + 800
        reposition_search_buttons
End If
End Sub

Private Sub Timer2_Timer()
If toolbar_frame = "off" Then
If frm_toolbar.Height > 1100 Then
    frm_toolbar.Height = 1200
    frm_toolbar.Top = Me.Height - 1510
    Timer2.Enabled = False
    frm_toolbar.Enabled = True
    toolbar_frame = "on"
    Timer3.Enabled = True
    Exit Sub
End If
frm_toolbar.Height = frm_toolbar.Height + 100
frm_toolbar.Top = frm_toolbar.Top - 100
ElseIf toolbar_frame = "on" Then
Timer3.Enabled = False
If frm_toolbar.Height < 635 Then
    frm_toolbar.Height = 535
    frm_toolbar.Top = Me.Height - 895
    Timer2.Enabled = False
    frm_toolbar.Enabled = True
    toolbar_frame = "off"
    Exit Sub
End If
frm_toolbar.Enabled = False
frm_toolbar.Height = frm_toolbar.Height - 100
frm_toolbar.Top = frm_toolbar.Top + 100
End If
Shape2.Height = frm_toolbar.Height
Shape2.Width = frm_toolbar.Width - 120
Line1.X2 = frm_toolbar.Width
End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
lbl_edit = "Edit: " & edit_frame & " " & "(" & edit_oxima & ")"
lbl_history = "History: " & history_type & " " & "(" & history_id & ")"
End Sub

Private Sub add_new_istoriko()
    If typos_oximatos(edit_oxima) = "ΚΙΟ" Then
            'οχημα
        db_query "SELECT * FROM istoriko WHERE 1=0"
        rs.AddNew
        rs!date_in = ergasia_date
        rs!history_type = "oxima"
        rs!history_id = edit_oxima
        rs!ergasia = edit_frame
        rs!simeioseis = txt_kio(kio_i)
        rs!date_out = Date
        rs.Update
        rs_close
            'κινητηρας1
        If over_kin1(kio_i) = "" And txt_kin1(kio_i) <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin1(kio_i), Len(kin1(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_kin1(kio_i)
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_kin1(kio_i) <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin1(kio_i), Len(kin1(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_kin1(kio_i)
            rs!simeioseis = txt_kin1(kio_i)
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του κινητηρα στα στοιχεια
            db_query "SELECT kinitiras1 FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!kinitiras1 = over_kin1(kio_i)
            rs.Update
            rs_close
                'εγγραφή κινητήρα στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_kinitires WHERE id_kinitires='" & over_kin1(kio_i) & "'"
            rs!id_kinitires = Right(kin1(kio_i), Len(kin1(kio_i)) - 11)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
            'κινητήρας2
        If over_kin2(kio_i) = "" And txt_kin2(kio_i) <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin2(kio_i), Len(kin2(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_kin2(kio_i)
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_kin2(kio_i) <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin2(kio_i), Len(kin2(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_kin2(kio_i)
            rs!simeioseis = txt_kin2(kio_i)
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του κινητηρα στα στοιχεια
            db_query "SELECT kinitiras2 FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!kinitiras2 = over_kin2(kio_i)
            rs.Update
            rs_close
                'εγγραφή κινητήρα στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_kinitires WHERE id_kinitires='" & over_kin2(kio_i) & "'"
            rs!id_kinitires = Right(kin2(kio_i), Len(kin2(kio_i)) - 11)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
            'κινητηρας3
        If over_kin3(kio_i) = "" And txt_kin3(kio_i) <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin3(kio_i), Len(kin3(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_kin3(kio_i)
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_kin3(kio_i) <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin3(kio_i), Len(kin3(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_kin3(kio_i)
            rs!simeioseis = txt_kin3(kio_i)
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του κινητηρα στα στοιχεια
            db_query "SELECT kinitiras3 FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!kinitiras3 = over_kin3(kio_i)
            rs.Update
            rs_close
                'εγγραφή κινητήρα στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_kinitires WHERE id_kinitires='" & over_kin3(kio_i) & "'"
            rs!id_kinitires = Right(kin3(kio_i), Len(kin3(kio_i)) - 11)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
            'κινητηρας4
        If over_kin4(kio_i) = "" And txt_kin4(kio_i) <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin4(kio_i), Len(kin4(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_kin4(kio_i)
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_kin4(kio_i) <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "kinitiras"
            rs!history_id = Right(kin4(kio_i), Len(kin4(kio_i)) - 11)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_kin4(kio_i)
            rs!simeioseis = txt_kin4(kio_i)
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του κινητηρα στα στοιχεια
            db_query "SELECT kinitiras4 FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!kinitiras4 = over_kin4(kio_i)
            rs.Update
            rs_close
                'εγγραφή κινητήρα στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_kinitires WHERE id_kinitires='" & over_kin4(kio_i) & "'"
            rs!id_kinitires = Right(kin4(kio_i), Len(kin4(kio_i)) - 11)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If

    ElseIf typos_oximatos(edit_oxima) = "ΙΟ" Then
            'οχημα
        db_query "SELECT * FROM istoriko WHERE 1=0"
        rs.AddNew
        rs!date_in = ergasia_date
        rs!history_type = "oxima"
        rs!history_id = edit_oxima
        rs!ergasia = edit_frame
        rs!simeioseis = txt_io
        rs!date_out = Date
        rs.Update
        rs_close
            'συμπιεστης
        If over_sym_io = "" And txt_sym_io <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "sympiestis"
            rs!history_id = Right(sym_io, Len(sym_io) - 12)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_sym_io
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_sym_io <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "sympiestis"
            rs!history_id = Right(sym_io, Len(sym_io) - 12)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_sym_io
            rs!simeioseis = txt_sym_io
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του συμπιεστη στα στοιχεια
            db_query "SELECT sympiestis FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!sympiestis = over_sym_io
            rs.Update
            rs_close
                'εγγραφή συμπιεστη στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_sympiestes WHERE id_sympiestes='" & over_sym_io & "'"
            rs!id_sympiestes = Right(sym_io, Len(sym_io) - 12)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
            'Η/Ζ
        If over_hz_io = "" And txt_hz_io <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "hz"
            rs!history_id = Right(hz_io, Len(hz_io) - 5)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_hz_io
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_hz_io <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "hz"
            rs!history_id = Right(hz_io, Len(hz_io) - 5)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_hz_io
            rs!simeioseis = txt_hz_io
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του Η/Ζ στα στοιχεια
            db_query "SELECT hz FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!hz = over_hz_io
            rs.Update
            rs_close
                'εγγραφή Η/Ζ στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_hz WHERE id_hz='" & over_hz_io & "'"
            rs!id_hz = Right(hz_io, Len(hz_io) - 5)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If

    ElseIf typos_oximatos(edit_oxima) = "ΡΟ" Then
            'οχημα
        db_query "SELECT * FROM istoriko WHERE 1=0"
        rs.AddNew
        rs!date_in = ergasia_date
        rs!history_type = "oxima"
        rs!history_id = edit_oxima
        rs!ergasia = edit_frame
        rs!simeioseis = txt_ro
        rs!date_out = Date
        rs.Update
        rs_close
            'συμπιεστης
        If over_sym_ro = "" And txt_sym_ro <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "sympiestis"
            rs!history_id = Right(sym_ro, Len(sym_ro) - 12)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_sym_ro
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_sym_ro <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "sympiestis"
            rs!history_id = Right(sym_ro, Len(sym_ro) - 12)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_sym_ro
            rs!simeioseis = txt_sym_ro
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του συμπιεστη στα στοιχεια
            db_query "SELECT sympiestis FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!sympiestis = over_sym_ro
            rs.Update
            rs_close
                'εγγραφή συμπιεστη στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_sympiestes WHERE id_sympiestes='" & over_sym_ro & "'"
            rs!id_sympiestes = Right(sym_ro, Len(sym_ro) - 12)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
            'Η/Ζ
        If over_hz_ro = "" And txt_hz_ro <> "" Then
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "hz"
            rs!history_id = Right(hz_ro, Len(hz_ro) - 5)
            rs!ergasia = edit_frame
            rs!simeioseis = txt_hz_ro
            rs!date_out = Date
            rs.Update
            rs_close
        ElseIf over_hz_ro <> "" Then
                'εγγραφή ιστορικού
            db_query "SELECT * FROM istoriko WHERE 1=0"
            rs.AddNew
            rs!date_in = ergasia_date
            rs!history_type = "hz"
            rs!history_id = Right(hz_ro, Len(hz_ro) - 5)
            rs!ergasia = edit_frame
            rs!antikatastasi_me = over_hz_ro
            rs!simeioseis = txt_hz_ro
            rs.Fields("8esi_prin") = edit_oxima
            rs.Fields("8esi_meta") = "d"
            rs!date_out = Date
            rs.Update
            rs_close
                'αντικατασταση του Η/Ζ στα στοιχεια
            db_query "SELECT hz FROM stoixeia WHERE oxima='" & edit_oxima & "'"
            rs!hz = over_hz_ro
            rs.Update
            rs_close
                'εγγραφή Η/Ζ στους διαθέσιμους
            db_query "SELECT * FROM dia8esimoi_hz WHERE id_hz='" & over_hz_ro & "'"
            rs!id_hz = Right(hz_ro, Len(hz_ro) - 5)
            rs!paralabi = paralabi_oximatos(edit_oxima)
            rs.Update
            rs_close
        End If
    End If
End Sub
