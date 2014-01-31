VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_unloading 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Αποθήκευση δεδομένων και δημιουργία αντίγραφων ασφαλείας..."
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frm_unloading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6120
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   794
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   794
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Δημιουργία αντίγραφων ασφαλείας:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αποθήκευση δεδομένων:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frm_unloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
     (ByVal hMenu As Long) _
     As Long
Private Declare Function RemoveMenu Lib "user32" _
     (ByVal hMenu As Long, ByVal _
     nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" _
     (ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&


Private Sub DisableX()
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(Me.hwnd, 0)
     nCount = GetMenuItemCount(hMenu)

     'Get rid of the Close menu and its separator
     Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)

     'Make sure the screen updates
     'our change
     DrawMenuBar Me.hwnd
End Sub


Private Sub Form_Load()
DisableX
ProgressBar1.Value = 20
frm_main.img_logo_Click
ProgressBar1.Value = 60
db_close
ProgressBar1.Value = 100
End Sub

Private Sub Timer1_Timer()
CopyFile App.Path & "\data1.mdb", backup_path
Timer1.Enabled = False
End Sub

