VERSION 5.00
Begin VB.Form frmhotkey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Hotkey..."
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   Icon            =   "frmhotkey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "F10"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "F12"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Scroll Lock"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmhotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
HOTKEY = Space(255)
Call GetPrivateProfileString("Hotkey", "Hotkey", "Error", HOTKEY, 255, "c:\WINDOWS\SYSTEM\SCPHK.ini")
If HOTKEY = "Error" Then Call WritePrivateProfileString("Hotkey", "Hotkey", "vbKeyPause", "c:\WINDOWS\SYSTEM\SCPHK.ini")
If HOTKEY = "vbKeyPause" Then Option1.Value = True
If HOTKEY = "vbKeyF10" Then Option2.Value = True
If HOTKEY = "vbKeyScrollLock" Then Option3.Value = True
If HOTKEY = "vbKeyF12" Then Option4.Value = True
End Sub

Private Sub Option1_Click()
Call WritePrivateProfileString("Hotkey", "Hotkey", "vbKeyPause", "c:\WINDOWS\SYSTEM\SCPHK.ini")
End Sub

Private Sub Option2_Click()
Call WritePrivateProfileString("Hotkey", "Hotkey", "vbKeyF10", "c:\WINDOWS\SYSTEM\SCPHK.ini")
End Sub

Private Sub Option3_Click()
Call WritePrivateProfileString("Hotkey", "Hotkey", "vbKeyScrollLock", "c:\WINDOWS\SYSTEM\SCPHK.ini")
End Sub

Private Sub Option4_Click()
Call WritePrivateProfileString("Hotkey", "Hotkey", "vbKeyF12", "c:\WINDOWS\SYSTEM\SCPHK.ini")
End Sub

