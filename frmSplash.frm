VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3810
      Left            =   280
      Picture         =   "frmSplash.frx":030A
      ScaleHeight     =   3750
      ScaleWidth      =   6750
      TabIndex        =   0
      Top             =   210
      Width           =   6810
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Load MDIFRM
    MDIFRM.Show
End Sub

Private Sub Picture1_Click()
    Unload Me
    Load MDIFRM
    MDIFRM.Show
End Sub
