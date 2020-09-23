VERSION 5.00
Begin VB.MDIForm MDIFRM 
   BackColor       =   &H8000000C&
   Caption         =   "Screen Capture Plus"
   ClientHeight    =   5760
   ClientLeft      =   2250
   ClientTop       =   2190
   ClientWidth     =   7770
   Icon            =   "MDIFRM.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   1440
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuhowto 
         Caption         =   "How to..."
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "MDIFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Double

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuabout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuhowto_Click()
Load frmHelp
frmHelp.Show
End Sub

Private Sub mnuquit_Click()
    End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Static y As Single
y = y + 1
    If GetAsyncKeyState(vbKeyPause) And y >= PAUSE_TIME Then
        Static count As Long
        Dim newss As New frmSSC
        count = count + 1
        newss.Caption = "Screenshot " & count
        newss.Show
        CaptureScreen newss
        y = 0
    End If
If frmsave.Visible = True Then Call BringWindowToTop(frmsave.hwnd)
End Sub
