VERSION 5.00
Begin VB.Form frmSSC 
   AutoRedraw      =   -1  'True
   Caption         =   "Screenshot"
   ClientHeight    =   3195
   ClientLeft      =   3030
   ClientTop       =   3090
   ClientWidth     =   4680
   Icon            =   "ScreenShot.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   1320
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
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
Attribute VB_Name = "frmSSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnucopy_Click()
Clipboard.Clear
Call Clipboard.SetData(Me.Image, 0)
End Sub

Private Sub mnuhowto_Click()
Load frmHelp
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo printerror
    Call Printer.PaintPicture(Me.Image, 0&, 0&)

Exit Sub

printerror:
Call MsgBox("Error while printing file.", vbCritical, "Error")
End Sub

Private Sub mnuquit_Click()
    End
End Sub

Private Sub mnuSave_Click()
Load frmsave
Call SavePicture(Me.Image, App.Path & "\" & "tempimg.bmp")
frmsave.Show
End Sub

