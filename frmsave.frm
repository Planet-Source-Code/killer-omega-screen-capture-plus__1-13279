VERSION 5.00
Begin VB.Form frmsave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save As..."
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmsave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Type..."
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      Begin VB.OptionButton optJPEG 
         Caption         =   "JPEG"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optbitmap 
         Caption         =   "Bitmap"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "C:\WINDOWS\Desktop\ScreenShot"
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "File save path:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmsave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filetypetosave As String

Private Sub Command1_Click()

Dim SavePath As String

SavePath = Text1.Text
Select Case filetypetosave
    Case ".bmp"
        If InStr(1, SavePath, ".bmp", vbBinaryCompare) Then
            Else
            SavePath = SavePath & ".bmp"
        End If
        Call FileCopy(App.Path & "\" & "tempimg.bmp", SavePath)
        Kill (App.Path & "\" & "tempimg.bmp")
        Unload Me
        Exit Sub
        
    Case ".jpeg"
        If InStr(1, SavePath, ".jpeg", vbBinaryCompare) Then
            Else
            SavePath = SavePath & ".jpeg"
        End If
    
        Call BmpToJpeg(App.Path & "\" & "tempimg.bmp", SavePath, 100)
        Kill (App.Path & "\" & "tempimg.bmp")
        Unload Me
        Exit Sub
    End Select
Unload Me
End Sub

Private Sub Form_Load()
optbitmap.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill (App.Path & "\" & "tempimg.bmp")
End Sub

Private Sub optbitmap_Click()
    If InStr(1, SavePath, ".bmp", vbBinaryCompare) Then
        Else
        SavePath = SavePath & ".bmp"
    End If
End Sub

Private Sub optJPEG_Click()
    If InStr(1, SavePath, ".jpeg", vbBinaryCompare) Then
        Else
        SavePath = SavePath & ".jpeg"
    End If
End Sub

Private Sub Timer1_Timer()
If Me.Visible Then Call BringWindowToTop(frmsave.hwnd)
If optbitmap.Value = True Then
    filetypetosave = ".bmp"
    Else
    filetypetosave = ".jpeg"
    End If
End Sub
