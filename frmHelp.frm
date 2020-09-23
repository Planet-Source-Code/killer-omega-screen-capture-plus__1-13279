VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Screen Capture Plus Help"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtinfo 
      Height          =   4575
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHelp.frx":0742
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Using Help"
List1.AddItem "About This Program"
List1.AddItem "Program Functions"
List1.AddItem "     - Screenshot"
List1.AddItem "     - Copy"
List1.AddItem "     - Save"
List1.AddItem "     - Print"
End Sub

Private Sub List1_Click()
If List1.Text = List1.List(0) Then
    txtinfo.Text = "Using this help file is easy.  All you have to do is click on the name of the help topic you want help with."
    End If
If List1.Text = List1.List(1) Then
    txtinfo.Text = "This program is a combined effort between Adam Orenstein & Jason Dorfman.  We have created what we hope to be a useful tool for you.  Please enjoy this program!"
    End If
If List1.Text = List1.List(2) Then
    txtinfo.Text = "This program has several functions for what you can actually do with your screenshots.  Click a funtion for information on using it."
    End If
If List1.Text = List1.List(3) Then
    txtinfo.Text = "To take a Screenshot press the pause key."
    End If
If List1.Text = List1.List(4) Then
    txtinfo.Text = "This will copy the screenshot from the selected window to the Windows clipboard."
    End If
If List1.Text = List1.List(5) Then
    txtinfo.Text = "This will save the image to a file on your computer.  Type in the file path then select the file type and click save."
    End If
If List1.Text = List1.List(6) Then
    txtinfo.Text = "This function prints the image of the selected screenshot window."
    End If
End Sub
