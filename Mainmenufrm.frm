VERSION 5.00
Begin VB.Form mainmenufrm 
   BackColor       =   &H00EDEDD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu of Chatting Software"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mainmenufrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnulogin 
      Caption         =   "&Login"
      Begin VB.Menu mnunewuser 
         Caption         =   "&New User"
      End
      Begin VB.Menu mnustart 
         Caption         =   "&Start Chatting"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "mainmenufrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnunewuser_Click()
chatmasterfrm.Show
End Sub

Private Sub mnustart_Click()
chattempfrm.Show
End Sub
