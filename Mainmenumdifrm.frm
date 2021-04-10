VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Main Menu of Chat Software"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "Mainmenumdifrm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Mainmenumdifrm.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnulogin 
      Caption         =   "&Login"
      Begin VB.Menu mnuuser 
         Caption         =   "&New User"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Chatting"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MsgBox "Please Read the readme.txt and then after making appropriate changes run the application. And after making the changes create users and start chatting..."

End Sub

Private Sub mnuabout_Click()
aboutfrm.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuStart_Click()
chattempfrm.SetFocus
chattempfrm.WindowState = 0
chattempfrm.Move 3500, 1750
chattempfrm.Show
End Sub

Private Sub mnuuser_Click()
chatmasterfrm.SetFocus
chatmasterfrm.WindowState = 0
chatmasterfrm.Height = 2775
chatmasterfrm.Width = 4860
chatmasterfrm.Move 3500, 1750
chatmasterfrm.Show
End Sub
