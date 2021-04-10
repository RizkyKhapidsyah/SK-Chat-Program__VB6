VERSION 5.00
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.tlb"
Begin VB.Form aboutfrm 
   BackColor       =   &H00EDEDD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Server Chat Software"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "aboutfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   1185
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "aboutfrm.frx":0442
      Top             =   1110
      Width           =   4905
   End
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   2475
      Scrollbar       =   0   'False
      URL             =   "mailto:n_indureddy@yahoo.com"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   465
      Left            =   1980
      TabIndex        =   0
      Top             =   2550
      Width           =   1425
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

