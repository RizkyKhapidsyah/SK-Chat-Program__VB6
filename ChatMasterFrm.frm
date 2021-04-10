VERSION 5.00
Begin VB.Form chatmasterfrm 
   BackColor       =   &H00EDEDD8&
   Caption         =   "Chat Registration Form"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   Icon            =   "ChatMasterFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2430
      TabIndex        =   3
      Top             =   1710
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   870
      TabIndex        =   2
      Top             =   1710
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2130
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2130
      TabIndex        =   0
      Top             =   510
      Width           =   1845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Chat Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   540
      Width           =   1545
   End
End
Attribute VB_Name = "chatmasterfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
On Error GoTo y

rs.Open "insert into chatmaster values('" & Text1 & "','" & Text2 & "')", adoconn, adOpenStatic
MsgBox "Registered Successfully"
Text1.Text = ""
Text2.Text = ""
Exit Sub
y:
    MsgBox "Select Another Chat Name, this chat name already exists"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call condb

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call conclose
End Sub
