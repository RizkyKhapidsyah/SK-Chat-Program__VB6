VERSION 5.00
Begin VB.Form chattempfrm 
   BackColor       =   &H00EDEDD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start Chatting......."
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChatTempFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   405
      Left            =   2550
      TabIndex        =   5
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Chatting"
      Default         =   -1  'True
      Height          =   405
      Left            =   690
      TabIndex        =   4
      Top             =   1920
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1290
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   630
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1350
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Chat Name"
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   690
      Width           =   1545
   End
End
Attribute VB_Name = "chattempfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
On Error GoTo y
rs.Open "select * from chatmaster where chat_name='" & Text1.Text & "' and pwd='" & Text2.Text & "'", adoconn, adOpenStatic
If rs.RecordCount <> 1 Then
    MsgBox "You are not a Registered User:", vbExclamation
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
Else
    rs.Close
    glousername = Text1.Text
    rs.Open "insert into online values('" & Text1.Text & "')", adoconn, adOpenStatic
    Call conclose
    ChatMediumfrm.Refresh
    ChatMediumfrm.SetFocus
    ChatMediumfrm.Show
End If
y:
If Err.Number = 3705 Then
    rs.Close
    Exit Sub
End If

    


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call condb
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo y


y:
If Err.Number = 3704 Or Err.Number = 3705 Then
rs.Close
Call conclose
End If
End Sub
