VERSION 5.00
Begin VB.Form ChatMediumfrm 
   BackColor       =   &H00EDEDD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Online........."
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChatMediumfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   1815
      Left            =   4590
      TabIndex        =   10
      Top             =   4050
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   90
      Width           =   2235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Logout"
      Height          =   345
      Left            =   2070
      TabIndex        =   7
      Top             =   4260
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   4260
      Width           =   1395
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1845
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   990
      Width           =   6075
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   5790
      Top             =   3210
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   150
      TabIndex        =   2
      Top             =   3540
      Width           =   4785
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "ChatMediumfrm.frx":030A
      Left            =   6840
      List            =   "ChatMediumfrm.frx":030C
      TabIndex        =   0
      Top             =   960
      Width           =   3105
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "My Chat Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OnLine Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7620
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages Received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message to Sent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   3030
      Width           =   2445
   End
End
Attribute VB_Name = "ChatMediumfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim lstr As String
Private Sub Command1_Click()
On Error GoTo y:
If Len(List1.Text) > 0 Then
Command1.Enabled = False
rs3.Open "insert into chattemp values('" & glousername & "','" & lstr & "','" & Text1.Text & "')", adoconn, adOpenStatic
Command1.Enabled = True
Text1.Text = ""
Text1.SetFocus
Else
Command1.Enabled = False
rs3.Open "insert into chattemp values('" & glousername & "','default','" & Text1.Text & "')", adoconn, adOpenStatic
Command1.Enabled = True
Text1.Text = ""
Text1.SetFocus
End If
y:
If Err.Number = 3705 Then
    rs3.Close
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call condb
rs.Open "select * from online", adoconn, adOpenStatic
List1.Clear

Do Until rs.EOF = True
    List1.AddItem rs(0)
    rs.MoveNext
Loop
rs.Close
Text3.Text = glousername
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo y:
Call condb
rs.Open "delete * from online where online_users='" & glousername & "'", adoconn, adOpenStatic
rs1.Open "delete * from chattemp where from='" & glousername & "'", adoconn, adOpenStatic
Call conclose
y:
If Err.Number = 3705 Then
    'rs.Close
    Call conclose
    'Exit Sub
    Resume
End If
End Sub

Private Sub List1_Click()
lstr = List1.Text
End Sub

Private Sub Timer1_Timer()
On Error GoTo y
Dim s As String
Dim k As Integer
rs.Open "select * from online", adoconn, adOpenStatic
List1.Clear

Do Until rs.EOF = True
    List1.AddItem rs(0)
    rs.MoveNext
Loop
rs.Close

rs.Open "select * from chattemp where to='" & glousername & "' or to= 'default'", adoconn, adOpenStatic
Text4.Text = ""
If rs.RecordCount > 0 Then
Do Until rs.EOF = True
    Text4.Text = Text4.Text & rs.Fields(0) & " : " & rs.Fields(2) & vbCrLf
    Text4.Refresh
    rs.MoveNext
Loop
If Text3.Text = Text4.Text Then
    Exit Sub
Else
Text2.Text = ""
Text2.Text = Text4.Text
Text2.Refresh
End If

'If rs.RecordCount > 1 Then
'    rs.MoveLast
'        s = rs.Fields(0) & ":" & rs.Fields(2)
'        k = InStr(s, Text2.Text)
'
'        If k > 0 Then
'        Exit Sub
'        Else
'        Text2.Text = Text2.Text + vbCrLf + rs.Fields(0) & ":" & rs.Fields(2)
'        End If
'    rs.Close
'ElseIf rs.RecordCount = 1 Then
'        rs.MoveFirst
'        s = rs.Fields(0) & ":" & rs.Fields(2)
'        k = InStr(0, s, Text2.Text)
'        If k > 0 Then
'        Exit Sub
'        Else
'        Text2.Text = Text2.Text + vbCrLf + rs.Fields(0) & ":" & rs.Fields(2)
'        End If
'
'    rs.Close
ElseIf rs.RecordCount = 0 Then
    DoEvents
End If
y:
If Err.Number = 3705 Then
    rs.Close
End If

End Sub
