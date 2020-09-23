VERSION 5.00
Begin VB.Form search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANAZHTHSH"
   ClientHeight    =   915
   ClientLeft      =   5220
   ClientTop       =   4785
   ClientWidth     =   3945
   Icon            =   "search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   915
   ScaleWidth      =   3945
   Begin VB.CommandButton Button1 
      Caption         =   "Search"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3960
      Width           =   495
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "search.frx":1CFA
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "search.frx":213C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You searched for: "
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1305
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iscar As Boolean
Public loadByCar As String

Private Sub Command1_Click()
List1_DblClick
End Sub

Private Sub Form_Resize()
Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMain.db.Close
End Sub

Private Sub List1_DblClick()
On Error GoTo exodos
If iscar = False Then
Dim name1 As String
name1 = TrimSpaces(List1.Text)
frmCustomers.Show
frmCustomers.showCustomer name1, "names"
frmCustomers.Width = 9000

End If
If iscar = True Then
name1 = TrimSpaces(List1.Text)
frmCustomers.Show
frmCustomers.showCustomer name1, "bycar"
frmCustomers.Width = 9000
End If
Unload Me
Exit Sub
exodos:
Unload Me

End Sub

Private Sub Button1_Click()
If Text1.Text = "1515151515" Then
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate "http://www.thehun.com"
End If
search.List1.Clear


findCustomer Text1.Text
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Me.Height = 1260
End Sub

Private Sub Form_Load()
Me.Move 0, 0
End Sub




Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Button1_Click
End If
End Sub
