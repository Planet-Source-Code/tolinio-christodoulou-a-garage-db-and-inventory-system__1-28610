VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox txtmainn 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Text            =   "ssss"
      Top             =   4680
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Button3 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Button4 
         Caption         =   "OK"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "^"
         TabIndex        =   10
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Tag             =   "&Password:"
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Tag             =   "&Password:"
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re type Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Tag             =   "&Password:"
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGING PASSWORD"
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
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Button2 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Button1 
         Caption         =   "OK"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         Picture         =   "frmLogin.frx":1272
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtUserName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Text            =   "Admin"
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         MouseIcon       =   "frmLogin.frx":13B8
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "&User Name:"
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "&Password:"
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.TextBox txtMain 
      Height          =   225
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   4770
      Width           =   495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Private Sub Button1_Click()
If txtPassword.Text = "" Then txtPassword.Text = " "
ENLG1 txtPassword.Text, txtMain.Text
With rs
If .RecordCount = 0 Then Exit Sub
  .MoveFirst
  
  If tme = !aaaaaaaaaaaaaaa Then

With rs2

.MoveFirst

!Okman = 1
.Update
End With
    Unload Me
    Load frmMain
    frmMain.Show
    frmMain.Enabled = True
    frmMain.Toolbar1.Enabled = True
    frmMain.menuFile.Enabled = True
    frmMain.mnuFile.Enabled = True
    frmMain.mnuView.Enabled = True
    frmMain.mnuTools.Enabled = True
    frmMain.mnuHelp.Enabled = True
    
  Else
    MsgBox "Invalid Password, try again!", , "Login"
    txtPassword.SetFocus
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
   txtMain.Text = "Text4"
    Exit Sub
  End If

End With
End Sub


Private Sub Button2_Click()
End
End Sub





Private Sub Button3_Click()
Frame2.Visible = False
Frame1.Visible = True
End Sub

Private Sub Button4_Click()
If Text1.Text = "" Then Text1.Text = " "
If Text2.Text = "" Then Text2.Text = " "
If Text3.Text = "" Then Text3.Text = " "
Dim pw As String
pw = Text2.Text

If Text2.Text <> Text3.Text Then
  MsgBox "Field 2 (new password) and Field3 (Re type password) do not match", vbInformation, "Misstyped password"
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Exit Sub
End If
ENLG11 Text1.Text, txtMain.Text
With rs
If .RecordCount = 0 Then Exit Sub
  .MoveFirst
  If tme = !aaaaaaaaaaaaaaa Then
  txtMain.Text = "Text4"

  ENLG1 txtPassword.Text, txtMain.Text
    !aaaaaaaaaaaaaaa = tme
  .Update
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
    MsgBox "Password has changed. The new password is : " & pw & vbCrLf & "Remember to write down your password"
    Frame2.Visible = False
    Frame1.Visible = True
    txtMain.Text = "Text4"
    Exit Sub
  Else
  MsgBox "Incorrect Password. Please Try again"
  
End If
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
createDsn


Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
db.Open "Provider=MSDASQL;DSN=TikkisDb;Password=1515151515;"
rs.Open "SELECT * FROM WestSide", db, adOpenKeyset, adLockPessimistic
rs2.Open "SELECT * FROM CustomerYN", db, adOpenKeyset, adLockPessimistic



End Sub

Private Sub createDsn()
Dim szDriverName As String
Dim szWantedDSN As String

szDriverName = String(255, Chr(32))
szWantedDSN = "TikkisDb"
'is access drivers installed?


If Not checkAccessDriver(szDriverName) Then
    MsgBox "You must Install Access ODBC Drivers before use this program.", vbOK + vbCritical
KeySection = "Times"
KeyKey = "TimesUsed"
KeyValue = 0
saveini

End If

'is our dsn exist?


If Not (checkWantedAccessDSN(szWantedDSN)) Then


    If szDriverName = "" Then
        MsgBox "Can't find access ODBC driver.", vbOK + vbCritical
        KeySection = "Times"
        KeyKey = "TimesUsed"
        KeyValue = 0
        saveini
    Else


        If Not createAccessDSN(szDriverName, szWantedDSN) Then
            MsgBox "Can't create database ODBC.", vbOK + vbCritical
        Else
         KeySection = "Times"
         KeyKey = "TimesUsed"
         KeyValue = 1
         saveini
        End If
    End If
End If

End Sub






Private Sub Form_Resize()
txtPassword.SetFocus
End Sub




'Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label1.ForeColor = vbBlue
'End Sub

Private Sub Label1_Click()
Frame2.Visible = True
Frame1.Visible = False

End Sub

'Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Label1.ForeColor = vbRed
'End Sub


Private Sub Text2_Change()
txtPassword.Text = Text2.Text
End Sub

'Private Sub txtPassword_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'Button1_Click
'End If





Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Button1_Click
End If
End Sub

Private Sub loadini()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\createDsn.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub


Private Sub saveini()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\createDsn.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub


