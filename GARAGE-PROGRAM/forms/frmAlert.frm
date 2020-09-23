VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debtors Alert"
   ClientHeight    =   2925
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4740
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   4740
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Button1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "Show Alert at Startup"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsService As ADODB.Recordset
Dim db1 As ADODB.Connection

Private Sub Button1_Click()
Unload Me
End Sub

Private Sub chkLoadTipsAtStartup_Click()
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub Form_Load()
Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
Set rsService = New ADODB.Recordset
rsService.Open "SELECT * FROM Customers", db1, adOpenKeyset, adLockPessimistic
With rsService
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
Dim leftover As Integer
leftover = Date - !Services
If leftover > 8 Then

List1.AddItem !LastName & !FirstName & ":" & " haven't paid for --" & leftover & "-- days"
.MoveNext
Else
.MoveNext
End If
Wend
End With
End Sub
