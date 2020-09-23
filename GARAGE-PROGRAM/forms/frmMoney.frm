VERSION 5.00
Begin VB.Form frmMoney 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MONEY"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5220
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Button2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Button1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   2880
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "£"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "£"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Debtors List"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomers As ADODB.Recordset
Dim rsPayments As ADODB.Recordset
Dim db1 As ADODB.Connection

'Dim rs As Recordset
'Dim au As Recordset
Dim getName, getLastName As String
Dim custNo As Integer
Private Sub Button1_Click()
'On Error GoTo errhandler
With rsCustomers
.MoveFirst
While Not .EOF

If !FirstName & " " & !LastName = List1.Text Then
getName = !FirstName
getLastName = !LastName
custNo = !CustomerNo
!Money = !Money - Text2.Text
!Services = Date
.Update
Text1.Text = !Money

.MoveNext
Else
.MoveNext
End If
Wend
End With
With rsPayments
.addNew
!FirstName = getName
!LastName = getLastName
!CustomerNumber = custNo
!Dates = Date
!AmountPaid = Text2.Text
If Text11.Text = "" Then Text11.Text = "__________"
!Memo = Text11.Text
.Update
MsgBox "Done", vbInformation
Text2.Text = ""
Text11.Text = ""
End With

End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
Set rsCustomers = New ADODB.Recordset
Set rsPayments = New ADODB.Recordset
rsCustomers.Open "SELECT * FROM CUSTOMERS", db1, adOpenKeyset, adLockPessimistic
rsPayments.Open "SELECT * FROM Pmds", db1, adOpenKeyset, adLockPessimistic
'Set rs = frmMain.db.OpenRecordset("SELECT * FROM Customers")
'Set au = frmMain.db.OpenRecordset("SELECT * FROM Pmds")


With rsCustomers
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
List1.AddItem !FirstName & " " & !LastName

.MoveNext
Wend
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'rs.Close
'au.Close

End Sub

Private Sub List1_Click()
Button1.Visible = True

With rsCustomers
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If !FirstName & " " & !LastName = List1.Text Then
Text1.Text = !Money
.MoveNext
Else
.MoveNext
End If
Wend
End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Numbers = KeyAscii
If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
Beep
KeyAscii = 0
End If

End Sub

