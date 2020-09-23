VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCustomers 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   6765
   ClientLeft      =   2670
   ClientTop       =   2610
   ClientWidth     =   9030
   Icon            =   "frmCustomers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   9030
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Button7 
         Caption         =   "OK"
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1200
         TabIndex        =   36
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   5040
         Width           =   3615
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1080
         Left            =   1680
         TabIndex        =   21
         Top             =   3720
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtCar 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   6480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Button5 
         Caption         =   "Car Info"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Car"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CARS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Money"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tel. Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   6480
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Customers:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   6480
         Width           =   1230
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Extit"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton Button6 
         Caption         =   "Info"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Button4 
         Caption         =   "Find"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Button2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Button1 
         Caption         =   "Edit"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Button3 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name and Customer Number"
            Object.Width           =   8467
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4800
         Top             =   6720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCustomers.frx":0E42
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public addCar As Boolean
Public q, r, s, m, model1
Dim rs1 As ADODB.Recordset
Dim rsCars As ADODB.Recordset
Dim rsService As ADODB.Recordset
Dim db1 As ADODB.Connection
Dim customerName

Private Sub Button1_Click()
Button7.Visible = False
Me.Width = 9200
customerName = TrimSpaces(ListView1.SelectedItem.Text)
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Combo1.Locked = False
Command7.Visible = True
Command8.Visible = True
Command9.Visible = True

Text1.SetFocus

End Sub

Private Sub Button2_Click()
Dim ret
ret = MsgBox("Are you Sure you want to delete the selected Customer?", vbYesNo, "Delete Part")
If ret = vbYes Then
Button7.Visible = False
Dim todelete As String
todelete = ListView1.SelectedItem.Text
'On Error GoTo err1

With rs1
.MoveFirst
While Not .EOF
If TrimSpaces(todelete) = !LastName & !FirstName & !CustomerNo Then
.Delete

ListView1.ListItems.Remove ListView1.SelectedItem.Index
Exit Sub
Else
.MoveNext
End If
Wend
End With
ElseIf ret = vbNo Then
MsgBox "Customer not deleted"
Exit Sub
End If
End Sub
Public Sub addNew()
Button3_Click
End Sub
Private Sub Button3_Click()
Command9.Visible = True
Me.Width = 9200
clearFields
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text8.Locked = False
Combo1.Locked = False
Command6.Visible = True
Command7.Visible = True
Command9.Visible = True
Text1.SetFocus
Button7.Visible = False
With rs1
Dim storeNum As Integer
storeNum = 0
If .RecordCount = 0 Then
Text8.Text = 1
Exit Sub
End If
.MoveFirst
While Not .EOF

If storeNum < !CustomerNo Or storeNum = !CustomerNo Then
storeNum = !CustomerNo + 1
.MoveNext
Else
.MoveNext
End If
Wend
Text8.Text = storeNum
End With
End Sub

Private Sub Button4_Click()

search.Show
End Sub

Private Sub Button5_Click()
Dim tinmanasougietisprizas As String
If Len(List1.Text) = 0 Then
  MsgBox "Select a car from the list"
  Exit Sub
End If
tinmanasougietisprizas = Text1.Text & Text2.Text
With rsService
  If .RecordCount = 0 Then GoTo noService
  .MoveFirst
  While Not .EOF
If List1.Text = !CarNumber Then
frmcarInfo.Text115.Text = !Date11
.MoveNext
Else
.MoveNext
End If
Wend
End With
noService:
With rsCars
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If tinmanasougietisprizas & List1.Text = !OwnerFirst & !OwnerLast & !Numbers Then
frmcarInfo.Label1.Caption = !OwnerFirst & " " & !OwnerLast
frmcarInfo.Text4.Text = !Numbers
frmcarInfo.Text2.Text = !Color
frmcarInfo.Text1.Text = !CarModel
frmcarInfo.Text3.Text = !Mileage
frmcarInfo.Show
.MoveNext
Else
.MoveNext
End If
Wend
End With
End Sub

Private Sub Button6_Click()
Me.Width = 9200
Button7.Visible = True
Command9.Visible = False
End Sub

Private Sub Button7_Click()
Me.Width = 4780
Button7.Visible = False

End Sub

Private Sub Combo1_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Text1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Command1_Click()
Unload Me
End Sub






Private Sub Command6_Click()
Command9.Visible = False
If Text5.Text = "" Then Text5.Text = 0
With rs1
.addNew
!FirstName = Text1.Text
!LastName = Text2.Text
!Address = Text3.Text
!City = Combo1.ListIndex
!Phone = Text4.Text
!Fax = Text5.Text
If Text6.Text = "" Then Text6.Text = 0
!Money = Text6.Text
!Services = Date
!CustomerNo = Text8.Text
!NumKeeper = .RecordCount + 1
Dim i As Integer
Dim da
For i = 0 To List1.ListCount - 1
da = da & List1.List(i) & ","
Next i

!Cars = da
!CarColor = r
!CarYear = s
!Mileage = m
If Text7.Text = "" Then Text7.Text = "No notes found for this customer..."
!Notes = Text7.Text
.Update
End With
With rsCars
.addNew
!Numbers = q
!Year = s
!Color = r
!OwnerFirst = Text1.Text
!OwnerLast = Text2.Text
!CustomerNum = Text8.Text
!Mileage = m
!CarModel = model1
.Update
End With
frmControls
ListView1.ListItems.Add , , Text2.Text & " " & Text1.Text & "               " & Text8.Text
Load frmProgress
frmProgress.Show
frmProgress.Timer1.Enabled = True
Dim theOwner As String
theOwner = Label6.Caption

With rsService
.addNew
!CarNumber = q
!Date11 = Date
!Owner = Text2.Text & " " & Text1.Text
!ServiceDescription = "FirstTime Serviced"

.Update


End With

End Sub

Private Sub Command7_Click()
frmControls
clearFields
Me.Width = 4780
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
End Sub

Private Sub Command8_Click()
Command9.Visible = False
With rs1
.MoveFirst
While Not .EOF
If .RecordCount = 0 Then Exit Sub
If !LastName & !FirstName & !CustomerNo = customerName Then

!FirstName = Text1.Text
!LastName = Text2.Text
!Address = Text3.Text
!City = Combo1.ListIndex
!Phone = Text4.Text
!Fax = Text5.Text
!Money = Text6.Text
!CustomerNo = Text8.Text
!Cars = !Cars & "," & q
!Notes = Text7.Text
.Update
.MoveNext
Else
.MoveNext
End If
Wend
End With
With rsCars
If addCar = True Then
  .addNew
  !Numbers = q
  !Year = s
  !Color = r
  !OwnerFirst = Text1.Text
  !OwnerLast = Text2.Text
  !CustomerNum = Text8.Text
  !Mileage = m
  !CarModel = model1
  .Update
ElseIf addCar = False Then
  Dim icar As Integer
  .MoveFirst
  While Not .EOF
  If .RecordCount = 0 Then Exit Sub
  For icar = 0 To List1.ListCount - 1
  List1.ListIndex = icar
  If !Numbers = List1.Text Then
  !OwnerFirst = Text1.Text
  !OwnerLast = Text2.Text
  !CustomerNum = Text8.Text

  .Update
  .MoveNext
  Else
  .MoveNext
  End If
  Next icar
  Wend
End If
End With
With rsService
.addNew
!CarNumber = q
!Date11 = Date
!Owner = Text2.Text & " " & Text1.Text
!ServiceDescription = "FirstTime Serviced"

.Update

End With
Me.Width = 4780
Command8.Visible = False
Command7.Visible = False
Command9.Visible = False
Button5.Visible = True
addCar = False

End Sub

Private Sub Command9_Click()
 List1.Clear
frmCarAdded.Show
addCar = True
End Sub

Private Sub Form_Load()
Me.Move 0, 0

Me.Width = 4780
ListView1.ColumnHeaders(1).Width = ListView1.Width - 20
frmControls

'Set db = OpenDatabase(App.Path & "\Database" & "\database.mdb", adLockOptimistic)
Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
Set rs1 = New ADODB.Recordset
Set rsCars = New ADODB.Recordset
Set rsService = New ADODB.Recordset
rs1.Open "SELECT * FROM CUSTOMERS", db1, adOpenKeyset, adLockPessimistic
rsCars.Open "SELECT * FROM Cars", db1, adOpenKeyset, adLockPessimistic

rsService.Open "SELECT * FROM Services", db1, adOpenKeyset, adLockPessimistic
LoadCustomers
Label11.Caption = ListView1.ListItems.Count


 End Sub
Private Sub LoadCustomers()
ListView1.ListItems.Clear
'If rs1.RecordCount > 0 Then
While Not rs1.EOF
    ListView1.ListItems.Add , , rs1!LastName & " " & rs1!FirstName & "               " & rs1!CustomerNo
    rs1.MoveNext
Wend
'End If
End Sub

Private Sub frmControls()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True

Combo1.Locked = True
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command8.Visible = False

'fill combo
Combo1.AddItem ""
Combo1.AddItem "LEFKOSIA"
Combo1.AddItem "LEMESOS"
Combo1.AddItem "LARNACA"
Combo1.AddItem "PAFOS"
Combo1.AddItem "AMMOXOSTOS"
Combo1.ListIndex = 0
End Sub





Private Sub List1_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
Text1.BackColor = vbWhite

End Sub

Private Sub ListView1_Click()

If ListView1.ListItems.Count = 0 Then Exit Sub
Dim name111 As String
name111 = TrimSpaces(ListView1.SelectedItem.Text)
List1.Clear
showCustomer name111, "names"

End Sub

Private Sub clearFields()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Combo1.ListIndex = 0
List1.Clear
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
ListView1_Click
End Sub



Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmPopUps.mnuCustomerList
End If
End Sub

Private Sub Text1_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = vbGreen
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Text2_Click()
Text1.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text2_GotFocus()
Text2.BackColor = vbGreen
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = vbWhite
End Sub

Private Sub Text3_Click()
Text2.BackColor = vbWhite
Text1.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text4_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text1.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = vbGreen
Text5.BackColor = vbWhite
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Numbers = KeyAscii
If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
Beep
KeyAscii = 0
End If

End Sub

Private Sub Text4_LostFocus()
Text5.BackColor = vbWhite
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Numbers = KeyAscii
If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
Beep
KeyAscii = 0
End If

End Sub

Private Sub Text5_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text1.BackColor = vbWhite
Text6.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = vbGreen
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = vbWhite
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Numbers = KeyAscii
If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
Beep
KeyAscii = 0
End If

End Sub

Private Sub Text6_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text1.BackColor = vbWhite
Text7.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = vbGreen
Text5.BackColor = vbWhite
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = vbWhite
End Sub

Private Sub Text7_Click()
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
Text1.BackColor = vbWhite
Combo1.BackColor = vbWhite
List1.BackColor = vbWhite

End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = vbGreen
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = vbWhite
End Sub
Private Sub Combo1_GotFocus()
Combo1.BackColor = vbGreen
End Sub
Private Sub Combo1_LostFocus()
Combo1.BackColor = vbWhite
End Sub
Private Sub List1_GotFocus()
List1.BackColor = vbGreen
End Sub
Private Sub List1_LostFocus()
List1.BackColor = vbWhite
End Sub




Private Sub Text3_GotFocus()
Text3.BackColor = vbGreen
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = vbWhite
End Sub

Public Sub showCustomer(name As String, loadby As String)
'On Error Resume Next
Dim txtcars As String
Dim CarsNo() As String
Dim X As Integer
If loadby = "names" Then
loadingbyname:
With rs1
If .RecordCount = 0 Then Exit Sub
.MoveFirst
Do While Not .EOF

If !LastName & !FirstName & !CustomerNo = name Then
GoTo endOfSearch
Else
.MoveNext
End If
Loop
Exit Sub
endOfSearch:
frmCustomers.Text1 = !FirstName
frmCustomers.Text2 = !LastName
frmCustomers.Text3 = !Address
frmCustomers.Text4 = !Phone
frmCustomers.Text5 = !Fax
frmCustomers.Text6 = !Money
frmCustomers.Text7 = !Notes
frmCustomers.Text8 = !CustomerNo

'frmCustomers.txtCar = !Cars
'cars
'txtcars = !Cars
'CarsNo = Split(txtcars, ",")
 'For x = LBound(CarsNo) To UBound(CarsNo)
 'frmCustomers.List1.AddItem CarsNo(x)
 'Next x
 frmCustomers.Combo1.ListIndex = !City
End With
With rsCars
.MoveFirst
While Not .EOF
If frmCustomers.Text1.Text & frmCustomers.Text2.Text = !OwnerFirst & !OwnerLast Then
frmCustomers.List1.AddItem !Numbers
.MoveNext
Else
.MoveNext
End If
Wend
Exit Sub
End With

ElseIf loadby = "bycar" Then
With rsCars
.MoveFirst
While Not .EOF

If !OwnerLast & !OwnerFirst & !Numbers = name Then
name = !OwnerLast & !OwnerFirst & !CustomerNum
GoTo loadingbyname
.MoveNext
Else
.MoveNext
End If
Wend
End With
End If
End Sub

