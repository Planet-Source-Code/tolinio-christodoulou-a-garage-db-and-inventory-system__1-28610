VERSION 5.00
Begin VB.Form frmParts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXARTHMATA"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Symbol"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   7680
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   3360
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Button4 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Button6 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Button5 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton Button9 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "NAI H OCI"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2760
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCars.frx":0442
         Left            =   2160
         List            =   "frmCars.frx":044C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIMH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PERIGRAFH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KATASTASH"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POSOTHTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STO STOK"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ARIQMOS KATALOGOU"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ONOMA EXARTHMATOS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton Button8 
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Button3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Button7 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Button1 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Button2 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As ADODB.Recordset
Dim dbParts As ADODB.Connection
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Button1_Click()
Me.Width = 8190
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Check1.Value = 0
Combo1.ListIndex = 0
Text5.Text = ""
Text1.Locked = False
Text2.Locked = False
Text4.Locked = False
Text3.Locked = False
Text5.Locked = False
Check1.Enabled = True
Combo1.Locked = False
Button5.Visible = True
Button6.Visible = True
Text1.SetFocus
End Sub

Private Sub Button2_Click()
If List1.Text = "" Then
MsgBox "Äéáëå÷ôå åîáñôçìá áðï ôçí ëéóôá ãéá äéïñèïóç", vbCritical
Exit Sub
End If
Me.Width = 8190
Button9.Visible = True
Button6.Visible = True
Button5.Visible = False
Text1.Locked = False
Text2.Locked = False
Text4.Locked = False
Text3.Locked = False
Text5.Locked = False
Check1.Enabled = False
Combo1.Locked = False
Label8.Visible = True
End Sub

Private Sub Button3_Click()
Dim ret
ret = MsgBox("Are you Sure you want to delete the selected part?", vbYesNo, "Delete Part")
If ret = vbYes Then
With rsParts
.MoveFirst
While Not .EOF
If List1.Text = !PartName Then
.Delete
List1.RemoveItem List1.ListIndex
.MoveNext
Else
.MoveNext
End If
Wend
End With
Else
Exit Sub
End If
End Sub

Private Sub Button4_Click()
Unload Me
End Sub

Private Sub Button5_Click()
With rsParts
.addNew
!PartNo = Text2.Text
!PartName = Text1.Text
!Price = Val(Text3.Text)
!PartCondition = Combo1.ListIndex
!Instock = Check1.Value
!Count = Text4.Text
If Text5.Text = "" Then Text5.Text = " "
!PartDescription = Text5.Text
.Update
List1.AddItem Text1.Text
'MsgBox !PartName & !PartNo & !Price & !PartCondition & !InStock
End With
Me.Width = 3375
End Sub

Private Sub Button6_Click()
Button5.Visible = False
Me.Width = 3375
End Sub

Private Sub Button7_Click()
On Error GoTo error12
If List1.Text = "" Then
MsgBox "Please Select from List", vbCritical
Exit Sub
End If
Dim box As Integer
box = InputBox("How Many " & List1.Text & " were added in the stock?" & vbCrLf & "(write the number below)")
With rsParts
.MoveFirst
While Not .EOF
If List1.Text = !PartName Then

!Count = !Count + box
.Update

Text4.Text = !Count
.MoveNext
Else
.MoveNext
End If
Wend
End With
Exit Sub
error12:
MsgBox Err.Description, vbCritical
Exit Sub
End Sub

Private Sub Button8_Click()
Me.Width = 8190
End Sub

Private Sub Button9_Click()
With rsParts
While Not .EOF
If Text6.Text = !PartNo And List1.Text = !PartName Then

!PartNo = Text2.Text
!PartName = Text1.Text
!Price = Text3.Text
!PartCondition = Combo1.ListIndex
!Instock = Check1.Value
!Count = Text4.Text
If Text5.Text = "" Then Text5.Text = " "
!PartDescription = Text5.Text
.Update
.MoveNext
Else
.MoveNext
End If
Wend
End With
Button5.Visible = True
Button5.Visible = False
Button9.Visible = False
Me.Width = 3375
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If

End Sub

Private Sub Form_Load()
Me.Move 0, 0, 3375
Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
Set rsParts = New ADODB.Recordset
rsParts.Open "SELECT * FROM Parts", db1, adOpenKeyset, adLockPessimistic

'Set db = OpenDatabase(App.Path & "\Database" & "\database.mdb", adLockOptimistic)
'Set rs = db.OpenRecordset("SELECT * FROM Parts")
With rsParts
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
List1.AddItem !PartName
.MoveNext
Wend
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'rs.Close

End Sub



Private Sub List1_Click()
With rsParts
.MoveFirst
While Not .EOF
If List1.Text = !PartName Then
Text2.Text = !PartNo
Text1.Text = !PartName
Text3.Text = !Price
Text4.Text = !Count
Text6.Text = !PartNo
Combo1.ListIndex = !PartCondition
If !Instock = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
Text5.Text = !PartDescription
Exit Sub
Else
.MoveNext
End If
Wend
End With
End Sub

Private Sub List1_DblClick()
List1_Click
Me.Width = 8190
End Sub



Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Numbers = KeyAscii
If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
Beep
KeyAscii = 0
End If

End Sub



