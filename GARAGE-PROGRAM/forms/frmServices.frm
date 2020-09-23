VERSION 5.00
Begin VB.Form frmServices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICES"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmServices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6165
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Button1 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   40
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton Button4 
         Caption         =   "Next>>"
         Height          =   375
         Left            =   3240
         TabIndex        =   39
         Top             =   4560
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CommandButton Command4 
            Caption         =   "Add More"
            Height          =   255
            Left            =   3480
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CASH"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   3600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CREDIT"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   3360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "+"
            Height          =   255
            Left            =   3480
            TabIndex        =   30
            Top             =   2640
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   28
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   25
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   24
            Top             =   3120
            Width           =   735
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   2175
            Left            =   2640
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   2175
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   2175
            Left            =   720
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Price £"
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Parts Used"
            Height          =   255
            Left            =   840
            TabIndex        =   17
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
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
            Left            =   4920
            TabIndex        =   31
            Top             =   3240
            Width           =   135
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "9% V.A.T"
            Height          =   255
            Left            =   3480
            TabIndex        =   27
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal"
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Labor"
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   2640
            Width           =   495
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            Height          =   3975
            Left            =   0
            Top             =   0
            Width           =   5535
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   5640
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5520
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Left            =   1920
            TabIndex        =   29
            Top             =   3480
            Width           =   615
         End
      End
      Begin VB.CommandButton Button5 
         Caption         =   "<<Back"
         Height          =   375
         Left            =   2160
         TabIndex        =   38
         Top             =   4560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   3975
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "_"
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "+"
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   960
            Width           =   495
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            Height          =   2760
            Left            =   3240
            TabIndex        =   7
            Top             =   360
            Width           =   2055
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Height          =   2760
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "<--------------------------->"
            Height          =   195
            Left            =   2160
            TabIndex        =   15
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "#"
            Height          =   195
            Left            =   2880
            TabIndex        =   13
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Parts Used"
            Height          =   195
            Left            =   3720
            TabIndex        =   11
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Parts List"
            Height          =   195
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmServices.frx":0442
         Left            =   1560
         List            =   "frmServices.frx":0444
         TabIndex        =   1
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label16 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Milia Autokinitou"
         Height          =   195
         Left            =   2760
         TabIndex        =   36
         Top             =   3000
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   14
         Top             =   120
         Width           =   90
      End
      Begin VB.Label lblPservice 
         AutoSize        =   -1  'True
         Caption         =   "SERVICE DESCRIPTION"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DIALEXTE AUTOKINHTO"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public partsPrice
Public paymentWay As String
Dim rs4Parts As ADODB.Recordset
Dim rs4Customers As ADODB.Recordset
Dim rs4Services As ADODB.Recordset
Dim rs4Invoices As ADODB.Recordset
Dim rs4Cars As ADODB.Recordset
Dim db1 As ADODB.Connection
Dim de As DataEnvironment1
Dim counter As Integer




Private Sub Button1_Click()
Unload Me
End Sub

Private Sub Button4_Click()
Dim partsNum() As String
Dim partUsed
Dim numOfpartUsed
If counter = 0 Then Button5.Visible = True
counter = counter + 1


If counter = 1 Then
If Combo2.Text <> "" Then
Combo2.Visible = False
Label3.Visible = False
Text1.Visible = True
Label15.Visible = True
Text6.Visible = True
lblPservice.Visible = True

Else
MsgBox "PLEASE SELECT A CUSTOMER FROM THE LIST", vbCritical
counter = counter - 1
Button5.Visible = True
End If
ElseIf counter = 2 Then
Frame2.Visible = True
Option1.Visible = True
Option2.Visible = True
Text1.Visible = False
Label15.Visible = False
Text6.Visible = False
lblPservice.Visible = False

ElseIf counter = 3 Then
Button4.Caption = "Finish"
Dim lstcount
For lstcount = 0 To List2.ListCount - 1
List2.ListIndex = lstcount
partsNum = Split(List2.Text, ",")
numOfpartUsed = partsNum(0)
partUsed = partsNum(1)
With rs4Parts
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If partUsed = !PartName Then
!Count = !Count - numOfpartUsed
.Update
.MoveNext
Else
.MoveNext
End If
Wend
End With
Next lstcount
Dim xx As Integer
Dim myArray() As String
Dim num As Integer
Dim name1 As String
For xx = 0 To List2.ListCount - 1
List2.ListIndex = xx
myArray = Split(List2.Text, ",")
num = myArray(0)
name1 = myArray(1)
List3.AddItem name1
List4.AddItem num

'=====get prices=======
getPrices name1, num
Next xx
Frame3.Visible = True
Frame2.Visible = False

ElseIf counter = 4 Then
If Option1.Value = False And Option2.Value = False Then
MsgBox "Please select the way of payment -- Cash Or PISTOSH", vbCritical
counter = counter - 1
ElseIf Option1.Value = True Then
saveService
paymentWay = "Cash"
Form1.Show

Unload Me
ElseIf Option2.Value = True Then
saveService
With rs4Customers
.MoveFirst
While Not .EOF
If Label6.Caption & Label16.Caption = !FirstName & " " & !LastName & !CustomerNo Then
!Money = !Money + Val(Text5.Text)
paymentWay = "PISTOSH"



.MoveNext
Else
.MoveNext
End If
Wend
End With
Form1.Show
Unload Me
End If
End If
End Sub
Private Sub saveService()
Dim theOwner As String
theOwner = Label6.Caption
With rs4Customers
.MoveFirst
While Not .EOF
If Label6.Caption & Label16.Caption = !FirstName & " " & !LastName & !CustomerNo Then
!Mileage = Text6.Text
.Update
.MoveNext
Else
.MoveNext
End If
Wend
End With

With rs4Services
.addNew
!CarNumber = Combo2.Text
!Date11 = Date
!Owner = theOwner
!ServiceDescription = Text1.Text
Dim i As Integer
For i = 0 To List3.ListCount - 1
 List3.ListIndex = i
!PartsUsed = !PartsUsed & "," & List3.Text
Next i
!PartsTotalPrice = Val(Text3.Text) - Val(Text4.Text)
!TheSubTotal = Val(Text3.Text)
!TheTotal = Val(Text5.Text)
!Cash = Option1.Value
.Update


End With

With rs4Invoices
If .RecordCount = 0 Then GoTo adding:
.MoveFirst
While Not .EOF
.Delete
.MoveNext
Wend
'====
adding:
Dim iii As Integer
  For iii = 0 To List3.ListCount - 1
  List3.ListIndex = iii
  List4.ListIndex = iii
  List5.ListIndex = iii
  .addNew
  !PartsUsed = List3.Text
  !QtyUsed = List4.Text
  !Price = List5.Text
  If Option1.Value = True Then
  !Payment = "CASH"
  Else
  !Payment = "PISTOSH"
  End If
  Next iii
  .MoveLast
  !TheSubTotal = Val(Text3.Text)
  !TheTotal = Val(Text5.Text)
  .Update
  End With
With rs4Cars
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If Combo2.Text = !Numbers Then
!Mileage = Text6.Text
.Update
.MoveNext
Else
.MoveNext
End If
Wend
End With

End Sub
Private Sub getPrices(name As String, qty As Integer)
Dim ttlPrice
With rs4Parts
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If name = !PartName Then
ttlPrice = Val(!Price * qty)
List5.AddItem ttlPrice
Exit Sub
Else
.MoveNext
End If
Wend
End With

End Sub
Private Sub Button5_Click()
If counter = 1 Then Button5.Visible = False

counter = counter - 1
If counter = 0 Then
  Combo2.Visible = True
  Text1.Visible = False
  lblPservice.Visible = False
  Label3.Visible = True
  Text6.Visible = False
  Button5.Visible = False
  
ElseIf counter = 1 Then
  Combo2.Visible = False
  Label3.Visible = False
  Text1.Visible = True
  lblPservice.Visible = True
  Frame2.Visible = False
  Text6.Visible = True

Else
  counter = 2
  Frame2.Visible = True
  Frame3.Visible = False
  Text1.Visible = False
  lblPservice.Visible = False
  Button4.Caption = "Next>>"
End If
End Sub

Private Sub Combo2_Change()
If Len(Combo2.Text) > 3 Then
Combo2_Click
End If
End Sub

Private Sub Combo2_Click()

With rs4Cars
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
If Combo2.Text = !Numbers Then
Label6.Caption = !OwnerFirst & " " & !OwnerLast
Label16.Caption = !CustomerNum
Exit Sub
Else
.MoveNext
End If
Wend
End With
End Sub


Private Sub Command1_Click()
If Len(Text2.Text) = 0 Then
MsgBox "Please enter the number of " & List1.Text & " used for the service", vbCritical
Else
If List1.ListCount < 1 Then
MsgBox "There are No items"
Exit Sub
End If
List2.AddItem Text2.Text & "," & List1.Text
End If
End Sub

Private Sub Command2_Click()
On Error GoTo err12
List2.RemoveItem List2.ListIndex
Exit Sub
err12:
If Err.Number = 5 Then
MsgBox "Please Select Item to remove", vbCritical
Exit Sub
End If
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim vat1
Dim subTotal
subTotal = 0
For i = 0 To List5.ListCount - 1
List5.ListIndex = i
subTotal = Val(subTotal) + Val(List5.Text)
Next i
Text3.Text = Val(subTotal) + Val(Text4.Text)
vat1 = Val(Text3.Text)
Text5.Text = vat1 + ((vat1 * 10) / 100)
List3.AddItem "Labor"
List5.AddItem Text4.Text
List4.AddItem "1"
partsPrice = Val(Text3.Text) - Val(Text4.Text)

End Sub

Private Sub Command4_Click()
frmResults.Show
End Sub

Private Sub Form_Load()
Label15.Caption = "Current Mileage"
Label3.Font = "Times New Roman"
Label3.Caption = "Choose Car being Serviced"
counter = 0
Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
Set rs4Parts = New ADODB.Recordset
Set rs4Customers = New ADODB.Recordset
Set rs4Cars = New ADODB.Recordset
Set rs4Services = New ADODB.Recordset
Set rs4Invoices = New ADODB.Recordset
rs4Parts.Open "SELECT * FROM Parts", db1, adOpenKeyset, adLockPessimistic
rs4Cars.Open "SELECT * FROM Cars", db1, adOpenKeyset, adLockPessimistic
rs4Services.Open "SELECT * FROM Services", db1, adOpenKeyset, adLockPessimistic
rs4Customers.Open "SELECT * FROM CUSTOMERS", db1, adOpenKeyset, adLockPessimistic
rs4Invoices.Open "SELECT * FROM Invoices", db1, adOpenKeyset, adLockPessimistic

With rs4Cars
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
Combo2.AddItem !Numbers
.MoveNext
Wend
End With
'====================
'parts
With rs4Parts
If .RecordCount = 0 Then Exit Sub
.MoveFirst
While Not .EOF
List1.AddItem !PartName
.MoveNext
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
