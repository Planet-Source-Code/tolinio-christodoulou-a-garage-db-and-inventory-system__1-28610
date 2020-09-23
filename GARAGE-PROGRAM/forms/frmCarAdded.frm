VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCarAdded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Car - Info"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4005
   Icon            =   "frmCarAdded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4005
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   1920
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37111
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Xronologia"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Last time Serviced"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Xiliometra"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Xroma"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Montelo Autokinitou"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Noumera Autokinitou"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmCarAdded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmCustomers.q = Text1.Text
frmCustomers.List1.AddItem q
frmCustomers.model1 = Text2.Text
frmCustomers.r = Text3.Text
frmCustomers.s = Text5.Text
frmCustomers.m = Text4.Text
Unload Me
End Sub

Private Sub Command2_Click()
frmCustomers.addCar = False
Unload Me
End Sub





Private Sub DTPicker_CloseUp()
   Text6.Text = Format(DTPicker.Value, "d/m/yy")
   Text6.SetFocus

End Sub

Private Sub Form_Load()
  Text6.Locked = True
   Text6.Text = Format(Now, "d/m/yy")
   
   DTPicker.Value = Text6
End Sub
