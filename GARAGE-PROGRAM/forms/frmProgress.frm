VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   0
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "Saving New Customer"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = frmCustomers.Command6.Top - frmCustomers.Command6.Height
Me.Left = frmCustomers.Frame2.Left

End Sub

Private Sub Timer1_Timer()
Bar1.Value = Bar1.Value + 10
Label2.Caption = Bar1.Value
If Bar1.Value = 100 Then
Timer1.Enabled = False

Unload Me
frmCustomers.Width = 4780
End If
End Sub
