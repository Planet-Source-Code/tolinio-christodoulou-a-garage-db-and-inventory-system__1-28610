VERSION 5.00
Begin VB.Form frmResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add service"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   3135
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Price  £"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Qty"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Desciption"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmServices.List3.AddItem Text1.Text
frmServices.List4.AddItem Text2.Text
frmServices.List5.AddItem Text3.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
