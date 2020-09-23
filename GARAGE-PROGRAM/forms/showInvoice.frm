VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   1095
   Begin VB.CommandButton Command1 
      Caption         =   "Show and print Invoice"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim de As DataEnvironment1
        Set de = New DataEnvironment1
        Set DataReport1.DataSource = de
         DataReport1.Show
    Unload Me
End Sub
