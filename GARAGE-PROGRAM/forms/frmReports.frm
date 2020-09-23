VERSION 5.00
Begin VB.Form frmReports 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   1980
   ClientLeft      =   5385
   ClientTop       =   2175
   ClientWidth     =   2775
   Icon            =   "frmReports.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2775
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Customers Balances"
List1.AddItem "Part's Quantity In Stock"
End Sub

