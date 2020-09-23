VERSION 5.00
Begin VB.Form frmcarInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAR INFORMATION"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmcarInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   4140
   Begin VB.CommandButton Command2 
      Caption         =   "ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text115 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Last Service"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Numbers"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Mileage"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Color"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Model"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CAR INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ONOMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmcarInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

