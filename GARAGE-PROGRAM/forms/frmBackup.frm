VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Database"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4035
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   4035
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3885
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Unpack Database"
      Height          =   855
      Left            =   120
      Picture         =   "frmBackup.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backup Database"
      Height          =   855
      Left            =   120
      Picture         =   "frmBackup.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      Caption         =   $"frmBackup.frx":0A70
      Height          =   1095
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Click This button to compress and save your database into a disk. It's recomended that you do that every day"
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub Command1_Click()

  Dim OldTimer As Single

ProgressBar1.Visible = True

  OldTimer = Timer
  
  Call Huffman.EncodeFile(App.Path & "\Database" & "\db.mdb", "a:\db.bkp")
  ProgressBar1.Value = 0
  StatusBar1.Panels(1).Text = ""
 Unload Me
  Exit Sub
  
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

  Set Huffman = New clsHuffman
  
End Sub


Private Sub Huffman_Progress(Procent As Integer)
  StatusBar1.Panels(1).Text = "Compressing Database"
 ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    StatusBar1.Panels(1).Text = "Saving to A:\"
    End If
  DoEvents

End Sub



