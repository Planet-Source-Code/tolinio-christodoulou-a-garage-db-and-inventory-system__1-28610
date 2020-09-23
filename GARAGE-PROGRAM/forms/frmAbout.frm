VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "About This Application"
   ClientHeight    =   4005
   ClientLeft      =   3555
   ClientTop       =   3330
   ClientWidth     =   4590
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Tag             =   "About Project1"
   Begin VB.CommandButton Button1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.cyprusgaming.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2610
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0442
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0563
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.Shape Shape1 
      Height          =   3975
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "A. Tikkis Garage Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrights CyprusGaming.com Ltd 2001"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   2820
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Move (frmMain.ScaleWidth - Me.Width) / 2, (frmMain.ScaleHeight - Me.Width) / 2
Label2.Caption = "http://www.codeauction.com"
Label3.Caption = "Copyrights CodeAuction.com Ltd 2001"
End Sub

Private Sub Label2_Click()
ShellExecute hwnd, "open", "http://www.codeauction.com", vbNullString, vbNullString, conSwNormal
End Sub
