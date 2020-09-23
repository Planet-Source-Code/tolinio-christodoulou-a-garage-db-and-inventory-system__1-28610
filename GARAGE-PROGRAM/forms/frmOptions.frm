VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   6150
   Tag             =   "Options"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "PELATES"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EXARTHMATA"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "EIDOPIEISH"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Combo2"
      Tab(2).Control(1)=   "Combo1"
      Tab(2).Control(2)=   "Line2"
      Tab(2).Control(3)=   "Label2"
      Tab(2).Control(4)=   "Line1"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -72120
         TabIndex        =   10
         Text            =   "10"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74880
         TabIndex        =   8
         Text            =   "10"
         Top             =   1680
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line2 
         X1              =   -75000
         X2              =   -69120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "POSA PREPEI NA EINAI TA EXARTHMATA STHN APOQHKH GIA NA EIDOPIEISTEOTI TO EXARTHMA KONTEUH NA EXANTLHQH ;"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72120
         TabIndex        =   9
         Top             =   720
         Width           =   2895
      End
      Begin VB.Line Line1 
         X1              =   -72240
         X2              =   -72240
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "META APO POSES MERES QELETE NA EIDOPIEISTE OTI ENAS PELATHS DEN PLHROSE ;"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   5
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   4
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   2
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.Label Label3 
      Caption         =   "UNDER CONSTRACTION"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   4560
      Width           =   3375
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
For i = 1 To 50
Combo1.AddItem i
Combo2.AddItem i
Next i
End Sub
