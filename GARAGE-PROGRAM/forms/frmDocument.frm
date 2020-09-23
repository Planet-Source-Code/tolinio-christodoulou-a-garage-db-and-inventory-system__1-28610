VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIMEIOSEIS"
   ClientHeight    =   5835
   ClientLeft      =   3045
   ClientTop       =   4035
   ClientWidth     =   7785
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   7785
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Button8 
         Caption         =   "Save New Notes"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton Button3 
         Caption         =   "Save New Notes"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5280
         Width           =   1935
      End
      Begin VB.CommandButton Button6 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   4935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   5520
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Button5 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton Button4 
         Caption         =   "Create New Notes"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton Button7 
         Caption         =   "Edit Selected"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Selected Notes"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         Pattern         =   "*.TXT;*.INI;*.lOG"
         TabIndex        =   2
         Top             =   1440
         Width           =   1920
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1560
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button2_Click()
Kill File1.Filename
MsgBox "file deleted"
File1.Refresh
End Sub

Private Sub Button3_Click()
Dim fname As String
Dim num
num = Format$(Now, "DD-MM-YYYY H MM AMPM")

fname = File1.Path & "\" & num & ".txt"
Close #1
'CommonDialog1.CancelError = True
'CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.* "
'CommonDialog1.DialogTitle = "Save As" ' set the commondialogs title
'CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
'On Error GoTo dialogerror
'CommonDialog1.ShowSave
'If CommonDialog1.FileName <> "" Then ' if The Commondialogs filename is equal too anything other than ""
Open fname For Output As #1 ' open the file for output
Print #1, Text1.Text
Close #1
File1.Refresh
File1.Enabled = True
Button3.Visible = False
Button6.Visible = False
Text1.Locked = True

End Sub



Private Sub Button4_Click()
Text1.Text = ""
File1.Enabled = False
Button3.Visible = True
Button6.Visible = True
Text1.Locked = False
End Sub


Private Sub Button5_Click()

Unload Me
End Sub

Private Sub Button6_Click()
Button3.Visible = False
Button6.Visible = False
Button8.Visible = False
File1.Enabled = True
Text1.Locked = True
End Sub

Private Sub Button7_Click()
Text1.Locked = False
Text1.SetFocus
File1.Enabled = False
Button8.Visible = True
Button6.Visible = True
End Sub

Private Sub Button8_Click()
'rtfText.SaveFile File1.FileName
File1.Enabled = True
Button3.Visible = False
Button6.Visible = False
Button8.Visible = False
'rtfText.Locked = True
End Sub

Private Sub File1_DblClick()
LoadNotes File1.Path & "\" & File1.Filename, Text1
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmPopUps.mnuf
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
File1.Path = App.Path & "\notes"
rr = 2
End Sub


