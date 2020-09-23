VERSION 5.00
Begin VB.Form frmPopUps 
   Caption         =   "Form1"
   ClientHeight    =   390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleWidth      =   4680
   Begin VB.Menu mnuf 
      Caption         =   "file"
      Begin VB.Menu mnuloadNote 
         Caption         =   "Load Selected Notes"
      End
      Begin VB.Menu mnudelNote 
         Caption         =   "Delete Selected Notes"
      End
      Begin VB.Menu mnuedtiNote 
         Caption         =   "Edit Selected Notes"
      End
   End
   Begin VB.Menu mnuCustomerList 
      Caption         =   "f"
      Begin VB.Menu mnuact 
         Caption         =   "Actions"
         Enabled         =   0   'False
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmPopUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnudelNote_Click()
Kill App.Path & "\notes" & "\" & frmDocument.File1.Filename
MsgBox "file deleted"
frmDocument.File1.Refresh
End Sub

Private Sub mnueditNote_Click()
'rtfText.Locked = False
'rtfText.SetFocus
End Sub


Private Sub mnuloadNote_Click()
LoadNotes frmDocument.File1.Path & "\" & frmDocument.File1.Filename, frmDocument.Text1
Unload Me
End Sub

