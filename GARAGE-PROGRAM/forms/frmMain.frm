VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Andreas Tikkis Garage"
   ClientHeight    =   7845
   ClientLeft      =   1905
   ClientTop       =   1185
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7590
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19606
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1429
      ButtonWidth     =   1535
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customers"
            Key             =   "pelates"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Money"
            Key             =   "lefta"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PARTS"
            Key             =   "parts"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            Key             =   "notes"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SERVICE"
            Key             =   "service"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back-Up"
            Key             =   "internet"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reminder"
            Key             =   "reminder"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Key             =   "reports"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "euresi"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lock"
            Key             =   "lock"
            ImageIndex      =   14
         EndProperty
      EndProperty
      BorderStyle     =   1
      Enabled         =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5580
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4920
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5006
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B36E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C3D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C722
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C87E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D55A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E7DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E902
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F756
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11462
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":118B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Enabled         =   0   'False
      Begin VB.Menu menuExit 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Customers"
      Enabled         =   0   'False
      Begin VB.Menu mnuFilemnuFilePelates 
         Caption         =   "Customer Database"
      End
      Begin VB.Menu mnuFilemnuFileNewcust 
         Caption         =   "Eggrafi Neou Pelati"
      End
      Begin VB.Menu mnufindCustomer 
         Caption         =   "EYRESH PELATH"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
      End
      Begin VB.Menu mnuViewmnuLefta 
         Caption         =   "Lefta"
      End
      Begin VB.Menu mnuViewmnuExartimata 
         Caption         =   "Exartimata"
      End
      Begin VB.Menu mnuViewmnuService 
         Caption         =   "Service"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu mnIntBrowser 
         Caption         =   "Internet Browser"
      End
      Begin VB.Menu mnuToolsmnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuToolsmnuReminder 
         Caption         =   "Reminder"
      End
      Begin VB.Menu menuBackUp 
         Caption         =   "Back-Up Database"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
      Begin VB.Menu mnuhlp 
         Caption         =   "Help Index"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db  As ADODB.Connection
Public rs As ADODB.Recordset
Public rs111 As ADODB.Recordset
Public rsCars As ADODB.Recordset

Private Sub MDIForm_Load()
loadAlert
Set db = New ADODB.Connection
Set rsCars = New ADODB.Recordset
Set rs111 = New ADODB.Recordset
Set rs = New ADODB.Recordset
db.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
rs.Open "SELECT * FROM CUSTOMERS", db, adOpenKeyset, adLockPessimistic, adCmdText
rs111.Open "SELECT * FROM CustomerYN", db, adOpenKeyset, adLockPessimistic, adCmdText
rsCars.Open "SELECT * FROM Cars", db, adOpenKeyset, adLockPessimistic, adCmdText

End Sub

Private Sub loadAlert()
Dim ShowAtStartup As Long
ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
If ShowAtStartup = 1 Then
  frmAlert.Show
  frmAlert.chkLoadTipsAtStartup.Value = vbChecked
End If
If ShowAtStartup = 0 Then
 'frmAlert.chkLoadTipsAtStartup.Value = vbUnchecked
  Unload frmAlert
 End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
With rs111
.MoveFirst
 !Okman = 0
 .Update
End
End With
db.Close


End Sub

Private Sub menuBackUp_Click()
frmBackup.Show
End Sub

Private Sub menuExit_Click()
End
End Sub

Private Sub mnufindCustomer_Click()
search.Show

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuhlp_Click()
frmHelp.Show
  'frmHelp.RichTextBox1.Filename = App.Path & "\Reports\Q100367.rtf"
End Sub

Private Sub mnuReports_Click()
frmReports.Show
End Sub



Private Sub mnuToolsmnuReminder_Click()
Load frmAlert
frmAlert.Show

End Sub

Private Sub mnuToolsmnuCalculator_Click()
   
  Shell "calc.exe", vbNormalFocusEnd
End Sub


Private Sub mnuViewmnuService_Click()
frmServices.Show

End Sub

Private Sub mnuViewmnuExartimata_Click()
frmParts.Show
End Sub

Private Sub mnuViewmnuLefta_Click()
   frmMoney.Show
End Sub

Private Sub mnuViewWebBrowser_Click()
    
    frmBrowser.StartingAddress = "http://www.carparts.com"
    frmBrowser.Show
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show
End Sub




Private Sub mnuFileSend_Click()
   
    
End Sub

Private Sub mnuFilePrint_Click()
   On Error Resume Next
    If rr = 0 Then Exit Sub
 
If rr = 1 Then
frmBrowser.brwWebBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
ElseIf rr = 2 Then
Printer.Print "" + frmDocument.Text1.Text + Str(Printer.Page)
Printer.NewPage
Printer.Print "" + frmDocument.Text1.Text + Str(Printer.Page)
Printer.EndDoc

End If
End Sub





Private Sub mnuFilemnuFileNewcust_Click()
frmCustomers.Show
frmCustomers.addNew

End Sub

Private Sub mnuFilemnuFilePelates_Click()
frmCustomers.Show
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key

Case "pelates"
  If frmCustomers.WindowState = vbMinimized Then frmCustomers.WindowState = vbNormal: Exit Sub
  frmCustomers.Show
  frmCustomers.SetFocus
  frmCustomers.Left = 0
  frmCustomers.Top = 0
Case "parts"
  If frmParts.WindowState = vbMinimized Then frmParts.WindowState = vbNormal: Exit Sub
  Load frmParts
  frmParts.Show
  frmParts.SetFocus
  frmParts.Left = 0
  frmParts.Top = 0
Case "lefta"
  If frmMoney.WindowState = vbMinimized Then frmMoney.WindowState = vbNormal: Exit Sub
  Load frmMoney
  frmMoney.Show
  frmMoney.SetFocus
  frmMoney.Left = 0
  frmMoney.Top = 0
Case "notes"
  If frmDocument.WindowState = vbMinimized Then frmDocument.WindowState = vbNormal: Exit Sub
  Load frmDocument
  frmDocument.Show
  frmDocument.SetFocus
 Case "internet"
  
  If frmBackup.WindowState = vbMinimized Then frmBackup.WindowState = vbNormal: Exit Sub
  Load frmBackup
  frmBackup.Show
  frmBackup.SetFocus
  
Case "lock"
  frmLogin.Show
  Me.Enabled = False
Case "reminder"
  If frmAlert.WindowState = vbMinimized Then frmAlert.WindowState = vbNormal
  Load frmAlert
  Dim ShowAtStartup As Long
  ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
  If ShowAtStartup = 0 Then
    frmAlert.chkLoadTipsAtStartup.Value = 0
  Else
    frmAlert.chkLoadTipsAtStartup.Value = 1
    frmAlert.Show
    frmAlert.SetFocus
  End If
Case "euresi"

 If search.WindowState = vbMinimized Then search.WindowState = vbNormal
  Load search
  search.Show
  search.SetFocus
Case "service"

If frmServices.WindowState = vbMinimized Then frmServices.WindowState = vbNormal
  Load frmServices
  frmServices.Show
  frmServices.SetFocus
Case "reports"
    
    frmReports.Show
End Select
End Sub
