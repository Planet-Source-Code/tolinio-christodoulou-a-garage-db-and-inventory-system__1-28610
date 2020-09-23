Attribute VB_Name = "dbFunctions"
Dim rsCars As ADODB.Recordset
Dim db1 As ADODB.Connection

Public Sub findCustomer(name As String)
Dim nom, lnom, car As String
Dim results As Integer
results = 0

With frmMain.rs
  .MoveFirst
  If .RecordCount = 0 Then Exit Sub
  Do While Not .EOF
    nom = !FirstName
    lnom = !LastName

      If LCase(nom) = LCase(name) Or LCase(lnom) = LCase(name) Or LCase(nom) & " " & LCase(lnom) = LCase(name) Then
        search.List1.AddItem !LastName & " " & !FirstName & " " & !CustomerNo
        search.Label3.Caption = UCase(name)
        search.iscar = False
        search.Show
        search.Height = 4800
       
        results = results + 1
        .MoveNext
      Else
        .MoveNext
      End If
    
   
    Loop
 
If results > 0 Then
 
 Exit Sub
Else
Set rsCars = New ADODB.Recordset
Set db1 = New ADODB.Connection
db1.Open "Provider=MSDASQL;DSN=TikkisDB;Password=1515151515;"
rsCars.Open "SELECT * FROM Cars", db1, adOpenKeyset, adLockPessimistic

With rsCars
  If .RecordCount = 0 Then Exit Sub
  .MoveFirst
 Do While Not .EOF
If LCase(!Numbers) = LCase(name) Then
search.List1.AddItem !OwnerLast & " " & !OwnerFirst & " " & !Numbers
search.Label3.Caption = UCase(name)
search.iscar = True
search.Show
search.Height = 4800

results = results + 1
.MoveNext
Else
.MoveNext
End If
Loop
End With
End If
    
    

End With
If results = 0 Or results < 1 Then
MsgBox "No Matches Found", vbExclamation
End If
End Sub
