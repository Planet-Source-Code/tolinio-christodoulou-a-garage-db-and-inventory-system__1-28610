Attribute VB_Name = "ListFunctions"
Public Sub LoadNotes(selectedfile As String, txt As TextBox)

On Error GoTo cannotfindfileerr:
If FileLen(selectedfile) > 65000 Then Pictoobig.Visible = True: Exit Sub

Close #1

On Error GoTo cannotfindfileerr:

Open selectedfile For Binary Access Read As #1
Text1.Text = Input(LOF(1), 1)
Close #1

cannotfindfileerr:
If Err.Number <> 0 Then
 On Error GoTo Erroutofmemory

  
    Close #1
     Open selectedfile For Binary Access Read As #1
      If FileLen(selectedfile) > 32000 Then Exit Sub
       txt.Text = Input(LOF(1), 1)
        Close #1
         
          Exit Sub
End If

Erroutofmemory:
    If Err.Number = 7 Then
      
      MsgBox "An Unexpected Error Has Occured" & vbNewLine & Err.Description, vbInformation, "NextPad"
       Exit Sub
    End If
    
    
    

End Sub

Public Function TrimSpaces(Text As String) As String
    Dim Loop1 As Long, SpaceCheck As String
    Dim FullString As String


    For Loop1& = 1 To Len(Text$)
        SpaceCheck$ = Mid(Text$, Loop1&, 1)


        If SpaceCheck$ <> " " Then
            FullString$ = FullString$ & SpaceCheck$
        End If
    Next Loop1&
    TrimSpaces = FullString$
End Function

