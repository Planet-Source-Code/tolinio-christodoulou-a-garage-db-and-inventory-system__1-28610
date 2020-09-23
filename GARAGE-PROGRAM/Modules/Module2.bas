Attribute VB_Name = "Module2"
Public tme As String



Public Sub ENLG1(txtpass As String, txtm As String)
    Dim Password As String
    Dim Words As String
    Dim Encrypted As String
    Dim counter As Integer
    Dim Tempchar As String
    counter = 1
        Password = txtpass
    Words = txtm


    For x = 1 To Len(Words)
        Tempchar1 = Mid(Password, counter, 1)
        Tempchar = Mid(Words, x, 1)
        
        TempAsc = Asc(Tempchar)
        TempAsc1 = Asc(Tempchar1)
        TempAsc = TempAsc + TempAsc1
        
        
        If TempAsc > 245 Then TempAsc = TempAsc - 245
        
        Tempchar = Chr(TempAsc)
        
        Encrypted = Encrypted & Tempchar
        counter = counter + 1
        
       
        If counter > Len(Password) Then counter = 1
        
    Next x
   
    'frmLogin.txtMain.Text = Encrypted
    tme = Encrypted
    
End Sub


Public Sub ENLG11(txtpass As String, txtm As String)
    Dim rasta As String
    Dim Words As String
    Dim Encrypted As String
    Dim counter As Integer
    Dim Tempchar As String
    counter = 1
    rasta = txtpass
    Words = txtm

    For x = 1 To Len(Words)
        Tempchar1 = Mid(rasta, counter, 1)
        Tempchar = Mid(Words, x, 1)
        
        TempAsc = Asc(Tempchar)
        TempAsc1 = Asc(Tempchar1)
        TempAsc = TempAsc + TempAsc1
        
        
        If TempAsc > 245 Then TempAsc = TempAsc - 245
        
        Tempchar = Chr(TempAsc)
        
        Encrypted = Encrypted & Tempchar
        counter = counter + 1
        
       
        If counter > Len(rasta) Then counter = 1
        
    Next x
   
    tme = Encrypted
    
End Sub


