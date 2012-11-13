Attribute VB_Name = "Module1"
Public Setlist(0 To 100) As String, a As Single, songname(255) As String, totalsongnumber As Integer

    Sub main()
    
    Dim tempname As String, pass As Integer, pos As Integer, k As Integer, ctr As Integer
    
    'loads songlist into an array
    
    Open App.Path & "\Songlist.txt" For Input As #1
    
    ctr = 0
    
    Do While Not EOF(1)
    
        ctr = ctr + 1
        
        Input #1, songname(ctr)
        
    Loop
    
    totalsongnumber = ctr
    
    Close #1
        
    'Sorts Alphabeticaly
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If songname(pos) > songname(pos + 1) Then
            tempname = songname(pos)
            songname(pos) = songname(pos + 1)
            songname(pos + 1) = tempname
            End If
        Next pos
    Next pass
    
    'Adds Song Names to Combo Box
    For k = 1 To ctr
    frmCreate.Combo1.AddItem songname(k)
    Next k

    frmMain.Show

    End Sub
