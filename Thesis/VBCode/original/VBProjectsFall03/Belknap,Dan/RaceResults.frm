VERSION 5.00
Begin VB.Form frmRace 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picpicture 
      Height          =   11055
      Left            =   0
      ScaleHeight     =   10995
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton cmdski 
         Caption         =   "Skiing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdbike 
         Caption         =   "Biking"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdquit 
         Caption         =   "Quit"
         Height          =   495
         Left            =   8280
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdsortage 
         Caption         =   "Sort by Age Group"
         Height          =   495
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdsortcity 
         Caption         =   "Sort by City"
         Height          =   495
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdsortname 
         Caption         =   "Sort by Name"
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdsortnumber 
         Caption         =   "Sort by Number"
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdsorttime 
         Caption         =   "Sort by Time/Place"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdresults 
         Caption         =   "Enter Results"
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdentry 
         Caption         =   "Begin Registration"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox picconfirm 
         Height          =   9015
         Left            =   360
         ScaleHeight     =   8955
         ScaleWidth      =   8595
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   8655
      End
      Begin VB.Label Label1 
         Caption         =   "Please Choose a Race Type:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Race Results by Daniel Belknap
'Form name = frmRace(M:\Desktop\Belknap, Dan\Race Results.frm)
'Finished on 10/27/03
'This program accepts input of name, age, and city for registration before the race.
'At the end of the race, the user inputs the person's number(printed during registration)
'and their time.  The user can then sort the finishers by time/place, number, name, hometown,
'or age, providing age groups to fit their race.
Option Explicit
Dim path As String
Dim registration(100, 6) As String, y As Integer, j As Integer, temparray(100, 6) As String, x As Integer

Private Sub cmdbike_Click()
    picpicture.Visible = True
    picpicture.Picture = LoadPicture(path & "\bike program.jpg")
    cmdski.Visible = False
    cmdbike.Visible = False
    cmdentry.Visible = True
    picconfirm.Visible = True
    Label1.Visible = False
End Sub

Private Sub cmdEntry_Click()
    Dim k As Integer
    Dim name As String, number As Integer, n As String, sp As String, city As String
    Dim regname As String, age As Single, first As String, last As String
    Open path & "\registration.txt" For Output As #1
    Open path & "\registration printable.txt" For Output As #2
        Do While name <> "i'm done"
        j = j + 1
            name = InputBox("Enter the participant's name. Type i'm done to finish.", "Entry") 'gets person's name
            If name = "" Then 'if the person enters a blank name
                picconfirm.Print "That is not a name"
                j = j - 1
            Else
                sp = InStr(name, " ") 'rearranging name from "first last" to "last, first"
                first = Left(name, sp - 1)
                last = Right(name, Len(name) - sp + 1)
                regname = last & ", " & first
                registration(j, 4) = regname
                registration(j, 3) = j 'assigning a person's number and putting it in the array
            End If
            If name = "i'm done" Then
                registration(j, 3) = ""
                registration(j, 4) = ""
                j = j - 1
            Else
                age = InputBox("Enter Age:", "Age")
                If age < 10 Then
                    picconfirm.Print "That person is too young."
                    j = j - 1
                ElseIf age = 0 Then
                picconfirm.Print "Please enter an age"
                Else
                    registration(j, 6) = age 'putting the person's name in the array
                End If
                city = InputBox("Enter City:", "City") 'putting a person's city in the array
                If city = "" Then
                    picconfirm.Print "Please enter a city."
                    j = j - 1
                Else
                    registration(j, 5) = city
                    picconfirm.Cls
                    picconfirm.Print name; "'s number is:"; j 'printing the name and number
                End If
            End If
       Loop
    picconfirm.Cls
    picconfirm.Print "Number"; Tab(15); "Name"; Tab(45); "City"; Tab(80); "Age"
    Print #2, "Number"; Tab(15); "Name"; Tab(45); "City"; Tab(80); "Age"
    For k = 1 To j
        Write #1, registration(k, 3); registration(k, 4); registration(k, 5); registration(k, 6)
        Print #2, registration(k, 3); Tab(14); 'printing a table of all the racers after registration
        Print #2, registration(k, 4); Tab(45);
        Print #2, registration(k, 5); Tab(80);
        Print #2, registration(k, 6)
        picconfirm.Print registration(k, 3); Tab(14); 'printing a table of all the racers after registration
        picconfirm.Print registration(k, 4); Tab(45);
        picconfirm.Print registration(k, 5); Tab(80);
        picconfirm.Print registration(k, 6)
    Next k
    cmdresults.Visible = True
    cmdentry.Visible = False
    'cmdprint.Visible = True
    Close #1
    Close #2
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdResults_Click()
    Dim number As Integer, temp1 As Integer, temp2 As String, temp3 As String, temp4 As Integer
    Dim found As Boolean, i As Integer, n As Integer, time As String, m As Integer, r As Integer, w As Integer
    Dim z As Integer, comp As Integer, pass
    Open path & "\sort by time.txt" For Output As #2
    Open path & "\registration.txt" For Input As #1
    Open path & "\time printable.txt" For Output As #3
    Do While Not EOF(1)
        x = x + 1
        Input #1, temparray(x, 3), temparray(x, 4), temparray(x, 5), temparray(x, 6)
    Loop
    y = 0
    For r = 1 To x
        number = InputBox("Enter the Finisher's number", "Finish Line")
        time = InputBox("Enter their time:", "Finish Line")
        found = False
        'temparray(r, 2) = time 'puts the person's time in the array
        If number < 1 Or number > x Then
            picconfirm.Cls
            picconfirm.Print "That person did not register in the race."
            r = r - 1
        End If
        temparray(r, 2) = time
    Next r
        'Else
    
    For n = 1 To x
            'temparray(n, 2) = time
            found = False
            i = 0
            'finding the number in the array
            Do Until found = True Or i = x
                i = i + 1
                If number = temparray(i, 3) Then found = True
            Loop
            'assigning which place they are in and moving data from row in array
            temp1 = temparray(n, 3)
            temp2 = temparray(n, 4)
            temp3 = temparray(n, 5)
            temp4 = temparray(n, 6)
            temparray(n, 3) = temparray(i, 3)
            temparray(n, 4) = temparray(i, 4)
            temparray(n, 5) = temparray(i, 5)
            temparray(n, 6) = temparray(i, 6)
            temparray(i, 3) = temp1
            temparray(i, 4) = temp2
            temparray(i, 5) = temp3
            temparray(i, 6) = temp4
            temparray(i, 1) = i
            Write #2, temparray(n, 1), temparray(n, 2), temparray(n, 3), temparray(n, 4), temparray(n, 5), temparray(n, 6)
    Next n
    For pass = 1 To x - 1
        If temparray(pass, 2) = temparray(pass + 1, 2) Then
            temparray(pass + 1, 1) = temparray(pass, 1)
        End If
    Next pass
'prints the results after all racers have finished
picconfirm.Cls
picconfirm.Print "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
Print #3, "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For y = 1 To x
    Print #3, Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
    picconfirm.Print Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
Next y
cmdresults.Visible = False
cmdsorttime.Visible = True
cmdsortnumber.Visible = True
cmdsortname.Visible = True
cmdsortage.Visible = True
cmdsortcity.Visible = True
cmdentry.Visible = False
Close #2
Close #3
End Sub



Private Sub cmdski_Click()
    picpicture.Visible = True
    picpicture.Picture = LoadPicture(path & "\skiier program.jpg")
    cmdski.Visible = False
    cmdbike.Visible = False
    cmdentry.Visible = True
    picconfirm.Visible = True
    Label1.Visible = False
End Sub

Private Sub cmdsortage_Click()
Dim ctr As Integer, temp1 As Single, temp2 As Single, temp3 As Single, temp4 As String, group1 As String
Dim temp5 As String, temp6 As Single, pass As Integer, n As Integer, group2 As String, group3 As String, g As Integer
Dim group4 As String, group5 As String, a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim groups1(100, 6) As String, groups2(100, 6) As String, groups3(100, 6) As String, groups4(100, 6) As String, groups5(100, 6) As String, groups6(100, 6) As String
Dim h As Integer, i As Integer, k As Integer, l As Integer, m As Integer, q As Integer
Dim r As Integer, s As Integer, t As Integer, u As Integer, v As Integer, w As Integer
picconfirm.Cls
group1 = InputBox("Enter the upper end of the youngest age group", "Groups")
group2 = InputBox("Enter the upper end of the second age group", "Groups")
group3 = InputBox("Enter the upper end of the third age group", "Groups")
group4 = InputBox("Enter the upper end of the fourth age group", "Groups")
group5 = InputBox("Enter the upper end of the fifth age group", "Groups")
Open path & "\sort by age.txt" For Output As #3
Open path & "\age printable.txt" For Output As #4
ctr = 0
    For ctr = 1 To x - 1
        For pass = 1 To x - ctr
            If temparray(pass, 2) > temparray(pass + 1, 2) Then 'sorts times
                temp1 = temparray(pass, 1)
                temp2 = temparray(pass, 2)
                temp3 = temparray(pass, 3)
                temp4 = temparray(pass, 4)
                temp5 = temparray(pass, 5)
                temp6 = temparray(pass, 6)
                temparray(pass, 1) = temparray(pass + 1, 1)
                temparray(pass, 2) = temparray(pass + 1, 2)
                temparray(pass, 3) = temparray(pass + 1, 3)
                temparray(pass, 4) = temparray(pass + 1, 4)
                temparray(pass, 5) = temparray(pass + 1, 5)
                temparray(pass, 6) = temparray(pass + 1, 6)
                temparray(pass + 1, 1) = temp1
                temparray(pass + 1, 2) = temp2
                temparray(pass + 1, 3) = temp3
                temparray(pass + 1, 4) = temp4
                temparray(pass + 1, 5) = temp5
                temparray(pass + 1, 6) = temp6
            End If
        Next pass
    Next ctr
    For n = 1 To x
        Write #3, temparray(n, 1), temparray(n, 2), temparray(n, 3), temparray(n, 4), temparray(n, 5), temparray(n, 6)
    Next n
    For q = 1 To x
        If temparray(q, 6) <= group1 Then
                a = a + 1
                groups1(a, 1) = temparray(q, 1)
                groups1(a, 2) = temparray(q, 2)
                groups1(a, 3) = temparray(q, 3)
                groups1(a, 4) = temparray(q, 4)
                groups1(a, 5) = temparray(q, 5)
                groups1(a, 6) = temparray(q, 6)
        ElseIf temparray(q, 6) <= group2 Then
                b = b + 1
                groups2(b, 1) = temparray(q, 1)
                groups2(b, 2) = temparray(q, 2)
                groups2(b, 3) = temparray(q, 3)
                groups2(b, 4) = temparray(q, 4)
                groups2(b, 5) = temparray(q, 5)
                groups2(b, 6) = temparray(q, 6)
        ElseIf temparray(q, 6) <= group3 Then
                c = c + 1
                groups3(c, 1) = temparray(q, 1)
                groups3(c, 2) = temparray(q, 2)
                groups3(c, 3) = temparray(q, 3)
                groups3(c, 4) = temparray(q, 4)
                groups3(c, 5) = temparray(q, 5)
                groups3(c, 6) = temparray(q, 6)
        ElseIf temparray(q, 6) <= group4 Then
                d = d + 1
                groups4(d, 1) = temparray(q, 1)
                groups4(d, 2) = temparray(q, 2)
                groups4(d, 3) = temparray(q, 3)
                groups4(d, 4) = temparray(q, 4)
                groups4(d, 5) = temparray(q, 5)
                groups4(d, 6) = temparray(q, 6)
        ElseIf temparray(q, 6) <= group5 Then
                e = e + 1
                groups5(e, 1) = temparray(q, 1)
                groups5(e, 2) = temparray(q, 2)
                groups5(e, 3) = temparray(q, 3)
                groups5(e, 4) = temparray(q, 4)
                groups5(e, 5) = temparray(q, 5)
                groups5(e, 6) = temparray(q, 6)
        Else
                g = g + 1
                groups6(g, 1) = temparray(q, 1)
                groups6(g, 2) = temparray(q, 2)
                groups6(g, 3) = temparray(q, 3)
                groups6(g, 4) = temparray(q, 4)
                groups6(g, 5) = temparray(q, 5)
                groups6(g, 6) = temparray(q, 6)
        End If
    Next q
picconfirm.Print "Age group 10 to"; group1
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For y = 1 To a
    If a = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); y; Tab(15); groups1(y, 1); Tab(25); groups1(y, 2); Tab(35); groups1(y, 3); Tab(42); groups1(y, 4); Tab(62); groups1(y, 5); Tab(77); groups1(y, 6)
Next y
picconfirm.Print
picconfirm.Print "Age group"; group1 + 1; "to "; group2
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For h = 1 To b
    If b = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); h; Tab(15); groups2(h, 1); Tab(25); groups2(h, 2); Tab(35); groups2(h, 3); Tab(42); groups2(h, 4); Tab(62); groups2(h, 5); Tab(77); groups2(h, 6)
Next h
picconfirm.Print
picconfirm.Print "Age group"; group2 + 1; "to "; group3
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For i = 1 To c
    If c = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); i; Tab(15); groups3(i, 1); Tab(25); groups3(i, 2); Tab(35); groups3(i, 3); Tab(42); groups3(i, 4); Tab(62); groups3(i, 5); Tab(77); groups3(i, 6)
Next i
picconfirm.Print
picconfirm.Print "Age group"; group3 + 1; "to "; group4
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For k = 1 To d
    If d = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); k; Tab(15); groups4(k, 1); Tab(25); groups4(k, 2); Tab(35); groups4(k, 3); Tab(42); groups4(k, 4); Tab(62); groups4(k, 5); Tab(77); groups4(k, 6)
Next k
picconfirm.Print
picconfirm.Print "Age group"; group4 + 1; "to "; group5
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For l = 1 To e
    If e = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); l; Tab(15); groups5(l, 1); Tab(25); groups5(l, 2); Tab(35); groups5(l, 3); Tab(42); groups5(l, 4); Tab(62); groups5(l, 5); Tab(77); groups5(l, 6)
Next l
picconfirm.Print
picconfirm.Print "Age group"; group5 + 1; "+"
picconfirm.Print "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
picconfirm.Print "age group"
For m = 1 To g
    If g = 0 Then
        picconfirm.Print
        picconfirm.Print
    End If
    picconfirm.Print Tab(2); m; Tab(15); groups6(m, 1); Tab(25); groups6(m, 2); Tab(35); groups6(m, 3); Tab(42); groups6(m, 4); Tab(62); groups6(m, 5); Tab(77); groups6(m, 6)
Next m

Print #4, "Age group 10 to"; group1
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For r = 1 To a
    Print #4, Tab(2); r; Tab(15); groups1(r, 1); Tab(25); groups1(r, 2); Tab(35); groups1(r, 3); Tab(42); groups1(r, 4); Tab(62); groups1(r, 5); Tab(77); groups1(r, 6)
Next r
Print #4, "Age group"; group1 + 1; "to "; group2
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For s = 1 To b
    Print #4, Tab(2); s; Tab(15); groups2(s, 1); Tab(25); groups2(s, 2); Tab(35); groups2(s, 3); Tab(42); groups2(s, 4); Tab(62); groups2(s, 5); Tab(77); groups2(s, 6)
Next s
Print #4, "Age group"; group2 + 1; "to "; group3
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For t = 1 To c
    Print #4, Tab(2); t; Tab(15); groups3(t, 1); Tab(25); groups3(t, 2); Tab(35); groups3(t, 3); Tab(42); groups3(t, 4); Tab(62); groups3(t, 5); Tab(77); groups3(t, 6)
Next t
Print #4, "Age group"; group3 + 1; "to "; group4
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For u = 1 To d
    Print #4, Tab(2); u; Tab(15); groups4(u, 1); Tab(25); groups4(u, 2); Tab(35); groups4(u, 3); Tab(42); groups4(u, 4); Tab(62); groups4(u, 5); Tab(77); groups4(u, 6)
Next u
Print #4, "Age group"; group4 + 1; "to "; group5
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For v = 1 To e
    Print #4, Tab(2); v; Tab(15); groups5(v, 1); Tab(25); groups5(v, 2); Tab(35); groups5(v, 3); Tab(42); groups5(v, 4); Tab(62); groups5(v, 5); Tab(77); groups5(v, 6)
Next v
Print #4, "Age group"; group5 + 1; "+"
Print #4, "Place in"; Tab(15); "Place"; Tab(25); "Time"; Tab(35); "#"; Tab(42); "Name"; Tab(62); "City"; Tab(77); "Age"
Print #4, "age group"
For w = 1 To g
    Print #4, Tab(2); w; Tab(15); groups6(w, 1); Tab(25); groups6(w, 2); Tab(35); groups6(w, 3); Tab(42); groups6(w, 4); Tab(62); groups6(w, 5); Tab(77); groups6(w, 6)
Next w
Close #3
Close #4
End Sub

Private Sub cmdsortcity_Click()
Dim ctr As Integer, temp1 As Single, temp2 As Single, temp3 As Single, temp4 As String
Dim temp5 As String, temp6 As Single, pass As Integer, n As Integer
Dim v As Integer
picconfirm.Cls
Open path & "\sort by city.txt" For Output As #4
Open path & "\city printable.txt" For Output As #3
ctr = 0
    For ctr = 1 To x - 1
        For pass = 1 To x - ctr
            If temparray(pass, 5) > temparray(pass + 1, 5) Then 'alphabetizes cities
                temp1 = temparray(pass, 1)
                temp2 = temparray(pass, 2)
                temp3 = temparray(pass, 3)
                temp4 = temparray(pass, 4)
                temp5 = temparray(pass, 5)
                temp6 = temparray(pass, 6)
                temparray(pass, 1) = temparray(pass + 1, 1)
                temparray(pass, 2) = temparray(pass + 1, 2)
                temparray(pass, 3) = temparray(pass + 1, 3)
                temparray(pass, 4) = temparray(pass + 1, 4)
                temparray(pass, 5) = temparray(pass + 1, 5)
                temparray(pass, 6) = temparray(pass + 1, 6)
                temparray(pass + 1, 1) = temp1
                temparray(pass + 1, 2) = temp2
                temparray(pass + 1, 3) = temp3
                temparray(pass + 1, 4) = temp4
                temparray(pass + 1, 5) = temp5
                temparray(pass + 1, 6) = temp6
            End If
        Next pass
    Next ctr
    For n = 1 To x
        Write #4, temparray(n, 1), temparray(n, 2), temparray(n, 3), temparray(n, 4), temparray(n, 5), temparray(n, 6)
    Next n
picconfirm.Print "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For y = 1 To x
    picconfirm.Print Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
Next y
Print #3, "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For v = 1 To x
    Print #3, Tab(2); temparray(v, 1); Tab(10); temparray(v, 2); Tab(20); temparray(v, 3); Tab(27); temparray(v, 4); Tab(47); temparray(v, 5); Tab(62); temparray(v, 6)
Next v
Close #3
Close #4
End Sub

Private Sub cmdsortname_Click()
Dim ctr As Integer, temp1 As Single, temp2 As Single, temp3 As Single, temp4 As String
Dim temp5 As String, temp6 As Single, pass As Integer, n As Integer
Dim v As Integer
picconfirm.Cls
Open path & "\sort by name.txt" For Output As #5
Open path & "\name printable.txt" For Output As #3
ctr = 0
    For ctr = 1 To x - 1
        For pass = 1 To x - ctr
            If temparray(pass, 4) > temparray(pass + 1, 4) Then 'sorts names
                temp1 = temparray(pass, 1)
                temp2 = temparray(pass, 2)
                temp3 = temparray(pass, 3)
                temp4 = temparray(pass, 4)
                temp5 = temparray(pass, 5)
                temp6 = temparray(pass, 6)
                temparray(pass, 1) = temparray(pass + 1, 1)
                temparray(pass, 2) = temparray(pass + 1, 2)
                temparray(pass, 3) = temparray(pass + 1, 3)
                temparray(pass, 4) = temparray(pass + 1, 4)
                temparray(pass, 5) = temparray(pass + 1, 5)
                temparray(pass, 6) = temparray(pass + 1, 6)
                temparray(pass + 1, 1) = temp1
                temparray(pass + 1, 2) = temp2
                temparray(pass + 1, 3) = temp3
                temparray(pass + 1, 4) = temp4
                temparray(pass + 1, 5) = temp5
                temparray(pass + 1, 6) = temp6
            End If
        Next pass
    Next ctr
    For n = 1 To x
        Write #5, temparray(n, 1), temparray(n, 2), temparray(n, 3), temparray(n, 4), temparray(n, 5), temparray(n, 6)
    Next n
    picconfirm.Print "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For y = 1 To x
    picconfirm.Print Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
Next y
Print #3, "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For v = 1 To x
    Print #3, Tab(2); temparray(v, 1); Tab(10); temparray(v, 2); Tab(20); temparray(v, 3); Tab(27); temparray(v, 4); Tab(47); temparray(v, 5); Tab(62); temparray(v, 6)
Next v
Close #3
Close #5
End Sub

Private Sub cmdsortnumber_Click()
Dim ctr As Integer, temp1 As Single, temp2 As Single, temp3 As Single, temp4 As String
Dim temp5 As String, temp6 As Single, pass As Integer, n As Integer
Dim v As Integer
Open path & "\sort by number.txt" For Output As #6
Open path & "\number printable.txt" For Output As #3
picconfirm.Cls
ctr = 0
    For ctr = 1 To x - 1
        For pass = 1 To x - ctr
            If temparray(pass, 3) > temparray(pass + 1, 3) Then 'sorts numbers
                temp1 = temparray(pass, 1)
                temp2 = temparray(pass, 2)
                temp3 = temparray(pass, 3)
                temp4 = temparray(pass, 4)
                temp5 = temparray(pass, 5)
                temp6 = temparray(pass, 6)
                temparray(pass, 1) = temparray(pass + 1, 1)
                temparray(pass, 2) = temparray(pass + 1, 2)
                temparray(pass, 3) = temparray(pass + 1, 3)
                temparray(pass, 4) = temparray(pass + 1, 4)
                temparray(pass, 5) = temparray(pass + 1, 5)
                temparray(pass, 6) = temparray(pass + 1, 6)
                temparray(pass + 1, 1) = temp1
                temparray(pass + 1, 2) = temp2
                temparray(pass + 1, 3) = temp3
                temparray(pass + 1, 4) = temp4
                temparray(pass + 1, 5) = temp5
                temparray(pass + 1, 6) = temp6
            End If
        Next pass
    Next ctr
    For n = 1 To x
        Write #6, temparray(n, 1), temparray(n, 2), temparray(n, 3), temparray(n, 4), temparray(n, 5), temparray(n, 6)
    Next n
    picconfirm.Print "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For y = 1 To x
    picconfirm.Print Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
Next y
Close #6
Print #3, "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For v = 1 To x
    Print #3, Tab(2); temparray(v, 1); Tab(10); temparray(v, 2); Tab(20); temparray(v, 3); Tab(27); temparray(v, 4); Tab(47); temparray(v, 5); Tab(62); temparray(v, 6)
Next v
Close #3
End Sub

Private Sub cmdsorttime_Click()
Dim ctr As Integer, temp1 As Single, temp2 As Single, temp3 As Single, temp4 As String
Dim temp5 As String, temp6 As Single, pass As Integer
Dim v As Integer
Open path & "\time printable.txt" For Output As #3
picconfirm.Cls
ctr = 0
    For ctr = 1 To x - 1
        For pass = 1 To x - ctr
            If temparray(pass, 2) > temparray(pass + 1, 2) Then 'sorts times
                temp1 = temparray(pass, 1)
                temp2 = temparray(pass, 2)
                temp3 = temparray(pass, 3)
                temp4 = temparray(pass, 4)
                temp5 = temparray(pass, 5)
                temp6 = temparray(pass, 6)
                temparray(pass, 1) = temparray(pass + 1, 1)
                temparray(pass, 2) = temparray(pass + 1, 2)
                temparray(pass, 3) = temparray(pass + 1, 3)
                temparray(pass, 4) = temparray(pass + 1, 4)
                temparray(pass, 5) = temparray(pass + 1, 5)
                temparray(pass, 6) = temparray(pass + 1, 6)
                temparray(pass + 1, 1) = temp1
                temparray(pass + 1, 2) = temp2
                temparray(pass + 1, 3) = temp3
                temparray(pass + 1, 4) = temp4
                temparray(pass + 1, 5) = temp5
                temparray(pass + 1, 6) = temp6
            End If
        Next pass
    Next ctr
    picconfirm.Print "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For y = 1 To x
    picconfirm.Print Tab(2); temparray(y, 1); Tab(10); temparray(y, 2); Tab(20); temparray(y, 3); Tab(27); temparray(y, 4); Tab(47); temparray(y, 5); Tab(62); temparray(y, 6)
Next y
Print #3, "Place"; Tab(10); "Time"; Tab(20); "#"; Tab(27); "Name"; Tab(47); "City"; Tab(62); "Age"
For v = 1 To x
    Print #3, Tab(2); temparray(v, 1); Tab(10); temparray(v, 2); Tab(20); temparray(v, 3); Tab(27); temparray(v, 4); Tab(47); temparray(v, 5); Tab(62); temparray(v, 6)
Next v
Close #3
End Sub

Private Sub Form_Load()
    path = "N:\CS130\handin\Belknap, Dan"
End Sub
