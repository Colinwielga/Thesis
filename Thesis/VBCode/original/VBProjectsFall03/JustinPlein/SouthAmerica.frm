VERSION 5.00
Begin VB.Form SouthAmerica 
   BackColor       =   &H00FFFF80&
   Caption         =   "SouthAmerica"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Countries of South America"
   ScaleHeight     =   8130
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton people 
      Caption         =   "Sort by Population"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Quitbox 
      Caption         =   "Quit"
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton freedom 
      Caption         =   "Independance of Countries"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton alpha 
      Caption         =   "All Countries Alphabetical"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox solve 
      BackColor       =   &H000000FF&
      FillColor       =   &H0000FFFF&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   4440
      Width           =   5295
   End
   Begin VB.PictureBox flag 
      BackColor       =   &H00FFFF80&
      Height          =   1695
      Left            =   3000
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Created By Justin Plein"
      Height          =   495
      Left            =   9960
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin VB.Line Line6 
      X1              =   8160
      X2              =   8400
      Y1              =   1680
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   8400
      X2              =   9000
      Y1              =   5520
      Y2              =   5880
   End
   Begin VB.Label frenchguiana 
      BackColor       =   &H00FFFF80&
      Caption         =   "French Guiana"
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line7 
      X1              =   8400
      X2              =   9240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label suriname 
      BackColor       =   &H00FFFF80&
      Caption         =   "Suriname"
      Height          =   255
      Left            =   8400
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Line Line5 
      X1              =   7920
      X2              =   7920
      Y1              =   1560
      Y2              =   1080
   End
   Begin VB.Label guyana 
      BackColor       =   &H00FFFF80&
      Caption         =   "Guyana"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Label venezuela 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Venezuela"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label colombia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Colombia"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label ecuador 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ecuador"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   6120
      X2              =   5520
      Y1              =   2520
      Y2              =   1920
   End
   Begin VB.Label peru 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Peru"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label bolivia 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bolivia"
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label paraguay 
      BackColor       =   &H00FFFF80&
      Caption         =   "Paraguay"
      Height          =   255
      Left            =   9480
      TabIndex        =   9
      Top             =   4800
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   8160
      X2              =   9480
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Label uruguay 
      BackColor       =   &H00FFFF80&
      Caption         =   "Uruguay"
      Height          =   255
      Left            =   9000
      TabIndex        =   8
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label chile 
      BackColor       =   &H00FFFF80&
      Caption         =   "Chile"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   5040
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   6720
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label argentina 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Argentina"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label brazil 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Brazil"
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   6900
      Left            =   5640
      Picture         =   "SouthAmerica.frx":0000
      Top             =   840
      Width           =   4890
   End
End
Attribute VB_Name = "SouthAmerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : South America (Justin Plein's VB Project.vbp)
'Form Name : South America (South America.frm)
'Author: Justin Plein
'Date Written: October 24, 2003
'Purpose of Form: To let someone pick a South American country
                  'and let them see the countries flag, independence
                  'date, and population of the country.
                  'also, they can arrange the countries
                  'alphabetically, by independence date,
                  'and by population

Option Explicit
Dim country(1 To 13) As String
Dim independence(1 To 13) As String
Dim population(1 To 13) As Long
Dim i As Integer, pass As Integer, n As Integer
Dim temp As String, tempa As String
Dim tempb As Long
Dim Found As Boolean
Dim PATH As String

Private Sub alpha_Click()
solve.Cls
    'clear the picture box of old information
flag.Picture = LoadPicture(PATH & "Flags\blank.gif")
    'clears any flag that was in teh flag picturebox
solve.Print "Country", , "Independence Date", "Population"
n = 13

For pass = 1 To n - 1
    For i = 1 To n - pass
        If country(i) > country(i + 1) Then
            'swap country(I) and country(I+1)
                temp = country(i)
            country(i) = country(i + 1)
            country(i + 1) = temp
            'swaps independence(i) and independence(i+1)
                tempa = independence(i)
            independence(i) = independence(i + 1)
            independence(i + 1) = tempa
            'swaps population(i) and population(i+1)
                tempb = population(i)
            population(i) = population(i + 1)
            population(i + 1) = tempb
        End If
    Next i
Next pass

For i = 1 To n
    'prints the results of countries alpahbetically
    solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
Next i

End Sub

Private Sub argentina_Click()
Dim name12 As String
solve.Cls
    'clears old info in picture box
name12 = argentina
Found = False

   i = 0

Do Until Found Or i >= 13
    i = i + 1
    'searches for argentina
    If name12 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
    flag.Picture = LoadPicture(PATH & "Flags\argentina.gif")
End Sub

Private Sub bolivia_Click()
Dim name9 As String
solve.Cls
name9 = bolivia
Found = False

   i = 0

Do Until Found Or i >= 13
    i = i + 1
    'searches for bolivia
    If name9 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
    flag.Picture = LoadPicture(PATH & "Flags\bolivia.gif")
End Sub

Private Sub brazil_Click()
Dim name As String
solve.Cls
name = brazil
Found = False

   i = 0
    'searches for brazil
Do Until Found Or i >= 13
    i = i + 1
    If name = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
'prints information on brazil
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints the flag of brazil
    flag.Picture = LoadPicture(PATH & "Flags\brazil.gif")
End Sub

Private Sub chile_Click()
Dim name13 As String
solve.Cls
name13 = chile
Found = False

   i = 0
'searches for chile
Do Until Found Or i >= 13
    i = i + 1
    If name13 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints chile information
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints the flag of chile
    flag.Picture = LoadPicture(PATH & "Flags\chile.gif")
End Sub

Private Sub colombia_Click()
Dim name5 As String
solve.Cls
name5 = colombia
Found = False

   i = 0
'searches for colombia
Do Until Found Or i >= 13
    i = i + 1
    If name5 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of colombia
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints the flag of colombia
    flag.Picture = LoadPicture(PATH & "Flags\colombia.gif")
End Sub

Private Sub ecuador_Click()
Dim name6 As String
solve.Cls
name6 = ecuador
Found = False

   i = 0
'searches for ecuador
Do Until Found Or i >= 13
    i = i + 1
    If name6 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of ecuador
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of ecuador
    flag.Picture = LoadPicture(PATH & "Flags\ecuador.gif")
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Misc\"
'opens notebook info of South America for use in project
Open PATH & "southamerica.txt" For Input As #1

For i = 1 To 13
        'loads the array in catagories
    Input #1, country(i), independence(i), population(i)
Next i
Close #1
'tells the user info on using the project
MsgBox "Click on a country's name to see individual information"
End Sub


Private Sub freedom_Click()
solve.Cls
flag.Picture = LoadPicture(PATH & "Flags\blank.gif")
solve.Print "Country", , "Independence Date", "Population"
n = 13

For pass = 1 To n - 1
    For i = 1 To n - pass
        If independence(i) > independence(i + 1) Then
            'swap independence(I) and independence(I+1)
                temp = country(i)
            country(i) = country(i + 1)
            country(i + 1) = temp
            ' swaps country(i) and country(i+1)
                tempa = independence(i)
            independence(i) = independence(i + 1)
            independence(i + 1) = tempa
            'swaps population(i) and population(i+1)
                tempb = population(i)
            population(i) = population(i + 1)
            population(i + 1) = tempb
        End If
    Next i
Next pass

For i = 1 To n
    solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
Next i
End Sub





Private Sub frenchguiana_Click()
Dim name1 As String
solve.Cls
flag.Cls
name1 = frenchguiana
Found = False

   i = 0
'searches for French Guiana
Do Until Found Or i >= 13
    i = i + 1
    If name1 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of French Guiana
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of French Guiana
    flag.Picture = LoadPicture(PATH & "Flags\frenchguiana.gif")
End Sub

Private Sub guyana_Click()
Dim name3 As String
solve.Cls
name3 = guyana
Found = False

   i = 0
'searches for Guyana
Do Until Found Or i >= 13
    i = i + 1
    If name3 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of Guyana
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of Guyana
    flag.Picture = LoadPicture(PATH & "Flags\guyana.gif")
End Sub


Private Sub paraguay_Click()
Dim name10 As String
solve.Cls
name10 = paraguay
Found = False

   i = 0
'searches for paraguay
Do Until Found Or i >= 13
    i = i + 1
    If name10 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info on paraguay
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of paraguay
    flag.Picture = LoadPicture(PATH & "Flags\paraguay.gif")
End Sub

Private Sub people_Click()
solve.Cls
flag.Picture = LoadPicture(PATH & "Flags\blank.gif")
solve.Print "Country", , "Independence Date", "Population"
n = 13

For pass = 1 To n - 1
    For i = 1 To n - pass
        If population(i) > population(i + 1) Then
            'swap population(I) and population(I+1)
                temp = country(i)
            country(i) = country(i + 1)
            country(i + 1) = temp
            'swap country(i) and country(i+1)
                tempa = independence(i)
            independence(i) = independence(i + 1)
            independence(i + 1) = tempa
            'swap population(i) and population(i+1)
                tempb = population(i)
            population(i) = population(i + 1)
            population(i + 1) = tempb
        End If
    Next i
Next pass

For i = 1 To n
        'prints info by independence of country
    solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
Next i
    
End Sub

Private Sub peru_Click()
Dim name7 As String
solve.Cls
name7 = peru
Found = False

   i = 0
'searches for peru
Do Until Found Or i >= 13
    i = i + 1
    If name7 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of peru
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of peru
    flag.Picture = LoadPicture(PATH & "Flags\peru.gif")
End Sub



Private Sub Quitbox_Click()
    'Ends the project
End
End Sub

Private Sub suriname_Click()
Dim name2 As String
solve.Cls
name2 = suriname
Found = False

   i = 0
'Searches for suriname
Do Until Found Or i >= 13
    i = i + 1
    If name2 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of suriname
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of suriname
    flag.Picture = LoadPicture(PATH & "Flags\suriname.gif")
End Sub

Private Sub uruguay_Click()
Dim name11 As String
solve.Cls
name11 = uruguay
Found = False

   i = 0
'searches for uruguay
Do Until Found Or i >= 13
    i = i + 1
    If name11 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of uruguay
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag or uruguay
    flag.Picture = LoadPicture(PATH & "Flags\uruguay.gif")
End Sub

Private Sub venezuela_Click()
Dim name4 As String
solve.Cls
name4 = venezuela
Found = False

   i = 0
'searches for Venezuela
Do Until Found Or i >= 13
    i = i + 1
    If name4 = country(i) Then
        Found = True
    End If
    
Loop

If Found Then
            'prints info of Venezuela
       solve.Print "Country", , "Independence Date", "Population"
       solve.Print i; country(i); Tab(23); , independence(i); Tab(46); , population(i)
    Else
      solve.Print "Country not found"
End If
        'prints flag of Venezuela
    flag.Picture = LoadPicture(PATH & "Flags\venezuela.gif")
End Sub
