VERSION 5.00
Begin VB.Form frmSex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton cmdTC 
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton cmdRoch 
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton cmdDuluth 
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox picResults 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   9975
      Left            =   7560
      ScaleHeight     =   9975
      ScaleWidth      =   5775
      TabIndex        =   4
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "Sort Duluth"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort St. Joe"
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "Sort Cities"
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort Rochester"
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "St. Joseph"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Twin Cities"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Rochester"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Duluth"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image frmSexoff 
      Height          =   7860
      Left            =   0
      Picture         =   "frmSex.frx":0000
      Top             =   0
      Width           =   7350
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Return 
         Caption         =   "Return To Menu"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmSex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Age(1 To 100) As Integer, CTR As Integer
 'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009
'Objective: to provide a easy to use map that allows the user to click a city and
' have a list of sex offenders displayed

Private Sub cmdDuluth_Click()
Dim Duluth(1 To 100) As String

 'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Duluth.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Duluth(CTR), Age(CTR)
        picResults.Print Duluth(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in Duluth is..."; CTR
    Close #1
End Sub
'offer the option to quit
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRoch_Click()
Dim Roch(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Rochester.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Roch(CTR), Age(CTR)
        picResults.Print Roch(CTR); Tab(30); Age(CTR)
        
 'repeat until end of file
    Loop

'print the number of Roch sex offenders

    picResults.Print " "
    picResults.Print "The total number of sex offenders in Rochester is..."; CTR
     Close #1
End Sub

Private Sub cmdTC_Click()
Dim Cities(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Cities.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Cities(CTR), Age(CTR)
        picResults.Print Cities(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in the Twin Cities is..."; CTR
     Close #1
End Sub

Private Sub Command1_Click()
Dim StJoe(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\st.joe.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, StJoe(CTR), Age(CTR)
        picResults.Print StJoe(CTR); Tab(30); Age(CTR)
        
        
    Loop
    picResults.Print " "
    picResults.Print "The total number of sex offenders in St. Joseph is..."; CTR
     Close #1
End Sub





Private Sub Command2_Click()
Dim Duluth(1 To 100) As String

 'initialize ctr to zero, to be used for position in the array
    CTR = 0
   picResults.Cls
    'Prepare the file to be read
    Open App.Path & "\Duluth.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Duluth(CTR), Age(CTR)
    
    Loop
    
  Dim pass As Integer, pos As Integer, j As Integer
Dim tempName As String, tempAge As Single
Dim i As Integer
'sort the names alphabetically
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Duluth(pos) > Duluth(pos + 1) Then
            tempName = Duluth(pos)
            Duluth(pos) = Duluth(pos + 1)
            Duluth(pos + 1) = tempName
            tempAge = Age(pos)
            Age(pos) = Age(pos + 1)
            Age(pos + 1) = tempAge
        End If
    Next pos
Next pass
    For i = 1 To CTR

'print the results alphabetically

             picResults.Print Duluth(i); Tab(30); Age(i)
    Next i
    picResults.Print " "
    picResults.Print "The total number of sex offenders in Duluth is..."; CTR
    Close #1
End Sub

Private Sub Command3_Click()
 Dim pass As Integer, pos As Integer, j As Integer
Dim tempName As String, tempAge As Single
Dim f As Integer

 picResults.Cls

'sort the names alphabetically

Dim StJoe(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
       'Prepare the file to be read
    Open App.Path & "\st.joe.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, StJoe(CTR), Age(CTR)
              
        
    Loop
  
'sort alphabetically

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If StJoe(pos) > StJoe(pos + 1) Then
            tempName = StJoe(pos)
            StJoe(pos) = StJoe(pos + 1)
            StJoe(pos + 1) = tempName
            tempAge = Age(pos)
            Age(pos) = Age(pos + 1)
            Age(pos + 1) = tempAge
        End If
    Next pos
Next pass
    For f = 1 To CTR

'print alphabetical list

             picResults.Print StJoe(f); Tab(30); Age(f)
                      
      
    Next f
    picResults.Print " "
    picResults.Print "The total number of sex offenders in St. Joe is..."; CTR
    Close #1

End Sub

Private Sub Command4_Click()
 Dim pass As Integer, pos As Integer, j As Integer
Dim tempName As String, tempAge As Single
Dim h As Integer

 picResults.Cls
'sort the names

Dim Cities(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
       'Prepare the file to be read
    Open App.Path & "\Cities.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Cities(CTR), Age(CTR)
              
        
    Loop
  
'sort alphabetically

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Cities(pos) > Cities(pos + 1) Then
            tempName = Cities(pos)
            Cities(pos) = Cities(pos + 1)
            Cities(pos + 1) = tempName
            tempAge = Age(pos)
            Age(pos) = Age(pos + 1)
            Age(pos + 1) = tempAge
        End If
    Next pos
Next pass
    For h = 1 To CTR

'print list alphabetically

             picResults.Print Cities(h); Tab(30); Age(h)
                      
      
    Next h
    picResults.Print " "
    picResults.Print "The total number of sex offenders in the Twin Cities is..."; CTR
    Close #1

End Sub

Private Sub Command5_Click()
 Dim pass As Integer, pos As Integer, j As Integer
Dim tempName As String, tempAge As Single
Dim a As Integer

 picResults.Cls
'sort the names

Dim Roch(1 To 100) As String
'initialize ctr to zero, to be used for position in the array
    CTR = 0
       'Prepare the file to be read
    Open App.Path & "\Rochester.txt" For Input As #1
    
    'print the header info
    picResults.Print "Name of Sex Offender"; Tab(30); "Years of Age"
    picResults.Print "***********************************************"
    picResults.Print " "
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Roch(CTR), Age(CTR)
        
        
    Loop
  
'sort alphabetically

For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Roch(pos) > Roch(pos + 1) Then
            tempName = Roch(pos)
            Roch(pos) = Roch(pos + 1)
            Roch(pos + 1) = tempName
            tempAge = Age(pos)
            Age(pos) = Age(pos + 1)
            Age(pos + 1) = tempAge
        End If
    Next pos
Next pass
    For a = 1 To CTR
    
'print results alphabetically

             picResults.Print Roch(a); Tab(30); Age(a)
                      
      
    Next a
    picResults.Print " "
    picResults.Print "The total number of sex offenders in Rochester is..."; CTR
    Close #1

End Sub


'offer user option to quit

Private Sub quit_Click()
End
End Sub

Private Sub return_Click()

'give option to switch forms

frmSex.Hide
frmHome.Show
End Sub

