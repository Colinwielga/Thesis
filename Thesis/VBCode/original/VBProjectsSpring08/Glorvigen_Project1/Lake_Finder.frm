VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFF80&
   Caption         =   "Lake Finder"
   ClientHeight    =   7965
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   9870
   LinkTopic       =   "Form7"
   Picture         =   "Lake_Finder.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   9870
   Visible         =   0   'False
   Begin VB.CommandButton cmdinput 
      BackColor       =   &H0080FFFF&
      Caption         =   "Input Lakes To Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdfindlake 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtcounty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CommandButton cmddisplaycounty 
      BackColor       =   &H0080FFFF&
      Caption         =   "List Countys To Search From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdrank 
      BackColor       =   &H0080FFFF&
      Caption         =   "List Lake In Rank of Popularity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdsubmitlake 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtlake 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   5640
      Width           =   3135
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   4320
      ScaleHeight     =   4155
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H0080FFFF&
      Caption         =   "List Lakes In Alpahbet Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Please Enter a Lake to See which County it is in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   11
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Please Enter a County you wish To Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   4680
      Width           =   3975
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rank(1 To 100) As Integer
Dim lake(1 To 100) As String, county(1 To 100) As String
Dim fish(1 To 100) As String
Dim ctr As Integer
'Minnesota Fisher
'Lake Finder
'Eric Glorvigen
'Date= March 5
' the main purpose of the Lake Finder form is to search and sort arrays, the page shows differnt
'ways for file input and searching techniques



Private Sub cmddisplaycounty_Click()

'opens county text to diplay names of counties, and lists them in alphabetical
'order
    Dim county(1 To 100) As String
    Dim find As String
    Dim pass As Integer, pos As Integer, temp As String, c As Integer
    Dim ctrtwo As Integer
    
     picoutput.Cls
     
        Open App.Path & "\countys.txt" For Input As #1
                Do Until EOF(1)
                    ' Get the data from the file
                    ctrtwo = ctrtwo + 1
                    Input #1, county(ctrtwo)
                Loop
        Close #1
  
        For pass = 1 To ctrtwo - 1
            For pos = 1 To ctrtwo - pass
                If county(pos) > county(pos + 1) Then
                    temp = county(pos)
                    county(pos) = county(pos + 1)
                    county(pos + 1) = temp
                End If
            Next pos
        Next pass
        
          picoutput.Print "Countys"
          picoutput.Print "*********************"
          
            For c = 1 To ctrtwo
               picoutput.Print county(c)
            Next c
End Sub

Private Sub cmdexit_Click()
    'return to main page
        form1.Show
        Form7.Hide
End Sub

Private Sub cmdfind_Click()
'orders the lakes in alphabet order, from the array already inputted

    Dim pass As Integer, pos As Integer, temp As String, c As Integer
    Dim temptwo As String, tempthree As Integer
    
     picoutput.Cls
  
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If lake(pos) > lake(pos + 1) Then
                    temp = lake(pos)
                    lake(pos) = lake(pos + 1)
                    lake(pos + 1) = temp
                    temptwo = county(pos)
                    county(pos) = county(pos + 1)
                    county(pos + 1) = temptwo
                    tempthree = rank(pos)
                    rank(pos) = rank(pos + 1)
                    rank(pos + 1) = tempthree
                End If
            Next pos
        Next pass
        
          picoutput.Print "Lakes", Tab(30); "Countys"
          picoutput.Print "**********************************************"
          
            For c = 1 To ctr
                picoutput.Print lake(c), Tab(30); county(c)
            Next c
End Sub

Private Sub cmdfindlake_Click()
'finds the county that the lake, which the user typed in, is located

    Dim find As String
    Dim countyname As String
    Dim ctrtwo As Integer
    Dim lakename As String
    Dim k As Integer
    Dim n As Integer
    
    
        picoutput.Cls

        ctrtwo = 0
        find = txtcounty.Text
        
        For n = 1 To ctr
            If LCase(find) = LCase(lake(n)) Then
               picoutput.Print UCase(lake(n)); " is in "; UCase(county(n)); " county. See you there!"
               k = k + 1
            End If
        Next n
        
        If k = 0 Then
            picoutput.Print "Sorry, I have Not Fished That Lake."
            picoutput.Print "I do not know what county "; UCase(find); " is in."
        End If
        
    

End Sub



Private Sub cmdinput_Click()
        
        'this button inputs the data
        'to make sure it is pushed first I have the two sorting buttons
        'as visisble=false, so once the input is in the user can sort
             
            cmdfind.Visible = True
            cmdrank.Visible = True
            cmdfindlake.Visible = True
            cmdsubmitlake.Visible = True
            
        
        ctr = 0
        Open App.Path & "\lakes.txt" For Input As #1
            Do Until EOF(1)
                ctr = ctr + 1
                Input #1, lake(ctr), county(ctr), rank(ctr), fish(ctr)
            Loop
        Close #1
        
End Sub

Private Sub cmdleave_Click()
    'exit program
        End
End Sub

Private Sub cmdrank_Click()
' this button take the array from the input button and sorts by rank of popularity
'between each lake

    Dim pass As Integer, pos As Integer, temp As String, m As Integer
    Dim temptwo As String, tempthree As String
    
    picoutput.Cls
    picoutput.Print "Lakes", Tab(30); "Countys"
    picoutput.Print "****************************************************"
    
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If rank(pos) > rank(pos + 1) Then
                    temp = rank(pos)
                    rank(pos) = rank(pos + 1)
                    rank(pos + 1) = temp
                    temptwo = county(pos)
                    county(pos) = county(pos + 1)
                    county(pos + 1) = temptwo
                    tempthree = lake(pos)
                    lake(pos) = lake(pos + 1)
                    lake(pos + 1) = tempthree
                End If
            Next pos
        Next pass
           
            For m = 1 To ctr
                picoutput.Print lake(m), Tab(30); county(m)
            Next m
End Sub

Private Sub cmdsubmitlake_Click()
    'this button will show the user what lakes are in the county they are
    'searching
    
    Dim find As String
    Dim k As Integer
    Dim n As Integer
    
            picoutput.Cls
            picoutput.Print "The best lake(s) in "; find; " are:"
            picoutput.Print "***********************************************"
            
            find = txtlake.Text
            For n = 1 To ctr
                If LCase(find) = LCase(county(n)) Then
                    picoutput.Print lake(n)
                    k = k + 1
                End If
            Next n
        
        If k = 0 Then
            picoutput.Print "Sorry, I haven't fished that county."
        End If
          
End Sub



