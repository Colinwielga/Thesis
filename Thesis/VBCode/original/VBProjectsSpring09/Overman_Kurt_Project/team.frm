VERSION 5.00
Begin VB.Form team 
   BackColor       =   &H000000FF&
   Caption         =   "roster"
   ClientHeight    =   8520
   ClientLeft      =   645
   ClientTop       =   1035
   ClientWidth     =   13695
   LinkTopic       =   "Form3"
   Picture         =   "team.frx":0000
   ScaleHeight     =   7573.333
   ScaleMode       =   0  'User
   ScaleWidth      =   14330.99
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton clr 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton number 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sort By Numbers"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton alpha 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Alphabetically"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton find 
      BackColor       =   &H00FF8080&
      Caption         =   "Click For Roster"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H008080FF&
      Height          =   8535
      Left            =   7320
      ScaleHeight     =   8475
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Twins Roster"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   7335
   End
End
Attribute VB_Name = "team"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dims variables
Dim Names(1 To 60) As String, CTR As Integer, numbers(1 To 60) As String, heights(1 To 60) As String, weight(1 To 60) As String, birth(1 To 60) As String, temp As TextBox

' alphabatizes roster text
Private Sub alpha_Click()
picResults.Cls
Dim pass As Integer, Pos As Integer, j As Integer
Dim Tempnum As String, Tempname As String, Tempheight As String, Tempweight As String, Tempbirth As String

For pass = 1 To CTR - 1
    For Pos = 1 To CTR - pass
        
        If Names(Pos) > Names(Pos + 1) Then
            
            Tempname = Names(Pos)
            Tempnum = numbers(Pos)
            Tempheight = heights(Pos)
            Tempweight = weight(Pos)
            Tempbirth = birth(Pos)
       
       Names(Pos) = Names(Pos + 1)
        Names(Pos + 1) = Tempname
       
       numbers(Pos) = numbers(Pos + 1)
        numbers(Pos + 1) = Tempnum
       
       heights(Pos) = heights(Pos + 1)
        heights(Pos + 1) = Tempheight
        
       weight(Pos) = weight(Pos + 1)
        weight(Pos + 1) = Tempweight
       
       birth(Pos) = birth(Pos + 1)
        birth(Pos + 1) = Tempbirth
        
       
 
       End If
    Next Pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "Name", Tab(30); "Number", "Height", "Weight", "birth"
    picResults.Print "**********************************"
    
'then print the list
    For j = 1 To CTR
             picResults.Print Names(j), Tab(30); numbers(j), heights(j), weight(j), birth(j)
       
    Next j

End Sub
'clears picResults
Private Sub clr_Click()
picResults.Cls
End Sub
'hides team form and then shows main form
Private Sub Command1_Click()
team.Hide
main.Show
End Sub
'opens roster text and then prints it
Private Sub find_Click()
CTR = 0
   
    
    Open App.Path & "\roster.txt" For Input As #1
    
    
    picResults.Print "Name", Tab(30); "Number", "Height", "Weight", "Birthday";
    picResults.Print Tab(1); "*******************************************************************************************************"
    picResults.Print
    Do While Not EOF(1)
       
        CTR = CTR + 1
        
       Input #1, Names(CTR), numbers(CTR), heights(CTR), weight(CTR), birth(CTR)
      
        
        picResults.Print Names(CTR), Tab(30); numbers(CTR), heights(CTR), weight(CTR), birth(CTR)
    Loop
    
   
    picResults.Print "*******************************************************************************************************************************"
    picResults.Print
    picResults.Print
    
    Close #1
End Sub
'takes roster text and then organizes it by numbers
Private Sub number_Click()
picResults.Cls
Dim pass As Integer, Pos As Integer, j As Integer
Dim Tempnum As String, Tempname As String, Tempheight As String, Tempweight As String, Tempbirth As String

For pass = 1 To CTR - 1
    For Pos = 1 To CTR - pass
        
        If numbers(Pos) > numbers(Pos + 1) Then
            
            Tempnum = numbers(Pos)
            Tempname = Names(Pos)
            Tempheight = heights(Pos)
            Tempweight = weight(Pos)
            Tempbirth = birth(Pos)
       
       
       Names(Pos) = Names(Pos + 1)
        Names(Pos + 1) = Tempname
       
       numbers(Pos) = numbers(Pos + 1)
        numbers(Pos + 1) = Tempnum
        
       heights(Pos) = heights(Pos + 1)
        heights(Pos + 1) = Tempheight
        
       weight(Pos) = weight(Pos + 1)
        weight(Pos + 1) = Tempweight
       
       birth(Pos) = birth(Pos + 1)
        birth(Pos + 1) = Tempbirth
        
       
 
       End If
    Next Pos
Next pass
 
'print the list
'first print the header info
    picResults.Print "Name", Tab(30); "Number", "Height", "Weight", "Birthday"
    picResults.Print "**************************************************************************************************************************************************************************"
    
'then print the list
    For j = 1 To CTR
             
             picResults.Print Names(j), Tab(30); numbers(j), heights(j), weight(j), birth(j)
       
    Next j
End Sub
