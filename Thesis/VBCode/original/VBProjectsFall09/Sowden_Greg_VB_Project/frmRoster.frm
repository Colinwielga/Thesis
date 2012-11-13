VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H000000FF&
   Caption         =   "Roster Maker"
   ClientHeight    =   11370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmRoster.frx":0000
   ScaleHeight     =   11370
   ScaleWidth      =   17355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSort2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort By Fastest Player"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmRoster.frx":5C620
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sort the players by their speed "
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search for Player"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmRoster.frx":5D1E3
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Search for a player in the list"
      Top             =   8280
      Width           =   2415
   End
   Begin VB.CommandButton cmdLearn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Learn More about the Offense"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmRoster.frx":5DDA6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Learn about each Offensive position"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdForm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Your Position!"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   14280
      Picture         =   "frmRoster.frx":5E969
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find what position you would be best at"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdPosPic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Offensive Positions"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmRoster.frx":5F52C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Find where every position on the field is"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort Positions in File!"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      Picture         =   "frmRoster.frx":600EF
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sort the players into positions"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   3840
      ScaleHeight     =   8115
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   2760
      Width           =   9495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Offensive Positions"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   9495
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Football: The Offense
'   Roster
'   Greg Sowden
'   10/9/09
'   This form is the "home page" for the project.  It has buttons for the user to sort the data from the file in several ways, and it
    'is the jumping off point for the information page, the search button, and the user input data function.  The picture box in the
    'middle is the display for much of the information.

'   This subcommand prompts an input box so the user can type in a name they would like to search
'   If the name is found it will display in the picture box the name and all statistics of that player
'   If the name is not found the picture box will say the player is not in the data file.

'   first dim the integers that will be used in the data file
Private Sub cmdFind_Click()
    Dim title As String, i As Integer
    Dim found As Boolean
'   since we are searching we need a boolean - found
    
    picResults.Cls '    clear the picture box
    
    title = InputBox("Enter the name of the player you are searching for.", "Find Player")
'   set the variable equal to the input so it can be used below
        
    found = False
    
    i = 0
    
'   the do while loop will search each entry in the data file for a mach with the name from the input box
    Do While (Not found) And (i < ctr)
        i = i + 1
            If title = names(i) Then
                found = True
            End If
    Loop
    
    If (Not found) Then '   if the player is not found, the display will say so
        picResults.Print title; "You did not input a player that is in the data file"
        Else    '   if the player is found, the picresults box will display name and all statistics including this heading
                picResults.Print Tab(0); "Name";
                picResults.Print Tab(22); "Height (in)";
                picResults.Print Tab(35); "Weight";
                picResults.Print Tab(45); "Forty Time";
                picResults.Print Tab(58); "Throwing Accuracy"
                
                picResults.Print "***********************************************************************************"
                
                picResults.Print Tab(0); names(i);
                picResults.Print Tab(22); heights(i);
                picResults.Print Tab(35); weights(i);
                picResults.Print Tab(45); forty(i);
                picResults.Print Tab(58); AccScore(i)
        End If
        
    
End Sub
'   this brings up the inputdata form
Private Sub cmdForm_Click()
    frmRoster.Hide
    frmInputData.Show
    
End Sub
'   this brings up the information page allowing the user to read about each position
Private Sub cmdLearn_Click()
    frmRoster.Hide
    frmLearn.Show
End Sub
'   this loads a picture of the offense in the picture box
Private Sub cmdPosPic_Click()
    picResults.Picture = LoadPicture(App.Path & "\Offense.jpg")
End Sub



Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSort_Click()
'   this button will sort the players into their individual positions based on criteria of their data.
'   an exhaustive search will be used
    Dim found As Boolean
    Dim i As Integer, K As Integer, J As Integer, L As Integer, M As Integer
'   tell the program the desired information has not yet been read
    found = False
    
'   set the background of the picbox to a white background to clear the possible previous picture
    picResults.Picture = LoadPicture(App.Path & "\WHITE%20WALL.jpg")
'   picresults.cls will clear any text in the picResults box
    picResults.Cls
    
'   again, print a header
    picResults.Print Tab(0); "Name";
                picResults.Print Tab(22); "Height (in)";
                picResults.Print Tab(35); "Weight";
                picResults.Print Tab(45); "Forty Time";
                picResults.Print Tab(58); "Throwing Accuracy"
                
                picResults.Print "***********************************************************************************"
    
'   each section recieves a header for the position
    picResults.Print "OFFENSIVE LINEMEN"
    picResults.Print "*********************************************************************************************"
    
'   begin the exhaustive search for each player who falls under the category.
    For i = 1 To ctr
        
        If forty(i) > 5.3 And weights(i) > 240 Then
            found = True
            
            picResults.Print Tab(0); names(i);
            picResults.Print Tab(25); heights(i);
            picResults.Print Tab(35); weights(i);
            picResults.Print Tab(45); forty(i);
            picResults.Print Tab(58); AccScore(i)
        End If
    Next i
    
    If (Not found) Then
        picResults.Print "There are no viable offensive linemen"
'   repeat the steps for all of the other positions.
    End If
    
    picResults.Print ""
    picResults.Print "FULLBACKS"
    picResults.Print "***********************************************************************************************"
    
    For K = 1 To ctr
    If forty(K) <= 5.3 And weights(K) <= 240 And weights(K) >= 210 Then
            found = True
            
            picResults.Print Tab(0); names(K);
            picResults.Print Tab(25); heights(K);
            picResults.Print Tab(35); weights(K);
            picResults.Print Tab(45); forty(K);
            picResults.Print Tab(58); AccScore(K)
        End If
    Next K
    
    If (Not found) Then
        picResults.Print "There are no viable fullbacks"

    End If
    picResults.Print ""
    picResults.Print "RUNNING BACKS"
    picResults.Print "***********************************************************************************************"
    
    For J = 1 To ctr
    If forty(J) < 5 And weights(J) < 240 Then
            found = True
            
            picResults.Print Tab(0); names(J);
            picResults.Print Tab(25); heights(J);
            picResults.Print Tab(35); weights(J);
            picResults.Print Tab(45); forty(J);
            picResults.Print Tab(58); AccScore(J)
        End If
    Next J
    
    If (Not found) Then
        picResults.Print "There are no viable Running Backs"

    End If
    
    picResults.Print ""
    picResults.Print "QUARTERBACKS"
    picResults.Print "***********************************************************************************************"
    
    For L = 1 To ctr
    If forty(L) <= 5.3 And weights(L) <= 230 And AccScore(L) >= 70 And heights(L) >= 68 Then
            found = True
            
            picResults.Print Tab(0); names(L);
            picResults.Print Tab(25); heights(L);
            picResults.Print Tab(35); weights(L);
            picResults.Print Tab(45); forty(L);
            picResults.Print Tab(58); AccScore(L)
        End If
    Next L
    
    If (Not found) Then
        picResults.Print "There are no viable Quarterbacks"
    End If
    
    picResults.Print ""
    picResults.Print "WIDE RECIEVERS"
    picResults.Print "***********************************************************************************************"
    
    For M = 1 To ctr
    If forty(M) <= 5 And weights(M) <= 230 And heights(M) >= 72 Then
            found = True
            
            picResults.Print Tab(0); names(M);
            picResults.Print Tab(25); heights(M);
            picResults.Print Tab(35); weights(M);
            picResults.Print Tab(45); forty(M);
            picResults.Print Tab(58); AccScore(M)
        End If
    Next M

'   in case none are found, tell the user there are no candidates
    If (Not found) Then
        picResults.Print "There are no viable Wide Recievers"
End If
    
End Sub



Private Sub cmdSort2_Click()
'   this command button will sort the players based on their forty yard dash time
'   they will be sorted in descending order starting with the fastest.
    Dim pass As Integer, pos As Integer, i As Integer
    Dim tempForty As Single, tempName As String, tempHeight As Single, tempWeight As Integer, tempAcc As Single
'   dim temporary variables to take the place of the actual variables.
    picResults.Cls
    picResults.Picture = LoadPicture(App.Path & "\WHITE%20WALL.jpg")
'   use a bubble sort to find the order
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If forty(pos) > forty(pos + 1) Then
                tempForty = forty(pos)
                forty(pos) = forty(pos + 1)
                forty(pos + 1) = tempForty
'   each variable needs to be bubble sorted to keep it with the variable being arranged
                tempName = names(pos)
                names(pos) = names(pos + 1)
                names(pos + 1) = tempName
                
                tempHeight = heights(pos)
                heights(pos) = heights(pos + 1)
                heights(pos + 1) = tempHeight
                
                tempWeight = weights(pos)
                weights(pos) = weights(pos + 1)
                weights(pos + 1) = tempWeight
                
                tempAcc = AccScore(pos)
                AccScore(pos) = AccScore(pos + 1)
                AccScore(pos + 1) = tempAcc
                
            End If
        Next pos
    Next pass
'   now print header and results
                picResults.Print Tab(0); "Name";
                picResults.Print Tab(22); "Height (in)";
                picResults.Print Tab(35); "Weight";
                picResults.Print Tab(45); "Forty Time";
                picResults.Print Tab(58); "Throwing Accuracy"
    
    
    For i = 1 To ctr
'   it has to be i = 1 all the way to ctr to display each find
            picResults.Print Tab(0); names(i);
            picResults.Print Tab(25); heights(i);
            picResults.Print Tab(35); weights(i);
            picResults.Print Tab(45); forty(i);
            picResults.Print Tab(58); AccScore(i)
    Next i

End Sub


Private Sub Form_Load()
'   this rearranges the forms so that the program will begin with the start form that I actually created later than the Roster form
    frmRoster.Hide
    frmStart.Show
    
End Sub
