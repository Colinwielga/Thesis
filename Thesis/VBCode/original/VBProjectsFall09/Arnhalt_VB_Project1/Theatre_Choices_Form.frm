VERSION 5.00
Begin VB.Form frmTheatre 
   BackColor       =   &H00000000&
   Caption         =   "London Theatre"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18870
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Copperplate Gothic Light"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11880
   ScaleWidth      =   18870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoToAttractions 
      BackColor       =   &H0080FFFF&
      Caption         =   "Return to Popular Attractions Page"
      Height          =   975
      Left            =   6668
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10680
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   10508
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10680
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H0080FFFF&
      Caption         =   "Return to the Home Page"
      Height          =   975
      Left            =   8588
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10680
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7568
      ScaleHeight     =   975
      ScaleWidth      =   4695
      TabIndex        =   12
      Top             =   7920
      Width           =   4695
   End
   Begin VB.CommandButton cmdAttend 
      BackColor       =   &H0080FFFF&
      Caption         =   "I would like to attend one of these productions."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5288
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   8295
   End
   Begin VB.CommandButton cmdInfoRockYou 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   15908
      Picture         =   "Theatre_Choices_Form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoLesMis 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   13388
      Picture         =   "Theatre_Choices_Form.frx":10C2B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoPhantom 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   10868
      Picture         =   "Theatre_Choices_Form.frx":200A1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoWicked 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8348
      Picture         =   "Theatre_Choices_Form.frx":2E91E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoMammaMia 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5828
      Picture         =   "Theatre_Choices_Form.frx":3A40C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoBillyElliot 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3308
      Picture         =   "Theatre_Choices_Form.frx":45C09
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfoLionKing 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   788
      Picture         =   "Theatre_Choices_Form.frx":51114
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoadData 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Load Data"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8228
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtResults 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   4328
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   6600
      Width           =   10215
   End
   Begin VB.Label lblLearnInstructions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the picture of the musical you would like to learn more about."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6008
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label lblTheatre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Few of London's Most Popular Musicals"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   4028
      TabIndex        =   0
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "frmTheatre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London
'Form Name: Theatre
'Author: Heather Arnhalt
'Date Written: October 16, 2009
'Form Objective: To allow the user to click on the picture of a musical in order to display general information about that production
'including a synopsis of the production, the average ticket price (this is not what is used to calculate the cost, there are specific amounts
'associated with the type of seating the day of the week that are read into arrays and used to calculate the cost of a party attending the theatre.
'The user can also click the corresponding button to the production they might like to attend in order to find out approximately how much
'it would cost depending on how many people are in their party, the time of the week they are going, and the type of seats they want.
'The program reads data from parallel arrays to calculate the total cost of attending the production based on the variables the user
'enters in the input box.

    'declare form level variables
    Dim ctr As Integer, musical(1 To 7) As String, theatre(1 To 7) As String, tube(1 To 7) As String
    Dim stallsDay(1 To 7) As Single, dressDay(1 To 7) As Single, grandDay(1 To 7) As Single
    Dim stallsEnd(1 To 7) As Single, dressEnd(1 To 7) As Single, grandEnd(1 To 7) As Single
    Dim people As Integer, production As String, day As String, seats As String

Private Sub cmdGoToAttractions_Click()
    'returns the user to the popular attractions page of the project
    frmPopularAttractions.Show
    frmTheatre.Hide
End Sub

Private Sub cmdGoToHome_Click()
    'returns the user to the home page of the project
    frmHomePage.Show
    frmTheatre.Hide
End Sub

Private Sub cmdQuit_Click()
    'end the program
    End
End Sub

Private Sub Form_Load()
    'disable the buttons that require the data to be read into the arrays first so the program works properly
    cmdAttend.Enabled = False
    cmdGoToAttractions.Enabled = False
    cmdGoToHome.Enabled = False
    cmdQuit.Enabled = False

End Sub
Private Sub cmdLoadData_Click()

    'open the data file
    Open App.Path & "/theatre.txt" For Input As #2
    
    'initialize the counter variable
    ctr = 0
    
    'read the data from the text file into nine parallel arrays
    Do While Not EOF(2)
        ctr = ctr + 1
        Input #2, musical(ctr), theatre(ctr), tube(ctr), stallsDay(ctr), dressDay(ctr), grandDay(ctr), stallsEnd(ctr), dressEnd(ctr), grandEnd(ctr)
    Loop
    
    'enable the other buttons

    
    'disabled the load data button
    cmdLoadData.Enabled = False
    cmdGoToAttractions.Enabled = True
    cmdGoToHome.Enabled = True
    cmdQuit.Enabled = True
    cmdAttend.Enabled = True
    
End Sub

Private Sub cmdAttend_Click()

    'declare the variables for this subroutine
    Dim I As Integer, found As Boolean, price As Single, totalPrice As Single
    
    'initiate the variables
    I = 0
    found = False
    totalPrice = 0

    'get the information from the user using a message box: the production they want to see, how many people are attending,
    'whether they would like to go during the week or weekend, and the type of seats they want
    production = InputBox("What production would you like to see?", "Theatre Production")
    people = InputBox("How many people are in your party?", "Number of People")
    day = InputBox("Would you like to go on a weekday or weekend?", "Day of the Week")
    seats = InputBox("What type of seats would you like? Enter stalls, dress circle, or grand circle. (Note: Stalls are the best seating, dress circle is next best, grand circle is farthest from the stage.)", "Seating Type")
    
    'Use a match/stop search to search the data that has been read into the arrays
    'Search for the production the user entered in the arrays until it is found, then stop searching
    Do While ((Not found) And (I < ctr))
        I = I + 1
        If production = musical(I) Then
            found = True
        End If
    Loop
    
    'find out how to calculate the price using if statements to find the proper price in the array according to the
    'day of the week the user wants to attend and the type of seats they would like.
    If day = "weekday" And seats = "stalls" Then
            price = stallsDay(I)
        ElseIf day = "weekday" And seats = "dress circle" Then
            price = dressDay(I)
        ElseIf day = "weekday" And seats = "grand circle" Then
            price = grandDay(I)
        ElseIf day = "weekend" And seats = "stalls" Then
            price = stallsEnd(I)
        ElseIf day = "weekend" And seats = "dress circle" Then
            price = dressEnd(I)
        ElseIf day = "weekend" And seats = "grand circle" Then
            price = grandEnd(I)
        Else
            MsgBox "A value you entered is invalid. Please try again.", , "Error"
    End If
    
    'calculate the total price by multiplying the number of people the user entered by the appropriate price found above.
    totalPrice = price * people
            
    
    'If not found, tell the user a value they entered is invalid and to try again.
    'If found then display the results of how much it would cost the user to attend the production based on their preferences.
    If Not found Then
            MsgBox "A value you entered is invalid. Please try again", , "Error"
        Else
            MsgBox "For a party of " & people & " to attend " & production & " on a " & day & " in " & seats & " seating, it would cost " & FormatCurrency(totalPrice) & ".", , "Price"
    End If
    
End Sub


Private Sub cmdInfoBillyElliot_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "Billy Elliot" in a text box and picture box
    txtResults.Text = "Billy Elliot is the tale of a motherless boy whose father wants him to take up boxing. Instead, the boy discovers a love for ballet that leads him from secret lessons to a place at the Royal Ballet School."
    picResults.Print "Playing at Victoria Palace Theatre"
    picResults.Print "Genre: Dramatic Comedy"
    picResults.Print "Tube Station: Victoria"
    picResults.Print "Average Ticket Price: £63"
End Sub

Private Sub cmdInfoLesMis_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "Les Miserables" in a box and a picture box
    txtResults.Text = "Les Miserables concerns love and bravery in 19th century France during the revolutionary struggles."
    picResults.Print "Playing at Queen's Theatre"
    picResults.Print "Genre: Drama"
    picResults.Print "Tube Station: Picadilly Circus"
    picResults.Print "Average Ticket Price: £52"
End Sub

Private Sub cmdInfoLionKing_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "The Lion King" in a box and a picture box
    txtResults.Text = "The Lion King is a colorful musical production suited for all ages. The story begins when the young lion prince Simba is born and his evil uncle Scar is pushed back to second in line to the throne. Scar plots to kill both Simba and his father, King Mufasa, and proclaim himself king. Simba survives but is led to believe that his father died because of him and he decides to flee the kingdom."
    picResults.Print "Playing at the Lyceum Theatre"
    picResults.Print "Genre: Family"
    picResults.Print "Tube Station: Covent Garden"
    picResults.Print "Average Ticket Price: £49"
End Sub

Private Sub cmdInfoMammaMia_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "Mamma Mia!" in a text box and picture box
    txtResults.Text = "Inspired by the songs of ABBA. This original musical is a story of a mother and daughter set on the eve of the daughter's wedding."
    picResults.Print "Playing at Prince of Wales Theatre"
    picResults.Print "Genre: Comedy"
    picResults.Print "Tube Station: Picadilly Circus"
    picResults.Print "Average Ticket Price: £54"
End Sub

Private Sub cmdInfoPhantom_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "Phantom of the Opera" in a text box and picture box
    txtResults.Text = "This haunting musical traces the tragic love story of a beautiful opera singer and a young composer shamed by his physical appearance into a shadowy existence beneath the majestic Opera Paris House."
    picResults.Print "Playing at Her Majesty's Theatre"
    picResults.Print "Genre: Drama"
    picResults.Print "Tube Station: Picadilly Circus"
    picResults.Print "Average Ticket Price: £62"
End Sub

Private Sub cmdInfoRockYou_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "We Will Rock You" in a text box and picture box
    txtResults.Text = "The musical does not chronicle the story of the band, Queen, rather it incorporates the songs of the rock group. The time is the future. Globalisation is complete. On Planet Mall all musical instruments are banned. But resistance is growing. Legend persists that somewhere on Planet Mall instruments still exist..."
    picResults.Print "Playing at Dominion Theatre"
    picResults.Print "Genre: Comedy"
    picResults.Print "Tube Station: Tottenham Court Road"
    picResults.Print "Average Ticket Price: £55"
End Sub

Private Sub cmdInfoWicked_Click()
    'clear the picture box
    picResults.Cls
    
    'print information about "Wicked" in a text box and picture box
    txtResults.Text = "Based on the acclaimed novel by Gregory Maguire that re-imagined the stories and characters created by L. Frank Baum in 'The Wonderful Wizard of Oz', WICKED tells the incredible untold story of an unlikely but profound friendship between two girls who first meet as sorcery students. Their extraordinary adventures in Oz ultimately see them fulfil their destinies as Glinda The Good and the Wicked Witch of the West."
    picResults.Print "Playing at Apollo Victoria Theatre"
    picResults.Print "Genre: Dramatic Comedy"
    picResults.Print "Tube Station: Victoria"
    picResults.Print "Average Ticket Price: £45"
End Sub

