VERSION 5.00
Begin VB.Form frmSUV 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   9795
   ClientLeft      =   1440
   ClientTop       =   840
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   12690
   Begin VB.CommandButton cmdColor5 
      Caption         =   "View a color"
      Height          =   495
      Left            =   10680
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor4 
      Caption         =   "View a color"
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor3 
      Caption         =   "View a color"
      Height          =   495
      Left            =   10080
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor2 
      Caption         =   "View a color"
      Height          =   495
      Left            =   11280
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "View a color"
      Height          =   495
      Left            =   10080
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picColor 
      Height          =   1815
      Left            =   10080
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Go back to Main Page"
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   9000
      Width           =   4335
   End
   Begin VB.CommandButton cmdType 
      Caption         =   "Go back to New Vehicles Page"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   9000
      Width           =   4335
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   600
      ScaleHeight     =   5475
      ScaleWidth      =   9075
      TabIndex        =   3
      Top             =   3240
      Width           =   9135
   End
   Begin VB.PictureBox picModels 
      Height          =   2655
      Left            =   5880
      ScaleHeight     =   2595
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H000080FF&
      Caption         =   "Choose your model"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000C000&
      Caption         =   "Load all SUV models"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmSUV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name- Stack's Car Lot
'Form Name- frmSUV
'Author- Nick Stackenwalt
'Date Written- Saturday March 08, 2008
'Objective-  This form shows the user all SUV models that Ford has to offer for the year of 2008
            'It then allows the user to choose one, and view information about that model
'Other comments-  The Load button allows users to view all models
                 'Then the user is promted to enter a number related to the model they want
                 'They then will view all information available with that model
                 
Private Sub Picture2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdChoose_Click()
Dim ctr As Integer      'Defines variables
Dim Model(1 To 100) As String
Dim J(1 To 100) As Integer
Dim ModelNumber As Integer
Dim color As String
picResults.Cls      'Clears the screen
picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\white.jpg")     'clears the picture screen
ctr = 0     'sets ctr to 0
X = 1       'sets X variable to 1
ModelNumber = InputBox("Enter the Number of the model you would like to view")      'Tells the user to choose a model
If ModelNumber = 1 Then     'Tells what should happen if you choose model number 1
    cmdColor.Visible = True     'Makes the option to view a color available
    cmdColor2.Visible = False
    cmdColor3.Visible = False
    cmdColor4.Visible = False
    cmdColor5.Visible = False
    Open App.Path & "\escape.txt" For Input As #1        'Opens the corresponding file
    Do Until EOF(1)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #1, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #1        'Closes the file
ElseIf ModelNumber = 2 Then     'Tells what should happen if you choose model number 2
    cmdColor2.Visible = True     'Makes the option to view a color available
    cmdColor.Visible = False
    cmdColor3.Visible = False
    cmdColor4.Visible = False
    cmdColor5.Visible = False
    Open App.Path & "\sporttrac.txt" For Input As #2       'Opens the corresponding file
    Do Until EOF(2)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #2, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #2        'Closes the file
ElseIf ModelNumber = 3 Then     'Tells what should happen if you choose model number 3
    cmdColor3.Visible = True     'Makes the option to view a color available
    cmdColor.Visible = False
    cmdColor2.Visible = False
    cmdColor5.Visible = False
    cmdColor4.Visible = False
    Open App.Path & "\escapehybrid.txt" For Input As #3      'Opens the corresponding file
    Do Until EOF(3)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #3, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #3        'Closes the file
ElseIf ModelNumber = 4 Then     'Tells what should happen if you choose model number 4
    cmdColor4.Visible = True     'Makes the option to view a color available
    cmdColor.Visible = False
    cmdColor2.Visible = False
    cmdColor3.Visible = False
    cmdColor5.Visible = False
    Open App.Path & "\explorer.txt" For Input As #4       'Opens the corresponding file
    Do Until EOF(4)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #4, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #4        'Closes the file
ElseIf ModelNumber = 5 Then     'Tells what should happen if you choose model number 5
    cmdColor5.Visible = True     'Makes the option to view a color available
    cmdColor.Visible = False
    cmdColor2.Visible = False
    cmdColor3.Visible = False
    cmdColor4.Visible = False
    Open App.Path & "\expedition.txt" For Input As #1        'Opens the corresponding file
    Do Until EOF(1)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #1, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #5        'Closes the file
End If
End Sub

Private Sub cmdColor_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blueescape.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackescape.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverescape.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redescape.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor2_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\bluetrac.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blacktrac.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silvertrac.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redtrac.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor3_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blueescape.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackescape.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverescape.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redescape.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor4_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blueexplorer.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackexplorer.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverexplorer.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redexplorer.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor5_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blueexpidition.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackexpidition.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverexpidition.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redexpidition.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdLoad_Click()
Dim ctr As Integer      'Delclares my variable
Dim Model(1 To 6) As String     'Delclares my variable
Dim J(1 To 6) As Integer        'Delclares my variable
picModels.Cls       'Clears the Models screen
picModels.Print "Ford SUV Models"     'Prints "Ford SUV Models"
picModels.Print "**********************"      'Prints "*********************"
ctr = 0     'Sets ctr to 0
X = 1       'Sets X to 1
Open App.Path & "\suv.txt" For Input As #1       'Opens suv models file
Do While Not EOF(1)     'Tells to read the entire file
    ctr = ctr + 1       'Adds one to my ctr so it won't read the same thing over again
    Input #1, Model(ctr)        'Reads the names in the file
Loop        'Starts the process over (back to "Do While Not EOF(1)
For X = 1 To ctr        'Reads first name
    picModels.Print Model(X)        'Prints first name
Next X      'Goes to next name
cmdChoose.Visible = True        'Makes the Choose button visible to the user
Close #1
End Sub

Private Sub cmdMain_Click()
frmSUV.Hide     'Hides SUV form
frmMain.Show    'Shows Main form
End Sub

Private Sub cmdType_Click()
frmSUV.Hide     'Hides SUV form
frmNew1.Show    'Shows New Vehicle form
End Sub

