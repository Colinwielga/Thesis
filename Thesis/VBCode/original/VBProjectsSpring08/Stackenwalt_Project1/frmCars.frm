VERSION 5.00
Begin VB.Form frmCars 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   2430
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   10845
   Begin VB.CommandButton cmdColor4 
      Caption         =   "View a color"
      Height          =   495
      Left            =   9240
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor3 
      Caption         =   "View a color"
      Height          =   495
      Left            =   8040
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor2 
      Caption         =   "View a color"
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "View a color"
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picColor 
      Height          =   1695
      Left            =   8040
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Go back to Main Page"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   9960
      Width           =   4335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Go back to New Vehicles Page"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   9960
      Width           =   4335
   End
   Begin VB.PictureBox picResults 
      Height          =   6375
      Left            =   480
      ScaleHeight     =   6315
      ScaleWidth      =   7035
      TabIndex        =   3
      Top             =   3360
      Width           =   7095
   End
   Begin VB.PictureBox picModels 
      Height          =   2655
      Left            =   5520
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H000000FF&
      Caption         =   "Choose your model"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load all Car models"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name- Stack's Car Lot
'Form Name- frmCars
'Author- Nick Stackenwalt
'Date Written- Saturday March 08, 2008
'Objective-  This form shows the user all Car models that Ford has to offer for the year of 2008
            'It then allows the user to choose one, and view information about that model
'Other comments-  The Load button allows users to view all models
                 'Then the user is promted to enter a number related to the model they want
                 'They then will view all information available with that model

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
    cmdColor.Visible = True 'Makes the color view option available
    cmdColor2.Visible = False
    cmdColor3.Visible = False
    cmdColor4.Visible = False
    Open App.Path & "\focus.txt" For Input As #1        'Opens the corresponding file
    Do Until EOF(1)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #1, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #1        'Closes the file
ElseIf ModelNumber = 2 Then     'Tells what should happen if you choose model number 2
    cmdColor2.Visible = True        'Makes the color view option available
    cmdColor.Visible = False
    cmdColor3.Visible = False
    cmdColor4.Visible = False
    Open App.Path & "\fusion.txt" For Input As #2       'Opens the corresponding file
    Do Until EOF(2)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #2, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #2        'Closes the file
ElseIf ModelNumber = 3 Then     'Tells what should happen if you choose model number 3
    cmdColor3.Visible = True     'Makes the color view option available
    cmdColor2.Visible = False
    cmdColor.Visible = False
    cmdColor4.Visible = False
    Open App.Path & "\mustang.txt" For Input As #3      'Opens the corresponding file
    Do Until EOF(3)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #3, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #3        'Closes the file
ElseIf ModelNumber = 4 Then     'Tells what should happen if you choose model number 4
    cmdColor4.Visible = True     'Makes the color view option available
    cmdColor.Visible = False
    cmdColor3.Visible = False
    cmdColor2.Visible = False
    Open App.Path & "\taurus.txt" For Input As #4       'Opens the corresponding file
    Do Until EOF(4)     'Tells it to read entire file
        ctr = ctr + 1       'Adds 1 to the counter
        Input #4, Model(ctr)        'Tells it what input to use
    Loop        'starts over at "Do Until EOF(1)"
            For X = 1 To ctr        'Prints off all lines in the file
                picResults.Print Model(X)
            Next X
    Close #4        'Closes the file
End If
End Sub

Private Sub cmdColor_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\bluefocus.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackfocus.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverfocus.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redfocus.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor2_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\bluefusion.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackfusion.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silverfusioin.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redfusion.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor3_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\bluemustang.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blackmustang.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silvermustang.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redmustang.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdColor4_Click()
Dim color As String         'Dims variable color
color = InputBox("Enter a color (Blue, Black, Silver, Red)")        'Asks for a choice of color
    If color = "Blue" Then      'Loads blue picture
        picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\bluetaurus.jpg")
        ElseIf color = "Black" Then     'Loads black picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\blacktaurus.jpg")
        ElseIf color = "Silver" Then        'Loads Silver picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\silvertaurus.jpg")
        ElseIf color = "Red" Then       'Loads Red picture
            picColor.Picture = LoadPicture(App.Path & "\ProjectPictures\redtaurus.jpg")
        Else        'If none availbe it tells you that
            MsgBox ("Color not available")
    End If
End Sub

Private Sub cmdLoad_Click()
Dim ctr As Integer      'Delclares my variable
Dim Model(1 To 6) As String     'Delclares my variable
Dim J(1 To 6) As Integer        'Delclares my variable
picModels.Cls       'Clears the Models screen
picModels.Print "Ford Car Models"     'Prints "Ford Car Models"
picModels.Print "**********************"      'Prints "*********************"
ctr = 0     'Sets ctr to 0
X = 1       'Sets X to 1
Open App.Path & "\cars.txt" For Input As #75       'Opens car models file
Do While Not EOF(75)     'Tells to read the entire file
    ctr = ctr + 1       'Adds one to my ctr so it won't read the same thing over again
    Input #75, Model(ctr)        'Reads the names in the file
Loop        'Starts the process over (back to "Do While Not EOF(1)
For X = 1 To ctr        'Reads first name
    picModels.Print Model(X)        'Prints first name
Next X      'Goes to next name
cmdChoose.Visible = True        'Makes the Choose button visible to the user
Close #75
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdMain_Click()
frmCars.Hide        'Hides Cars form
frmMain.Show        'Shows Main form
End Sub

Private Sub cmdNew_Click()
frmCars.Hide        'Hides Cars form
frmNew1.Show        'Shows New Vehicles form
End Sub

