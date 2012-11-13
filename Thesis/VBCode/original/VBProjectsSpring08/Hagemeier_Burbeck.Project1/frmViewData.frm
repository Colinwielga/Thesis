VERSION 5.00
Begin VB.Form frmViewData 
   Caption         =   "View Data"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcleartxt 
      Caption         =   "Clear Input Box"
      Height          =   975
      Left            =   3240
      TabIndex        =   35
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Results"
      Height          =   1215
      Left            =   3240
      TabIndex        =   34
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdukblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdswitzerlandblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdnetherlandsblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdportugalblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdspainblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmditalyblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdirelandblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdgermanyblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdfranceblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdswitzerland 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdnetherlands 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmduk 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdportugal 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdspain 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmditaly 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdireland 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdgermany 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdfrance 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3015
   End
   Begin VB.CheckBox Checkuk 
      Caption         =   "The United Kingdom"
      Height          =   615
      Left            =   3240
      TabIndex        =   15
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CheckBox Checknetherlands 
      Caption         =   "The Netherlands"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CheckBox Checkswitzerland 
      Caption         =   "Switzerland"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Checkportugal 
      Caption         =   "Portugal"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox Checkspain 
      Caption         =   "Spain"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox Checkitaly 
      Caption         =   "Italy"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.CheckBox Checkireland 
      Caption         =   "Ireland"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox Checkgermany 
      Caption         =   "Germany"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Checkfrance 
      Caption         =   "France"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox Checkbelgium 
      Caption         =   "Belgium"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdbelgiumblog1 
      Caption         =   "Save Blog"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return To Main Menu"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdbelgium 
      Caption         =   "Display Personal Travel Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtinput 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   6960
      Width           =   6015
   End
   Begin VB.PictureBox picdisplay 
      Height          =   9375
      Left            =   6360
      ScaleHeight     =   9315
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label lblchoose 
      Caption         =   "Please choose a country :"
      Height          =   375
      Left            =   1680
      TabIndex        =   36
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lbltravelblog 
      Caption         =   "Please enter your personal travel information below"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   6015
   End
End
Attribute VB_Name = "frmViewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmViewData
'Author: Brad Hagemeier
'Date Written:  March 26th, 2008
'Objective: This form allows the user to input a blog about any country they wish to.  They just check a country form the check boxes and
'enter the information into the text box.  They use the save blog button in order to save what they have writtem into a text file for that country.
'They can also display data about any of the countries. by checking the box and then clicking display data.  This shows the user what they have written about a country.
'It also shows if they have taken a survey for that country or not and if they have, it shows what their results were. It also shows the overall
'ratings of that particular place.

Option Explicit                 'Declare the form as option explicit in order to make us declare the variables
Dim blog As String
Dim N As Integer                'Declare all the form level variables
Dim bel As Integer              'The string variables are used in order to save the data entered in the text box into a text file
Dim belgiumblog As String       'the single variables are used in order to store data temporarily while it does calculations for the results
Dim ukblog As String
Dim franceblog As String
Dim germanyblog As String
Dim italyblog As String
Dim irelandblog As String
Dim spainblog As String
Dim portugalblog As String
Dim switzerlandblog As String
Dim netherlandsblog As String
Dim tempT As Single
Dim tempU As Single
Dim tempV As Single
Dim tempW As Single
Dim tempX As Single
Dim tempY As Single
Dim tempZ As Single
Dim myString As String, stringLength As Integer, tempString As String, lineLength As Integer, i As Integer
Dim space As String, pos As Integer






Private Sub Checkbelgium_Click()
cmdbelgium.Visible = True               'When the Belgium check box is clicked this makes all the other check boxes clear
cmdfrance.Visible = False               'It also makes only the save blog button for the Belgium active as well as the display results button for Belgium visible
cmdgermany.Visible = False              'This is done in this way because there are stacks of buttons for both and it is the easiest way for me to create the program
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = True
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False



End Sub



Private Sub Checkfrance_Click()
cmdbelgium.Visible = False              'the code for this check box is the same as the code for the other check boxes except for France
cmdfrance.Visible = True
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = True
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False


End Sub


Private Sub Checkgermany_Click()
cmdbelgium.Visible = False               'the code for this check box is the same as the code for the other check boxes except for Germany
cmdfrance.Visible = False
cmdgermany.Visible = True
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = True
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False

End Sub

Private Sub Checkireland_Click()
cmdbelgium.Visible = False                     'the code for this check box is the same as the code for the other check boxes except for Ireland
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = True
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = True
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False


End Sub

Private Sub Checkitaly_Click()
cmdbelgium.Visible = False               'the code for this check box is the same as the code for the other check boxes except for Italy
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = True
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = True
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False

End Sub

Private Sub Checknetherlands_Click()
cmdbelgium.Visible = False               'the code for this check box is the same as the code for the other check boxes except for Netherlands
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = True
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = True
cmdukblog1.Visible = False


End Sub

Private Sub Checkportugal_Click()
cmdbelgium.Visible = False               'the code for this check box is the same as the code for the other check boxes except for Portugal
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = True
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = True
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False

End Sub

Private Sub Checkspain_Click()
cmdbelgium.Visible = False           'the code for this check box is the same as the code for the other check boxes except for Spain
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = True
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkportugal = False
Checkuk = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = True
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False


End Sub

Private Sub Checkswitzerland_Click()
cmdbelgium.Visible = False                   'the code for this check box is the same as the code for the other check boxes except for Switzerland
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = True
cmdnetherlands.Visible = False
cmduk.Visible = False

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkuk = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = True
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = False


End Sub

Private Sub Checkuk_Click()
cmdbelgium.Visible = False           'the code for this check box is the same as the code for the other check boxes except for The United Kingdom
cmdfrance.Visible = False
cmdgermany.Visible = False
cmdireland.Visible = False
cmditaly.Visible = False
cmdspain.Visible = False
cmdportugal.Visible = False
cmdswitzerland.Visible = False
cmdnetherlands.Visible = False
cmduk.Visible = True

Checkbelgium = False
Checkfrance = False
Checkgermany = False
Checkitaly = False
Checkireland = False
Checkspain = False
Checkportugal = False
Checkswitzerland = False
Checknetherlands = False

cmdbelgiumblog1.Visible = False
cmdfranceblog1.Visible = False
cmdgermanyblog1.Visible = False
cmdirelandblog1.Visible = False
cmditalyblog1.Visible = False
cmdspainblog1.Visible = False
cmdportugalblog1.Visible = False
cmdswitzerlandblog1.Visible = False
cmdnetherlandsblog1.Visible = False
cmdukblog1.Visible = True



End Sub

Private Sub cmdbelgium_Click()
Open App.Path & ("\belgiumblog.txt") For Input As #1    'This opens the belgium blog text file and inputs the data in order to work with the previously stored data
    Input #1, belgiumblog                               'The data is input as a variable
Close #1

Open App.Path & ("\belgium.txt") For Input As #1        'This is needed in order to calculate the overall average it causes an error otherwise,it si also why I have all the files open at once in th esurvey folder, it was my original attempt to fix this
    Input #1, belgiumX, belgiumY, belgiumZ, belgiumV, belgiumW, belgiumU, belgiumT, belgiumCTR
Close #1
'This Next part was done with help from The TA, Mr. Kerber, I was able to print what I wanted but he helped with this so it would print on multiple lines.
space = " "
pos = 0
lineLength = 50                                         '#number of characters you want on each line
    myString = belgiumblog
    stringLength = Len(myString)                        'length of string
    i = 1                                               'basically a counter
    While i + lineLength < stringLength                 'write another line until counter > stringLength
       tempString = Mid(myString, i, lineLength)        'computes the next line to write using the mid function
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString                      'prints that computed line
       i = i + pos                                      'increments the counter by the lineLenth
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)
'This is the end of what i had help with.  The next line prints it how I originaly had it.  I use this code to print multiple lines in ten spots(one for each country)
'picdisplay.Print belgiumblog                            'This prints what the user has previously written about the Belgium
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedbelgium = True Then                         'This statment is used in another form, once the user completes a survey of Belgium it is made true
    picdisplay.Print "Your survey results for Belgium were: "  'When a survey has been completed it prints the following information
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; belX 'It gives the users answers to each of the seven questions answered
    picdisplay.Print "How would you rate the transportation? "; belY
    picdisplay.Print "How would you rate the lodging? "; belZ
    picdisplay.Print "How would you rate the food? "; belV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; belW
    picdisplay.Print "How would you rate the local attractions? "; belU
    picdisplay.Print "How would you rate the night life? "; belT
Else
    picdisplay.Print "You have not completed a survey of Belgium yet"     'If a survey has not been completed It will tell the user they have not completed a survey of Belgium yet
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Belgium are:"                       'IT also prints the overal ratings of Belgium by all users
  
tempT = belgiumT / belgiumCTR                                               'THese take the total of everyones answer to each of the questions and divides them
tempU = belgiumU / belgiumCTR                                               'by the total number of times the survey has been taken
tempV = belgiumV / belgiumCTR                                               'The results is the overal average rating of each of the 7 questions
tempW = belgiumW / belgiumCTR                                               'The resutls is stored as a temporary variable
tempX = belgiumX / belgiumCTR
tempY = belgiumY / belgiumCTR
tempZ = belgiumZ / belgiumCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)     'The questions and results for the overal rating of Belgium are printed
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; belgiumCTR; "Total ratings"                                                      'The total number of times the survey has been taken is also printed
End Sub

Private Sub cmdbelgiumblog1_Click()
belgiumblog = txtinput.Text                                                         'This button takes whatever is typed in the text box and converts it to a variable
Open App.Path & ("\belgiumblog.txt") For Output As #1                               'This data is then outputed to a text file in order to be viewed later
   Print #1, belgiumblog                                                            'The data is stored in a special file just for the blog about Belgium
Close #1
End Sub

Private Sub cmdclear_Click()
picdisplay.Cls                                                                      'This button clears the picture box
End Sub

Private Sub cmdcleartxt_Click()
txtinput.Text = ""                                                                  'This button clears the text box by overwrittin it with nothing
End Sub

Private Sub cmdfrance_Click()
Open App.Path & ("\franceblog.txt") For Input As #1                                 'This is the same code as for the other display buttons except for France
    Input #1, franceblog
Close #1
Open App.Path & ("\france.txt") For Input As #1
   Input #1, franceX, franceY, franceZ, franceV, franceW, franceU, franceT, franceCTR
Close #1
space = " "
pos = 0
lineLength = 50


    myString = franceblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)

'picdisplay.Print franceblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedfrance = True Then
    picdisplay.Print "Your survey results for France were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; fraX
    picdisplay.Print "How would you rate the transportation? "; fraY
    picdisplay.Print "How would you rate the lodging? "; fraZ
    picdisplay.Print "How would you rate the food? "; fraV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; fraW
    picdisplay.Print "How would you rate the local attractions? "; fraU
    picdisplay.Print "How would you rate the night life? "; fraT
Else
    picdisplay.Print "You have not completed a survey of France yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of France are:"
  
tempT = franceT / franceCTR
tempU = franceU / franceCTR
tempV = franceV / franceCTR
tempW = franceW / franceCTR
tempX = franceX / franceCTR
tempY = franceY / franceCTR
tempZ = franceZ / franceCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; franceCTR; "Total ratings"
End Sub

Private Sub cmdfranceblog1_Click()
franceblog = txtinput.Text                                  'This is the same as the other save blog buttons except for France
Open App.Path & ("\franceblog.txt") For Output As #1
   Print #1, franceblog
Close #1
End Sub

Private Sub cmdgermany_Click()
Open App.Path & ("\germanyblog.txt") For Input As #1        'This is the same code as for the other display buttons except for Germany
    Input #1, germanyblog
Close #1
Open App.Path & ("\germany.txt") For Input As #1
    Input #1, germanyX, germanyY, germanyZ, germanyV, germanyW, germanyU, germanyT, germanyCTR
Close #1
space = " "
pos = 0
lineLength = 50


    myString = germanyblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)


'picdisplay.Print germanyblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedgermany = True Then
    picdisplay.Print "Your survey results for Germany were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; gerX
    picdisplay.Print "How would you rate the transportation? "; gerY
    picdisplay.Print "How would you rate the lodging? "; gerZ
    picdisplay.Print "How would you rate the food? "; gerV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; gerW
    picdisplay.Print "How would you rate the local attractions? "; gerU
    picdisplay.Print "How would you rate the night life? "; gerT
Else
    picdisplay.Print "You have not completed a survey of Germany yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Germany are:"
  
tempT = germanyT / germanyCTR
tempU = germanyU / germanyCTR
tempV = germanyV / germanyCTR
tempW = germanyW / germanyCTR
tempX = germanyX / germanyCTR
tempY = germanyY / germanyCTR
tempZ = germanyZ / germanyCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; germanyCTR; "Total ratings"
End Sub

Private Sub cmdgermanyblog1_Click()
germanyblog = txtinput.Text                                         'This is the same as the other save blog buttons except for Germany
Open App.Path & ("\germanyblog.txt") For Output As #1
   Print #1, germanyblog
Close #1
End Sub

Private Sub cmdireland_Click()
Open App.Path & ("\irelandblog.txt") For Input As #1                'This is the same code as for the other display buttons except for Ireland
    Input #1, irelandblog
Close #1
Open App.Path & ("\ireland.txt") For Input As #1
    Input #1, irelandX, irelandY, irelandZ, irelandV, irelandW, irelandU, irelandT, irelandCTR
Close #1
             
space = " "
pos = 0
lineLength = 50

    myString = irelandblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)


'picdisplay.Print irelandblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedireland = True Then
    picdisplay.Print "Your survey results for Ireland were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; ireX
    picdisplay.Print "How would you rate the transportation? "; ireY
    picdisplay.Print "How would you rate the lodging? "; ireZ
    picdisplay.Print "How would you rate the food? "; ireV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; ireW
    picdisplay.Print "How would you rate the local attractions? "; ireU
    picdisplay.Print "How would you rate the night life? "; ireT
Else
    picdisplay.Print "You have not completed a survey Ireland yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Ireland are:"
  
tempT = irelandT / irelandCTR
tempU = irelandU / irelandCTR
tempV = irelandV / irelandCTR
tempW = irelandW / irelandCTR
tempX = irelandX / irelandCTR
tempY = irelandY / irelandCTR
tempZ = irelandZ / irelandCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; irelandCTR; "Total ratings"
End Sub

Private Sub cmdirelandblog1_Click()
irelandblog = txtinput.Text                                             'This is the same as the other save blog buttons except for Ireland
Open App.Path & ("\irelandblog.txt") For Output As #1
   Print #1, irelandblog
Close #1
End Sub

Private Sub cmditaly_Click()
Open App.Path & ("\italyblog.txt") For Input As #1          'This is the same code as for the other display buttons except for Italy
    Input #1, italyblog
Close #1
Open App.Path & ("\italy.txt") For Input As #1
    Input #1, italyX, italyY, italyZ, italyV, italyW, italyU, italyT, italyCTR
Close #1
space = " "
pos = 0
lineLength = 50

    myString = italyblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)

'picdisplay.Print italyblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completeditaly = True Then
    picdisplay.Print "Your survey results for Italy were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; itaX
    picdisplay.Print "How would you rate the transportation? "; itaY
    picdisplay.Print "How would you rate the lodging? "; itaZ
    picdisplay.Print "How would you rate the food? "; itaV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; itaW
    picdisplay.Print "How would you rate the local attractions? "; itaU
    picdisplay.Print "How would you rate the night life? "; itaT
Else
    picdisplay.Print "You have not completed a survey of Italy yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Italy are:"
  
tempT = italyT / italyCTR
tempU = italyU / italyCTR
tempV = italyV / italyCTR
tempW = italyW / italyCTR
tempX = italyX / italyCTR
tempY = italyY / italyCTR
tempZ = italyZ / italyCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; italyCTR; "Total ratings"
End Sub

Private Sub cmditalyblog1_Click()
italyblog = txtinput.Text                                               'This is the same as the other save blog buttons except for Italy
Open App.Path & ("\italyblog.txt") For Output As #1
   Print #1, italyblog
Close #1
End Sub

Private Sub cmdnetherlands_Click()
Open App.Path & ("\netherlandsblog.txt") For Input As #1            'This is the same code as for the other display buttons except for The Netherlands
    Input #1, netherlandsblog
Close #1
Open App.Path & ("\netherlands.txt") For Input As #1
    Input #1, netherlandsX, netherlandsY, netherlandsZ, netherlandsV, netherlandsW, netherlandsU, netherlandsT, netherlandsCTR
Close #1
space = " "
pos = 0
lineLength = 50
    myString = netherlandsblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)

'picdisplay.Print netherlandsblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completednetherlands = True Then
    picdisplay.Print "Your survey results for The Netherlands were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; netX
    picdisplay.Print "How would you rate the transportation? "; netY
    picdisplay.Print "How would you rate the lodging? "; netZ
    picdisplay.Print "How would you rate the food? "; netV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; netW
    picdisplay.Print "How would you rate the local attractions? "; netU
    picdisplay.Print "How would you rate the night life? "; netT
Else
    picdisplay.Print "You have not completed a survey of The Netherlands yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of The Netherlands are:"
  
tempT = netherlandsT / netherlandsCTR
tempU = netherlandsU / netherlandsCTR
tempV = netherlandsV / netherlandsCTR
tempW = netherlandsW / netherlandsCTR
tempX = netherlandsX / netherlandsCTR
tempY = netherlandsY / netherlandsCTR
tempZ = netherlandsZ / netherlandsCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; netherlandsCTR; "Total ratings"
End Sub

Private Sub cmdnetherlandsblog1_Click()
netherlandsblog = txtinput.Text                                 'This is the same as the other save blog buttons except for The Netherlands
Open App.Path & ("\netherlandsblog.txt") For Output As #1
   Print #1, netherlandsblog
Close #1
End Sub

Private Sub cmdportugal_Click()
Open App.Path & ("\portugalblog.txt") For Input As #1       'This is the same code as for the other display buttons except for Portugal
    Input #1, portugalblog
Close #1
Open App.Path & ("\portugal.txt") For Input As #1
    Input #1, portugalX, portugalY, portugalZ, portugalV, portugalW, portugalU, portugalT, portugalCTR
Close #1
space = " "
pos = 0
lineLength = 50
    myString = portugalblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)

'picdisplay.Print portugalblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedportugal = True Then
    picdisplay.Print "Your survey results for Portugal were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; porX
    picdisplay.Print "How would you rate the transportation? "; porY
    picdisplay.Print "How would you rate the lodging? "; porZ
    picdisplay.Print "How would you rate the food? "; porV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; porW
    picdisplay.Print "How would you rate the local attractions? "; porU
    picdisplay.Print "How would you rate the night life? "; porT
Else
    picdisplay.Print "You have not completed a survey of Portugal yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Portugal are:"
  
tempT = portugalT / portugalCTR
tempU = portugalU / portugalCTR
tempV = portugalV / portugalCTR
tempW = portugalW / portugalCTR
tempX = portugalX / portugalCTR
tempY = portugalY / portugalCTR
tempZ = portugalZ / portugalCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; portugalCTR; "Total ratings"
End Sub

Private Sub cmdportugalblog1_Click()
portugalblog = txtinput.Text                                'This is the same as the other save blog buttons except for Portugal
Open App.Path & ("\portugalblog.txt") For Output As #1
   Print #1, portugalblog
Close #1
End Sub

Private Sub cmdreturn_Click()
frmViewData.Hide                                           'This button allows you to return tot the main page by hiding this page and showing the main page
frmMainMenu.Show

End Sub

Private Sub cmdspain_Click()
Open App.Path & ("\spainblog.txt") For Input As #1          'This is the same code as for the other display buttons except for Spain
    Input #1, spainblog
Close #1
Open App.Path & ("\spain.txt") For Input As #1
    Input #1, spainX, spainY, spainZ, spainV, spainW, spainU, spainT, spainCTR
Close #1

space = " "
pos = 0
lineLength = 50
    myString = spainblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)


'picdisplay.Print spainblog
picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedspain = True Then
    picdisplay.Print "Your survey results for Spain were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; spaX
    picdisplay.Print "How would you rate the transportation? "; spaY
    picdisplay.Print "How would you rate the lodging? "; spaZ
    picdisplay.Print "How would you rate the food? "; spaV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; spaW
    picdisplay.Print "How would you rate the local attractions? "; spaU
    picdisplay.Print "How would you rate the night life? "; spaT
Else
    picdisplay.Print "You have not completed a survey of Spain yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Spain are:"
  
tempT = spainT / spainCTR
tempU = spainU / spainCTR
tempV = spainV / spainCTR
tempW = spainW / spainCTR
tempX = spainX / spainCTR
tempY = spainY / spainCTR
tempZ = spainZ / spainCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; spainCTR; "Total ratings"
End Sub

Private Sub cmdspainblog1_Click()
spainblog = txtinput.Text                                       'This is the same as the other save blog buttons except for Spain
Open App.Path & ("\spainblog.txt") For Output As #1
   Print #1, spainblog
Close #1
End Sub

Private Sub cmdswitzerland_Click()
Open App.Path & ("\switzerlandblog.txt") For Input As #1        'This is the same code as for the other display buttons except for Switzerland
    Input #1, switzerlandblog
Close #1
Open App.Path & ("\switzerland.txt") For Input As #1
    Input #1, switzerlandX, switzerlandY, switzerlandZ, switzerlandV, switzerlandW, switzerlandU, switzerlandT, switzerlandCTR
Close #1
space = " "
pos = 0
lineLength = 50
    myString = switzerlandblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)
'picdisplay.Print switzerlandblog

picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completedswitzerland = True Then
    picdisplay.Print "Your survey results for Switzerland were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; swiX
    picdisplay.Print "How would you rate the transportation? "; swiY
    picdisplay.Print "How would you rate the lodging? "; swiZ
    picdisplay.Print "How would you rate the food? "; swiV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; swiW
    picdisplay.Print "How would you rate the local attractions? "; swiU
    picdisplay.Print "How would you rate the night life? "; swiT
Else
    picdisplay.Print "You have not completed a survey of Switzerland yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of Switzerland are:"
  
tempT = switzerlandT / switzerlandCTR
tempU = switzerlandU / switzerlandCTR
tempV = switzerlandV / switzerlandCTR
tempW = switzerlandW / switzerlandCTR
tempX = switzerlandX / switzerlandCTR
tempY = switzerlandY / switzerlandCTR
tempZ = switzerlandZ / switzerlandCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; switzerlandCTR; "Total ratings"
End Sub

Private Sub cmdswitzerlandblog1_Click()
switzerlandblog = txtinput.Text                                 'This is the same as the other save blog buttons except for switzerland
Open App.Path & ("\switzerlandblog.txt") For Output As #1
   Print #1, switzerlandblog
Close #1
End Sub

Private Sub cmduk_Click()
Open App.Path & ("\ukblog.txt") For Input As #1             'This is the same code as for the other display buttons except for The Untied Kingdom
    Input #1, ukblog
Close #1
Open App.Path & ("\uk.txt") For Input As #1
     Input #1, ukX, ukY, ukZ, ukV, ukW, ukU, ukT, ukCTR
Close #1
space = " "
pos = 0
lineLength = 50
    myString = ukblog
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       pos = InStrRev(tempString, space)
       tempString = Mid(myString, i, pos)
       picdisplay.Print tempString
       i = i + pos
    Wend
    picdisplay.Print Right(myString, stringLength - i + 1)
    
'picdisplay.Print ukblog

picdisplay.Print "-----------------------------------------------------------------------------------------------------------------------------"
If completeduk = True Then
    picdisplay.Print "Your survey results for The United Kingdom were: "
    picdisplay.Print ""
    picdisplay.Print "How would you rate the destination you traveled to overall? "; engX
    picdisplay.Print "How would you rate the transportation? "; engY
    picdisplay.Print "How would you rate the lodging? "; engZ
    picdisplay.Print "How would you rate the food? "; engV
    picdisplay.Print "How would you rate the helpfulness of the local people? "; engW
    picdisplay.Print "How would you rate the local attractions? "; engU
    picdisplay.Print "How would you rate the night life? "; engT
Else
    picdisplay.Print "You have not completed a survey of The United Kingdom yet"
End If
picdisplay.Print ""
picdisplay.Print "The overal ratings of The United Kingdom are:"
  
tempT = ukT / ukCTR
tempU = ukU / ukCTR
tempV = ukV / ukCTR
tempW = ukW / ukCTR
tempX = ukX / ukCTR
tempY = ukY / ukCTR
tempZ = ukZ / ukCTR
picdisplay.Print "-----------------------------------------------------------------------------------------------------"
picdisplay.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
picdisplay.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
picdisplay.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
picdisplay.Print "How would you rate the food? "; FormatNumber(tempV, 1)
picdisplay.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
picdisplay.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
picdisplay.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
picdisplay.Print "Out of"; ukCTR; "Total ratings"

End Sub

Private Sub cmdukblog1_Click()
ukblog = txtinput.Text                                  'This is the same as the other save blog buttons except for The United Kingdom
Open App.Path & ("\ukblog.txt") For Output As #1
   Print #1, ukblog
Close #1
End Sub




        

        


        

        

        

        

        


