VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   13860
   ClientLeft      =   240
   ClientTop       =   270
   ClientWidth     =   23130
   LinkTopic       =   "Form1"
   ScaleHeight     =   13860
   ScaleWidth      =   23130
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H8000000E&
      Caption         =   "Reload Data"
      Enabled         =   0   'False
      Height          =   975
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdEngsort 
      BackColor       =   &H80000009&
      Caption         =   "Sort English Words Alphabetically"
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdGersort 
      BackColor       =   &H80000009&
      Caption         =   "Sort German Words Alphabetically"
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdForm2 
      BackColor       =   &H80000009&
      Caption         =   "Exchange Rate"
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdinfo 
      BackColor       =   &H80000009&
      Caption         =   "More Information"
      Enabled         =   0   'False
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000009&
      Caption         =   "Search for Data Information"
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransportation 
      BackColor       =   &H80000009&
      Caption         =   "Transportation"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuildings 
      BackColor       =   &H80000009&
      Caption         =   "Important Places"
      Height          =   855
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSights 
      BackColor       =   &H80000009&
      Caption         =   "Sights to See"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdFrage 
      BackColor       =   &H80000009&
      Caption         =   "Phrases you might use"
      Height          =   855
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit Program"
      Height          =   855
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   855
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show me more about Austria and Germany"
      Height          =   855
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   12255
      Left            =   9720
      ScaleHeight     =   12195
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton CMdShowall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Foods"
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CMDTranslate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Translate Words"
      Enabled         =   0   'False
      Height          =   855
      Left            =   15120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Lblpurpose 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Project Form.frx":0000
      BeginProperty Font 
         Name            =   "Gentium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   480
      TabIndex        =   15
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Information for German/Austria tourists
'Form 1
'Joseph Reinhardt
'March 15, 2009
'Give any tourist of Germany/Austria tidbits of information on what to say, foods to eat, activities, exchange rates, and show them interesting things about the countries
'show important words and phrases and their translations
'Give tidbits of information on each words or phrase

Option Explicit
'declare variables for array used in all cmd buttons
'Declare Variables
Dim German(1 To 100) As String, English(1 To 100) As String, Said(1 To 100) As String
Dim Ctr As Single, J As Single
Private Sub cmdForm2_Click()
'Load the second form
Form2.Show

End Sub
Private Sub cmdPicture_Click()
'Load the third form
Form3.Show

End Sub
'Reload file into array after sorted and reenable cmd buttons
Private Sub cmdLoad_Click()
'Define Variables
Dim W As Integer

're-open data file
Open App.Path & "\GermanWordschanged.txt" For Input As #1

'Load the File back into an Array
Do Until EOF(1)
   W = W + 1
    Input #1, German(W), English(W)
Loop

'close the file
Close #1

'enable the buttons again
CMdShowall.Enabled = True
cmdBuildings.Enabled = True
cmdTransportation.Enabled = True
cmdSights.Enabled = True
cmdFrage.Enabled = True
cmdSearch.Enabled = True
CMDTranslate.Enabled = True
cmdinfo.Enabled = True
End Sub
'Load file and load into an array
Private Sub Form_Load()
'open data file
Open App.Path & "\GermanWordschanged.txt" For Input As #1

'Load the File into an Array
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, German(Ctr), English(Ctr)
Loop

'Close File
Close #1

End Sub
'Find a certain data by number and show translation
Private Sub CMDTranslate_Click()
'Declare Variables
Dim Word As Single

'Get number from user
Word = InputBox("Enter the number of a word you'd like to translate", "German Word Number")

'Find number in array and show errors
Select Case Word
    Case Is > 51
        MsgBox "Sorry that number doesn't represent a word", , "Error"
    Case Is < 0
        MsgBox "Sorry that number doesn't represent a word", , "Error"
    Case Else
        MsgBox "The translation of " & English(Word) & " is " & German(Word), , "Translation"
End Select

End Sub
'Show Data for information on food
Private Sub CMdShowall_Click()
'Define Variables
Dim I As Integer

'Print title to data
picresults.Print ""
picresults.Print "Good Food when in Austria/Germany"
picresults.Print "************************************************"

'Find and print array data 1 to 19
For I = 1 To 19
    picresults.Print I & ". "; English(I)
Next I

'Enable Translation  and Information buttons
CMDTranslate.Enabled = True
cmdinfo.Enabled = True

End Sub
'show data for phrases and words
Private Sub cmdFrage_Click()
'Define Variables
Dim K As Single

'Print title to data
picresults.Print ""
picresults.Print "Useful Phrases When in Austria/Germany"
picresults.Print "************************************************"

'Find and print array data 42 to 51
For K = 42 To 51
    picresults.Print K & ". "; English(K)
Next K

'Enable Translation  and Information buttons
CMDTranslate.Enabled = True
cmdinfo.Enabled = True

End Sub
'show data for sights to see
Private Sub cmdSights_Click()
'Define Variables
Dim H As Single

'Print title to data
picresults.Print ""
picresults.Print "Sights to See When in Austria/Germany"
picresults.Print "************************************************"

'Find and print array data 30 to 41
For H = 30 To 41
    picresults.Print H & ". "; English(H)
Next H

'Enable Translation and Information buttons
CMDTranslate.Enabled = True
cmdinfo.Enabled = True

End Sub
'show data for transportation information
Private Sub cmdTransportation_Click()
'Define Variables
Dim Y As Single

'Print title to data
picresults.Print ""
picresults.Print "Transporatation When in Austria/Germany"
picresults.Print "************************************************"

'Find and print array data 24 to 29
For Y = 24 To 29
    picresults.Print Y & ". "; English(Y)
Next Y

'Enable Translation  and Information buttons
CMDTranslate.Enabled = True
cmdinfo.Enabled = True

End Sub
'Show data for important places information
Private Sub cmdBuildings_Click()
'Define Variables
Dim x As Single

'Print title to data
picresults.Print ""
picresults.Print "Important Buildings When in Austria/Germany"
picresults.Print "************************************************"

'Find and print array data 20 to 23
For x = 20 To 23
    picresults.Print x & ". "; English(x)
Next x

'Enable Translation  and Information buttons
CMDTranslate.Enabled = True
cmdinfo.Enabled = True

End Sub

'take data number from user and show information on that data
Private Sub cmdinfo_Click()
'Define Variables
Dim T As Integer

'Take user-input from inout box a give variable T it's value
T = InputBox("Enter the Number what you would like more information on", "Information")

'Check if number has information
'Print in a message box tid bit of information
If T > 51 Then
        MsgBox "That is not a number for information"
    ElseIf T = 51 Then
        MsgBox "A polite way to ask if someone knows to speak English. Said Like Spreken Zie Aenglisch", , "Information"
    ElseIf T = 50 Then
        MsgBox "This has information for travelling. Is usually in train and bus stations", , "Information"
    ElseIf T = 49 Then
         MsgBox "The usual symbol for a restroom", , "Information"
    ElseIf T = 48 Then
         MsgBox "An exit in case of emergency in the building", , "Information"
    ElseIf T = 47 Then
         MsgBox "An exit to a building", , "Information"
    ElseIf T = 46 Then
         MsgBox "An entrance to a building", , "Information"
    ElseIf T = 45 Then
         MsgBox "A polite way to excuse yourself", , "Information"
    ElseIf T = 44 Then
         MsgBox "Polite way to ask for a bill at a restaraunt", , "Information"
    ElseIf T = 43 Then
         MsgBox "Polite way to say thank you", , "Information"
    ElseIf T = 42 Then
         MsgBox "Way to ask please and also say your welcome. Use often when asking for anything", , "Information"
    ElseIf T = 41 Then
         MsgBox "There are many well preserved, interesting, and historical castles throughout Germany and Austria", , "Information"
    ElseIf T = 40 Then
         MsgBox "The oldest and most historical area of any town. Usually the center of any city.", , "Information"
    ElseIf T = 39 Then
         MsgBox "A park usually with wilderness, gardens, and playgrounds.", , "Information"
    ElseIf T = 38 Then
         MsgBox "Usually near the city center with many vendors. Lot's of interesting shopping and handmade goods.", , "Information"
    ElseIf T = 37 Then
         MsgBox "Western Austria and Bavaria is known for it's Mountanious culture. Lot's of hiking and sight seeing to be done in the alps.", , "Information"
    ElseIf T = 36 Then
         MsgBox "There are lot's of great hiking trails in Bavaria and Austria. Have large culture in hiking.", , "Information"
    ElseIf T = 35 Then
         MsgBox "The city hall. Usually a large and beautifully architectured building.", , "Information"
    ElseIf T = 34 Then
         MsgBox "Many festivals are held year-round in the Altstadts of many German/Austrian cities. Many attractions are there like special markets, rides, and beer tents.", , "Information"
    ElseIf T = 33 Then
         MsgBox "There are many historical museums throughout Germany and Austria.", , "Information"
    ElseIf T = 32 Then
         MsgBox "Cathedrals in Austria and Germany are usually very extravagent and are amazing to see in person.", , "Information"
    ElseIf T = 31 Then
         MsgBox "The Abby is the main church of a order of monks. Usually a very extravagent church to see.", , "Information"
    ElseIf T = 30 Then
         MsgBox "There are many churches throughout Austria and Germany that are spectacular.", , "Information"
    ElseIf T = 29 Then
         MsgBox "Hiking has a large part in Bavrian and Austrian culture with many great hiking throughout both countries.", , "Information"
    ElseIf T = 28 Then
         MsgBox "Biking is a great way to get around with many bike paths and ways around the cities.", , "Information"
    ElseIf T = 27 Then
         MsgBox "Planes are a great way to get around Germany and Austria quick.", , "Information"
    ElseIf T = 26 Then
         MsgBox "Trains run throughout Austria and Germany and there is rarely somewhere you can't go with a train.", , "Information"
    ElseIf T = 25 Then
         MsgBox "Many cities have buses as public transport around cities for relatively cheap.", , "Information"
    ElseIf T = 24 Then
         MsgBox "Most cities have Taxi services which can be a good way for those who have trouble reading maps to get around.", , "Information"
    ElseIf T = 23 Then
         MsgBox "A good thing to know about wherever you are in case of emergency.", , "Information"
    ElseIf T = 22 Then
         MsgBox "A good thing to know about wherever you are in case of emergency.", , "Information"
    ElseIf T = 21 Then
         MsgBox "Place to send and recieve mail.Symbol of post horn outside of building.", , "Information"
    ElseIf T = 20 Then
         MsgBox "Place where you can exchange currency and get money from your bank accounts.", , "Information"
    ElseIf T = 19 Then
         MsgBox "A German/Austrian classic in food and they do it the best.", , "Information"
    ElseIf T = 18 Then
         MsgBox "Many of delicious mustards are made by Germans and Austrians but be careful cause they can get spicy.", , "Information"
    ElseIf T = 17 Then
         MsgBox "Coffee shops offer a great atmosphere and some of the best espresso in the world.", , "Information"
    ElseIf T = 16 Then
         MsgBox "A side in many different German/Austrian meals in a variety of different styles.", , "Information"
    ElseIf T = 15 Then
         MsgBox "Some of best Milk chocolates are made by the amazing dairy industry of Germany/Austria.", , "Information"
    ElseIf T = 14 Then
         MsgBox "A good word to know for all onion lovers.", , "Information"
    ElseIf T = 13 Then
         MsgBox "A good word to know for all mushroom lovers.", , "Information"
    ElseIf T = 12 Then
         MsgBox "A traditional German food that is amazingly good in the area.", , "Information"
    ElseIf T = 11 Then
         MsgBox "A word to know when you want a salad as an appetizer.", , "Information"
    ElseIf T = 10 Then
         MsgBox "Austrians/Germans make amazing soups are definately worth trying.", , "Information"
    ElseIf T = 9 Then
         MsgBox "A good word to know when ordering food.", , "Information"
    ElseIf T = 8 Then
         MsgBox "Many different wines come to Austria/Germany from Italy but local treasure Sturm wine is a very interesting and different wine.", , "Information"
    ElseIf T = 7 Then
         MsgBox "Germans/Austrians are know for their wonderful beers and are a must try for any beer lover.", , "Information"
    ElseIf T = 6 Then
         MsgBox "Lemon is a common flavor of Schnapps, a delicious liquor made with fruit flavoring.", , "Information"
    ElseIf T = 5 Then
         MsgBox "A good word to know when ordering food.", , "Information"
    ElseIf T = 4 Then
         MsgBox "A often used ingredient in many foods. My favorite is kurbiskremesuppe or pumpkin cream soup.", , "Information"
    ElseIf T = 3 Then
         MsgBox "A baked good made with either fruit or cream cheese that is native to Austria/Germany.", , "Information"
    ElseIf T = 2 Then
         MsgBox "A good word to know when ordering food.", , "Information"
    ElseIf T = 1 Then
         MsgBox "Austria/Germany have many bakerys with fresh baked bakery goods.", , "Information"
    Else:
        MsgBox "That is not a number for information", , "Error"
End If
End Sub
'Sort German words and shows all for use in search command
Private Sub cmdGersort_Click()
'Define Variables
Dim Pass As Integer, Pos As Integer, Temp As String
Dim Q As Single

'Use bubble sort to sort the information
'this sort will sort ascending
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If German(Pos) > German(Pos + 1) Then
            Temp = German(Pos)
            German(Pos) = German(Pos + 1)
            German(Pos + 1) = Temp
        End If
    Next Pos
Next Pass

'Print the sorted list
For Q = 1 To Ctr
    picresults.Print German(Q)
Next Q

'Disable buttons using file data
cmdLoad.Enabled = True
CMdShowall.Enabled = False
cmdBuildings.Enabled = False
cmdTransportation.Enabled = False
cmdSights.Enabled = False
cmdFrage.Enabled = False
cmdSearch.Enabled = False
CMDTranslate.Enabled = False
cmdinfo.Enabled = False


End Sub
'Sort english words and show all for use in Search command
Private Sub cmdEngsort_Click()
'Define Variables
Dim Pas As Integer, Poss As Integer, Tempe As String
Dim e As Single

'Use bubble sort to sort the information
'this sort will sort ascending
For Pas = 1 To Ctr - 1
    For Poss = 1 To Ctr - Pas
        If English(Poss) > English(Poss + 1) Then
            Tempe = English(Poss)
            English(Poss) = English(Poss + 1)
            English(Poss + 1) = Tempe
        End If
    Next Poss
Next Pas

'Print the sorted list
For e = 1 To Ctr
    picresults.Print English(e)
Next e

'Disable buttons using file data
cmdLoad.Enabled = True
CMdShowall.Enabled = False
cmdBuildings.Enabled = False
cmdTransportation.Enabled = False
cmdSights.Enabled = False
cmdFrage.Enabled = False
cmdSearch.Enabled = False
CMDTranslate.Enabled = False
cmdinfo.Enabled = False
End Sub
'Search for english or german word and display data for that word
Private Sub cmdSearch_Click()
'Define variables and give them Value
Dim Found As Boolean
Dim m As Integer, Search As String
m = 0
Found = False

'Ask user for String to search in input box
Search = InputBox("What would you like to find?", "Search")

'search array for string from user
Do While (Not Found) And (m < Ctr)
    m = m + 1
    If Search = English(m) Then
        Found = True
    ElseIf Search = German(m) Then
        Found = True
    Else:
    End If
Loop

'Determine if found and display info
If (Not Found) Then
        MsgBox "There is no information on that, Sorry", , Error
    Else
        picresults.Print ""
        picresults.Print "English", , "German"
        picresults.Print "****************************************************"
        picresults.Print m & ". " & English(m), Tab(29); German(m)
End If

End Sub
'clear data from picture box
Private Sub cmdclear_Click()

'Clear Picresults Picture box
picresults.Cls

End Sub
'End program
Private Sub cmdEnd_Click()

'Close Form
End

End Sub

