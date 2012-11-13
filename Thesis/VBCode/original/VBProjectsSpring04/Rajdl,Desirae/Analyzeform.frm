VERSION 5.00
Begin VB.Form Analyzeform 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhelp 
      Caption         =   "Help"
      Height          =   615
      Left            =   1800
      TabIndex        =   16
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display "
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   4560
      Picture         =   "Analyzeform.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2115
      TabIndex        =   14
      Top             =   5520
      Width           =   2175
   End
   Begin VB.PictureBox picresults4 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   5040
      ScaleHeight     =   675
      ScaleWidth      =   3795
      TabIndex        =   13
      Top             =   4200
      Width           =   3855
   End
   Begin VB.TextBox txtdisplay 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   10080
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get Data"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdpercent 
      Caption         =   "Percentages"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8760
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdalpha 
      Caption         =   "Alphabetical order"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Totals"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picresults3 
      BackColor       =   &H00C0FFFF&
      Height          =   2655
      Left            =   7560
      ScaleHeight     =   2595
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H00C0FFFF&
      Height          =   2655
      Left            =   3960
      ScaleHeight     =   2595
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0FFFF&
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2595
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lbldisplay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "by itself"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblpercent 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Percentages and Totals:"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblalpha 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Totals in Alphabetical Order:"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lbltotal 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Totals:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Analyzeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'These are all the global variables that are used throughout
Dim I As Integer
Dim Category(1 To 9) As String
Dim sum As Single
Dim Path As String
Dim Votes(1 To 9) As Integer
Dim percentage(1 To 9) As Single
'JEC Voting Analysis  (JECvoting)
'Analytical Page (Analyzeform)
'By: Desirae Rajdl
'Written: March 10, 2004
'This form was written to help analyze voting information gathered from a
'ficticious vote by students on what events they want to see go on at
'CSBSJU.  It will gather the information, display the totals, order the
'totals in alphabetical order and display that, display the totals and
'percentages, and also display individual categories if wanted.
Private Sub cmdalpha_Click()
Dim pass As Single
Dim comp As Single
Dim tempcategory As String
Dim tempvotes As Single
'I clear all of the picture boxes with each command button to
'keep the display area free from distraction by having too many
'numbers all over
picresults.Cls
picresults2.Cls
picresults3.Cls
picresults4.Cls
I = 9
'Here is where the comparison between categories is done to put
'them in alphabetical order.
For pass = 1 To (I - 1)
    For comp = 1 To (I - pass)
        'switches to alphabetize
        If Category(comp) > Category(comp + 1) Then
            tempcategory = Category(comp)
            Category(comp) = Category(comp + 1)
            Category(comp + 1) = tempcategory
            tempvotes = Votes(comp)
            Votes(comp) = Votes(comp + 1)
            Votes(comp + 1) = tempvotes
        End If
    Next comp
Next pass
sum = 0
picresults2.Print "Category"; Tab(20); "Votes"
picresults2.Print "************************************"
'categories are displayed and sum is found to display
For I = 1 To 9
    picresults2.Print Category(I); Tab(20); Votes(I)
    sum = sum + Votes(I)
Next I
picresults2.Print "************************************"
picresults2.Print "Total"; Tab(20); sum
End Sub

Private Sub cmddisplay_Click()
Dim display As String
Dim found As Boolean
picresults.Cls
picresults2.Cls
picresults3.Cls
picresults4.Cls
found = False
position = 0
sum = 0
'The sum of all of the votes is found
For I = 1 To 9
    sum = sum + Votes(I)
Next I
'Here the percentage for each category is found
For I = 1 To 9
    percentage(I) = Votes(I) / sum
Next I
'Here the category entered by the user is gotten and then
'compared through all of the categories and the number of votes
'and percentage is displayed if found
display = txtdisplay.Text
Do While Not found And position < 9
    position = position + 1
    If Category(position) = display Then
        picresults4.Print "Category"; Tab(13); "Votes"; Tab(23); "Percentage"
        picresults4.Print "******************************************"
        picresults4.Print display; Tab(13); Votes(position); Tab(23); FormatPercent(percentage(position))
        found = True
    End If
Loop
'If the category entered was not found this is displayed
If Not found Then
    picresults4.Print "Sorry, your category was not found, please try again."
End If
End Sub

Private Sub cmdget_Click()
'Here the file is brought up and put into the array
Open Path & "mydata.txt" For Input As #1
I = 0
Do While Not EOF(1)
    I = I + 1
    Input #1, Category(I), Votes(I)
Loop
'the other command buttons not enabled before are enabled
'and then able to be pushed and used.
cmdtotal.Enabled = True
cmdalpha.Enabled = True
cmdpercent.Enabled = True
cmddisplay.Enabled = True
cmdget.Enabled = False
End Sub

Private Sub cmdhelp_Click()
'This will bring up the introductary form and hide the general
'form with the calculations on it.
Introform.Show
Analyzeform.Hide
End Sub

Private Sub cmdpercent_Click()
picresults.Cls
picresults2.Cls
picresults3.Cls
picresults4.Cls
picresults3.Print "Category"; Tab(18); "Votes"; Tab(32); "Percentage"
picresults3.Print "*****************************************************"
sum = 0
'The sum of the votes are found
For I = 1 To 9
    sum = sum + Votes(I)
Next I
'The percentage of each category is found and displayed.
For I = 1 To 9
    percentage(I) = Votes(I) / sum
    picresults3.Print Category(I); Tab(18); Votes(I); Tab(33); FormatPercent(percentage(I))
Next I
picresults3.Print "*****************************************************"
picresults3.Print "Total:"; Tab(18); sum
End Sub

Private Sub cmdquit_Click()
Close #1
End
End Sub

Private Sub cmdtotal_Click()
picresults.Cls
picresults2.Cls
picresults3.Cls
picresults4.Cls
sum = 0
'Here the categories and votes are displayed as they were brought
'up from the file, and the sum of the votes is found and displayed.
picresults.Print "Category"; Tab(20); "Votes"
picresults.Print "************************************"
For I = 1 To 9
    picresults.Print Category(I); Tab(20); Votes(I)
    sum = sum + Votes(I)
Next I
picresults.Print "************************************"
picresults.Print "Total"; Tab(20); sum
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\Rajdl, Desirae\"
End Sub
