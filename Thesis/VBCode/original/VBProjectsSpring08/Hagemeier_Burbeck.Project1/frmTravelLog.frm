VERSION 5.00
Begin VB.Form frmTravelLog 
   BackColor       =   &H008080FF&
   Caption         =   "Your Travel Log: Western Europe"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   Picture         =   "frmTravelLog.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Results"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmdfrance 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      Picture         =   "frmTravelLog.frx":1DCB8
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdgermany 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdireland 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmditaly 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdportugal 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdspain 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdswitzerland 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmduk 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdnetherlands 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Back To Main Menu"
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdbelgium 
      Caption         =   "VIEW OTHER PEOPLES RATINGS"
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picresults 
      Height          =   8175
      Left            =   3480
      ScaleHeight     =   8115
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   240
      Width           =   7455
   End
   Begin VB.CommandButton cmdsurvey 
      BackColor       =   &H000000C0&
      Caption         =   "Complete Survey"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmTravelLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmTravelLog
'Author: Brad Hagemeier
'Date Written:  March 26th, 2008
'Objective: This form is used in order to allow the user to take a survey about any of The Western European Countries.  It will display the overal user rating as well as allow the user
'to see what otehr peopel have rated the place like
Option Explicit
Dim tempX As Single                         'Declaring temporary variabls in order to allow for calcualtions later on
Dim tempY As Single
Dim tempZ As Single
Dim tempT As Single
Dim tempU As Single
Dim tempV As Single
Dim tempW As Single

Private Sub cmdback_Click()
    frmMainMenu.Show                            'This button returns to the main menu by showing that form and hiding this one
    frmTravelLog.Hide
    
End Sub

Private Sub cmdbelgium_Click()
                                                    'This button allows the user to see the overal ratings for belgium
    picResults.Cls                                  'First it clears the picture box
    tempT = belgiumT / belgiumCTR                   'The seven calculations at the side take the running variable and divide it by the number of times the survey has been taken and store it as a temporary variable
    tempU = belgiumU / belgiumCTR
    tempV = belgiumV / belgiumCTR
    tempW = belgiumW / belgiumCTR
    tempX = belgiumX / belgiumCTR
    tempY = belgiumY / belgiumCTR
    tempZ = belgiumZ / belgiumCTR
    picResults.Print "BELGIUM"                      'This prints the name of the Country
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)     'The next seven lines print the average user rating to each of the seven questions
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; belgiumCTR; "Total ratings"                                                      'Prints the total numer of times teh survey ahs been taken
End Sub

Private Sub cmdclear_Click()
    picResults.Cls                                    'This clears the picture screen
        cmdbelgium.Visible = False                     'It aslo sets all the Show overal rating buttons to visible=false
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
End Sub

Private Sub cmdfrance_Click()
    picResults.Cls                              'This button is the same as the other overal ratings buttons except for France
    tempT = franceT / franceCTR
    tempU = franceU / franceCTR
    tempV = franceV / franceCTR
    tempW = franceW / franceCTR
    tempX = franceX / franceCTR
    tempY = franceY / franceCTR
    tempZ = franceZ / franceCTR
    picResults.Print "FRANCE"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; franceCTR; "Total ratings"
End Sub

Private Sub cmdgermany_Click()
    picResults.Cls                          'This button is the same as the other overal ratings buttons except for Germany
    tempT = germanyT / germanyCTR
    tempU = germanyU / germanyCTR
    tempV = germanyV / germanyCTR
    tempW = germanyW / germanyCTR
    tempX = germanyX / germanyCTR
    tempY = germanyY / germanyCTR
    tempZ = germanyZ / germanyCTR
    picResults.Print "GERMANY"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; germanyCTR; "Total ratings"
End Sub

Private Sub cmdireland_Click()
    picResults.Cls                         'This button is the same as the other overal ratings buttons except for Ireland
    tempT = irelandT / irelandCTR
    tempU = irelandU / irelandCTR
    tempV = irelandV / irelandCTR
    tempW = irelandW / irelandCTR
    tempX = irelandX / irelandCTR
    tempY = irelandY / irelandCTR
    tempZ = irelandZ / irelandCTR
    picResults.Print "IRELAND"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; irelandCTR; "Total ratings"
End Sub

Private Sub cmditaly_Click()
    picResults.Cls                          'This button is the same as the other overal ratings buttons except for Italy
    tempT = italyT / italyCTR
    tempU = italyU / italyCTR
    tempV = italyV / italyCTR
    tempW = italyW / italyCTR
    tempX = italyX / italyCTR
    tempY = italyY / italyCTR
    tempZ = italyZ / italyCTR
    picResults.Print "ITALY"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; italyCTR; "Total ratings"
End Sub

Private Sub cmdnetherlands_Click()
    picResults.Cls                                      'This button is the same as the other overal ratings buttons except for The Netherlands
    tempT = netherlandsT / netherlandsCTR
    tempU = netherlandsU / netherlandsCTR
    tempV = netherlandsV / netherlandsCTR
    tempW = netherlandsW / netherlandsCTR
    tempX = netherlandsX / netherlandsCTR
    tempY = netherlandsY / netherlandsCTR
    tempZ = netherlandsZ / netherlandsCTR
    picResults.Print "THE NETHERLANDS"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; netherlandsCTR; "Total ratings"
End Sub

Private Sub cmdportugal_Click()
    picResults.Cls                                  'This button is the same as the other overal ratings buttons except for Portugal
    tempT = portugalT / portugalCTR
    tempU = portugalU / portugalCTR
    tempV = portugalV / portugalCTR
    tempW = portugalW / portugalCTR
    tempX = portugalX / portugalCTR
    tempY = portugalY / portugalCTR
    tempZ = portugalZ / portugalCTR
    picResults.Print "PORTUGAL"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; portugalCTR; "Total ratings"
End Sub

Private Sub cmdspain_Click()
    picResults.Cls                              'This button is the same as the other overal ratings buttons except for Spain
    tempT = spainT / spainCTR
    tempU = spainU / spainCTR
    tempV = spainV / spainCTR
    tempW = spainW / spainCTR
    tempX = spainX / spainCTR
    tempY = spainY / spainCTR
    tempZ = spainZ / spainCTR
    picResults.Print "SPAIN"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; spainCTR; "Total ratings"
End Sub

Private Sub cmdsurvey_Click()

   Dim Country As String           'This declares a variable to store the country which has been entered by the user and will be searched for
   Dim found As Boolean             'This declares a variable which stops the loop when found

'The ten codes below for opening a file are used in order to load the ratings from each of the ten countries text files.
'It is done this way so that a total count and counter can be kept in a text file for finding a average user rating
'The first seven pieces of data loaded are the total of the answers to each of the seven questions asked, the last variable is the counter of how many times the survey was taken.
'This could have been done in each part of the if/Elseif/else statements but for ease and a troubleshooting matter they are put here.
Open App.Path & ("\belgium.txt") For Input As #1
    Input #1, belgiumX, belgiumY, belgiumZ, belgiumV, belgiumW, belgiumU, belgiumT, belgiumCTR
Close #1

Open App.Path & ("\france.txt") For Input As #1
   Input #1, franceX, franceY, franceZ, franceV, franceW, franceU, franceT, franceCTR
Close #1
        
Open App.Path & ("\germany.txt") For Input As #1
    Input #1, germanyX, germanyY, germanyZ, germanyV, germanyW, germanyU, germanyT, germanyCTR
Close #1
        
Open App.Path & ("\ireland.txt") For Input As #1
    Input #1, irelandX, irelandY, irelandZ, irelandV, irelandW, irelandU, irelandT, irelandCTR
Close #1
             
Open App.Path & ("\italy.txt") For Input As #1
    Input #1, italyX, italyY, italyZ, italyV, italyW, italyU, italyT, italyCTR
Close #1
        
Open App.Path & ("\spain.txt") For Input As #1
    Input #1, spainX, spainY, spainZ, spainV, spainW, spainU, spainT, spainCTR
Close #1
        
Open App.Path & ("\portugal.txt") For Input As #1
    Input #1, portugalX, portugalY, portugalZ, portugalV, portugalW, portugalU, portugalT, portugalCTR
Close #1
        
Open App.Path & ("\netherlands.txt") For Input As #1
    Input #1, netherlandsX, netherlandsY, netherlandsZ, netherlandsV, netherlandsW, netherlandsU, netherlandsT, netherlandsCTR
Close #1
        
Open App.Path & ("\switzerland.txt") For Input As #1
    Input #1, switzerlandX, switzerlandY, switzerlandZ, switzerlandV, switzerlandW, switzerlandU, switzerlandT, switzerlandCTR
Close #1
        
Open App.Path & ("\uk.txt") For Input As #1
     Input #1, ukX, ukY, ukZ, ukV, ukW, ukU, ukT, ukCTR
Close #1
  
   
 Do While found = False                                                                         'This is the start of a loop which asks the user for a country name and searches for that name if the name is not found it asks again
    Country = InputBox("please enter the name of the country you traveled to")                  'This is the input box which asks for a country name
    Country = LCase(Country)                                                                    'This converts the input to lower case in order to allow for more correct answers
    
    If Country = "belgium" Or InStr(Country, "bel") > 0 Then                                    'This is the start of a loop which searches for a match to the country entered, it uses an instring function so if the first three letters are entered it still allows the statematn to be true
        belgiumCTR = belgiumCTR + 1                                                             'If Belgium is a match to the country entered it adds one to a counter for the number of times the Belgium survey has been taken
        completedbelgium = True                                                                 'This sets a boolean variable to true, this variable is used on the view data form to tell if a survey has been taken or not
        cmdbelgium.Visible = True                                                               'This set of visible statements makes the button which allows the user to see the overall ratings visible and makes the other countries buttons not visible
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        
        TOTAL = 0                                                                               'Set the total (the sum of the users seven answers)equal to zero
        MsgBox ("Please enter ratings between 0 and 10")                                        'A message box to ask teh user to enter ratings between 0 adn 10
        belX = -1                                                                               'Set the answer to the first question to a negative number in order for the loop to work
        Do While belX < 0 Or belX > 10                                                          'Create a loop which works while the number entered is out of the range of 0 to 10.
            belX = InputBox("How would you rate the destination you traveled to overall?")      'Ask the user the first question and set that answer equal to a variable which is independent of the other answers for this quiz and the otehrs.
            If belX < 0 Or belX > 10 Then                                                       'If the rating is out of range, use a message box to tell the user to enter a rating between 0 and 10
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belX                                                            'When an appropriate answer si given add that answer to the total which is used for this idiviudal survey and reset for each survey
                belgiumX = (belgiumX + belX)                                                    'Add the answer to the global variable for this particular question in order to find the average rating overall.
                
            End If
        Loop                                                                                    'go to start of loop in case the ratign was out of range
    
        belY = -1                                                                               'This is the same as the previous part except for teh second question of the survey
        Do While belY < 0 Or belY > 10
            belY = InputBox("How would you rate the transportation?")
            If belY < 0 Or belY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belY
                belgiumY = (belgiumY + belY)
            End If
        Loop
    
        belZ = -1                                                                               'This is the same as the previous part except for the third question of the survey
        Do While belZ < 0 Or belZ > 10
            belZ = InputBox("How would you rate the lodging?")
            If belZ < 0 Or belZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belZ
                belgiumZ = (belgiumZ + belZ)
            End If
        Loop
    
        belV = -1                                                                               'This is the same as the previous part except for the fourth question of the survey
        Do While belV < 0 Or belV > 10
            belV = InputBox("How would you rate the food?")
            If belV < 0 Or belV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belV
                belgiumV = (belgiumV + belV)
            End If
        Loop
    
        belW = -1                                                                               'This is the same as the previous part except for the fifth question of the survey
        Do While belW < 0 Or belW > 10
            belW = InputBox("How would you rate the helpfulness of the local people?")
            If belW < 0 Or belW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belW
                belgiumW = (belgiumW + belW)
            End If
        Loop
   
        belU = -1                                                                               'This is the same as the previous part except for the sixth question of the survey
        Do While belU < 0 Or belU > 10
            belU = InputBox("How would you rate the local attractions?")
            If belU < 0 Or belU > 10 Then
             MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belU
                belgiumU = (belgiumU + belU)
            End If
        Loop
            belT = -1                                                                           'This is the same as the previous part except for the seventh question of the survey
        Do While belT < 0 Or belT > 10
            belT = InputBox("How would you rate the night life?")
            If belT < 0 Or belT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + belT
                belgiumT = (belgiumT + belT)
            End If
        Loop
    found = True                                                                                'Set found equal to true in order to stop loop
    TOTAL = TOTAL / 7                                                                           'find this users average rating by dividing the total by the number of questions
    picResults.Print "Your overall rating of Belgium is "; FormatNumber(TOTAL, 2)               'Print the usser rating formated to two decimal places






    ElseIf Country = "france" Or InStr(Country, "fra") > 0 Then                                 'THis is the same as teh otehr surveys except for France
        franceCTR = franceCTR + 1
        completedfrance = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = True
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        fraX = -1
        Do While fraX < 0 Or fraX > 10
            fraX = InputBox("How would you rate the destination you traveled to overall?")
            If fraX < 0 Or fraX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraX
                franceX = (franceX + fraX)
            End If
        Loop
    
        fraY = -1
        Do While fraY < 0 Or fraY > 10
        fraY = InputBox("How would you rate the transportation?")
            If fraY < 0 Or fraY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraY
                franceY = (franceY + fraY)
            End If
        Loop
    
        fraZ = -1
        Do While fraZ < 0 Or fraZ > 10
            fraZ = InputBox("How would you rate the lodging?")
            If fraZ < 0 Or fraZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraZ
                franceZ = (franceZ + fraZ)
            End If
        Loop
    
        fraV = -1
        Do While fraV < 0 Or fraV > 10
            fraV = InputBox("How would you rate the food?")
            If fraV < 0 Or fraV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraV
                franceV = (franceV + fraV)
            End If
        Loop
    
        fraW = -1
        Do While fraW < 0 Or fraW > 10
            fraW = InputBox("How would you rate the helpfulness of the local people?")
            If fraW < 0 Or fraW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraW
                franceW = (franceW + fraW)
            End If
        Loop
   
        fraU = -1
        Do While fraU < 0 Or fraU > 10
            fraU = InputBox("How would you rate the local attractions?")
            If fraU < 0 Or fraU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraU
                franceU = (franceU + fraU)
            End If
        Loop
            fraT = -1
        Do While fraT < 0 Or fraT > 10
            fraT = InputBox("How would you rate the night life?")
            If fraT < 0 Or fraT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + fraT
                franceT = (franceT + fraT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of France is "; FormatNumber(TOTAL, 2)
    
    ElseIf Country = "germany" Or InStr(Country, "ger") > 0 Then                            'THis is the same as the other surveys except for Germany
        germanyCTR = germanyCTR + 1
        completedgermany = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = True
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        gerX = -1
        Do While gerX < 0 Or gerX > 10
            gerX = InputBox("How would you rate the destination you traveled to overall?")
            If gerX < 0 Or gerX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerX
                germanyX = (germanyX + gerX)
            End If
        Loop
    
        gerY = -1
        Do While gerY < 0 Or gerY > 10
            gerY = InputBox("How would you rate the transportation?")
            If gerY < 0 Or gerY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerY
                germanyY = (germanyY + gerY)
            End If
        Loop
    
        gerZ = -1
        Do While gerZ < 0 Or gerZ > 10
            gerZ = InputBox("How would you rate the lodging?")
            If gerZ < 0 Or gerZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerZ
                germanyZ = (germanyZ + gerZ)
            End If
        Loop
    
        gerV = -1
        Do While gerV < 0 Or gerV > 10
            gerV = InputBox("How would you rate the food?")
            If gerV < 0 Or gerV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerV
                germanyV = (germanyV + gerV)
            End If
        Loop
    
        gerW = -1
        Do While gerW < 0 Or gerW > 10
            gerW = InputBox("How would you rate the helpfulness of the local people?")
            If gerW < 0 Or gerW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerW
                germanyW = (germanyW + gerW)
            End If
        Loop
   
        gerU = -1
        Do While gerU < 0 Or gerU > 10
            gerU = InputBox("How would you rate the local attractions?")
            If gerU < 0 Or gerU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerU
                germanyU = (germanyU + gerU)
            End If
        Loop
        
        gerT = -1
        Do While gerT < 0 Or gerT > 10
            gerT = InputBox("How would you rate the night life?")
            If gerT < 0 Or gerT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + gerT
                germanyT = (germanyT + gerT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of Germany is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "ireland" Or InStr(Country, "ire") > 0 Then                                'THis is the same as the other surveys except for Ireland
        irelandCTR = irelandCTR + 1
        completedireland = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = True
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        ireX = -1
        Do While ireX < 0 Or ireX > 10
            ireX = InputBox("How would you rate the destination you traveled to overall?")
            If ireX < 0 Or ireX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireX
                irelandX = (irelandX + ireX)
            End If
        Loop
    
        ireY = -1
     
        Do While ireY < 0 Or ireY > 10
            ireY = InputBox("How would you rate the transportation?")
            If ireY < 0 Or ireY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireY
                irelandY = (irelandY + ireY)
            End If
        Loop
    
        ireZ = -1
        Do While ireZ < 0 Or ireZ > 10
            ireZ = InputBox("How would you rate the lodging?")
            If ireZ < 0 Or ireZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireZ
                irelandZ = (irelandZ + ireZ)
            End If
        Loop
    
        ireV = -1
        Do While ireV < 0 Or ireV > 10
            ireV = InputBox("How would you rate the food?")
            If ireV < 0 Or ireV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireV
                irelandV = (irelandV + ireV)
            End If
        Loop
    
        ireW = -1
        Do While ireW < 0 Or ireW > 10
            ireW = InputBox("How would you rate the helpfulness of the local people?")
            If ireW < 0 Or ireW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireW
                irelandW = (irelandW + ireW)
            End If
        Loop
   
        ireU = -1
        Do While ireU < 0 Or ireU > 10
            ireU = InputBox("How would you rate the local attractions?")
            If ireU < 0 Or ireU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireU
                irelandU = (irelandU + ireU)
            End If
        Loop
        
        ireT = -1
        Do While ireT < 0 Or ireT > 10
            ireT = InputBox("How would you rate the night life?")
            If ireT < 0 Or ireT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + ireT
                irelandT = (irelandT + ireT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of Ireland is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "italy" Or InStr(Country, "ita") > 0 Then                                  'THis is the same as the other surveys except for Italy
        italyCTR = italyCTR + 1
        completeditaly = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = True
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        itaX = -1
        Do While itaX < 0 Or itaX > 10
            itaX = InputBox("How would you rate the destination you traveled to overall?")
            If itaX < 0 Or itaX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaX
                italyX = (italyX + itaX)
            End If
        Loop
    
        itaY = -1
        Do While itaY < 0 Or itaY > 10
            itaY = InputBox("How would you rate the transportation?")
            If itaY < 0 Or itaY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaY
                italyY = (italyY + itaY)
            End If
        Loop
    
        itaZ = -1
        Do While itaZ < 0 Or itaZ > 10
            itaZ = InputBox("How would you rate the lodging?")
            If itaZ < 0 Or itaZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaZ
                italyZ = (italyZ + itaZ)
            End If
        Loop
    
        itaV = -1
        Do While itaV < 0 Or itaV > 10
            itaV = InputBox("How would you rate the food?")
            If itaV < 0 Or itaV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaV
                italyV = (italyV + itaV)
            End If
        Loop
    
        itaW = -1
        Do While itaW < 0 Or itaW > 10
            itaW = InputBox("How would you rate the helpfulness of the local people?")
            If itaW < 0 Or itaW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaW
                italyW = (italyW + itaW)
            End If
        Loop
   
        itaU = -1
        Do While itaU < 0 Or itaU > 10
            itaU = InputBox("How would you rate the local attractions?")
            If itaU < 0 Or itaU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaU
                italyU = (italyU + itaU)
            End If
        Loop
            
        itaT = -1
        Do While itaT < 0 Or itaT > 10
            itaT = InputBox("How would you rate the night life?")
            If itaT < 0 Or itaT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + itaT
                italyT = (italyT + itaT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of ITALY is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "netherlands" Or InStr(Country, "net") > 0 Or InStr(Country, "hol") > 0 Then    'THis is the same as the other surveys except for The Netherlands
        netherlandsCTR = netherlandsCTR + 1
        completednetherlands = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = True
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        netX = -1
        Do While netX < 0 Or netX > 10
            netX = InputBox("How would you rate the destination you traveled to overall?")
            If netX < 0 Or netX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netX
                netherlandsX = (netherlandsX + netX)
            End If
        Loop
    
        netY = -1
        Do While netY < 0 Or netY > 10
            netY = InputBox("How would you rate the transportation?")
            If netY < 0 Or netY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netY
                netherlandsY = (netherlandsY + netY)
            End If
        Loop
    
        netZ = -1
        Do While netZ < 0 Or netZ > 10
            netZ = InputBox("How would you rate the lodging?")
            If netZ < 0 Or netZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netZ
                netherlandsZ = (netherlandsZ + netZ)
            End If
        Loop
    
        netV = -1
        Do While netV < 0 Or netV > 10
            netV = InputBox("How would you rate the food?")
            If netV < 0 Or netV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netV
                netherlandsV = (netherlandsV + netV)
            End If
        Loop
    
        netW = -1
        Do While netW < 0 Or netW > 10
            netW = InputBox("How would you rate the helpfulness of the local people?")
            If netW < 0 Or netW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netW
                netherlandsW = (netherlandsW + netW)
            End If
        Loop
   
        netU = -1
        Do While netU < 0 Or netU > 10
            netU = InputBox("How would you rate the local attractions?")
            If netU < 0 Or netU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netU
                netherlandsU = (netherlandsU + netU)
            End If
        Loop
            
        netT = -1
        Do While netT < 0 Or netT > 10
            netT = InputBox("How would you rate the night life?")
            If netT < 0 Or netT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + netT
                netherlandsT = (netherlandsT + netT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of The Netherlands is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "portugal" Or InStr(Country, "por") > 0 Then                               'THis is the same as the other surveys except Portugal
        portugalCTR = portugalCTR + 1
        completedportugal = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = True
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        porX = -1
        Do While porX < 0 Or porX > 10
            porX = InputBox("How would you rate the destination you traveled to overall?")
            If porX < 0 Or porX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porX
                portugalX = (portugalX + porX)
            End If
        Loop
    
        porY = -1
        Do While porY < 0 Or porY > 10
            porY = InputBox("How would you rate the transportation?")
            If porY < 0 Or porY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porY
                portugalY = (portugalY + porY)
            End If
        Loop
    
        porZ = -1
        Do While porZ < 0 Or porZ > 10
            porZ = InputBox("How would you rate the lodging?")
            If porZ < 0 Or porZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porZ
                portugalZ = (portugalZ + porZ)
            End If
        Loop
    
        porV = -1
        Do While porV < 0 Or porV > 10
            porV = InputBox("How would you rate the food?")
            If porV < 0 Or porV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porV
                portugalV = (portugalV + porV)
            End If
        Loop
    
        porW = -1
        Do While porW < 0 Or porW > 10
            porW = InputBox("How would you rate the helpfulness of the local people?")
            If porW < 0 Or porW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porW
                portugalW = (portugalW + porW)
            End If
        Loop
   
        porU = -1
        Do While porU < 0 Or porU > 10
            porU = InputBox("How would you rate the local attractions?")
            If porU < 0 Or porU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porU
                portugalU = (portugalU + porU)
            End If
        Loop
            
        porT = -1
        Do While porT < 0 Or porT > 10
            porT = InputBox("How would you rate the night life?")
            If porT < 0 Or porT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + porT
                portugalT = (portugalT + porT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of Portugal is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "spain" Or InStr(Country, "spa") > 0 Then                              'THis is the same as the other surveys except for Spain
        spainCTR = spainCTR + 1
        completedspain = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = True
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        spaX = -1
        Do While spaX < 0 Or spaX > 10
            spaX = InputBox("How would you rate the destination you traveled to overall?")
            If spaX < 0 Or spaX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaX
                spainX = (spainX + spaX)
            End If
        Loop
    
        spaY = -1
        Do While spaY < 0 Or spaY > 10
            spaY = InputBox("How would you rate the transportation?")
            If spaY < 0 Or spaY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaY
                spainY = (spainY + spaY)
            End If
        Loop
    
        spaZ = -1
        Do While spaZ < 0 Or spaZ > 10
            spaZ = InputBox("How would you rate the lodging?")
            If spaZ < 0 Or spaZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaZ
                spainZ = (spainZ + spaZ)
            End If
        Loop
    
        spaV = -1
        Do While spaV < 0 Or spaV > 10
            spaV = InputBox("How would you rate the food?")
            If spaV < 0 Or spaV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaV
                spainV = (spainV + spaV)
            End If
        Loop
    
        spaW = -1
        Do While spaW < 0 Or spaW > 10
            spaW = InputBox("How would you rate the helpfulness of the local people?")
            If spaW < 0 Or spaW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaW
                spainW = (spainW + spaW)
            End If
        Loop
   
        spaU = -1
        Do While spaU < 0 Or spaU > 10
            spaU = InputBox("How would you rate the local attractions?")
            If spaU < 0 Or spaU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaU
                spainU = (spainU + spaU)
            End If
        Loop
            
        spaT = -1
        Do While spaT < 0 Or spaT > 10
            spaT = InputBox("How would you rate the night life?")
            If spaT < 0 Or spaT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + spaT
                spainT = (spainT + spaT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of Spain is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "switzerland" Or InStr(Country, "swi") > 0 Then                        'THis is the same as the other surveys except for Switzerland
        switzerlandCTR = switzerlandCTR + 1
        completedswitzerland = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = False
        cmdswitzerland.Visible = True
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        swiX = -1
        Do While swiX < 0 Or swiX > 10
            swiX = InputBox("How would you rate the destination you traveled to overall?")
            If swiX < 0 Or swiX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiX
                switzerlandX = (switzerlandX + swiX)
            End If
        Loop
    
        swiY = -1
        Do While swiY < 0 Or swiY > 10
            swiY = InputBox("How would you rate the transportation?")
            If swiY < 0 Or swiY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiY
                switzerlandY = (switzerlandY + swiY)
            End If
        Loop
    
        swiZ = -1
        Do While swiZ < 0 Or swiZ > 10
            swiZ = InputBox("How would you rate the lodging?")
            If swiZ < 0 Or swiZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiZ
                switzerlandZ = (switzerlandZ + swiZ)
            End If
        Loop
    
            swiV = -1
        Do While swiV < 0 Or swiV > 10
            swiV = InputBox("How would you rate the food?")
            If swiV < 0 Or swiV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiV
                switzerlandV = (switzerlandV + swiV)
            End If
        Loop
    
        swiW = -1
        Do While swiW < 0 Or swiW > 10
            swiW = InputBox("How would you rate the helpfulness of the local people?")
            If swiW < 0 Or swiW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiW
               switzerlandW = (switzerlandW + swiW)
            End If
        Loop
   
        swiU = -1
        Do While swiU < 0 Or swiU > 10
            swiU = InputBox("How would you rate the local attractions?")
            If swiU < 0 Or swiU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiU
                switzerlandU = (switzerlandU + swiU)
            End If
        Loop
            
        swiT = -1
        Do While swiT < 0 Or swiT > 10
            swiT = InputBox("How would you rate the night life?")
            If swiT < 0 Or swiT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + swiT
                switzerlandT = (switzerlandT + swiT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of Switzerland is "; FormatNumber(TOTAL, 2)

    ElseIf Country = "united kingdom" Or InStr(Country, "uni") > 0 Or InStr(Country, "uk") > 0 Or InStr(Country, "eng") > 0 Then        'THis is the same as the other surveys except for The United Kingdom
        ukCTR = ukCTR + 1
        completeduk = True
        cmdbelgium.Visible = False
        cmdfrance.Visible = False
        cmdgermany.Visible = False
        cmditaly.Visible = False
        cmdireland.Visible = False
        cmdspain.Visible = False
        cmdportugal.Visible = False
        cmduk.Visible = True
        cmdswitzerland.Visible = False
        cmdnetherlands.Visible = False
        TOTAL = 0
        
        MsgBox ("Please enter ratings between 0 and 10")
        engX = -1
        Do While engX < 0 Or engX > 10
            engX = InputBox("How would you rate the destination you traveled to overall?")
            If engX < 0 Or engX > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engX
                ukX = (ukX + engX)
            End If
        Loop
    
        engY = -1
        Do While engY < 0 Or engY > 10
            engY = InputBox("How would you rate the transportation?")
            If engY < 0 Or engY > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engY
                ukY = (ukY + engY)
            End If
        Loop
    
        engZ = -1
        Do While engZ < 0 Or engZ > 10
            engZ = InputBox("How would you rate the lodging?")
            If engZ < 0 Or engZ > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engZ
                ukZ = (ukZ + engZ)
            End If
        Loop
    
        engV = -1
        Do While engV < 0 Or engV > 10
            engV = InputBox("How would you rate the food?")
            If engV < 0 Or engV > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engV
                ukV = (ukV + engV)
            End If
        Loop
    
        engW = -1
        Do While engW < 0 Or engW > 10
            engW = InputBox("How would you rate the helpfulness of the local people?")
            If engW < 0 Or engW > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engW
               ukW = (ukW + engW)
            End If
        Loop
   
        engU = -1
        Do While engU < 0 Or engU > 10
            engU = InputBox("How would you rate the local attractions?")
            If engU < 0 Or engU > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engU
                ukU = (ukU + engU)
            End If
        Loop
            
        engT = -1
        Do While engT < 0 Or engT > 10
            engT = InputBox("How would you rate the night life?")
            If engT < 0 Or engT > 10 Then
                MsgBox ("please enter a positve rating between 0 and 10")
            Else
                TOTAL = TOTAL + engT
                ukT = (ukT + engT)
            End If
        Loop
    found = True
    TOTAL = TOTAL / 7
    picResults.Print "Your overall rating of The United Kingdom is "; FormatNumber(TOTAL, 1)
    Else
        MsgBox ("Please enter a western European country")                                      'If there is no match use a message box to tell the user to enter a western European counry
        found = False                                                                           'Set found to false in order to loop to the beginning
End If

Loop

'Output all the running variables to the appropriate text files in order to store them for the next use and for the overall ratings
'This could have been done in each part of the quiz but for ease I made them all at the end and loaded them all at the begining
Open App.Path & ("\belgium.txt") For Output As #1
    Print #1, belgiumX; belgiumY; belgiumZ; belgiumV; belgiumW; belgiumU; belgiumT; belgiumCTR;
Close #1

Open App.Path & ("\france.txt") For Output As #1
    Print #1, franceX; franceY; franceZ; franceV; franceW; franceU; franceT; franceCTR;
Close #1
        
Open App.Path & ("\germany.txt") For Output As #1
    Print #1, germanyX; germanyY; germanyZ; germanyV; germanyW; germanyU; germanyT; germanyCTR;
Close #1
        
Open App.Path & ("\ireland.txt") For Output As #1
    Print #1, irelandX; irelandY; irelandZ; irelandV; irelandW; irelandU; irelandT; irelandCTR;
Close #1
             
Open App.Path & ("\italy.txt") For Output As #1
    Print #1, italyX; italyY; italyZ; italyV; italyW; italyU; italyT; italyCTR;
Close #1
        
Open App.Path & ("\spain.txt") For Output As #1
    Print #1, spainX; spainY; spainZ; spainV; spainW; spainU; spainT; spainCTR;
Close #1
        
Open App.Path & ("\portugal.txt") For Output As #1
    Print #1, portugalX; portugalY; portugalZ; portugalV; portugalW; portugalU; portugalT; portugalCTR;
Close #1
        
Open App.Path & ("\netherlands.txt") For Output As #1
    Print #1, netherlandsX; netherlandsY; netherlandsZ; netherlandsV; netherlandsW; netherlandsU; netherlandsT; netherlandsCTR;
Close #1
        
Open App.Path & ("\switzerland.txt") For Output As #1
    Print #1, switzerlandX; switzerlandY; switzerlandZ; switzerlandV; switzerlandW; switzerlandU; switzerlandT; switzerlandCTR;
Close #1
        
Open App.Path & ("\uk.txt") For Output As #1
    Print #1, ukX; ukY; ukZ; ukV; ukW; ukU; ukT; ukCTR;
Close #1
End Sub

Private Sub cmduk_Click()
    picResults.Cls                          'This button is the same as the other overal ratings buttons except for the United Kingdom
    tempT = ukT / ukCTR
    tempU = ukU / ukCTR
    tempV = ukV / ukCTR
    tempW = ukW / ukCTR
    tempX = ukX / ukCTR
    tempY = ukY / ukCTR
    tempZ = ukZ / ukCTR
    picResults.Print "THE UNITED KINGDOM"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; ukCTR; "Total ratings"
End Sub

Private Sub cmdswitzerland_Click()
    picResults.Cls                                  'This button is the same as the other overal ratings buttons except for Switzerland
    tempT = switzerlandT / switzerlandCTR
    tempU = switzerlandU / switzerlandCTR
    tempV = switzerlandV / switzerlandCTR
    tempW = switzerlandW / switzerlandCTR
    tempX = switzerlandX / switzerlandCTR
    tempY = switzerlandY / switzerlandCTR
    tempZ = switzerlandZ / switzerlandCTR
    picResults.Print "SWITZERLAND"
    picResults.Print "-----------------------------------------------------------------------------------------------------"
    picResults.Print "How would you rate the destination you traveled to overall? "; FormatNumber(tempX, 1)
    picResults.Print "How would you rate the transportation? "; FormatNumber(tempY, 1)
    picResults.Print "How would you rate the lodging? "; FormatNumber(tempZ, 1)
    picResults.Print "How would you rate the food? "; FormatNumber(tempV, 1)
    picResults.Print "How would you rate the helpfulness of the local people? "; FormatNumber(tempW, 1)
    picResults.Print "How would you rate the local attractions? "; FormatNumber(tempU, 1)
    picResults.Print "How would you rate the night life? "; FormatNumber(tempT, 1)
    picResults.Print "Out of"; switzerlandCTR; "Total ratings"
End Sub

