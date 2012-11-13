VERSION 5.00
Begin VB.Form frmpositions 
   BackColor       =   &H00000080&
   Caption         =   "Positions of Rugby"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form2"
   Picture         =   "frmpositions.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdlearn 
      Caption         =   "Position Description"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8040
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmddisplaypos 
      Caption         =   "Display Positions"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.PictureBox picbox 
      Height          =   4695
      Left            =   5520
      ScaleHeight     =   4635
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdreturntomainmenu 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      TabIndex        =   0
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   2640
      Picture         =   "frmpositions.frx":24CB9
      Top             =   1680
      Width           =   2880
   End
End
Attribute VB_Name = "frmpositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : RugbyVBProject (Rugby.vbp)
'Form Name : frmpositions(frmpositions.frm)
'Author: Alex Forrest
'purpose of the form: This form is designed to give brief explanations of each position
    'in rugby by first having the user click on the display positions button and then
    'inputting which position they wish to learn about in an input box.

Option Explicit
Dim numarray(1 To 15) As Integer
Dim positionarray(1 To 15) As String



Private Sub cmddisplaypos_Click()
Dim I As Single
Dim j As Single
picbox.Cls

I = 0 'sets the counter equal to 0
    Open App.Path & "\rugbypositions.txt" For Input As #1 'opens the specified file from Path = M:\csi130\VB Project
    Do Until EOF(1) 'instructs the program to begin loop and read the whole file
        I = I + 1
        Input #1, numarray(I), positionarray(I)
    Loop 'loops and reads the next part of the array
    picbox.Print "The Positions of Rugby are:"
    picbox.Print 'prints blank line
    For j = 1 To I 'begins the loop
        picbox.Print numarray(j); positionarray(j) 'prints specified part of both arrays
    Next j 'loops to read and print next part of array
    Close #1 'closes the file when done reading array
End Sub

Private Sub cmdlearn_Click()
Dim position As String
position = InputBox("Input the number of the position you wish to learn about corresponding to the list (1-15)") 'sets the position variable equal to the user's input in the input box
Select Case position 'selects the proper case to be tested
    Case 1 'all of the following case codes test the users input, finds the correct match, and then prints out the corresponding information
        MsgBox "The Loosehead Prop is Usually the stockiest member of the team, whose head typically joins his shoulders without recourse to a neck. His job is to support the hooker in the scrums and the jumpers in the lineout. He also thrives on physically intimidating his opposite number. The only difference between the Loosehead Prop and the Tighthead Prop is the their position in the scrum in realation to where the Scrumhalf feed the ball in. They also each have certain responsibilities in scrum and ruck play.", , "Loosehead Prop"
    Case 2
        MsgBox "The hooker uses their feet to 'hook' the ball in the scrum, because of the pressure put on the body by the scrum it is considered to be one of the most dangerous positions to play. They also normally throw the ball in at line-outs, partly because they are normally the shortest of the forwards, but more usually beacuse they are the most skillfull of the forwards.", , "Hooker"
    Case 3
        MsgBox "The Tighthead Prop is Usually the stockiest member of the team, whose head typically joins his shoulders without recourse to a neck. His job is to support the hooker in the scrums and the jumpers in the lineout. He also thrives on physically intimidating his opposite number. The only difference between the Loosehead Prop and the Tighthead Prop is the their position in the scrum in realation to where the Scrumhalf feed the ball in. They also each have certain responsibilities in scrum and ruck play.", , "Tighthead Prop"
    Case 4 To 5
        MsgBox "Second Rows are almost always the tallest players on the team and so are the primary targets at line-outs. At line-outs, second rows must jump aggressively to catch the ball and get it to the scrum half or at least get the first touch so that the ball comes down on their own side. The two second rows stick their heads between the two props and the hooker in the scrums. They are also responsible for keeping the scrum square and provide the power to shift it forward.", , "Second Row"
    Case 6 To 7
        MsgBox "Flankers are the players with the fewest set responsibilities and therefore the position where the player should have all round attributes, speed, strength, fitness, handling skills, along with many other skills. Flankers are always involved in the game, as they are the real ball winners in broken play especially the no. 7. Flankers do less pushing in the scrum than the tight five, but need to be fast as their task is to break quickly and cover the opposing half-backs if the opponents win the scrum. At one time flankers were allowed to break away from the scrum with the ball.", , "Flanker"
    Case 8
        MsgBox "The modern Eight-man has the physical strength of a forward along with the speed and skill of a back. The eight-man packs down at the rear of the scrum, controlling the movement and feeding the ball to the scrum-half. The eight-man is the position where the ball enters the backline from the scrum and hence both fly half and inside centre take their role from the eight-man who, as the last player in the scrum, can elect to pick and run with the ball like a back. No other forward player from a scrum can legally do this. As a result the eight-man has the opportunities as a back to run from set plays.", , "Eight-Man"
    Case 9
        MsgBox "The Scrum Half forms the all-important link between the forwards and the backs. They normally act as the 'General' for the forwards and are always in the spot of the action. A scrum half is normally quite small, with a high degree of vision and able to react to situations very quickly. A scrum half, pound-for-pound, is very strong as they will spend a large percentage of their time up with the forwards, and with superb handling skills, enabling more time on the ball, which results in less 'pressure' for his 'inside' backs.", , "Scrum Half"
    Case 10
        MsgBox "Fly half is short for flying half back because they take the ball on the run. They are probably the most influential players on the pitch. The fly half is the person who makes key decisions during a game such as whether to kick for space, move the ball wide or run with the ball themselves. They should be very fast, able to kick with both feet, have brilliant handling skills, and operate well under pressure.", , "Fly Half"
    Case 11
        MsgBox "The wings act as finishers to finish movements by scoring tries. The idea being that the space should be created by the forwards and backs inside the wingers so once they receive the ball they have a clear run to use their speed and agility to score tries. They are often the quickest members of the team and need to able to jink and side step to finish off scoring situations. The differentiation between weak and strong side wings are their position in relative to where the ball is on the field.  If the back-line is lined up on the right side of the field, the left-side wing is the weak side wing and vise versa.", , "Weak-side Wing"
    Case 12
        MsgBox "Centres need to have a strong all-round game: they need to be able to break through opposition lines and pass the ball accurately. When attack turns into defence they need to be strong in the tackle. The inside centre tends to be the larger of the two centres and the largest back. In defence or attack, the inside centre is always in the thick of the action, drawing the opposition's defence, making the breaks to make the space for the outside centre and dishing out the tackles in defence along with the forwards. Some of the skills of the fly-half, such as distribution and kicking, can be advantageous to inside centres, as they may be expected to act as fly-halves if the normal fly-half is involved in a ruck or maul.", , "Inside Centre"
    Case 13
        MsgBox "Centres need to have a strong all-round game: they need to be able to break through opposition lines and pass the ball accurately. When attack turns into defence they need to be strong in the tackle. The outside is typically the lighter, more agile of the two centres. They are the rapiers that are given the ball, normally via the fly half, to make breaks through the opposition backs before offloading to the wingers after drawing the last line of defence. An outside centre should be very strong, fast and able to pass with pinpoint accuracy under pressure.", , "Outside Centre"
    Case 14
        MsgBox "The wings act as finishers to finish movements by scoring tries. The idea being that the space should be created by the forwards and backs inside the wingers so once they receive the ball they have a clear run to use their speed and agility to score tries. They are often the quickest members of the team and need to able to jink and side step to finish off scoring situations. The differentiation between weak and strong side wings are their position in relative to where the ball is on the field.  If the back-line is lined up on the right side of the field, the left-side wing is the weak side wing and vise versa.", , "Strong-side Wing"
    Case 15
        MsgBox "The player responsible for the last line of defence against both running attacks and tactical kicks. Must be rock solid under the high ball and unconcerned about the prospect of being gang tackled the moment he takes the catch. Can pop up unexpectedly to create an overlap in an attacking back line. An excellent kicker out of hand and also frequently performs goal-kicking duties", , "Fullback"
    Case Else 'if the user's input does not match any of the cases, it prints the following output
        MsgBox "Sorry, you did not input a position on the list!", , "Error"
    End Select 'ends the selected case
End Sub

Private Sub cmdquit_Click()
    End 'ends the program
End Sub

Private Sub cmdreturntomainmenu_Click()
    frmpositions.Hide
    frmMainmenu.Show 'returns the user to the main menu
End Sub
