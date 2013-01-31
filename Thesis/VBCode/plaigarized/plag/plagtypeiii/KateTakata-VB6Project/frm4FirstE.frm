VERSION 5.00
Begin VB.Form frmFirstE
   BackColor       =   &H00808000&
   Caption         =   "First Electric"
   ClientHeight    =   13155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   13155
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3000
      ScaleHeight     =   4515
      ScaleWidth      =   3315
      TabIndex        =   19
      Top             =   8280
      Width           =   3375
   End
   Begin VB.CommandButton cmdFinished
      Caption         =   "I'm finished with the first electric!"
      Height          =   735
      Left            =   7320
      TabIndex        =   18
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   735
      Left            =   7320
      TabIndex        =   17
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton cmdInformation
      Caption         =   "Here's some helpful information!"
      Height          =   735
      Left            =   7320
      TabIndex        =   16
      Top             =   8400
      Width           =   2055
   End
   Begin VB.PictureBox picFirstPlot
      Height          =   495
      Left            =   1320
      Picture         =   "frm4FirstE.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   9915
      TabIndex        =   3
      Top             =   6480
      Width           =   9975
      Begin VB.CheckBox Check12
         Caption         =   "Check12"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check11
         Caption         =   "Check11"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check10
         Caption         =   "Check10"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check9
         Caption         =   "Check9"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check8
         Caption         =   "Check8"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check7
         Caption         =   "Check7"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check6
         Caption         =   "Check6"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check5
         Caption         =   "Check5"
         Height          =   255
         Left            =   5760
         TabIndex        =   8
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check4
         Caption         =   "Check4"
         Height          =   255
         Left            =   6960
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check3
         Caption         =   "Check3"
         Height          =   255
         Left            =   7800
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check2
         Caption         =   "Check2"
         Height          =   255
         Left            =   8760
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox Check1
         Caption         =   "Check1"
         Height          =   255
         Left            =   9240
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picPhoto
      Height          =   5175
      Left            =   1680
      Picture         =   "frm4FirstE.frx":2658
      ScaleHeight     =   5115
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   840
      Width           =   9255
   End
   Begin VB.Label lblInstr
      BackStyle       =   0  'Transparent
      Caption         =   $"frm4FirstE.frx":17073
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   20
      Top             =   7200
      Width           =   9375
   End
   Begin VB.Label lblFirstCap
      BackStyle       =   0  'Transparent
      Caption         =   "The first electric in Gorecki"
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label lblTitle
      BackStyle       =   0  'Transparent
      Caption         =   "First Electric"
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmFirstE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is where the user works on the first electric.
Dim instChoose(1 To 100) As String
Dim gelChoose(1 To 100) As String
Dim instString(1 To 100) As String
Dim gelString(1 To 100) As String
Dim CTR As Integer

Private Sub Method1()
'The first method for dealing with data the user feeds into the program.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            If Not instChoose(CTR) <> LCase("Ellipsoidal") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Ellipsoidals should be used for ONLY side-light on the first electric. Make your placement choices wisely..."
            ElseIf Not instChoose(CTR) <> LCase("Fresnel") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Fresnels are great for adding top-light! Nice choice!"
            ElseIf Not instChoose(CTR) <> LCase("PAR") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "PARs are a good choice for top-light! Good choice!"
            ElseIf Not instChoose(CTR) <> LCase("Scoop") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Nice try, but there's are lights which would be put to better use on the first electric. Try again!"
                instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            ElseIf True
                MsgBox "That is not a valid option, please try again!"
                instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            End If

            If Not instString(CTR) = LCase("Scoop") Then
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for top-light and side-light, you can choose from all five gel colors! Your choices are R09 (amber), R25 (red), R62 (light blue), R80 (blue), and R89 (green). (Remember, it's a good idea to mainly use R09 and R62 to further your washes on stage, and keep the other three limited to use as accent colors.)", "Gel Choices")
                        If Not gelChoose(CTR) <> "R09" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R09, a warm wash color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R25" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R25, a red accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R62" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R62, a cool wash color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R80" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R80, a blue accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R89" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R89, a green accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf True
                            MsgBox "For now, your options are limited to R09, R25, R62, R80, and R89. Please try again!", , "Oops!"
                            gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for top-light and side-light, you can choose from all five gel colors! Your choices are R09 (amber), R25 (red), R62 (light blue), R80 (blue), and R89 (green). (Remember, it's a good idea to mainly use R09 and R62 to further your washes on stage, and keep the other three limited to use as accent colors.)", "Gel Choices")
                        End If
            End If

        Open App.Path & "\First.txt" For Output As #1
            Write #1, CTR, instString(CTR), gelString(CTR)
        Close #1

End Sub

Private Sub Method2()
'The second method for dealing with data the user feeds into the program.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            If Not instChoose(CTR) <> LCase("Ellipsoidal") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Ellipsoidals should be used for ONLY side-light on the first electric. Make your placement choices wisely..."
            ElseIf Not instChoose(CTR) <> LCase("Fresnel") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Fresnels are great for adding top-light! Nice choice!"
            ElseIf Not instChoose(CTR) <> LCase("PAR") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "PARs are a good choice for top-light! Good choice!"
            ElseIf Not instChoose(CTR) <> LCase("Scoop") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "Nice try, but there's are lights which would be put to better use on the first electric. Try again!"
                instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            ElseIf True
                MsgBox "That is not a valid option, please try again!"
                instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
            End If

            If Not instString(CTR) = LCase("Scoop") Then
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for top-light and side-light, you can choose from all five gel colors! Your choices are R09 (amber), R25 (red), R62 (light blue), R80 (blue), and R89 (green). (Remember, it's a good idea to mainly use R09 and R62 to further your washes on stage, and keep the other three limited to use as accent colors.)", "Gel Choices")
                        If Not gelChoose(CTR) <> "R09" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R09, a warm wash color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R25" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R25, a red accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R62" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R62, a cool wash color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R80" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R80, a blue accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf Not gelChoose(CTR) <> "R89" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R89, a green accent color!"
                            picResults.Print instChoose(CTR) & " - " & gelString(CTR)
                        ElseIf True
                            MsgBox "For now, your options are limited to R09, R25, R62, R80, and R89. Please try again!", , "Oops!"
                            gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for top-light and side-light, you can choose from all five gel colors! Your choices are R09 (amber), R25 (red), R62 (light blue), R80 (blue), and R89 (green). (Remember, it's a good idea to mainly use R09 and R62 to further your washes on stage, and keep the other three limited to use as accent colors.)", "Gel Choices")
                        End If
            End If

        Open App.Path & "\First.txt" For Append As #1
            Write #1, CTR, instString(CTR), gelString(CTR)
        Close #1

End Sub

Private Sub Check1_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method1
End Sub

Private Sub Check2_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check3_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check4_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check5_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check6_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check7_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check8_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check9_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check10_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check11_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub Check12_Click()
'Clicking on this check box prompts the user to input data about what instrument to hang and what gel color to use.
    Method2
End Sub

Private Sub cmdFinished_Click()
'This button advances the program to the First Arbor form.

    MsgBox "You have completed hanging, gelling, and circuiting the first electric! Let's make sure it won't fall onto the stage and head over to the arbor!"

    frmFirstE.Hide
    frmFirstArbor.Show

End Sub

Private Sub cmdInformation_Click()
'Displays information about the first electric.
    frmFirstInfo.Show
End Sub

Private Sub cmdQuit_Click()
'Ends the program
    End
End Sub
