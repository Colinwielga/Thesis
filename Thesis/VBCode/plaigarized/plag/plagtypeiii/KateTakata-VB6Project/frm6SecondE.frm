VERSION 5.00
Begin VB.Form frmSecondE
   BackColor       =   &H00008080&
   Caption         =   "Second Electric"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2400
      ScaleHeight     =   4275
      ScaleWidth      =   3675
      TabIndex        =   17
      Top             =   7200
      Width           =   3735
   End
   Begin VB.CommandButton cmdEnd
      Caption         =   "Quit"
      Height          =   735
      Left            =   6840
      TabIndex        =   16
      Top             =   10680
      Width           =   2175
   End
   Begin VB.CommandButton cmdFinished
      Caption         =   "I'm finished with the second electric!"
      Height          =   735
      Left            =   6840
      TabIndex        =   15
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplayInfo
      Caption         =   "Here's some helpful information!"
      Height          =   735
      Left            =   6840
      TabIndex        =   14
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CheckBox Check12
      Caption         =   "Check12"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check11
      Caption         =   "Check11"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check10
      Caption         =   "Check10"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check9
      Caption         =   "Check9"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check8
      Caption         =   "Check8"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check7
      Caption         =   "Check7"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check6
      Caption         =   "Check6"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check5
      Caption         =   "Check5"
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check4
      Caption         =   "Check4"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check3
      Caption         =   "Check3"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check2
      Caption         =   "Check2"
      Height          =   255
      Left            =   9720
      TabIndex        =   3
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox Check1
      Caption         =   "Check1"
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblInstruct
      BackStyle       =   0  'Transparent
      Caption         =   $"frm6SecondE.frx":0000
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   18
      Top             =   6000
      Width           =   9615
   End
   Begin VB.Image Image2
      Height          =   495
      Left            =   960
      Picture         =   "frm6SecondE.frx":010A
      Top             =   5040
      Width           =   9765
   End
   Begin VB.Label lblCap1
      BackStyle       =   0  'Transparent
      Caption         =   "The farthest upstage electric in Escher"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Image ImageElectric
      Height          =   3405
      Left            =   3360
      Picture         =   "frm6SecondE.frx":2762
      Top             =   960
      Width           =   5145
   End
   Begin VB.Label lblTitle
      BackStyle       =   0  'Transparent
      Caption         =   "Second Electric"
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmSecondE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is where the user works on the second electric.
Dim instChoose(1 To 100) As String
Dim gelChoose(1 To 100) As String
Dim instString(1 To 100) As String
Dim gelString(1 To 100) As String
Dim CTR As Integer

Private Sub cmdDisplayInfo_Click()
'Displays information about the second electric.
    frmSecondInfo.Show
End Sub

Private Sub Method1()
'The first method for dealing with data the user feeds into the program.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")

            Do Until Not instChoose(CTR) <> LCase("Scoop")
                Select Case instChoose(CTR)
                Case LCase("Fresnel")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("PAR")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("Ellipsoidal")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case Else
                    MsgBox "That is not a valid option, please try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                End Select
            Loop

            If Not instChoose(CTR) <> LCase("Scoop") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "You chose a Scoop! Nice choice!"
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for accenting scenery, your choices will be limited to R25 (red), R80 (blue), and R89 (green).", "Gel Choices")
                        If gelChoose(CTR) = "R25" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R25! Remember, to create a solid wall of red light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        ElseIf gelChoose(CTR) = "R80" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R80! Remember, to create a solid wall of blue light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        ElseIf gelChoose(CTR) = "R89" Then
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R89! Remember, to create a solid wall of green light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Else
                            MsgBox "For now, your options are limited to R25, R80, and R89, please try again!", , "Oops!"
                            gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for accenting scenery, your choices will be limited to R25 (red), R80 (blue), and R89 (green).", "Gel Choices")
                        End If
            End If

        Open App.Path & "\Second.txt" For Output As #1
            Write #1, CTR, instString(CTR), gelString(CTR)
        Close #1

End Sub

Private Sub Method2()
'The second method for dealing with data.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")

            While Not instChoose(CTR) = LCase("Scoop")
                Select Case instChoose(CTR)
                Case LCase("Fresnel")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("PAR")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("Ellipsoidal")
                    MsgBox "Nice try, but there's a light which would work better for lighting a cyc. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case Else
                    MsgBox "That is not a valid option, please try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                End Select
            End While

            If instChoose(CTR) = LCase("Scoop") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "You chose a Scoop! Nice choice!"
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for accenting scenery, your choices will be limited to R25 (red), R80 (blue), and R89 (green).", "Gel Choices")
                        Select Case gelChoose(CTR)
                        Case "R25"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R25! Remember, to create a solid wall of red light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case "R80"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R80! Remember, to create a solid wall of blue light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case "R89"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "You chose R89! Remember, to create a solid wall of green light, try to divide the lights up evenly in terms of the number of gel colors you'll use."
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case Else
                            MsgBox "For now, your options are limited to R25, R80, and R89, please try again!", , "Oops!"
                            gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for accenting scenery, your choices will be limited to R25 (red), R80 (blue), and R89 (green).", "Gel Choices")
                        End Select
            End If

        Open App.Path & "\Second.txt" For Append As #1
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
'This button advances the program to the Second Arbor form.

    MsgBox "You have completed hanging, gelling, and circuiting the second electric! Let's make sure it won't fall onto the stage and head over to the arbor!"

    frmSecondE.Hide
    frmSecondArbor.Show

End Sub

Private Sub cmdEnd_Click()
'Ends the program.
    End
End Sub


