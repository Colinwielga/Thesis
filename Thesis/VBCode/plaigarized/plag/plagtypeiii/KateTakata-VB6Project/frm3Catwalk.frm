VERSION 5.00
Begin VB.Form frmCatwalk
   BackColor       =   &H00800080&
   Caption         =   "Catwalk"
   ClientHeight    =   11955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleWidth      =   10890
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
      Height          =   3855
      Left            =   2520
      ScaleHeight     =   3795
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   7440
      Width           =   3375
   End
   Begin VB.CommandButton cmdEnd
      Caption         =   "Quit"
      Height          =   735
      Left            =   6480
      TabIndex        =   6
      Top             =   10560
      Width           =   2295
   End
   Begin VB.CommandButton cmdFinished
      Caption         =   "Finished with the catwalk!"
      Height          =   735
      Left            =   6480
      TabIndex        =   5
      Top             =   8520
      Width           =   2295
   End
   Begin VB.PictureBox picCatPlot
      Height          =   855
      Left            =   1440
      Picture         =   "frm3Catwalk.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   8355
      TabIndex        =   4
      Top             =   5160
      Width           =   8415
      Begin VB.CheckBox Check6
         Caption         =   "Check6"
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check5
         Caption         =   "Check5"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check4
         Caption         =   "Check4"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check3
         Caption         =   "Check3"
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check2
         Caption         =   "Check2"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check1
         Caption         =   "Check1"
         Height          =   255
         Left            =   7200
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdInfo
      Caption         =   "Here's some helpful information!"
      Height          =   735
      Left            =   6480
      TabIndex        =   1
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Image ImageEscher
      Height          =   4005
      Left            =   960
      Picture         =   "frm3Catwalk.frx":1B3C
      Top             =   840
      Width           =   7110
   End
   Begin VB.Label lblInstruct
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm3Catwalk.frx":A72F
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
      Left            =   960
      TabIndex        =   3
      Top             =   6360
      Width           =   9615
   End
   Begin VB.Label lblEscher
      BackStyle       =   0  'Transparent
      Caption         =   "One of the Escher Auditorium catwalks from stage level."
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblCatwalk
      BackStyle       =   0  'Transparent
      Caption         =   "The Catwalk"
      BeginProperty Font
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmCatwalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is where the user works on the catwalk.
Dim instChoose(1 To 100) As String
Dim gelChoose(1 To 100) As String
Dim instString(1 To 100) As String
Dim gelString(1 To 100) As String
Dim CTR As Integer

Private Sub Method1()
'The first method for dealing with data the user feeds into the program.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")

            While Not instChoose(CTR) = LCase("Ellipsoidal")
                Select Case instChoose(CTR)
                Case LCase("Fresnel")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("PAR")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("Scoop")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case Else
                    MsgBox "That is not a valid option, please try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                End Select
            End While

            If Not instChoose(CTR) <> LCase("Ellipsoidal") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "You chose an Ellipsoidal! Nice choice!"
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for front light, your choices will be limited to R09 (amber) and R62 (light blue).", "Gel Choices")
                        Select Case gelChoose(CTR)
                        Case "R09"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "By choosing R09, you are beginning to create a warm wash. To continue the warm wash, every other light afterwards should have R09 in it as well!"
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case "R62"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "By choosing R62, you are beginning to create a cool wash. To continue the cool wash, every other light afterwards should have R62 in it as well!"
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case Else
                            MsgBox "For now, your options are limited to R09 and R62, please try again!", , "Oops!"
                        End Select
            End If

        Open App.Path & "\Catwalk.txt" For Output As #1
            Write #1, CTR, instString(CTR), gelString(CTR)
        Close #1

End Sub

Private Sub Method2()
'The second method for dealing with data the user feeds into the program.

    CTR = CTR + 1

    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")

            Do Until Not instChoose(CTR) <> LCase("Ellipsoidal")
                Select Case instChoose(CTR)
                Case LCase("Fresnel")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("PAR")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case LCase("Scoop")
                    MsgBox "Nice try, but there's a light which would work better for creating front light. Try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                Case Else
                    MsgBox "That is not a valid option, please try again!"
                    instChoose(CTR) = InputBox("Which light would you like to use? Your choices are Ellipsoidal, Fresnel, PAR, or Scoop.", "Instrument Choices")
                End Select
            Loop

            If Not instChoose(CTR) <> LCase("Ellipsoidal") Then
                instString(CTR) = instChoose(CTR)
                MsgBox "You chose an Ellipsoidal! Nice choice!"
                    gelChoose(CTR) = InputBox("What color gel would you like to use? Since this is for front light, your choices will be limited to R09 (amber) and R62 (light blue).", "Gel Choices")
                        Select Case gelChoose(CTR)
                        Case "R09"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "By choosing R09, you are beginning to create a warm wash. To continue the warm wash, every other light afterwards should have R09 in it as well!"
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case "R62"
                            gelString(CTR) = gelChoose(CTR)
                            MsgBox "By choosing R62, you are beginning to create a cool wash. To continue the cool wash, every other light afterwards should have R62 in it as well!"
                            picResults.Print instString(CTR) & " - " & gelString(CTR)
                        Case Else
                            MsgBox "For now, your options are limited to R09 and R62, please try again!", , "Oops!"
                        End Select
            End If

        Open App.Path & "\Catwalk.txt" For Append As #1
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

Private Sub cmdInfo_Click()
'Displays information about the catwalk.
    frmCatInfo.Show
End Sub

Private Sub cmdFinished_Click()
'This button advances the program to the First Electric form.

    MsgBox "You have completed hanging, gelling, and circuiting the catwalk! Let's move on to the first electric!"

    frmFirstE.Show
    frmCatwalk.Hide

End Sub

Private Sub cmdEnd_Click()
'Ends the program.
    End
End Sub

