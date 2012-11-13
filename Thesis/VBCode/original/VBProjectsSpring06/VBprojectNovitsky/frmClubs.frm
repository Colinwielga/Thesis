VERSION 5.00
Begin VB.Form frmClubs 
   BackColor       =   &H000000FF&
   Caption         =   "Information on ""Your top 5""  sports!"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   3840
      ScaleHeight     =   6435
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CmdThree 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdFour 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Calculations!"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   1935
   End
End
Attribute VB_Name = "frmClubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click() 'returns to previous form
    frmClubs.Hide
    frmSecond.Show
End Sub


Private Sub cmdFive_Click()
PicDisplay.Cls      'clears pic box
   PicDisplay.Print "here are your "; rankname(5); " Results" 'loads basic info onto pic box
    If InStr(rankname(5), "Track-") <> 0 Then   'searches for name in array for match of 1st result
            Open App.Path & "\trackinfo.txt" For Input As #1 'if the statement is true it loads a pre-designated file into an array
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print    'displays results
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Soccer") <> 0 Then 'searches for name in array for match of 1st result
            Open App.Path & "\soccerinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print 'displays results
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Basket") <> 0 Then
            Open App.Path & "\basketinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print 'displays results
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Football") <> 0 Then
            Open App.Path & "\footballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print 'displays results
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Hockey") <> 0 Then
            Open App.Path & "\hockeyinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print  'displays results
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Walking") <> 0 Then
            Open App.Path & "\Walkinginfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)   'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Video") <> 0 Then
               Open App.Path & "\videoinfo.txt" For Input As #1 'opens file with proper data name.
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          ElseIf InStr(rankname(5), "Danc") <> 0 Then
            Open App.Path & "\danceinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Golf") <> 0 Then
               Open App.Path & "\golfinfo.txt" For Input As #1 'opens file with proper data name.
               Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          Close #1
          ElseIf InStr(rankname(5), "Weight") <> 0 Then
            Open App.Path & "\weightinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Base") <> 0 Then
            Open App.Path & "\baseballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

Private Sub cmdFour_Click()
PicDisplay.Cls
   PicDisplay.Print "here are your "; rankname(4); " Results" 'does the same as above button, searhces for neams in rankname array
    If InStr(rankname(4), "Track-") <> 0 Then
            Open App.Path & "\trackinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Soccer") <> 0 Then
            Open App.Path & "\soccerinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Basket") <> 0 Then
            Open App.Path & "\basketinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Football") <> 0 Then
            Open App.Path & "\footballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Hockey") <> 0 Then
            Open App.Path & "\hockeyinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Walking") <> 0 Then
            Open App.Path & "\Walkinginfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Video") <> 0 Then
               Open App.Path & "\videoinfo.txt" For Input As #1 'opens file with proper data name.
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          ElseIf InStr(rankname(2), "Danc") <> 0 Then
            Open App.Path & "\danceinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Golf") <> 0 Then
               Open App.Path & "\golfinfo.txt" For Input As #1 'opens file with proper data name.
               Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          Close #1
          ElseIf InStr(rankname(4), "Weight") <> 0 Then
            Open App.Path & "\weightinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Base") <> 0 Then
            Open App.Path & "\baseballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

Private Sub cmdOne_Click()
  Dim A(1 To 100) As String
   Dim pos As Integer
   PicDisplay.Cls
   PicDisplay.Print "here are your "; rankname(1); " Results" 'does the same as above button, searches for neams in rankname array
    If InStr(rankname(1), "Track-") <> 0 Then
            Open App.Path & "\trackinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Soccer") <> 0 Then
            Open App.Path & "\soccerinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Basket") <> 0 Then
            Open App.Path & "\basketinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Football") <> 0 Then
            Open App.Path & "\footballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Hockey") <> 0 Then
            Open App.Path & "\hockeyinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Walking") <> 0 Then
            Open App.Path & "\Walkinginfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Video") <> 0 Then
               Open App.Path & "\videoinfo.txt" For Input As #1 'opens file with proper data name.
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\danceinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Golf") <> 0 Then
               Open App.Path & "\golfinfo.txt" For Input As #1 'opens file with proper data name.
               Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          Close #1
          ElseIf InStr(rankname(1), "Weight") <> 0 Then
            Open App.Path & "\weightinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Base") <> 0 Then
            Open App.Path & "\baseballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
        
End Sub

Private Sub CmdThree_Click()
PicDisplay.Cls
   PicDisplay.Print "here are your "; rankname(3); " Results"
    If InStr(rankname(3), "Track-") <> 0 Then 'does the same as above button, searhces for neams in rankname array
            Open App.Path & "\trackinfo.txt" For Input As #1
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Soccer") <> 0 Then
            Open App.Path & "\soccerinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Basket") <> 0 Then
            Open App.Path & "\basketinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Football") <> 0 Then
            Open App.Path & "\footballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Hockey") <> 0 Then
            Open App.Path & "\hockeyinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Walking") <> 0 Then
            Open App.Path & "\Walkinginfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Video") <> 0 Then
               Open App.Path & "\videoinfo.txt" For Input As #1 'opens file with proper data name.
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos) 'displays results
                    PicDisplay.Print A(pos)
                    PicDisplay.Print
                Loop
          ElseIf InStr(rankname(3), "Danc") <> 0 Then
            Open App.Path & "\danceinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Golf") <> 0 Then
               Open App.Path & "\golfinfo.txt" For Input As #1 'opens file with proper data name.
               Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos) 'displays results
                    PicDisplay.Print A(pos)
                    PicDisplay.Print
                Loop
          Close #1
          ElseIf InStr(rankname(3), "Weight") <> 0 Then
            Open App.Path & "\weightinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Base") <> 0 Then
            Open App.Path & "\baseballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

Private Sub cmdTwo_Click()
  Dim A(1 To 100) As String
   Dim pos As Integer
   PicDisplay.Cls
   PicDisplay.Print "here are your "; rankname(2); " Results"
    If InStr(rankname(2), "Track-") <> 0 Then 'does the same as above button, searhces for neams in rankname array
            Open App.Path & "\trackinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Soccer") <> 0 Then
            Open App.Path & "\soccerinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Basket") <> 0 Then
            Open App.Path & "\basketinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Football") <> 0 Then
            Open App.Path & "\footballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Hockey") <> 0 Then
            Open App.Path & "\hockeyinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos) 'displays results
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Walking") <> 0 Then
            Open App.Path & "\Walkinginfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Video") <> 0 Then
               Open App.Path & "\videoinfo.txt" For Input As #1 'opens file with proper data name.
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          ElseIf InStr(rankname(2), "Danc") <> 0 Then
            Open App.Path & "\danceinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Golf") <> 0 Then
               Open App.Path & "\golfinfo.txt" For Input As #1 'opens file with proper data name.
               Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    PicDisplay.Print A(pos) 'displays results
                    PicDisplay.Print
                Loop
          Close #1
          ElseIf InStr(rankname(2), "Weight") <> 0 Then
            Open App.Path & "\weightinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Base") <> 0 Then
            Open App.Path & "\baseballinfo.txt" For Input As #1 'opens file with proper data name.
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos) 'displays results
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr" 'used to see if file working... confusses the user on purpose, because the program isnt working.
        End If
End Sub

Private Sub Form_Load() 'assigns the user inputed data as the caption for each button
cmdOne.Caption = rankname(1)
cmdTwo.Caption = rankname(2)
CmdThree.Caption = rankname(3)
cmdFour.Caption = rankname(4)
cmdFive.Caption = rankname(5)
End Sub
