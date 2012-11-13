VERSION 5.00
Begin VB.Form frmLocal 
   BackColor       =   &H000000FF&
   Caption         =   "Local team records for your top 5!"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdfifth 
      BackColor       =   &H000000FF&
      Caption         =   "Display #5!"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdfourth 
      BackColor       =   &H000000FF&
      Caption         =   "Display #4!"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdThird 
      BackColor       =   &H000000FF&
      Caption         =   "Display #3!"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplaysecond 
      BackColor       =   &H000000FF&
      Caption         =   "Display #2!"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   7455
      ScaleWidth      =   6015
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdDisplayfirst 
      BackColor       =   &H000000FF&
      Caption         =   "Display #1!"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to Calculations!"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
End
Attribute VB_Name = "frmLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click() 'goes back to frmSecond
    frmLocal.Hide
    frmSecond.Show
End Sub

Private Sub cmdDisplayfirst_Click() 'searches for first ranked value in array based on name
   Dim A(1 To 100) As String
   Dim pos As Integer
    PicDisplay.Cls
     PicDisplay.Cls
    PicDisplay.Print "here are your "; rankname(1); " Results"
        If InStr(rankname(1), "Track-") <> 0 Then 'searches for value based on string value
            Open App.Path & "\track.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Soccer") <> 0 Then 'searches for value based on string value
            Open App.Path & "\soccer.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Basket") <> 0 Then 'searches for value based on string value
            Open App.Path & "\basket.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Football") <> 0 Then 'searches for value based on string value
            Open App.Path & "\football.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Hockey") <> 0 Then 'searches for value based on string value
            Open App.Path & "\hockey.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Walking") <> 0 Then 'searches for value based on string value
            Open App.Path & "\Walking.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Video") <> 0 Then
                PicDisplay.Print "Uh, there aren't any teams that we know of!"
          ElseIf InStr(rankname(1), "Danc") <> 0 Then 'searches for value based on string value
            Open App.Path & "\dance.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Golf") <> 0 Then 'searches for value based on string value
                PicDisplay.Print "We're Working on this one!"
                PicDisplay.Print "Try back later!"
        
          ElseIf InStr(rankname(1), "Weight") <> 0 Then 'searches for value based on string value
            Open App.Path & "\weight.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Base") <> 0 Then 'searches for value based on string value
            Open App.Path & "\baseball.txt" For Input As #1
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
        
        End Sub

Private Sub cmdDisplaysecond_Click()
   Dim A(1 To 100) As String
   Dim pos As Integer
    PicDisplay.Cls
    PicDisplay.Print "here are your "; rankname(2); " Results"
        If InStr(rankname(2), "Track") <> 0 Then 'searches for value based on string value
            Open App.Path & "\track.txt" For Input As #1
            pos = 0
            Do Until EOF(1)   'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Soccer") <> 0 Then 'searches for value based on string value
            Open App.Path & "\soccer.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Basket") <> 0 Then 'searches for value based on string value
            Open App.Path & "\basket.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Football") <> 0 Then 'searches for value based on string value
            Open App.Path & "\football.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Hockey") <> 0 Then 'searches for value based on string value
            Open App.Path & "\hockey.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Walking") <> 0 Then 'searches for value based on string value
            Open App.Path & "\Walking.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Video") <> 0 Then 'searches for value based on string value
                PicDisplay.Print "Uh, there aren't any teams that we know of!"
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\dance.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Golf") <> 0 Then
                PicDisplay.Print "We're Working on this one!"
                PicDisplay.Print "Try back later!"
        
          ElseIf InStr(rankname(2), "Weight") <> 0 Then 'searches for value based on string value
            Open App.Path & "\weight.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Base") <> 0 Then 'searches for value based on string value
            Open App.Path & "\baseball.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
        
End Sub

Private Sub cmdfifth_Click()
       Dim A(1 To 100) As String
   Dim pos As Integer
    PicDisplay.Cls
    PicDisplay.Print "here are your "; rankname(5); " Results"
        If InStr(rankname(5), "Track") <> 0 Then 'searches for value based on string value
            Open App.Path & "\track.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Soccer") <> 0 Then 'searches for value based on string value
            Open App.Path & "\soccer.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Basket") <> 0 Then 'searches for value based on string value
            Open App.Path & "\basket.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Football") <> 0 Then 'searches for value based on string value
            Open App.Path & "\football.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Hockey") <> 0 Then 'searches for value based on string value
            Open App.Path & "\hockey.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Walking") <> 0 Then 'searches for value based on string value
            Open App.Path & "\Walking.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Video") <> 0 Then 'searches for value based on string value
                PicDisplay.Print "Uh, there aren't any teams that we know of!"
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\dance.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Golf") <> 0 Then
                PicDisplay.Print "We're Working on this one!"
                PicDisplay.Print "Try back later!"
         
          ElseIf InStr(rankname(5), "weight") <> 0 Then 'searches for value based on string value
            Open App.Path & "\weight.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Base") <> 0 Then 'searches for value based on string value
            Open App.Path & "\baseball.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

Private Sub cmdfourth_Click()
       Dim A(1 To 100) As String
   Dim pos As Integer
    PicDisplay.Cls
    PicDisplay.Print "here are your "; rankname(4); " Results"
        If InStr(rankname(4), "Track") <> 0 Then 'searches for value based on string value
            Open App.Path & "\track.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Soccer") <> 0 Then 'searches for value based on string value
            Open App.Path & "\soccer.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Basket") <> 0 Then 'searches for value based on string value
            Open App.Path & "\basket.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Football") <> 0 Then 'searches for value based on string value
            Open App.Path & "\football.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Hockey") <> 0 Then 'searches for value based on string value
            Open App.Path & "\hockey.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Walking") <> 0 Then 'searches for value based on string value
            Open App.Path & "\Walking.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Video") <> 0 Then
                PicDisplay.Print "Uh, there aren't any teams that we know of!"
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\dance.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Golf") <> 0 Then 'searches for value based on string value
                PicDisplay.Print "We're Working on this one!"
                PicDisplay.Print "Try back later!"
        
          ElseIf InStr(rankname(4), "Weight") <> 0 Then 'searches for value based on string value
            Open App.Path & "\weight.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Base") <> 0 Then 'searches for value based on string value
            Open App.Path & "\baseball.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

Private Sub cmdThird_Click()
   Dim A(1 To 100) As String
   Dim pos As Integer
    PicDisplay.Cls
    PicDisplay.Print "here are your "; rankname(3); " Results"
        If InStr(rankname(3), "Track") <> 0 Then 'searches for value based on string value
            Open App.Path & "\track.txt" For Input As #1
            pos = 0
            Do Until EOF(1)   'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Soccer") <> 0 Then 'searches for value based on string value
            Open App.Path & "\soccer.txt" For Input As #1
            pos = 0
            Do Until EOF(1)   'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Basket") <> 0 Then 'searches for value based on string value
            Open App.Path & "\basket.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Football") <> 0 Then 'searches for value based on string value
            Open App.Path & "\football.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Hockey") <> 0 Then 'searches for value based on string value
            Open App.Path & "\hockey.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Walking") <> 0 Then 'searches for value based on string value
            Open App.Path & "\Walking.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Video") <> 0 Then
                PicDisplay.Print "Uh, there aren't any teams that we know of!"
          ElseIf InStr(rankname(3), "danc") <> 0 Then 'searches for value based on string value
            Open App.Path & "\dance.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Golf") <> 0 Then 'searches for value based on string value
                PicDisplay.Print "We're Working on this one!"
                PicDisplay.Print "Try back later!"
        
          ElseIf InStr(rankname(3), "Weight") <> 0 Then 'searches for value based on string value
            Open App.Path & "\weight.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Base") <> 0 Then 'searches for value based on string value
            Open App.Path & "\baseball.txt" For Input As #1
            pos = 0
            Do Until EOF(1)   'opens file to load into an array
                pos = pos + 1
                Input #1, A(pos)
                PicDisplay.Print A(pos)
                PicDisplay.Print
            Loop
        Close #1
        Else
        
            PicDisplay.Print "grrrr"
        End If
End Sub

