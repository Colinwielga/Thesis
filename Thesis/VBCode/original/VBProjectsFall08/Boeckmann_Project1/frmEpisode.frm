VERSION 5.00
Begin VB.Form frmEpisode 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Episodes Alphabetically"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H0080C0FF&
      Caption         =   "List Episodes by Season"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0080C0FF&
      Caption         =   "Load Episodes"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   3720
      ScaleHeight     =   7275
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   240
      Width           =   8175
   End
   Begin VB.CommandButton cmdNum 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort Episodes By Episode Number"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
   End
End
Attribute VB_Name = "frmEpisode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Episode List Form (frmEpisode)
'Ann Boeckmann
'November 1, 2008
'The purpose of this form is to allow users to type in a season and the episode list for the given
'season will be displayed in the picture box.  The user can sort the episodes by episode number
'(order aired on TV) or in alphabetical order
Dim Num(1 To 200) As Integer, Season(1 To 200) As Integer, Airdate(1 To 200) As String, Title(1 To 200) As String, CTR As Integer, N As Integer, SeasonNum As Integer

Private Sub cmdAlpha_Click()
'puts episodes in alphabetical order

Dim Pass As Integer, Pos As Integer, NumTemp As Integer, SeasonTemp As Integer, AirdateTemp As String, TitleTemp As String

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
    If Title(Pos) > Title(Pos + 1) Then
     TitleTemp = Title(Pos)
     Title(Pos) = Title(Pos + 1)
     Title(Pos + 1) = TitleTemp
     
     NumTemp = Num(Pos)
     Num(Pos) = Num(Pos + 1)
     Num(Pos + 1) = NumTemp
     
     SeasonTemp = Season(Pos)
     Season(Pos) = Season(Pos + 1)
     Season(Pos + 1) = SeasonTemp
     
     AirdateTemp = Airdate(Pos)
     Airdate(Pos) = Airdate(Pos + 1)
     Airdate(Pos + 1) = AirdateTemp
     
End If
    Next Pos
        Next Pass
        
picResults.Cls
picResults.Print Tab(36); "Season "; SeasonNum
picResults.Print "                             "
picResults.Print "Episode #"; Tab(20); "Title"; Tab(60); "Airdate"
picResults.Print "------------------------------------------------------------------------------------------------------------------------------------------"

For N = 1 To CTR
If SeasonNum = Season(N) Then
picResults.Print Num(N); Tab(20); Title(N); Tab(60); Airdate(N)
End If
Next N
        

     
End Sub

Private Sub cmdBack_Click()

frmEpisode.Hide
frmOptions.Show

End Sub

Private Sub cmdLoad_Click()
'loads all episodes from file

CTR = 0
Open App.Path & "\episodelist.txt" For Input As #1

Do Until EOF(1)
 CTR = CTR + 1
 Input #1, Num(CTR), Season(CTR), Airdate(CTR), Title(CTR)
Loop
Close #1

End Sub

Private Sub cmdList_Click()
'displays episodes from a user given season

SeasonNum = InputBox("Please enter a season number (1 - 7)", "Enter Season Number") 'allows user to select a season

If SeasonNum > 7 Or SeasonNum <= 0 Then 'Error message pops up if user enters an invalid season number
MsgBox "This is not a valid season number", , "Error!"
End If

picResults.Cls

If SeasonNum >= 1 And SeasonNum <= 7 Then
picResults.Cls
picResults.Print Tab(36); "Season "; SeasonNum
picResults.Print "                             "
picResults.Print "Episode #"; Tab(20); "Title"; Tab(60); "Airdate"
picResults.Print "------------------------------------------------------------------------------------------------------------------------------------------"

For N = 1 To CTR
If SeasonNum = Season(N) Then
picResults.Print Num(N); Tab(20); Title(N); Tab(60); Airdate(N)
End If
Next N

End If

End Sub

Private Sub cmdNum_Click()
'sorts episodes by episode number (order aired)

Dim Pass As Integer, Pos As Integer, NumTemp As Integer, SeasonTemp As Integer, AirdateTemp As String, TitleTemp As String

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
    If Num(Pos) > Num(Pos + 1) Then
     NumTemp = Num(Pos)
     Num(Pos) = Num(Pos + 1)
     Num(Pos + 1) = NumTemp
     
     TitleTemp = Title(Pos)
     Title(Pos) = Title(Pos + 1)
     Title(Pos + 1) = TitleTemp
     
     SeasonTemp = Season(Pos)
     Season(Pos) = Season(Pos + 1)
     Season(Pos + 1) = SeasonTemp
     
     AirdateTemp = Airdate(Pos)
     Airdate(Pos) = Airdate(Pos + 1)
     Airdate(Pos + 1) = AirdateTemp
     
End If
    Next Pos
        Next Pass
        
picResults.Cls
picResults.Print Tab(36); "Season "; SeasonNum
picResults.Print "                             "
picResults.Print "Episode #"; Tab(20); "Title"; Tab(60); "Airdate"
picResults.Print "------------------------------------------------------------------------------------------------------------------------------------------"

For N = 1 To CTR
If SeasonNum = Season(N) Then
picResults.Print Num(N); Tab(20); Title(N); Tab(60); Airdate(N)
End If
Next N
        
End Sub

