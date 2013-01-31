VERSION 5.00
Begin VB.Form frmERA
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdList
      Caption         =   "Go Back To List Page"
      Height          =   735
      Left            =   9120
      TabIndex        =   7
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   735
      Left            =   9120
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCompute
      Caption         =   "compute"
      Height          =   1215
      Left            =   7200
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.PictureBox picResults
      Height          =   1215
      Left            =   7200
      ScaleHeight     =   1155
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtInnings
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtRuns
      Height          =   1215
      Left            =   1920
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Learn To Calculate Your Favorite Pitchers ERA"
      BeginProperty Font
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label Label2
      Caption         =   "Enter the number of innings that your pitcher has thrown throughout the year =======>"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1
      Caption         =   "Enter the number of runs that your pitcher has given up ====>"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image Image1
      Height          =   8415
      Left            =   0
      Picture         =   "frmERA.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10950
   End
End
Attribute VB_Name = "frmERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Cy Young Award Winners Over the Last 30 Years
'Form Name: frmERA
'Author: Anthony and Cameron
'Date Written: February 13, 2010
'Objective: Teaches the user how to calculate the ERA stat.

Private Sub cmdCompute_Click()
    Dim Innings As Single, ER As Single, ERA As Single
    picResults.Cls 'clear print screen
    Innings = txtInnings.Text
    ER = txtRuns.Text
    ERA = (ER / Innings) * 9 'calculations needed to find out a pitchers ERA
    picResults.Print "Your Pitchers ERA is", FormatNumber(ERA, 2) 'print the ERA and format it to two decimal places


     'different cases to be printed depending on what the pitchers ERA is
        If ERA > 9 Then
            picResults.Print "Struggling"
        ElseIf ERA >= 7 Then
            picResults.Print "Minor Leaguer"
        ElseIf ERA >= 5 Then
            picResults.Print "Rookie"
        ElseIf ERA >= 4 Then
            picResults.Print "Seasoned Vet"
        ElseIf ERA >= 3 Then
            picResults.Print "Great Year"
        ElseIf ERA < 3 Then
            picResults.Print "CY Young Worthy"
        End If

End Sub

Private Sub cmdList_Click()
    frmList.Show 'go to the list page
    frmERA.Hide 'hide the ERA calculations page

End Sub

Private Sub cmdQuit_Click()
    End 'quit the program

End Sub
