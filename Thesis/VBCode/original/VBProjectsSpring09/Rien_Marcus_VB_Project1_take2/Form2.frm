VERSION 5.00
Begin VB.Form DiveList 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12150
   LinkTopic       =   "Form2"
   ScaleHeight     =   8730
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDiveList 
      Height          =   6255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   10815
   End
   Begin VB.CommandButton cmdLoadDives 
      BackColor       =   &H00FF8080&
      Caption         =   "Click for Diving List"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Enter 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter"
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "After you have found the correct Degree of Difficulty for your dive, select Enter."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Description"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Degree of Difficulty"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Dive Number"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "DiveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Competitive Diving Form
'Form Name: DiveList
'Marcus Rien
'3/22/09
' This Form is an array of all of the Dives possible with their Degree of Difficulty and Dive Description
' Once the user has found the information they were looking for, they select Enter which brings them back to the Dive Sheet (DiveSheet form) in order to input the correct Degree of Difficulty.
Option Explicit
Dim DiveNumber(1 To 200) As String
Dim DiveDD(1 To 200) As String
Dim DiveDescription(1 To 200) As String
Dim CTR As Integer
Dim I As Integer


'This loads the dives from the array
Private Sub cmdLoadDives_Click()
Dim DiveList As String

CTR = 0
Open App.Path & "/1MeterDD.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescription(CTR)
Loop
txtDiveList.Text = "DIVE NUMBER" & "DEGREE OF DIFFICULTY" & "DESCRIPTION"
txtDiveList.Text = "*******************************************************************************************************************"

For I = 1 To CTR
DiveList = DiveList & DiveNumber(I) & "                          " & DiveDD(I) & "                        " & DiveDescription(I) & vbCrLf
txtDiveList.Text = DiveList
Next

Close #1
End Sub

'This goes back to the original form when correct information is found
Private Sub Enter_Click()
              DiveList.Hide
              DiveSheet.Show

End Sub


'This creates a scroll for the printed Array
Private Sub VScroll1_Change()
         picDiveList.Top = -VScroll1.Value
End Sub


