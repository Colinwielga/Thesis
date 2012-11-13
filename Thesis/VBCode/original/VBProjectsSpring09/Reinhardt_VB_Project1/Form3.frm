VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000FF&
   Caption         =   "Form3"
   ClientHeight    =   13950
   ClientLeft      =   990
   ClientTop       =   645
   ClientWidth     =   23175
   LinkTopic       =   "Form3"
   ScaleHeight     =   13950
   ScaleWidth      =   23175
   Begin VB.PictureBox pictitle 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   16395
      TabIndex        =   4
      Top             =   240
      Width           =   16455
   End
   Begin VB.CommandButton cmddone 
      BackColor       =   &H80000009&
      Caption         =   "Go Back to Main Form"
      Height          =   1455
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdpictures 
      BackColor       =   &H80000009&
      Caption         =   "Show me pictures"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox picpictures 
      BackColor       =   &H00FFFFFF&
      Height          =   12615
      Left            =   1320
      ScaleHeight     =   12555
      ScaleWidth      =   16395
      TabIndex        =   0
      Top             =   960
      Width           =   16455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   $"Form3.frx":0000
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Define Variables
Dim Pictures(1 To 20) As String, ctr3 As Single
'Return to main Form
Private Sub cmddone_Click()

'Show Form 1 again
Form1.Show

End Sub
'Information for German/Austria tourists
'Form 3
'Joseph Reinhardt
'March 20, 2009
'To show the user some pictures of Germany/Austria and give them more information
'Can only be accesed through the Form 1

'Take user input to find and display picture and title
Private Sub cmdpictures_Click()
'Define Variables
Dim L As Single

'Give Variable L value of user-input from the Input box
L = InputBox("Enter number of picture you would like to see.", "Picture")



'Check if number is applicable
'IF number has a number show that picture in the picpictures picture box
If L > 16 Then
        MsgBox "Sorry, that is an invalid number.", , "Error"
    ElseIf L < 1 Then
        MsgBox "Sorry, that is an invalid number.", , "Error"
    Else:
        'Clear the title picture box
        pictitle.Cls
            Select Case L
                Case Is = 1
                    pictitle.Print "The Melk Abbey in Upper Austria. A beautiful abbey but one of many in Austria/Germany"
                Case Is = 2
                    pictitle.Print "The Atlstadt of any Austrian/German city is the historical center"
                Case Is = 3
                    pictitle.Print "The Mountains of Bavaria and Western Austria offer beautiful secenery and mountian activities"
                Case Is = 4
                    pictitle.Print "The Bakeries have wonderful, fresh baked goods. A good place to get a strudel"
                Case Is = 5
                    pictitle.Print "The Beer tents at festivals are a place of merriment with good food, beer, and music"
                Case Is = 6
                    pictitle.Print "The beer maids work at the beer tents during festivals and are very strong"
                Case Is = 7
                    pictitle.Print "The Berlin Wall still stand partly in Berlin and is a reminder of the tough times of cold war"
                Case Is = 8
                    pictitle.Print "The Famous Gate of Berlin that holds that statue of Victoria, the roman goddes of victory"
                Case Is = 9
                    pictitle.Print "The Cafes of Austria/Germany are a great place to go for a good atmosphere. This Cafe, Cafe Sacher of Vienna, is world famous"
                Case Is = 10
                    pictitle.Print "The Markets of any town offer a plethora of interesting goods"
                Case Is = 11
                    pictitle.Print "The churches of Germany/Austrian are gourgeously architectured. This church, the Stephensdom is in Vienna"
                Case Is = 12
                    pictitle.Print "The trains around all of Europe can take you where ever you would like"
                Case Is = 13
                    pictitle.Print "Biking is a commmon activity of Austrians/Germans and is a great way to get around anywhere in a city"
                Case Is = 14
                    pictitle.Print "The WC is the symbol for a bathroom in Europe"
                Case Is = 15
                    pictitle.Print "The Austrian Flag"
                Case Is = 16
                    pictitle.Print "The German Flag"
                Case Else
                End Select
        picpictures.Picture = LoadPicture(App.Path & "\vb project pics\" & Pictures(L))
        
End If

End Sub
'load picture data for display command button
Private Sub Form_Load()

'Open Data File
Open App.Path & "\pictures.txt" For Input As #2

'Load the File into an Array
Do Until EOF(2)
    ctr3 = ctr3 + 1
    Input #2, Pictures(ctr3)
Loop

End Sub

