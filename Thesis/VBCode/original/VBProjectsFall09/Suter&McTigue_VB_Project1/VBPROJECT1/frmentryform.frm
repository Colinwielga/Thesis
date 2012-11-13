VERSION 5.00
Begin VB.Form frmentryform 
   Caption         =   "Welcome Page"
   ClientHeight    =   11895
   ClientLeft      =   3360
   ClientTop       =   1605
   ClientWidth     =   18900
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "frmentryform.frx":0000
   ScaleHeight     =   11895
   ScaleWidth      =   18900
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00404040&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   14520
      Picture         =   "frmentryform.frx":44B49
      TabIndex        =   1
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton cmdusteam 
      BackColor       =   &H00404040&
      Caption         =   "See the U.S. Soccer Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2160
      TabIndex        =   0
      Top             =   8880
      Width           =   2295
   End
End
Attribute VB_Name = "frmentryform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'USAsoccer
'frmentryform
'author Sean
'October 14
'the purpose of this program is for the user to learn more about the USA soccer team see pictures
'also it is meant for fans to buy tickets and find out about different venues
Private Sub cmdquit_Click()
    'this button will quit the program
    End
    
End Sub



Private Sub cmdusteam_Click()
    'This button will load the main form
     Form1.Show
     frmentryform.Hide
    
End Sub



Private Sub Form_Load()
Dim numTimes As Double
numTimes = getLoadedCount()
'this shows when the project loads what how many times its been loaded
MsgBox "View Count is " & numTimes & " time(s)"
End Sub
