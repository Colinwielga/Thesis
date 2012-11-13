VERSION 5.00
Begin VB.Form FrmTheDesert 
   BackColor       =   &H0080C0FF&
   Caption         =   "The Desert"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form3"
   ScaleHeight     =   6660
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturntoHeadquarters 
      BackColor       =   &H0080C0FF&
      Caption         =   "Return to Safari HQ"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.PictureBox picdesert 
      Height          =   3495
      Left            =   360
      Picture         =   "FrmTheDesert.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   3120
      Width           =   6615
   End
   Begin VB.PictureBox picImage 
      Height          =   4455
      Left            =   7200
      Picture         =   "FrmTheDesert.frx":1331C
      ScaleHeight     =   4395
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton cmdWeather 
      BackColor       =   &H0080C0FF&
      Caption         =   "Whats the weather?"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      DrawMode        =   4  'Mask Not Pen
      FillColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   5055
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
   End
   Begin VB.TextBox txttemp 
      Height          =   735
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblentertemp 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter the Temperature outside:"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmTheDesert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Great Safari Adventure
'Frm The Desert
'Kit and Liz Chambers
'February 23rd 2010
'Objective: The purpose of this program is to:
           'Say what to do depending on the temperature entered
           
Private Sub Cmdquit_Click()
End
End Sub

Private Sub cmdReturntoHeadquarters_Click()
FrmTheDesert.Hide  'hides the desert form
FrmWelcome.Show ' shows the welcome page

End Sub

Private Sub cmdWeather_Click()
'This program is used to calculate the temperature,
'given the degress entered by the user into a textbox on the form.
'declare the variables used
Dim Temp As Single, Weather As String
    
    'Clear the picturebox used for output
    picResults.Cls
    'get degrees from textbox and assign to variable
    
    Temp = txttemp.Text
    
    
    'assign the correct pharse
    If Temp >= 180 Then
            Weather = "Too Hot! Crank the AC"
        ElseIf Temp >= 135 Then
            Weather = "Feeling Hot! Hot! Hot!"
        ElseIf Temp >= 100 Then
            Weather = "Average Desert Day,remember to bring water"
        ElseIf Temp >= 90 Then
            Weather = "Cool desert day, enjoy it while you can"
        ElseIf Temp >= 70 Then
            Weather = "Its cold for the desert"
        ElseIf Temp >= 0 Then
            Weather = "Where are you? Minnesota?"
        Else: picResults.Print "Error in temp value off the radar."
            Weather = " impossible."
              
    End If

    picResults.Print Temp; "Degrees fahrenheit"
    picResults.Print "the weather is "; Weather
    txttemp.Text = ""
End Sub





