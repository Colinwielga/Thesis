VERSION 5.00
Begin VB.Form frmsick 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   Picture         =   "frmsick.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   7320
      Width           =   2895
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   9600
      ScaleHeight     =   5835
      ScaleWidth      =   7755
      TabIndex        =   6
      Top             =   2880
      Width           =   7815
   End
   Begin VB.CommandButton cmdsick 
      Caption         =   "Should I See A Doctor?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4200
      TabIndex        =   0
      Top             =   6600
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmsick.frx":5C4B
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5040
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3.  Take thermometer out after time has expired and see where the top of the red line lies"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   9960
      TabIndex        =   4
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2.  Keep in mouth for at least two minutes"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1.  Take thermometer and place it in mouth"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Directions for taking a temperature:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmsick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'Form frmsick
'Joel Coleman
'March 29, 2008
'To help people figure out if they should see a doctor by looking at their symptoms
Option Explicit

Private Sub cmdsick_click()
'Print the symptoms of fatigue as found from http://www.cancerhelp.org.uk/help/default.asp?page=10269
picresults.Print "SYMPTOMS OF FATIGUE:"
picresults.Print "Lack of energy - you may just want to stay in bed all day "
picresults.Print "Feeling ‘I just cannot be bothered to do much’ "
picresults.Print "Problems sleeping"
picresults.Print "Finding it hard to get up in the morning"
picresults.Print "Feeling; anxious Or depressed"
picresults.Print "Pain in your muscles - you may find it hard to climb stairs or walk short distances"
picresults.Print "Being short of breath after doing small tasks, like having a shower or making your bed"
picresults.Print "Finding it hard to concentrate, even if watching TV or talking to a good friend"
picresults.Print "Being unable to think clearly or make decisions easily"
picresults.Print "picresults.print Loss of interest in doing things you usually enjoy"
picresults.Print "Negative feelings about yourself and others"

'Declaring variables
Dim temp As Single, throat As String, fatigue As String, CTR As Single
temp = InputBox("Input Current Temperature in Farenheit")
throat = InputBox("Are your lymph nodes inflammed?  Yes/No?")
fatigue = InputBox("Are you feeling fatigued?  Yes/No?")
CTR = 0

'If then statement to determine if given temp is abnormal
If temp > 99.5 Then
MsgBox ("Your temperature exceeds normal standards")
'Counts up how many questions are answered in a negative way
CTR = CTR + 1
Else
End If

'If then statement to determine if given temp is abnormal
If temp < 96 Then
'Counts up how many questions are answered in a negative way
CTR = CTR + 1
Else
End If

'If then statement to determine if throat is swollen
If throat = "Yes" Then
MsgBox ("Drink large amount of liquids, and consider using cough drops!")
'Counts up how many questions are answered in a negative way
CTR = CTR + 1
Else
End If

'If then statement to determine if there is fatigue
If fatigue = "Yes" Then
MsgBox ("You better keep sleeping and hope you don't have Mono!")
'Counts up how many questions are answered in a negative way
CTR = CTR + 1
Else
End If

'Select case to print the total being counted up, the more negative answers the better chance the person is sick
Select Case CTR
    Case 0
    MsgBox ("There is nothing to be worried about!")
    Case 1
    MsgBox ("Get good rest and you should be fine")
    Case 2
    MsgBox ("You should consider seeing a doctor")
    Case 3
    MsgBox ("You should see a doctor as soon as possible")
End Select


End Sub

Private Sub Command1_Click()
frmsick.Hide
frmMainpage.Show
End Sub

