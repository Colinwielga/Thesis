VERSION 5.00
Begin VB.Form frmyourcountry 
   BackColor       =   &H00008000&
   Caption         =   "Where Does Your Favorite Country Stand?"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   Picture         =   "frmyourcountry.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00008000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008000&
      Height          =   5055
      Left            =   720
      ScaleHeight     =   4995
      ScaleWidth      =   9795
      TabIndex        =   1
      Top             =   3120
      Width           =   9855
   End
   Begin VB.CommandButton cmdfindout 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To Choose Your Winning Country "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   720
      Picture         =   "frmyourcountry.frx":38ADE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmyourcountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
    frmyourcountry.Hide
    frmwhichfact.Show
End Sub

Private Sub cmdclear_Click()
    picResult.Cls
End Sub
'declare all variables
Private Sub cmdfindout_Click()
    Dim InputCountry As String
    Dim CountryArray(1 To 100) As String
    Dim MedalArray(1 To 100) As Single
    Dim CTR As Integer
    Dim Pos As Integer
    
    'read file into two arrays
    Open App.Path & "\YourCountry.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, CountryArray(CTR), MedalArray(CTR)
    Loop
    Close #1
    'ask user for a country and determine if it matches with one in the list of gold medal winners
    'if it does make a match, print the result
    Pos = 0
    InputCountry = InputBox("Enter The Name Of Your Favorite Country (Use 3 Letter All Caps Abbreviation ex. USA)", "FavoriteCountry")
    For Pos = 1 To CTR
        Pos = Pos + 1
        If InputCountry = CountryArray(Pos) Then
            picResult.Print InputCountry; "Has Earned A Gold Medal."
        End If
    Next Pos
    'if it doesnt make a match, report that also
    If InputCountry <> CountryArray(Pos) Then
        picResult.Print InputCountry; , "Has Never Won A Medal In The 100 Meter Sprint. C'mon Pick A Winner!"
    End If
    
End Sub

