VERSION 5.00
Begin VB.Form frmSched 
   BackColor       =   &H00000000&
   Caption         =   "Check to see if your schedule works!"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Restart!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdWeekdays 
      Caption         =   "Weekdays"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdWeekends 
      Caption         =   "Weekends"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdMornings 
      Caption         =   "Mornings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAfternoon 
      Caption         =   "Afternoons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdNights 
      Caption         =   "Nights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nights, Afternoons, Mornings, Weekends, Weekdays As Boolean

Private Sub cmdAfternoon_Click() 'sets value to true for calculations
    Afternoons = True
    cmdAfternoon.Visible = False
End Sub

Private Sub cmdBack_Click()
    frmSched.Cls
    frmSched.Hide
    frmSecond.Show
End Sub

Private Sub cmdCalculate_Click()
    frmSched.Cls
    frmSched.Print "Your options according to your Schedule are:"
    If Weekends = True And Weekdays = True Then ' searches for values to display
        If Nights = True Then
            frmSched.Print "Whole week nights:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Football, Track"
            frmSched.Print , "Baseball, Hockey, Soccer"
        End If
        If Afternoons = True Then
            frmSched.Print "Whole week afternoons:"
                frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Track"
                frmSched.Print , "Baseball, Soccer, Golf"
        End If
        If Mornings = True Then
            frmSched.Print "Whole week mornings:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing,"
            frmSched.Print , "Hockey, Soccer"
        End If
   
    ElseIf Weekdays = True Then ' searches for values to display
        If Nights = True Then
            frmSched.Print "Weekday nights:"
            frmSched.Print , "Weight lifting, Walking, Video games, Dancing, Indoor Track"
        End If
        If Afternoons = True Then
            frmSched.Print "Weekday afternoons:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Golf"
        End If
        If Mornings = True Then
            frmSched.Print "Weekday mornings:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Golf"
        End If
    ElseIf Weekends = True Then ' searches for values to display
        If Nights = True Then
            frmSched.Print "Weekend nights:"
              frmSched.Print , "Walking, Weight Lifting, Intramural Basketball,Video Games, Dancing"
        End If
        If Afternoons = True Then
            frmSched.Print "Weekend afternoons:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Golf"
        End If
        If Mornings = True Then
            frmSched.Print "Weekend mornings:"
            frmSched.Print , "Walking, Weight Lifting, Video Games, Dancing, Golf"
        End If
    Else
        frmSched.Print "Please select a time of the week (weekends or weekdays... or both!)"
    End If
End Sub

Private Sub cmdClear_Click() 'resets all variables to false and clears form
    Nights = False
    Afternoons = False
    Mornings = False
    Weekdays = False
    Weekends = False
    frmSched.Cls
    cmdNights.Visible = True
    cmdAfternoon.Visible = True
    cmdMornings.Visible = True
    cmdWeekdays.Visible = True
    cmdWeekends.Visible = True
End Sub

Private Sub cmdMornings_Click() 'sets value to true for calculations
    Mornings = True
    cmdMornings.Visible = False
End Sub

Private Sub cmdNights_Click() 'sets value to true for calculations
    Nights = True
   cmdNights.Visible = False
End Sub
 
Private Sub cmdWeekdays_Click() 'sets value to true for calculations
    Weekdays = True
    cmdWeekdays.Visible = False
End Sub

Private Sub cmdWeekends_Click() 'sets value to true for calculations
    Weekends = True
    cmdWeekends.Visible = False
End Sub

Private Sub Form_Load() 'sets all variable to false, to prevent data from being displayed
Nights = False
Afternoons = False
Mornings = False
Weekdays = False
Weekends = False
End Sub
