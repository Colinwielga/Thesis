VERSION 5.00
Begin VB.Form frmbestcountries 
   BackColor       =   &H00008000&
   Caption         =   "Which Country Has The Most Gold Medals?"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   Picture         =   "frmbestcountries.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
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
      Height          =   2775
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008000&
      Height          =   4455
      Left            =   2280
      ScaleHeight     =   4395
      ScaleWidth      =   6555
      TabIndex        =   2
      Top             =   3720
      Width           =   6615
   End
   Begin VB.CommandButton cmdmedals 
      Caption         =   "Most Wins"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2280
      Picture         =   "frmbestcountries.frx":38ADE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Click Below For Top Country"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmbestcountries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
    frmbestcountries.Hide
    frmwhichfact.Show
End Sub

Private Sub cmdmedals_Click()
    'declare all variables
    Dim CountryArray(1 To 30) As Single
    Dim CTR As Integer
    Dim Pass As Integer
    Dim Comp As Integer
    Dim Pos As Integer
    Dim Temp As Integer
    Dim USA As Integer
    Dim SAG As Integer
    Dim GBR As Integer
    Dim CAN As Integer
    Dim GER As Integer
    Dim URS As Integer
    Dim TRI As Integer
    USA = 0
    SAG = 0
    GBR = 0
    CAN = 0
    GER = 0
    URS = 0
    TRI = 0
    
    'read this file into one array
    Open App.Path & "\medals.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, CountryArray(CTR)
    Loop
    Close #1
    'total all gold medalist countries and print the outcome
    For Pos = 1 To CTR
        If CountryArray(Pos) = "USA" Then
            USA = USA + 1
        End If
            If CountryArray(Pos) = "SAG" Then
            SAG = SAG + 1
        End If
        If CountryArray(Pos) = "GBR" Then
            GBR = GBR + 1
        End If
        If CountryArray(Pos) = "CAN" Then
            CAN = CAN + 1
        End If
        If CountryArray(Pos) = "GER" Then
            GER = GER + 1
        End If
        If CountryArray(Pos) = "URS" Then
            URS = URS + 1
        End If
        If CountryArray(Pos) = "TRI" Then
            TRI = TRI + 1
        End If
    Next Pos
       
    picResult.Print "USA Has", ; USA; , "Gold Medals"
    picResult.Print "SAG Has", ; SAG; , "Gold Medals"
    picResult.Print "GBR Has", ; GBR; , "Gold Medals"
    picResult.Print "CAN Has", ; CAN; , "Gold Medals"
    picResult.Print "GER Has", ; GER; , "Gold Medals"
    picResult.Print "URS Has", ; URS; , "Gold Medals"
    picResult.Print "TRI Has", ; TRI; , "Gold Medals"
    
End Sub
