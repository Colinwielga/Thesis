VERSION 5.00
Begin VB.Form frmForm3 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   2910
   ClientTop       =   3075
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11145
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "First Select Your Gender..."
      Top             =   360
      Width           =   4335
   End
   Begin VB.ComboBox Combobox1 
      Height          =   315
      ItemData        =   "frmForm3.frx":0000
      Left            =   360
      List            =   "frmForm3.frx":000A
      TabIndex        =   6
      Text            =   "Show your gender and click enter....."
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Enter"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   7455
      Left            =   4800
      ScaleHeight     =   7395
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   360
      Width           =   6135
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   720
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1575
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lblCheckFemale 
      Caption         =   "Females click here - - - - ->"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblCheckMale 
      Caption         =   "Males click here - - - - - - - >"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "frmForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
frmForm3.Hide
frmFormFemales.Show
End Sub

Private Sub Check2_Click()
frmForm3.Hide
frmFormMales.Show
End Sub

Private Sub cmdContinue_Click()
Dim List As String
If Combobox1.Text = "Female" Then
    Open App.Path & "\Images\musclewoman.jpg" For Input As #1
        picResults.Cls
        picResults.Picture = LoadPicture(App.Path & "\Images\musclewoman.jpg")
Else
    Open App.Path & "\Images\muscleMan.jpg" For Input As #2
        picResults.Cls
        picResults.Picture = LoadPicture(App.Path & "\Images\muscleMan.jpg")
End If
Close #1
Close #2
End Sub

Private Sub cmdQuit_Click()
End
End Sub

