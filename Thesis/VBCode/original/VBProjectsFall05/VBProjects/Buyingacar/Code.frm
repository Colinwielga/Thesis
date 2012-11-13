VERSION 5.00
Begin VB.Form Ford 
   BackColor       =   &H80000011&
   Caption         =   "Ford"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmsPush 
      Caption         =   "Push"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox picOutput 
      Height          =   2895
      Left            =   3480
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox txtInput 
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   8040
      Picture         =   "Code.frx":0000
      Top             =   8040
      Width           =   1140
   End
   Begin VB.Image Image11 
      Height          =   675
      Left            =   9360
      Picture         =   "Code.frx":07B6
      Top             =   8040
      Width           =   1065
   End
   Begin VB.Image Image10 
      Height          =   855
      Left            =   6720
      Picture         =   "Code.frx":0EDA
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Image Image9 
      Height          =   675
      Left            =   5400
      Picture         =   "Code.frx":2EA7
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   720
      Left            =   4080
      Picture         =   "Code.frx":3339
      Top             =   8040
      Width           =   1140
   End
   Begin VB.Image Image7 
      Height          =   675
      Left            =   2760
      Picture         =   "Code.frx":3AEF
      Top             =   8040
      Width           =   1065
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   1440
      Picture         =   "Code.frx":4213
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   120
      Picture         =   "Code.frx":61E0
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   9360
      Picture         =   "Code.frx":6672
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   8040
      Picture         =   "Code.frx":6E28
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6720
      Picture         =   "Code.frx":754C
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5400
      Picture         =   "Code.frx":9519
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Convertible 
      Height          =   720
      Left            =   4080
      Picture         =   "Code.frx":99AB
      Top             =   240
      Width           =   1140
   End
   Begin VB.Image SUV 
      Height          =   675
      Left            =   2760
      Picture         =   "Code.frx":A161
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image sedan 
      Height          =   855
      Left            =   1440
      Picture         =   "Code.frx":A885
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Truck 
      Height          =   675
      Left            =   120
      Picture         =   "Code.frx":C852
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      Caption         =   "Here are some of the cars that match your requirements:"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblstar 
      Caption         =   "************************************************"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lblStyles 
      Caption         =   "Trucks  Sedans  SUVs  Convertibles"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblStyle 
      Caption         =   "Which style of car are you interested in?"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "Ford"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(John Boruff's VB-project.vbp)
'Form Name : Ford(NewSnowForm.frm)
'Author: Katie Lee
'purpose of the form:  Since the user has selected a Brand, the purpose
                    'of this form is for the user to determine which style of car
                    'he/she needs depending on their preference
Option Explicit

Dim Style(1 To 39) As Integer

Private Sub cmsPush_Click()
Dim Make As String
Make = txtInput.Text
Dim CTR As Integer
For CTR = 1 To 39
    If Company(CTR) = "Ford" And Style(CTR) = Make Then
        picOutput.Print "Ford", Name(CTR), Price(CTR)
    End If
Next CTR
End
End Sub

