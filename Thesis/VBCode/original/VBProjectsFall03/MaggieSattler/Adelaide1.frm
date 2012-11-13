VERSION 5.00
Begin VB.Form Adelaide1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton GoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   3
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Image"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   4920
      Width           =   3135
   End
   Begin VB.PictureBox Imagebox 
      Height          =   3375
      Left            =   1680
      ScaleHeight     =   3315
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Maggie Sattler"
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Adelaide Images"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Adelaide1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Just an FYI: Please refer to this form as a model for my other forms with names
'of Australian cities for the sake of saving myself from redundance.  Thank you.

'declare all variables on this form
'declare P as a counter to keep track of number of clicks being made
Dim P As Integer




Private Sub Command1_Click()
    'with each successive click, make sure the variable goes in order to next image
    P = P + 1
    'if user reaches the end of the available pictures, send user back to the first image
    If P = 6 Then P = 1
    'for each click, send the user to a different image.  Once the maximum amount of
    'images have been shown, the user will be allowed to continue through the program
    'by being sent through the order of images again.
    If P = 1 Then
            Imagebox.Picture = LoadPicture(Australia1.Path & "Adelaide1.jpg")
        ElseIf P = 2 Then
            Imagebox.Picture = LoadPicture(Australia1.Path & "Adelaide2.jpg")
        ElseIf P = 3 Then
            Imagebox.Picture = LoadPicture(Australia1.Path & "Adelaide3.jpg")
        ElseIf P = 4 Then
            Imagebox.Picture = LoadPicture(Australia1.Path & "Adelaide4.jpg")
        ElseIf P = 5 Then
            Imagebox.Picture = LoadPicture(Australia1.Path & "Adelaide5.jpg")
    End If
            
End Sub

Private Sub GoBack_Click()
    'send user back to first form
    Australia1.Show
    Adelaide1.Hide
End Sub
