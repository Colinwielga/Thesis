VERSION 5.00
Begin VB.Form LevelSelect 
   BackColor       =   &H00FF0000&
   Caption         =   "Select your skiing level"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton search 
      Caption         =   "Search for a specific ski company"
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   5280
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   2640
      Picture         =   "LevelSelect2.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton cmdBegniner 
      Caption         =   "Beginner"
      Height          =   855
      Left            =   7320
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdExpert 
      BackColor       =   &H000000C0&
      Caption         =   "Expert"
      Height          =   855
      Left            =   240
      MaskColor       =   &H80000000&
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Select your skiing ability"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "LevelSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: LevelSelect (LevelSelect2.frm)
'Author David Meacham
'Date Written: Wednesday, October 22
'Purpose of form: Allows user to select their ski level and move
    'them to that form. Also allows them to search for a
    'specific brand to see if it is sold by me
    
Option Explicit

Private Sub cmdBegniner_Click()
'Goes to the Beginner ski form
LevelSelect.Hide
BeginnerSki.Show
End Sub

Private Sub cmdExpert_Click()
'Goes to the expert ski form
LevelSelect.Hide
ExpertSki.Show
End Sub

Private Sub Form_Load()
'Opens file immediatley for whole project
strPath = "N:\CS130\handin\dbmeacham\"
End Sub

Private Sub search_Click()
'Search's for a specific company and tells the shopper if we sell that brand
Dim found As Boolean
Dim company As String
Dim skicompany(1 To 5) As String
found = False
Dim names As String
names = strPath & "company.txt"         'creates a path to open company.txt
Dim i As Integer
    Open names For Input As #1
        For i = 1 To 5
            Input #1, skicompany(i)     'inputs ski company names from the file to the array
        Next i
company = InputBox("Enter a ski company to see if we sell their products")      'asks user to input a company name
i = 0
    Do Until found Or i = 5             'searches for the name
        i = i + 1
        If company = skicompany(i) Then
            found = True
        End If
    Loop
        If found = True Then
            MsgBox "Yes, we sell " & skicompany(i) & "'s products."             'prints company name if found
        Else
            MsgBox "I'm sorry.  We do not sell any of " & company & "'s products"   'prints a message telling user that name was not found
        End If
Close #1
End Sub
