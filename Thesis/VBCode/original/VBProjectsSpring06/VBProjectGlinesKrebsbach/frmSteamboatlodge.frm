VERSION 5.00
Begin VB.Form frmSteamboatlodge 
   BackColor       =   &H000080FF&
   Caption         =   "Steamboat Lodging"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display Top 5 Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdsort 
      Caption         =   "Sort Lodging Alphabetically"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdprice 
      Caption         =   "Sort Lodging by Price"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2280
      ScaleHeight     =   4755
      ScaleWidth      =   8475
      TabIndex        =   2
      Top             =   1560
      Width           =   8535
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Learn more about Steamboat's Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   1
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Steamboat"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblsteamboat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Steamboat Lodging"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label lblper 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "These prices are based on per night"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   10320
      Width           =   2775
   End
End
Attribute VB_Name = "frmSteamboatlodge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmSteamboatlodge(frmSteamboatlodge.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to look up the top resorts
'available for the ski resort. this form allows the user to sort the resorts
'alphabetically, and by price. this form loads information from a text file using
'notepad and loads it into an array. this form also allows the user acces to another
'form which goes into more depth about each resort that we loaded.
Option Explicit
Dim I As Integer
Dim resorts(1 To 5) As String
Dim prices(1 To 5) As Single
Dim Temp As String
Dim CTR As Integer



Private Sub cmdback_Click()
    frmSteamboatlodge.Hide
    frmSteamboat.Show

End Sub

Private Sub cmddisplay_Click()
   
    picResults.Cls   'clears the picture box if there is any information in it.
    Open App.Path & "\steamboatlodge.txt" For Input As #1   'opens file steamboatlodge.txt
    For I = 1 To 5
        Input #1, resorts(I), prices(I)
    Next I    'goes to the next CTR, until 5
    picResults.Print "These are the Resorts Keystone has to offer:"
    picResults.Print
    For I = 1 To 5
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I))

    Next I
    Close #1
End Sub



Private Sub cmdMore_Click()
    frmSteamboatlodge.Hide
    frmSteamboatresorts.Show
End Sub

Private Sub cmdprice_Click()
picResults.Cls
    Dim F As Integer, Pass As Integer
    Dim resorts(1 To 5) As String
    Dim prices(1 To 5) As Single
    Dim Temp As Single
    Dim Temp2 As String
    Dim CTR As Integer
    Dim Size As Integer
    Open App.Path & "\steamboatlodge.txt" For Input As #1   'opens the Steamboatlodge.txt file
    CTR = 0
        Do Until EOF(1)    'Loads resorts and prices until the end of the file
        CTR = CTR + 1   'Everytime the program loops through the CTR in increments of one, goes to the next state.
        Input #1, resorts(CTR), prices(CTR)
    Loop
    Close #1     'close file when done reading the array
    Size = CTR
    For Pass = 1 To Size - 1     'sorts the states array into numerical order
        For CTR = 1 To Size - Pass
            If prices(CTR) > prices(CTR + 1) Then
                Temp = prices(CTR)
                prices(CTR) = prices(CTR + 1)
                prices(CTR + 1) = Temp
                Temp2 = resorts(CTR)
                resorts(CTR) = resorts(CTR + 1)
                resorts(CTR + 1) = Temp2
            End If
        Next CTR
    Next Pass
    For I = 1 To 5
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I)) 'prints all 5 resorts from the file
    Next I
End Sub

Private Sub cmdsort_Click()
picResults.Cls
    Dim F As Integer, Pass As Integer
    Dim resorts(1 To 5) As String
    Dim prices(1 To 5) As Single
    Dim Temp As String
    Dim Temp2 As Single
    Dim CTR As Integer
    Dim Size As Integer
    Open App.Path & "\steamboatlodge.txt" For Input As #1   'opens the Steamboatlodge.txt file
    CTR = 0
        Do Until EOF(1)    'Loads resorts until the end of the file
        CTR = CTR + 1   'Everytime the program loops through the CTR in increments of one, goes to the next state.
        Input #1, resorts(CTR), prices(CTR)
    Loop
    Close #1     'close file when done reading the array
    Size = CTR
    For Pass = 1 To Size - 1     'sorts the resorts array into alphabetical order (A-Z)
        For CTR = 1 To Size - Pass
            If resorts(CTR) > resorts(CTR + 1) Then
                Temp = resorts(CTR)
                resorts(CTR) = resorts(CTR + 1)
                resorts(CTR + 1) = Temp
                Temp2 = prices(CTR)
                prices(CTR) = prices(CTR + 1)
                prices(CTR + 1) = Temp2
            End If
        Next CTR
    Next Pass
    For I = 1 To 5
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I)) 'prints all 5 resorts from the file
    Next I

End Sub
