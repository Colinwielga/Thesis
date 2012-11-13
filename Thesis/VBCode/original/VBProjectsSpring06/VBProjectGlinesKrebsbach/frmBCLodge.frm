VERSION 5.00
Begin VB.Form frmBCLodge 
   BackColor       =   &H000080FF&
   Caption         =   "Beaver Creek Lodging"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search for a lodge"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Beaver Creek"
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
      Left            =   120
      TabIndex        =   7
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Learn more about Beaver Creek's Resorts"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   5880
      Width           =   3135
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
      Height          =   3375
      Left            =   2640
      ScaleHeight     =   3315
      ScaleWidth      =   6555
      TabIndex        =   4
      Top             =   1800
      Width           =   6615
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
      Left            =   600
      TabIndex        =   3
      Top             =   4200
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
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "Display Top 8 Resorts"
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
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   10680
      Width           =   2775
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
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblBC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beaver Creek Lodging"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmBCLodge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBCLodge(frmBCLodge.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to look up the top resorts
'available for the ski resort. this form allows the user to sort the resorts
'alphabetically, and by price. this form loads information from a text file using
'notepad and loads it into an array. this form also allows the user acces to another
'form which goes into more depth about each resort that we loaded.
Option Explicit
Dim I As Integer
Dim resorts(1 To 8) As String
Dim prices(1 To 8) As Single
Dim Temp As String
Dim CTR As Integer

Private Sub cmddisplay_Click()
    picResults.Cls   'clears the picture box if there is any information in it.
    Open App.Path & "\BClodge.txt" For Input As #1   'opens file BCLodge.txt
    For I = 1 To 8
        Input #1, resorts(I), prices(I)
    Next I    'goes to the next I, until 8
    picResults.Print "These are the Resorts Beaver Creek has to offer:"
    picResults.Print
    For I = 1 To 8
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I))
    Next I
    Close #1 'closes the text file
End Sub



Private Sub cmdprice_Click()
    picResults.Cls 'clears the picture box if there is any info in it
    Dim F As Integer, Pass As Integer
    Dim resorts(1 To 8) As String
    Dim prices(1 To 8) As Single
    Dim Temp As Single
    Dim Temp2 As String
    Dim CTR As Integer
    Dim Size As Integer
    Open App.Path & "\BClodge.txt" For Input As #1   'opens the BClodge.txt file
    CTR = 0
        Do Until EOF(1)    'Loads Resorts until the end of the file
        CTR = CTR + 1   'Everytime the program loops through the CTR in increments of one, goes to the next Resort.
        Input #1, resorts(CTR), prices(CTR)
    Loop
    Close #1     'close file when done reading the array
    Size = CTR
    For Pass = 1 To Size - 1     'sorts the Resorts array into numerical order
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
    For I = 1 To 8
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I)) 'prints all 8 Resorts and prices from the file
    Next I
End Sub

Private Sub cmdSearch_Click()
    Dim found As Boolean
    Dim I As Integer, N As String
    Dim resorts(1 To 100) As String
    Dim pos As Integer
    pos = 0
    Open App.Path & "\BClodge.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, resorts(pos)
    Loop
    Close #1
    
    N = InputBox("enter the name of the resort you would like to look up", "Resort")
    pos = 0
    found = False
    
    Do While found = False And pos < 100
        pos = pos + 1
        If N = resorts(pos) Then
        found = True
        End If
    Loop
    
    If found = True Then
        MsgBox "yes we have more information about that resort please click on learn more about Beaver Creek resorts", , "Search"
    Else
        MsgBox "sorry we dont have information on the resort you requested, but there are other nice ones to choose from that have many great features.", , "Sorry"
    End If
   
End Sub

Private Sub cmdsort_Click()
picResults.Cls 'clears all info in the picture box
    Dim F As Integer, Pass As Integer
    Dim resorts(1 To 8) As String
    Dim prices(1 To 8) As Single
    Dim Temp As String
    Dim Temp2 As Single
    Dim CTR As Integer
    Dim Size As Integer
    Open App.Path & "\BClodge.txt" For Input As #1   'opens the BClodge.txt file
    CTR = 0
        Do Until EOF(1)    'Loads Resorts until the end of the file
        CTR = CTR + 1   'Everytime the program loops through the CTR in increments of one, goes to the next resort.
        Input #1, resorts(CTR), prices(CTR)
    Loop
    Close #1     'close file when done reading the array
    Size = CTR
    For Pass = 1 To Size - 1     'sorts the Resorts array into alphabetical order (A-Z)
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
    For I = 1 To 8
        picResults.Print resorts(I); Tab(45); FormatCurrency(prices(I)) 'prints all 8 Resorts and prices from the file
    Next I

End Sub

Private Sub cmdMore_Click()
    frmBCLodge.Hide 'hides the BClodge form
    frmBCresorts.Show 'shows the BCresorts form

End Sub

Private Sub cmdback_Click()
    frmBCLodge.Hide 'hides the BClodge form
    frmBeaver.Show 'shows the Beaver form
End Sub


