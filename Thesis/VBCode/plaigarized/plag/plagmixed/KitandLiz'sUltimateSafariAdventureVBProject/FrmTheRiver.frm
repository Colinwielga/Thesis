VERSION 5.00
Begin VB.Form FrmTheRiver 
   BackColor       =   &H00FF0000&
   Caption         =   "The Rushing River"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Find the average age!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   2655
   End
   Begin VB.PictureBox picImage 
      Height          =   4335
      Left            =   6360
      Picture         =   "FrmTheRiver.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Safari Headquarters"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   2880
      ScaleHeight     =   4275
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton cmdNumerical 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Numerically!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdAlphabatize 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Alphabatize them!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdEnterData 
      BackColor       =   &H00C0FFC0&
      Caption         =   " See all the animals names and ages that you passed on the river"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblRiverbend 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "play just around the riverbend.."
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   5160
      Width           =   6015
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   495
      Left            =   6360
      OleObjectBlob   =   "FrmTheRiver.frx":1C174
      SourceDoc       =   "M:\CS130\Kit and Liz's Ultimate Safari Adventure VB Project\Unknown Artist\Unknown Album (2-24-2010 6-02-05 PM)\05 Track 5.wma"
      TabIndex        =   7
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblWelcometotheRiver 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome to the River...."
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "FrmTheRiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Great Safari Adventure
'Frm The River
'Kit and Liz Chambers
'February 22nd 2010
'Objective: The purpose of this form is twofold
                   'Display a Picture
                   'Read an Array and Sort it(about animals names and ages)
                   
Private Sub cmdAlphabatize_Click()
 'Declares Variables
Dim Ages(1 To 10) As Integer
Dim Pass As Integer, Pos As Integer
Dim TempX As String, TempY As Integer
Dim I As Integer, CTR As Integer
Dim Animals(1 To 10) As String
 'Clears the picture box

picResults.Cls
 'Opens the file
Open App.Path & "\JungleAnimals.txt" For Input As #1
 'Reads the file
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Animals(CTR), Ages(CTR)
    picResults.Print Animals(CTR); Tab(35); Ages(CTR)
Loop


         'Sorts the file according to the alphabetical order of the animals
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Animals(Pos) > Animals(Pos + 1) Then
          'Arranges the ages accordingly
                    TempY = Ages(Pos)
                    Ages(Pos) = Ages(Pos + 1)
                    Ages(Pos + 1) = TempY
                    TempX = Animals(Pos)
                    Animals(Pos) = Animals(Pos + 1)
                    Animals(Pos + 1) = TempX
                End If
            Next Pos
        Next Pass
        'Prints the results
            picResults.Print "Animals"; Tab(35); "Ages"
            picResults.Print "*******************************"

        For I = 1 To CTR
            picResults.Print Animals(I); Tab(35); Ages(I)
        Next I
        Close #1
End Sub

Private Sub cmdAverage_Click()
'establish variables
Dim Ages(1 To 15) As Integer
Dim Animals(1 To 15) As String
Dim Sum As Integer
Dim CTR As Integer
Dim Average As Single

'get data from file
Open App.Path & "\JungleAnimals.txt" For Input As #1
'read file into two paralell arrays
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Animals(CTR), Ages(CTR)
    Sum = Sum + Ages(CTR)
    Average = Sum / CTR
   'calculate the average
Loop
    'print results
    picResults.Print "The Average Age of your "; CTR; " Animals is "; Round(Average)
    'close file
Close #1
End Sub

Private Sub cmdEnterData_Click()
 'Declares the Variables
Dim Animals(1 To 15) As String
Dim Ages(1 To 15) As Integer
Dim CTR As Integer
 'Clears the picture box
picResults.Cls
    'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\JungleAnimals.txt" For Input As #1
    
    'print the header info
    picResults.Print "Animals"; Tab(35); "Ages"
    picResults.Print "**********************************"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = 1 + CTR
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Animals(CTR), Ages(CTR)
        
        picResults.Print Animals(CTR); Tab(35); Ages(CTR)
        
    Loop
    
    'closes the file
    Close #1
    
    

End Sub


Private Sub cmdNumerical_Click()
 'Declares Variables
Dim Animals(1 To 15) As String
Dim Ages(1 To 15) As Integer
Dim Pass As Integer, Pos As Integer
Dim Temp1 As String, Temp2 As Integer
Dim I As Integer, CTR As Integer
 'Clears the picture box
picResults.Cls
 'Opens the file
Open App.Path & "\JungleAnimals.txt" For Input As #1
 'Reads the file
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Animals(CTR), Ages(CTR)
    
Loop
 'Sorts the file according to the numerical order of the ages
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Ages(Pos) > Ages(Pos + 1) Then
   'arranges the aminals along with the ages
            Temp1 = Animals(Pos)
            Animals(Pos) = Animals(Pos + 1)
            Animals(Pos + 1) = Temp1
  'Arranges the ages accordingly
            Temp2 = Ages(Pos)
            Ages(Pos) = Ages(Pos + 1)
            Ages(Pos + 1) = Temp2
        End If
    Next Pos
Next Pass
'Prints the results
    picResults.Print "Animals"; Tab(35); "Ages"
    picResults.Print "*******************************"
'print loop
For I = 1 To CTR
    picResults.Print Animals(I); Tab(35); Ages(I)
Next I
'closes the file
Close #1
End Sub

Private Sub cmdReturn_Click()
    FrmTheRiver.Hide 'hides river page from user
    FrmWelcome.Show 'shows main page to user
End Sub



Private Sub picImage_Click()

picImage.Show LitteRiver.jpg

End Sub
