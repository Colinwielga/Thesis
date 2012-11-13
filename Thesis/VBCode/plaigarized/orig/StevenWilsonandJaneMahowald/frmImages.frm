VERSION 5.00
Begin VB.Form FrmImages 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWeight 
      Caption         =   "Organize by Weight"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdLength 
      Caption         =   "Organize by Length"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Search for Image"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Menu"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   7560
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   2415
   End
   Begin VB.PictureBox PicImage 
      Height          =   7335
      Left            =   3000
      ScaleHeight     =   7275
      ScaleWidth      =   8355
      TabIndex        =   2
      Top             =   240
      Width           =   8415
   End
   Begin VB.PictureBox PicDino 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dinosaur(1 To 100) As String, Length(1 To 100) As Double
Dim Weight(1 To 100) As Double, Fossil(1 To 100) As String, Real(1 To 100) As String, CTR As Integer
Dim Dino As String, Found As Boolean, RealisticImage As String, PictureType As String, ImageType As String
Dim carnivore As String, herbivore As String, Pos As Integer, TempLength As Double, TempWeight As Double
Dim TempName As String, Pass As Integer
'takes user back to the main menu

Private Sub cmdMain_Click()
 frmMain.Visible = True
    frmGame.Visible = False
    frmFacts.Visible = False
    FrmImages.Visible = False
End Sub
'quits program
Private Sub cmdQuit_Click()
End
End Sub


'this button processes the data from a text file and reads it into an array.
'it then prints the names of the dinosaurs avalible so the user can pick a dinosaur from the list
Private Sub cmdStart_Click()
PicDino.Cls
Open App.Path & "/Dinosaurs.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dinosaur(CTR), Length(CTR), Weight(CTR), Fossil(CTR), Real(CTR)
        PicDino.Print Dinosaur(CTR)
    Loop
    Close #1
End Sub
Private Sub cmdShow_Click()
'This button allows the user to search though the dinosaur image database and pick
'either an image of a fossil or an artistic representation of the real dinosaur.
        Dino = InputBox("What Dinosaur Would You Like To See?")
        PictureType = InputBox("Fossil or Realistic?")
        Found = False
        CTR = 0
        
        Do While (Found = False And CTR < 11)
            CTR = CTR + 1
            If LCase(Dinosaur(CTR)) = LCase(Dino) Then
            Found = True
            End If
        Loop
         If Found = True Then
                If LCase(PictureType) = LCase("Fossil") Then
                    PicImage.Cls
                    PicImage.Picture = LoadPicture(App.Path & Fossil(CTR))
                ElseIf LCase(PictureType) = LCase("Realistic") Then
                    PicImage.Cls
                    PicImage.Picture = LoadPicture(App.Path & Real(CTR))
                End If
        ElseIf Found = False Then
            MsgBox "Sorry, There has been an error", , "ERROR"
        End If
End Sub

Private Sub cmdLength_Click()
'This button allows the user to orginze the dinosaurs based on length
PicDino.Cls
Open App.Path & "/Dinosaurs.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dinosaur(CTR), Length(CTR), Weight(CTR), Fossil(CTR), Real(CTR)
    Loop
    Close #1
For Pass = 1 To CTR - 1
        For Pos = 1 To (CTR - Pass)
            If Length(Pos) > Length(Pos + 1) Then
                TempLength = Length(Pos)
                Length(Pos) = Length(Pos + 1)
                Length(Pos + 1) = TempLength
                
                TempName = Dinosaur(Pos)
                Dinosaur(Pos) = Dinosaur(Pos + 1)
                Dinosaur(Pos + 1) = TempName
                End If
            Next Pos
    Next Pass
    PicDino.Print "Orginized By Length"
    For Pos = 1 To CTR
        PicDino.Print Dinosaur(Pos); " are "; Length(Pos); " ft. long"
    Next Pos
End Sub

Private Sub cmdWeight_Click()
'This allows the user to organize the dinosaurs based on weight.
PicDino.Cls
Open App.Path & "/Dinosaurs.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dinosaur(CTR), Length(CTR), Weight(CTR), Fossil(CTR), Real(CTR)
    Loop
    Close #1
For Pass = 1 To CTR - 1
        For Pos = 1 To (CTR - Pass)
            If Weight(Pos) > Weight(Pos + 1) Then
            TempWeight = Weight(Pos)
            Weight(Pos) = Weight(Pos + 1)
            Weight(Pos + 1) = TempWeight
            
            TempName = Dinosaur(Pos)
            Dinosaur(Pos) = Dinosaur(Pos + 1)
            Dinosaur(Pos + 1) = TempName
            End If
        Next Pos
    Next Pass
For Pass = 1 To CTR - 1
        For Pos = 1 To (CTR - Pass)
            If Weight(Pos) > Weight(Pos + 1) Then
                TempWeight = Weight(Pos)
                Weight(Pos) = Weight(Pos + 1)
                Weight(Pos + 1) = TempWeight
                
                TempName = Dinosaur(Pos)
                Dinosaur(Pos) = Dinosaur(Pos + 1)
                Dinosaur(Pos + 1) = TempName
            End If
            Next Pos
    Next Pass
    PicDino.Print "Orginized By Weight"
    For Pos = 1 To CTR
        PicDino.Print Dinosaur(Pos); " are "; Weight(Pos); " tons"
    Next Pos
End Sub
