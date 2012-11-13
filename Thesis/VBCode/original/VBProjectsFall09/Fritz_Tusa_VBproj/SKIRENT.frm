VERSION 5.00
Begin VB.Form FormRent 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtotitlerent 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   4680
      ScaleHeight     =   6075
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   4080
      Width           =   8295
   End
   Begin VB.CommandButton cmdGetppl 
      BackColor       =   &H00FF8080&
      Caption         =   "START RENTING"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "SKI EQUIPMENT RENTAL INFORMATION"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   $"SKIRENT.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP
'SKI RENTAL
'HOLLIS FRITTS
'8-18
'THIS IS INFORMATION ON SKI EQUIPMENT RENTAL

Option Explicit
'dim all variables
Dim People As Integer, Skis As Integer, Item As String, daystayed As Integer
Private Sub cmdGetppl_Click()

'establish a function for running total
runningtotal = runningtotal + Skis

'make an input box to ask user for the number of days staying
daystayed = InputBox("Enter the Number of Days You Will Be Staying")
'make an input box to ask user for amount of people renting
People = InputBox("Enter The Number Of People That Will Be Renting Equipment")

'set runningTotal to zero
runningtotal = 0

'make a nested If statement to ask user what he/she wants to rent, and add it to the runningTotal
    If People = 1 Then
        Skis = InputBox("Enter What You Wish To Rent (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
   
    ElseIf People = 2 Then
        Skis = InputBox("Enter What You Wish To Rent (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
               
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
              
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
      
    ElseIf People = 3 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
            
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        
    ElseIf People = 4 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
       
    ElseIf People = 5 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        
            runningtotal = runningtotal * 0.8
       
    ElseIf People = 6 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 6")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
            runningtotal = runningtotal * 0.8
   
    ElseIf People = 7 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 6")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 7")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
    
            runningtotal = runningtotal * 0.8

    ElseIf People = 8 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 6")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 7")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 8")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
            runningtotal = runningtotal * 0.8

    ElseIf People = 9 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 6")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 7")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 8")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 9")
    
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
    
            runningtotal = runningtotal * 0.8
    
    
    ElseIf People = 10 Then
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 1")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 2")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 3")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 4")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 5")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 6")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 7")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 8")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 9")
        
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
        Skis = InputBox("Enter What You Wish To Rent  (1-3)                                                                                                                                                                                                                     1) Junior Skis            2) Adult Skis              3) Snowboard", , "Person 10")
            
            If Skis = 1 Then
                runningtotal = runningtotal + 20
                Item = "Junior Skis"
            ElseIf Skis = 2 Then
                runningtotal = runningtotal + 40
                Item = "Adult Skis"
            ElseIf Skis = 3 Then
                runningtotal = runningtotal + 40
                Item = "Snowboard"
            End If
            
            runningtotal = runningtotal * 0.8
    End If
'Multiply the runningTotal by the number of days stayed
runningtotal = daystayed * runningtotal
'make a header'
picResults.Print "Total for"; People; "Person(s)"
picResults.Print "***********************************************"

'print what items he/she chose and the total of the items
picResults.Print FormatCurrency(runningtotal)

End Sub

Private Sub cmdtotitlerent_Click()
FormRent.Hide
Title.Show
End Sub

