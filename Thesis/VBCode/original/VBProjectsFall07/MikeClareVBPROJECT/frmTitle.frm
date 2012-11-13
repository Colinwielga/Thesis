VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FF0000&
   Caption         =   "Alien Invasion"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   2
      Top             =   3960
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   1920
      Width           =   9135
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFF00&
      Caption         =   "CLICK ME!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
    MsgBox ("So the world is pretty much destroyed, your house is a wreck, and you'll die if you stay here.  You decide to go outside in your yard and explore."), , ("The world is destroyed")
    
    
    frmStart.Hide   'shows the yard form
    frmYard.Show
    
    MsgBox ("You see a few items laying in the yard.  Choose one to pick up and help you on your way."), , ("Pick up an item")
    
End Sub

Private Sub cmdStart_Click()
          
          
    Open App.Path & "\Life.txt" For Input As #1   'load the file Life.txt
    cmdContinue.Enabled = True
    cmdStart.Enabled = False
    
    CTR = 0
    
    Do Until EOF(1)     'make the array
        CTR = CTR + 1
        Input #1, posArray(CTR), lifeArray(CTR), hpArray(CTR), moneyArray(CTR), attackArray(CTR)
        Loop
    Close #1
    
    Do Until Number = 1 Or Number = 2 Or Number = 3 Or Number = 4 Or Number = 5 Or Number = 6 Or Number = 7  'pick a number and based on that, you get a life
        Number = InputBox("Let's gamble with your life...pick a number (1-7)", "Your life")
        Pos = 0
        Found = False
        
        Do While Found = False And Pos < CTR
            Pos = Pos + 1
            If Number = posArray(Pos) Then
                Found = True
            
            End If
        Loop
        
        picResults.Cls
        
        If Found = True Then        'display the life you've picked and the stats
            HP = hpArray(Pos)
            Money = moneyArray(Pos)
            Attack = attackArray(Pos)
            picResults.Print "Based on the number you picked, your life has randomly been laid out for you"
            picResults.Print "as "; lifeArray(Pos); " with "; hpArray(Pos); " H.P., "; FormatCurrency(moneyArray(Pos)); " and"; attackArray(Pos); " attack points."
            picResults.Print "Obviously, your job doesn't matter much anymore,"
            picResults.Print "but your H.P., money and attack points vary based on your life."
        Else
            picResults.Print "This is not a valid number"
        End If
    Loop
    
    
   
End Sub
