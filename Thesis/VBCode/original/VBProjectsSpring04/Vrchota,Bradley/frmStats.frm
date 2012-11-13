VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00008000&
   Caption         =   "Take a Look"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   Picture         =   "frmStats.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optzito 
      BackColor       =   &H0080FFFF&
      Caption         =   "Barry Zito"
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   4920
      Width           =   1575
   End
   Begin VB.OptionButton optmaddux 
      BackColor       =   &H0080FFFF&
      Caption         =   "Greg Maddux"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin VB.OptionButton optortiz 
      BackColor       =   &H0080FFFF&
      Caption         =   "Russ Ortiz"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   4440
      Width           =   1575
   End
   Begin VB.OptionButton optcolon 
      BackColor       =   &H0080FFFF&
      Caption         =   "Bartolo Colon"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton optbrown 
      BackColor       =   &H0080FFFF&
      Caption         =   "Kevin Brown"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   3960
      Width           =   1575
   End
   Begin VB.OptionButton optmorris 
      BackColor       =   &H0080FFFF&
      Caption         =   "Matt Morris"
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OptionButton optpettitte 
      BackColor       =   &H0080FFFF&
      Caption         =   "Andy Pettitte"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.OptionButton optjohnson 
      BackColor       =   &H0080FFFF&
      Caption         =   "Randy Johnson"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton optvazquez 
      BackColor       =   &H0080FFFF&
      Caption         =   "Javier Vazquez"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.OptionButton optoswalt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Roy Oswalt"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton optmussina 
      BackColor       =   &H0080FFFF&
      Caption         =   "Mike Mussina"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.OptionButton optbeckett 
      BackColor       =   &H0080FFFF&
      Caption         =   "Josh Beckett"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton optmulder 
      BackColor       =   &H0080FFFF&
      Caption         =   "Mark Mulder"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton opthudson 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tim Hudson"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton opthalladay 
      BackColor       =   &H0080FFFF&
      Caption         =   "Roy Halladay"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton optwood 
      BackColor       =   &H0080FFFF&
      Caption         =   "Kerry Wood"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optschmidt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Jason Schmidt"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton optmartinez 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pedro Martinez"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton optSchilling 
      BackColor       =   &H0080FFFF&
      Caption         =   "Curt Schilling"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optPrior 
      BackColor       =   &H0080FFFF&
      Caption         =   "Mark Prior"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox picShowpic 
      BackColor       =   &H0080FFFF&
      Height          =   4215
      Left            =   720
      ScaleHeight     =   4155
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to the starting diamond"
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   480
      ScaleHeight     =   675
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "The top 20 MLB Fantasy Starters, listed in order. Choose which pitcher you want to see and their 2003 stats.>>>"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MLBPitchers (MLBPitchers.vbp)
'Form Name: frmStats (frmStats.frm)
'Author: Bradley Vrchota
'Date: March 14, 2004
'Purpose: This form's purpose is to allow the user to choose one of
        'the 20 pitchers at a time and see a picture of him and
        'his stats for wins, losses, ERA, and strikeouts for 2003
        
Option Explicit

Private Sub cmdBack_Click()
    'Takes user back to start form and hides stats form
    frmStats.Hide
    frmStart.Show
    
End Sub

'Each subroutine does the same thing, just for each different pitcher
Private Sub optbeckett_Click()
    'clear the picture box
    picList.Cls
    
    'print the header
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    
    'print the stats and picture for the pitcher selected
    picList.Print wins(9), losses(9), FormatNumber(ERA(9)), strikeouts(9)
    picShowpic.Picture = LoadPicture(PATH & "pics\9beckett.jpg")
End Sub

Private Sub optbrown_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(16), losses(16), FormatNumber(ERA(16)), strikeouts(16)
    picShowpic.Picture = LoadPicture(PATH & "pics\16brown.jpg")
End Sub

Private Sub optcolon_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(17), losses(17), FormatNumber(ERA(17)), strikeouts(17)
    picShowpic.Picture = LoadPicture(PATH & "pics\17colon.jpg")
End Sub

Private Sub opthalladay_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(6), losses(6), FormatNumber(ERA(6)), strikeouts(6)
    picShowpic.Picture = LoadPicture(PATH & "pics\6halladay.jpg")
End Sub

Private Sub opthudson_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(7), losses(7), FormatNumber(ERA(7)), strikeouts(7)
    picShowpic.Picture = LoadPicture(PATH & "pics\7hudson.jpg")
End Sub

Private Sub optjohnson_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(13), losses(13), FormatNumber(ERA(13)), strikeouts(13)
    picShowpic.Picture = LoadPicture(PATH & "pics\13johnson.jpg")
End Sub

Private Sub optmaddux_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(19), losses(19), FormatNumber(ERA(19)), strikeouts(19)
    picShowpic.Picture = LoadPicture(PATH & "pics\19maddux.jpg")
End Sub

Private Sub optmartinez_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(3), losses(3), FormatNumber(ERA(3)), strikeouts(3)
    picShowpic.Picture = LoadPicture(PATH & "pics\3martinez.jpg")
End Sub

Private Sub optmorris_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(15), losses(15), FormatNumber(ERA(15)), strikeouts(15)
    picShowpic.Picture = LoadPicture(PATH & "pics\15morris.jpg")
End Sub

Private Sub optmulder_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(8), losses(8), FormatNumber(ERA(8)), strikeouts(8)
    picShowpic.Picture = LoadPicture(PATH & "pics\8mulder.jpg")
End Sub

Private Sub optmussina_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(10), losses(10), FormatNumber(ERA(10)), strikeouts(10)
    picShowpic.Picture = LoadPicture(PATH & "pics\10mussina.jpg")
End Sub

Private Sub optortiz_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(18), losses(18), FormatNumber(ERA(18)), strikeouts(18)
    picShowpic.Picture = LoadPicture(PATH & "pics\18ortiz.jpg")
End Sub

Private Sub optoswalt_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(11), losses(11), FormatNumber(ERA(11)), strikeouts(11)
    picShowpic.Picture = LoadPicture(PATH & "pics\11oswalt.jpg")
End Sub

Private Sub optpettitte_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(14), losses(14), FormatNumber(ERA(14)), strikeouts(14)
    picShowpic.Picture = LoadPicture(PATH & "pics\14pettitte.jpg")
End Sub

Private Sub optPrior_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(1), losses(1), FormatNumber(ERA(1)), strikeouts(1)
    picShowpic.Picture = LoadPicture(PATH & "pics\1prior.jpg")
End Sub

Private Sub optSchilling_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(2), losses(2), FormatNumber(ERA(2)), strikeouts(2)
    picShowpic.Picture = LoadPicture(PATH & "pics\2schilling.jpg")
End Sub

Private Sub optschmidt_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(4), losses(4), FormatNumber(ERA(4)), strikeouts(4)
    picShowpic.Picture = LoadPicture(PATH & "pics\4schmidt.jpg")
End Sub

Private Sub optvazquez_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(12), losses(12), FormatNumber(ERA(12)), strikeouts(12)
    picShowpic.Picture = LoadPicture(PATH & "pics\12vazquez.jpg")
End Sub

Private Sub optwood_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(5), losses(5), FormatNumber(ERA(5)), strikeouts(5)
    picShowpic.Picture = LoadPicture(PATH & "pics\5wood.jpg")
End Sub

Private Sub optzito_Click()
    picList.Cls
    picList.Print "Wins", "Losses", "ERA", "Strikeouts"
    picList.Print "*****************************************************************"
    picList.Print wins(20), losses(20), FormatNumber(ERA(20)), strikeouts(20)
    picShowpic.Picture = LoadPicture(PATH & "pics\20zito.jpg")
End Sub
