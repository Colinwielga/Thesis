VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Andy Bestler"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "form1.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "go back"
      Height          =   375
      Left            =   13920
      TabIndex        =   19
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "quit"
      Height          =   375
      Left            =   13080
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdgoalie 
      Caption         =   "goalie pics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10920
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmddefense 
      Caption         =   "defense pics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdright 
      Caption         =   "Right Wing pics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox playerpic1 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic4 
      Height          =   1455
      Left            =   4560
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic7 
      Height          =   1455
      Left            =   9000
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic10 
      Height          =   1455
      Left            =   13680
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic3 
      Height          =   1455
      Left            =   3000
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic2 
      Height          =   1455
      Left            =   1560
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic6 
      Height          =   1455
      Left            =   7560
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic5 
      Height          =   1455
      Left            =   6000
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic9 
      Height          =   1455
      Left            =   12120
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox playerpic8 
      Height          =   1455
      Left            =   10560
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   9240
      Width           =   1095
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdcenter 
      Caption         =   "center pics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdleft 
      Caption         =   "left wing pics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdreaddata 
      Caption         =   "load data"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "pictures arranged from list left to right"
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   8880
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Wild (minnesota wild.vbp)
'Andy Bestler (form 2)
'Andy Bestler
'March 15, 2004

'The general purpose of this form is to allow the user to view the
'Minnesota Wild's individual pictures





Option Explicit
Dim x As Integer
Dim player(1 To 100) As String
Dim position(1 To 100) As String
Dim path As String





Private Sub cmdback_Click()
Form2.Hide  'hides second form and shows original
Form1.Show
End Sub

Private Sub cmdcenter_Click()
picresults.Cls          'clears pic box

For x = 1 To 30
    If position(x) = "Center" Then      'searches for and prints all centers
    picresults.Print player(x)
    End If
    Next x

    'prints players, centers, pics
    
    playerpic1.Picture = LoadPicture(path & "pics\Center\dowd.jpg")
    playerpic2.Picture = LoadPicture(path & "pics\Center\hendrickson.jpg")
    playerpic3.Picture = LoadPicture(path & "pics\Center\ronning.jpg")
    playerpic4.Picture = LoadPicture(path & "pics\Center\veilleux.jpg")
    playerpic5.Picture = LoadPicture(path & "pics\Center\wallin.jpg")
    playerpic6.Picture = LoadPicture(path & "pics\Center\walz.jpg")
    playerpic7.Picture = LoadPicture(path & "pics\Center\zholtok.jpg")
    playerpic8.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic9.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic10.Picture = LoadPicture(path & "pics\logo.gif")
    

End Sub

Private Sub cmddefense_Click()
picresults.Cls

For x = 1 To 30
    If position(x) = "Defense" Then     'searches for and prints out defensemen
    picresults.Print player(x)
    End If
    Next x
    
    'prints pics of defensemen
   
    playerpic1.Picture = LoadPicture(path & "pics\Defense\benysek.jpg")
    playerpic2.Picture = LoadPicture(path & "pics\Defense\bombardir.jpg")
    playerpic3.Picture = LoadPicture(path & "pics\Defense\brown.jpg")
    playerpic4.Picture = LoadPicture(path & "pics\Defense\kuba.jpg")
    playerpic5.Picture = LoadPicture(path & "pics\Defense\marshall.jpg")
    playerpic6.Picture = LoadPicture(path & "pics\Defense\mitchell.jpg")
    playerpic7.Picture = LoadPicture(path & "pics\Defense\schultz.jpg")
    playerpic8.Picture = LoadPicture(path & "pics\Defense\sekeras.jpg")
    playerpic9.Picture = LoadPicture(path & "pics\Defense\zyuzin.jpg")
    playerpic10.Picture = LoadPicture(path & "pics\logo.gif")
    
End Sub

Private Sub cmdgoalie_Click()
picresults.Cls


For x = 1 To 30
    If position(x) = "Goalie" Then      'searches for and prints goalies names
    picresults.Print player(x)
    End If
    Next x
    
    'prints pics of goalies
   
    playerpic1.Picture = LoadPicture(path & "pics\Goalie\fernandez.jpg")
    playerpic2.Picture = LoadPicture(path & "pics\Goalie\roloson.jpg")
    playerpic3.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic4.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic5.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic6.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic7.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic8.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic9.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic10.Picture = LoadPicture(path & "pics\logo.gif")

End Sub

Private Sub cmdleft_Click()

picresults.Cls


For x = 1 To 30
    If position(x) = "Left" Then        'searches for and prints left wingers
    picresults.Print player(x)
    End If
    Next x
    
    'prints pics of left wingers
   
    playerpic1.Picture = LoadPicture(path & "pics\Left Wing\bouchard.jpg")
    playerpic2.Picture = LoadPicture(path & "pics\Left Wing\brunette.jpg")
    playerpic3.Picture = LoadPicture(path & "pics\Left Wing\domen.jpg")
    playerpic4.Picture = LoadPicture(path & "pics\Left Wing\dupuis.jpg")
    playerpic5.Picture = LoadPicture(path & "pics\Left Wing\johnson.jpg")
    playerpic6.Picture = LoadPicture(path & "pics\Left Wing\laaksonen.jpg")
    playerpic7.Picture = LoadPicture(path & "pics\Left Wing\stevenson.jpg")
    playerpic8.Picture = LoadPicture(path & "pics\Left Wing\trudel.jpg")
    playerpic9.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic10.Picture = LoadPicture(path & "pics\logo.gif")
    
    
End Sub

Private Sub cmdquit_Click()
End         'quits program
End Sub

Private Sub cmdreaddata_Click()

Open path & "position.txt" For Input As #2
For x = 1 To 30
    
    'loads data into array
    
    Input #2, player(x), position(x)
    Next x
    Close (2)
                            'enables buttons
cmdleft.Enabled = True
cmdcenter.Enabled = True
cmdright.Enabled = True
cmddefense.Enabled = True
cmdgoalie.Enabled = True

    
End Sub



Private Sub cmdright_Click()
picresults.Cls


For x = 1 To 30
    If position(x) = "Right Wing" Then          'searches for and prints names of right wingers
    picresults.Print player(x)
    End If
    Next x
    
    'prints pics of right wingers
   
    playerpic1.Picture = LoadPicture(path & "pics\Right Wing\gaborik.jpg")
    playerpic2.Picture = LoadPicture(path & "pics\Right Wing\muckalt.jpg")
    playerpic3.Picture = LoadPicture(path & "pics\Right Wing\park.jpg")
    playerpic4.Picture = LoadPicture(path & "pics\Right Wing\wanvig.jpg")
    playerpic5.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic6.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic7.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic8.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic9.Picture = LoadPicture(path & "pics\logo.gif")
    playerpic10.Picture = LoadPicture(path & "pics\logo.gif")
    
End Sub

Private Sub Form_Load()
'path = "M:\CS130\VB project\"       'path for all files
path = "N:\CS130\handin\Bestler, Andy\"
End Sub
