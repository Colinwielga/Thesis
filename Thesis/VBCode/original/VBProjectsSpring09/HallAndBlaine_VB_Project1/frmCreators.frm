VERSION 5.00
Begin VB.Form frmCreators 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Creators"
   ClientHeight    =   8715
   ClientLeft      =   6690
   ClientTop       =   4020
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12495
   Begin VB.CommandButton cmdAndreandColin 
      BackColor       =   &H00C000C0&
      Caption         =   "Meet Andre Blaine and Colin Hall at the same time"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return To Main Page"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdAndre 
      BackColor       =   &H0000FF00&
      Caption         =   "Meet Andre Blaine"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdColin 
      BackColor       =   &H00FFFF00&
      Caption         =   "Meet Colin Hall"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   5400
      ScaleHeight     =   7695
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmCreators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmCreators
'Colin Hall and Andre Blaine
'March 15
     'This form will clear any pictures in picCreators,
     'as well as load and print a picture of one of the creators, Colin Hall or Andre Blaine,
     'or will load and print a picture containing both of the creators,
     'and will return you to the Main Page form or will quit.


Private Sub cmdColin_Click()

     'This will clear any previous pictures in picCreators.
     picResults.Cls
     
     'This will load and print the picture called ColinHall.
     picResults.Picture = LoadPicture(App.Path & "\ColinHall.jpg")

End Sub

Private Sub cmdAndre_Click()
     
     'This will clear any previous pictures in picCreators.
     picResults.Cls
     
     'This will load and print the picture called AndreBlaine.
     picResults.Picture = LoadPicture(App.Path & "\AndreBlaine.jpg")

End Sub

Private Sub cmdAndreandColin_Click()

     'This will clear any previous pictures in the picture box.
     picResults.Cls
     
     'This will load and print the picture called AndreandColin.
     picResults.Picture = LoadPicture(App.Path & "\AndreandColin.jpg")
     
End Sub

Private Sub cmdReturn_Click()

    'This button will hide the Creators form and will show the Main Page form.
    frmMainPage.Show
    frmCreators.Hide
    
End Sub

Private Sub cmdQuit_Click()

    'This button will end the Visual Basic Program.
    End
    
End Sub
