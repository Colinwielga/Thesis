VERSION 5.00
Begin VB.Form ChinaPhoto 
   BackColor       =   &H80000016&
   Caption         =   "Pictures of China"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   8
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtChoose 
      BackColor       =   &H80000016&
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
      Left            =   5280
      TabIndex        =   0
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image picImage 
      Height          =   4215
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Label lblFood 
      BackColor       =   &H80000016&
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblGreatWall 
      BackColor       =   &H80000016&
      Caption         =   "The Great Wall"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblPeople 
      BackColor       =   &H80000016&
      Caption         =   "People"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblHongKong 
      BackColor       =   &H80000016&
      Caption         =   "Hong Kong"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblBeijing 
      BackColor       =   &H80000016&
      Caption         =   "Beijing"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblShanghai 
      BackColor       =   &H80000016&
      Caption         =   "Shanghai"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblChoose 
      Caption         =   "What would you like to see?      (choose from the above)"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
End
Attribute VB_Name = "ChinaPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: ChinaPhoto.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form: In this form, the user will type one of the listed words
'3 pictures will display, each seperated by a message box that displays useful information
Option Explicit
'Shows the China Form and hides the ChinaPhoto Form
Private Sub cmdBack_Click()
ChinaPhoto.Hide
China.Show
End Sub
'When the user enters one of the listed words, a Picture comes up along with
'a message box with some useful information
Private Sub cmdGo_Click()
Dim h As String
Dim go As String            'Dim Variables



h = txtChoose.Text      'h is the text that is typed in by the user to select from list
go = h
    
    If go = "Shanghai" Then         'If the user enters "Shaghai" then the picImage loads a picture accossiated with it
        picImage.Picture = LoadPicture(App.Path & "\China\Shanghai.jpg")
        MsgBox ("Shanghai during the day")      'After the message box, the next picture displays
           
        picImage.Picture = LoadPicture(App.Path & "\China\Shanghai1.jpg")
        MsgBox ("Shanghai at night")
                
        picImage.Picture = LoadPicture(App.Path & "\China\Shanghai2.jpg")
        MsgBox ("The Bond Street")
        
    ElseIf go = "Beijing" Then      'If the user enters "Beijing" then a picture assocciated with it displays
    
        picImage.Picture = LoadPicture(App.Path & "\China\Beijing.jpg")
        MsgBox ("Mao")              'after the message box is viewed, the next picture will display
           
        picImage.Picture = LoadPicture(App.Path & "\China\tiananmen-square.jpg")
        MsgBox ("Tiananmen Square")
                
        picImage.Picture = LoadPicture(App.Path & "\China\Forbidden_City.jpg")
        MsgBox ("Forbidden City")
   
   ElseIf go = "Hong Kong" Then        'If the user enters "Hong Kong" then a picture assocciated with it displays
   
        picImage.Picture = LoadPicture(App.Path & "\China\hk.jpg")
        MsgBox ("Hong Kong used to be a British Colony")
                                        'after the message box is viewed, the next picture will display
        picImage.Picture = LoadPicture(App.Path & "\China\hk2.jpg")
        MsgBox ("Hong Kong was handed over to China in 1997")
                
        picImage.Picture = LoadPicture(App.Path & "\China\hk3.jpg")
        MsgBox ("Hong Kong is very humid")
        
    ElseIf go = "People" Then           'If the user enters "People" then a picture assocciated with it displays
    
        picImage.Picture = LoadPicture(App.Path & "\China\chinese2.jpg")
        MsgBox ("Family is very important for the Chinese")
                                         'after the message box is viewed, the next picture will display
        picImage.Picture = LoadPicture(App.Path & "\China\chinese.jpg")
        MsgBox ("The Chinese are considered to be quite reserved with new people")
                
        picImage.Picture = LoadPicture(App.Path & "\China\chinese1.jpg")
        MsgBox ("Philosophy of GUANXI is widespread in China. GUANXI = relationship")
    
    ElseIf go = "The Great Wall" Then       'If the user enters "The Great Wall" then a picture assocciated with it displays
    
        picImage.Picture = LoadPicture(App.Path & "\China\gw.jpg")
        MsgBox ("The only construction seen from space that was made by men")
                                            'after the message box is viewed, the next picture will display
        picImage.Picture = LoadPicture(App.Path & "\China\gw1.jpg")
        MsgBox ("The Great Wall is 2450 kilometer-long. It equals the distance from London to Moscow. Quite impressive, isn't it?")
                
     ElseIf go = "Food" Then                'If the user enters "Food" then a picture assocciated with it displays
     
        picImage.Picture = LoadPicture(App.Path & "\China\food.jpg")
        MsgBox ("Chinese Cuisine is considered to be one of the healthiest in the world")
                                            'after the message box is viewed, the next picture will display
        picImage.Picture = LoadPicture(App.Path & "\China\food1.jpg")
        MsgBox ("Chopsticks are Chinese silverware")
                
        picImage.Picture = LoadPicture(App.Path & "\China\food2.jpg")
        MsgBox ("Meals are often served with peanuts")
          
        
    End If
    

End Sub

