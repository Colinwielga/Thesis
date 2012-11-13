VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Kim Nguyen"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame info 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Automobile Information"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   4800
         TabIndex        =   20
         Top             =   6240
         Width           =   4215
      End
      Begin VB.CommandButton cmdmatchprice 
         Caption         =   "Find A Car That Match Your Price"
         Height          =   615
         Left            =   4800
         TabIndex        =   19
         Top             =   5280
         Width           =   4215
      End
      Begin VB.CommandButton cmdPurchase 
         Caption         =   "Which Car Do You Want To Buy"
         Height          =   615
         Left            =   4800
         TabIndex        =   18
         Top             =   4440
         Width           =   4215
      End
      Begin VB.CommandButton cmdshowcar 
         Caption         =   "Show All 10 Pictures Of the Cars"
         Height          =   615
         Left            =   4800
         TabIndex        =   17
         Top             =   3600
         Width           =   4215
      End
      Begin VB.PictureBox picDrivetrain 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   16
         Top             =   5760
         Width           =   2655
      End
      Begin VB.PictureBox picStandardTransmission 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   14
         Top             =   4920
         Width           =   2655
      End
      Begin VB.PictureBox picHorsepower 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   12
         Top             =   4200
         Width           =   2655
      End
      Begin VB.PictureBox picCylinders 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   10
         Top             =   3480
         Width           =   2655
      End
      Begin VB.PictureBox picClass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   8
         Top             =   2760
         Width           =   2655
      End
      Begin VB.PictureBox picYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   6
         Top             =   2160
         Width           =   2655
      End
      Begin VB.PictureBox picPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ComboBox cboSelect 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "CarInfo1.frx":0000
         Left            =   240
         List            =   "CarInfo1.frx":0022
         TabIndex        =   2
         Text            =   "Select A Car"
         Top             =   840
         Width           =   3855
      End
      Begin VB.PictureBox picShowCar 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   4560
         ScaleHeight     =   2475
         ScaleWidth      =   4755
         TabIndex        =   1
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lbDrivetrain 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Drivetrain"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label lbStandardTransmission 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Standard Transmission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lbHorsepower 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Horsepower"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lbcylinder 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cylinders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lbClass 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lbYear 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbPrice 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Online Car Dealer
'Form Name : frmInfo (CarInfo.frm)
'Author: Kim Nguyen
'Date Written: October 29, 2003
'Purpose of Form: To let the customer to view all of the cars available for sell
                'this form also serve as a homepage
                
                

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.Option Explicit
Dim Year(1 To 10) As Integer
Dim CarName(1 To 10) As String
Dim Price(1 To 10) As Long
Dim Class(1 To 10) As String
Dim Cylinders(1 To 10) As String
Dim Horsepower(1 To 10) As Single
Dim StandardTransmission(1 To 10) As String
Dim Drivetrain(1 To 10) As String
Dim PicName(1 To 10) As String
Dim I As Integer

'Open file and put it into an array

Public Sub Form_Load()
Path = "N:\CS130\handin\nguyen_kim\"
Open Path & "cars.txt" For Input As #1
For I = 1 To 10
    Input #1, Year(I), CarName(I), Price(I), Class(I), Cylinders(I), Horsepower(I), StandardTransmission(I), Drivetrain(I), PicName(I)
Next I
End Sub
'End
Private Sub cmdQuit_Click()
End 'End this program now
End Sub

Private Sub cmdmatchprice_Click()
frmInfo.Hide    'Hide the form that is using right now which is Car Information
frmMatchPrice.Show  'Show the form that indicated which is MatchPrice
End Sub

Private Sub cmdPurchase_Click()
frmInfo.Hide    'Hide the form that is using right now which is Car Infomation
frmPurchase.Show    'Hide the form that is using right now which is Purchase
End Sub

Private Sub cmdshowcar_Click()
frmInfo.Hide    'Hide the form that is using right now which is Car Information
frmBestCar.Show 'Hide the form that is using right now which is best car
End Sub



Private Sub cboSelect_Click()
    picYear.Cls 'clear the picture box
    picPrice.Cls 'clear the picture box
    picClass.Cls    'clear the picture box
    picCylinders.Cls    'clear the picture box
    picHorsepower.Cls   'clear the picture box
    picStandardTransmission.Cls 'clear the picture box
    picDrivetrain.Cls   'clear the picture box
    


I = cboSelect.ListIndex + 1 'I is the counter, so everytime a user click on the combo dropdown
                            'and click on the name of the car on the Indexlist, the position on the indexlist
                            'starts from 0 that's why I need to add 1 so that it would match with the array in the file
    picShowCar.Picture = LoadPicture(Path & PicName(I))    'load the picture up and show in name of the picture in position I
    
    picYear.Print Year(I)   'load the year in position I up and show in the picture box
    
    picPrice.Print FormatCurrency(Price(I), 2) 'load the price in position I up and show in picturebox in currency format
    
    picClass.Print Class(I) 'load the Class of the car in position I up and show in picture box
    
    picCylinders.Print Cylinders(I) 'load the cylinders of the car in position I up and print out in picture box
    
    picHorsepower.Print Horsepower(I) 'load the Horsepower of the car in position I up and print it out in the picture box
    
    picStandardTransmission.Print StandardTransmission(I) 'load the transmssiion of the car in position (I) and print it out in the picture box
    
    picDrivetrain.Print Drivetrain(I) 'load the Drivetrain of the car in position I up and print it out in the picture box
    
End Sub

Private Sub info_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
