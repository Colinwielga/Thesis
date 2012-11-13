VERSION 5.00
Begin VB.Form frmArt 
   BackColor       =   &H00404080&
   Caption         =   "Paintings"
   ClientHeight    =   10590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14850
   FillColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   14850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "View Chronologically "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   11280
      Picture         =   "frmArt.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   7
      Top             =   7680
      Width           =   1815
   End
   Begin VB.PictureBox picArtwork 
      Height          =   7455
      Left            =   360
      ScaleHeight     =   7395
      ScaleWidth      =   10830
      TabIndex        =   6
      Top             =   2640
      Width           =   10890
   End
   Begin VB.CommandButton cmdReturnMain 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   4
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdYear 
      Caption         =   "Year created=>"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdMedium 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSlideshow 
      Caption         =   "Slide Show of all Paintings"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11640
      TabIndex        =   0
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   10440
      Width           =   3855
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H00404080&
      Caption         =   "Search for Paintings by:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'The Artist's Multimedia Portfolio
'frmArt
'Ashley Thompson
'Friday March 20, 2009
'This form uses two commands to search for paintings read in the module
'It then allows users to search for paintings according to medium or year painted using Input Boxes
'Searches include Do and For Next Loops
'There is also a Slide Show command that allows the user to view all paintings using the Timer function
'The sort button uses a bubble sort to order the paintings chronologically and then displays them in this order
'It includes buttons to show the previous form and main form.


Private Sub cmdForm2_Click()
    frmMain.Show
    frmArt.Hide
End Sub





Private Sub cmdBack_Click()
frmArtMain.Show
frmArt.Hide
End Sub


Private Sub cmdFind_Click()
picArtwork.Cls


Dim Found As Boolean, j As Integer
    
    Found = False
    
    For j = 1 To ctr
        If cboYear = Year(j) Then
           picArtwork.Picture = LoadPicture(App.Path & "\" & Art(j))
            Found = True
            Sleep (1000)
       End If
    Next j
    



Close #1
End Sub

Private Sub cmdMedium_Click()
picArtwork.Cls


Dim UserMedium As String
Dim t As Double
Dim ctr2 As Double
t = Timer

UserMedium = InputBox("Enter Medium: Acrylic, Pen and Ink, Mixed Media, Oil", "Medium")


Dim Found As Boolean, j As Integer
    Found = False
    
    For j = 1 To ctr
            If UserMedium = Medium(j) Then
                picArtwork.Picture = LoadPicture(App.Path & "\" & Art(j))
                Found = True
                Sleep (1000)
            End If
    Next j
    
If Found = False Then
    MsgBox ("Invalid Medium")
End If


Close #1

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturnMain_Click()
frmMain.Show
frmArt.Hide
End Sub




Private Sub cmdSlideshow_Click()
picArtwork.Cls

Dim whichOne As Integer, stopper As Integer, t As Double, oldOne As Integer, ctr2 As Double


whichOne = 1

stopper = 0

Do While (stopper < 19)

    picArtwork.Picture = LoadPicture(App.Path & "\" & Art(whichOne))
 
    
    
    t = Timer
    Do While (Timer - t) < 1
        ctr2 = ctr2 + 1
        If ctr2 = 1000000 Then
            
            ctr2 = 0
        End If
    Loop
    
   
    stopper = stopper + 1
    
    oldOne = whichOne
    whichOne = (stopper Mod ctr) + 1
    
Loop

lblFileName.Caption = Art(oldOne)
lblFileName.Visible = True
 
Close #1

End Sub

Private Sub cmdSort_Click()
picArtwork.Cls

Dim pass As Integer, pos As Integer, j As Integer
Dim tempYear As Integer, tempArt As String, tempDraft As String, tempMedium As String


For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Year(pos) > Year(pos + 1) Then
            tempYear = Year(pos)
            Year(pos) = Year(pos + 1)
            Year(pos + 1) = tempYear
            tempArt = Art(pos)
            Art(pos) = Art(pos + 1)
            Art(pos + 1) = tempArt
            tempMedium = Medium(pos)
            Medium(pos) = Medium(pos + 1)
            Medium(pos + 1) = tempMedium
            tempDraft = Draft(pos)
            Draft(pos) = Draft(pos + 1)
            Draft(pos + 1) = tempDraft
        End If
    Next pos
Next pass

For j = 1 To ctr
             picArtwork.Picture = LoadPicture(App.Path & "\" & Art(j))
             Sleep (1000)
Next j

End Sub


'Private Sub cmdYear_Click()
'picArtwork.Cls


'Dim UserYear As String

'Dim Found As Boolean, j As Integer
    'UserYear = InputBox("Enter Year")
   'Found = False
    
    'For j = 1 To ctr
        'If UserYear = Year(j) Then
           ' picArtwork.Picture = LoadPicture(App.Path & "\" & Art(j))
            'Found = True
           ' Sleep (1000)
       ' End If
    'Next j
    
'If Found = False Then
    'MsgBox ("Invalid Year")
'End If

'Close #1



'End Sub

Private Sub cmdYear_Click()
cboYear.Clear

cboYear.AddItem "2004"
cboYear.ItemData(cboYear.NewIndex) = 2004
cboYear.AddItem "2005"
cboYear.ItemData(cboYear.NewIndex) = 2005
cboYear.AddItem "2006"
cboYear.ItemData(cboYear.NewIndex) = 2006
cboYear.AddItem "2007"
cboYear.ItemData(cboYear.NewIndex) = 2007
cboYear.AddItem "2009"
cboYear.ItemData(cboYear.NewIndex) = 2009

End Sub
