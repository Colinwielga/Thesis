VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FF0000&
   Caption         =   "Housing Options"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H000000FF&
      Caption         =   "Purchase Supplies For Your Dorm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H000000FF&
      Caption         =   "Update Availablitiy of Houses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000000FF&
      Caption         =   "Find Where Your Friend Is Living"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox picVirgilOcc 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9000
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   6480
      Width           =   2535
   End
   Begin VB.PictureBox picVincentOcc 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6000
      ScaleHeight     =   1635
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   6480
      Width           =   2775
   End
   Begin VB.PictureBox picMettenOcc 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3120
      ScaleHeight     =   1635
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   6480
      Width           =   2655
   End
   Begin VB.PictureBox picMaurOcc 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   9
      Top             =   6480
      Width           =   2535
   End
   Begin VB.PictureBox picVirgil 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8880
      Picture         =   "housing.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.PictureBox picVincent 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5880
      Picture         =   "housing.frx":25B1
      ScaleHeight     =   2115
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   2880
      Width           =   2895
   End
   Begin VB.PictureBox picMetten 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      Picture         =   "housing.frx":4735
      ScaleHeight     =   2115
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.PictureBox picMaur 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      Picture         =   "housing.frx":6FB6
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdVirgil 
      BackColor       =   &H000000FF&
      Caption         =   "Virgil Michel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdVincent 
      BackColor       =   &H000000FF&
      Caption         =   "Vincent Court"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdMetten 
      BackColor       =   &H000000FF&
      Caption         =   "Metten Court"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdMaur 
      BackColor       =   &H000000FF&
      Caption         =   "St. Maur House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FF0000&
      Caption         =   "Project Created By:                 Kyle Johnson"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblInformation 
      BackColor       =   &H00FF0000&
      Caption         =   "Click On the Picture of The House For                         Additional Information"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Label lblheading 
      BackColor       =   &H00FF0000&
      Caption         =   "Where Would You Like To Live??"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Housing Options Form
' Written By Kyle Johnson
' 3/22/06
'This form is where most of the user interaction takes place
' It allows the user to select thier house, search for friends,
' see availability of housing
' it also allows navigation to the additional information pages,
' and the school store,




Option Explicit
    'dim variables that are used by more than one sub-routine

    Dim maur, vincent, virgil, metten As Integer




Private Sub cmdMaur_Click()
    'keeps track of the number of kids in maur
    maur = maur + 1
    'only show maur as option if maur is not full
    If maur < 6 Then
        picMaur.Visible = True
        Else: picMaur.Visible = False
        cmdMaur.Visible = False
        picMaurOcc.Visible = False
    End If
     'saves Maur in the array parrallel with the person who drafted it
    houseArray(K) = "Maur"
    'inform the person where they are living
    MsgBox namesArray(K) & " You are living in " & houseArray(K)
    'move from options form to draft form
    frmOptions.Visible = False
    frmDraft.Visible = True
    
End Sub

Private Sub cmdMetten_Click()
    'keep track of the number of people in metten
    metten = metten + 1
    'only displays the option for metten if it is not full
    If metten < 6 Then
        picMetten.Visible = True
        Else: picMetten.Visible = False
        cmdMetten.Visible = False
        picMettenOcc.Visible = False
    End If
    'saves Metten in a parrallel array next to the person who selected it
    houseArray(K) = "Metten"
    'inform the person where they are living
    MsgBox namesArray(K) & " You are living in " & houseArray(K)
    'moves user from options page to draft page
    frmOptions.Visible = False
    frmDraft.Visible = True

End Sub


Private Sub cmdPurchase_Click()
    frmStore.Visible = True
    frmOptions.Visible = False
End Sub

Private Sub cmdSearch_Click()
    'dim local variables
    Dim txt As String
    Dim A, C As Integer
    Dim found As Boolean
    
    C = 0
    'saves the name entered into the text box as txt
    txt = InputBox("Enter the name of the friend you wish to find", "Friend Finder")
    'a loop that searches for the name entered into the text box
    Do While found = False And C <= K
        C = C + 1
        A = InStr(LCase(txt), LCase(namesArray(C)))
            If A <> 0 Then
                    found = True
            End If
        Loop
    'if the name is found, then print a message box saying where the student is living
    If found = True Then
        MsgBox namesArray(C) & " is living in " & houseArray(C), , "Friend Found"
    Else
        'student is not found, and message box reports that
        MsgBox txt & " Was not found in our records", , "Sorry"
    End If

End Sub

Private Sub cmdUpdate_Click()
    'dim local variables
    Dim vincentavail As Integer
    Dim mettenavail As Integer
    Dim mauravail As Integer
    Dim viravail As Integer

    'display in picture boxes the number of rooms occupied and the number of rooms avaliable for each of the housing options
    picVincentOcc.Cls
    vincentavail = 4 - vincent
    picVincentOcc.Print "# of rooms occupied"; Tab(23); vincent
    picVincentOcc.Print
    picVincentOcc.Print "# of rooms available"; Tab(23); vincentavail
    
    
    picMettenOcc.Cls
    mettenavail = 6 - metten
    picMettenOcc.Print "# of rooms occupied"; Tab(23); metten
    picMettenOcc.Print
    picMettenOcc.Print "# of rooms available"; Tab(23); mettenavail
    
    picMaurOcc.Cls
    mauravail = 6 - maur
    picMaurOcc.Print "# of rooms occupied"; Tab(23); maur
    picMaurOcc.Print
    picMaurOcc.Print "# of rooms avaliable"; Tab(23); mauravail
    
    picVirgilOcc.Cls
    
    viravail = 8 - virgil
    picVirgilOcc.Print "# of rooms occupied"; Tab(23); virgil
    picVirgilOcc.Print
    picVirgilOcc.Print "# of rooms avaliable"; Tab(23); viravail

End Sub

Private Sub cmdVincent_Click()

    'keeps count of the number of kids living in vincent
    vincent = vincent + 1
    'only displays vincent as an option if it is not full
    If vincent < 4 Then
        picVincent.Visible = True
    Else: picVincent.Visible = False
        cmdVincent.Visible = False
        picVincentOcc.Visible = False
    End If
    'saves Vincent in a parrallel array with the name of the person who drafted it
    houseArray(K) = "Vincent"
    'message box informing student of the house they have selected
    MsgBox namesArray(K) & " You are living in " & houseArray(K)
    'navigates from the options page to the draft page
    frmOptions.Visible = False
    frmDraft.Visible = True
 
 
End Sub

Private Sub cmdVirgil_Click()
    'keeps count of the number of people in virgil
    virgil = virgil + 1
    'only displays virgil as an option if it is not full
    If virgil < 8 Then
        picVirgil.Visible = True
    Else: picVirgil.Visible = False
        cmdVirgil.Visible = False
        picVirgilOcc.Visible = False
    End If
    'saves the name of the house in a parallel array along with the studen who drafted it
    houseArray(K) = "Virgil Michel"
    'message box informing student of what they selected
    MsgBox namesArray(K) & " You are living in " & houseArray(K)
    'navigates from the options form to the draft form
    frmOptions.Visible = False
    frmDraft.Visible = True


End Sub


Private Sub Form_Load()
    vincent = 0
    metten = 0
    maur = 0
    virgil = 0
    
    'navigates from the options form to the welcome form immediately when the project loads
    frmWelcome.Visible = True
    frmOptions.Visible = False

End Sub

Private Sub picMaur_Click()
    'navigates from options page to maur page
    frmMaur.Visible = True
    frmOptions.Visible = False


End Sub

Private Sub picMetten_Click()
    'navigates from options page to the metten page
    frmMetten.Visible = True
    frmOptions.Visible = False

End Sub

Private Sub picVincent_Click()
    'navigates from the options page to the vincent page
    frmOptions.Visible = False
    frmVincent.Visible = True

End Sub

Private Sub picVirgil_Click()
    'navigates from the options form to the virgil form
    frmVirgil.Visible = True
    frmOptions.Visible = False

End Sub
