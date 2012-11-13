VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdWork 
      Caption         =   "Sort List By Type of Work"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Sort List By Location of Work"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort List By Name"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load list of Artists"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "~frank verdeja"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   7080
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the viewer to load a list of my favorite graffiti artists, and sort
'the list alphabetically, by type of work, and by the region in which they work.

'There is also a button that will take the user back to the main menu.

Option Explicit
Dim strArtist(1 To 28) As String
Dim strType(1 To 28) As String
Dim strRegion(1 To 28) As String
Dim strFile As String
Private Sub cmdLoad_Click()
Dim i As Integer

strFile = strPath + "data.txt"
pbxResults.Cls

pbxResults.Print Tab(41); "MY FAVORITE GRAFFITI ARTISTS"
pbxResults.Print "Artist's Tag Name(Pen-Name)"; Tab(30); "Type of Work"; Tab(50); "Location of Work(U.S. if not specified)"
pbxResults.Print "**************************************************************************************************************"
Open strFile For Input As #1
    For i = 1 To 28
        Input #1, strArtist(i), strType(i), strRegion(i)
        pbxResults.Print strArtist(i); Tab(30); strType(i); Tab(50); strRegion(i)
    Next i
Close #1

End Sub


Private Sub cmdHome_Click()
Dim temp As String
Dim pass As Integer
Dim n As Integer
Dim i As Integer

pbxResults.Cls

pbxResults.Print Tab(41); "MY FAVORITE GRAFFITI ARTISTS"
pbxResults.Print "Artist's Tag Name(Pen-Name)"; Tab(30); "Type of Work"; Tab(50); "Location of Work(U.S. if not specified)"
pbxResults.Print "*************************************************************************************************************"

n = 28
For pass = 1 To n
    For i = 1 To n - pass
        If strRegion(i) < strRegion(i + 1) Then
            temp = strArtist(i + 1)
           strArtist(i + 1) = strArtist(i)
            strArtist(i) = temp
            
            temp = strType(i + 1)
            strType(i + 1) = strType(i)
            strType(i) = temp
            
            temp = strRegion(i + 1)
            strRegion(i + 1) = strRegion(i)
            strRegion(i) = temp
            
            
        End If
    Next i
            pbxResults.Print strArtist(i); Tab(30); strType(i); Tab(50); strRegion(i)
Next pass

End Sub



Private Sub cmdMenu_Click()
Form1.Show
Form2.Hide

End Sub

Private Sub cmdSort_Click()
Dim temp As String
Dim pass As Integer
Dim n As Integer
Dim i As Integer

pbxResults.Cls

pbxResults.Print Tab(41); "MY FAVORITE GRAFFITI ARTISTS"
pbxResults.Print "Artist's Tag Name(Pen-Name)"; Tab(30); "Type of Work"; Tab(50); "Location of Work(U.S. if not specified)"
pbxResults.Print "**************************************************************************************************************"

n = 28
For pass = 1 To n
    For i = 1 To n - pass
        If strArtist(i) < strArtist(i + 1) Then
            temp = strArtist(i + 1)
           strArtist(i + 1) = strArtist(i)
            strArtist(i) = temp
            
            temp = strType(i + 1)
            strType(i + 1) = strType(i)
            strType(i) = temp
            
            temp = strRegion(i + 1)
            strRegion(i + 1) = strRegion(i)
            strRegion(i) = temp
            

        End If
    Next i
            pbxResults.Print strArtist(i); Tab(30); strType(i); Tab(50); strRegion(i)
Next pass

End Sub

Private Sub cmdWork_Click()
Dim temp As String
Dim pass As Integer
Dim n As Integer
Dim i As Integer


pbxResults.Cls

pbxResults.Print Tab(41); "MY FAVORITE GRAFFITI ARTISTS"
pbxResults.Print "Artist's Tag Name(Pen-Name)"; Tab(30); "Type of Work"; Tab(50); "Location of Work(U.S. if not specified)"
pbxResults.Print "*************************************************************************************************************"

n = 28
For pass = 1 To n
    For i = 1 To n - pass
        If strType(i) < strType(i + 1) Then
            temp = strArtist(i + 1)
           strArtist(i + 1) = strArtist(i)
            strArtist(i) = temp
            
            temp = strType(i + 1)
            strType(i + 1) = strType(i)
            strType(i) = temp
            
            temp = strRegion(i + 1)
            strRegion(i + 1) = strRegion(i)
            strRegion(i) = temp
            

        End If
    Next i
            pbxResults.Print strArtist(i); Tab(30); strType(i); Tab(50); strRegion(i)
Next pass
End Sub
