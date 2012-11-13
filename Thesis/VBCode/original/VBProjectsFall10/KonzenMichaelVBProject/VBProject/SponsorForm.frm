VERSION 5.00
Begin VB.Form frmsponsor 
   Caption         =   "Sponsors"
   ClientHeight    =   11235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   Picture         =   "Sponsor Form.frx":0000
   ScaleHeight     =   11235
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdopensponsorlist 
      Caption         =   "Open Sponsor List"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   9615
      Left            =   2400
      ScaleHeight     =   9555
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   1440
      Width           =   7215
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdinputsponsors 
      Caption         =   "Veiw 2010 UFC Sponsors"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmsponsor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclear_Click()
picresults.Cls
End Sub

Private Sub cmdgoback_Click()
frmsponsor.Hide
frmmainscreen.Show
End Sub

Private Sub cmdinputsponsors_Click()
Dim sponsor As String
Dim sponsornumber As Integer, sponsorofwhat As String, count As Integer
count = 0
Open App.Path & "\Allsponsors.txt" For Input As #1 'opening file to draw from
Open App.Path & "\UFCsponsors.txt" For Output As #2 'opening file to write to
    picresults.Print "Sponsor#"; Tab(20); "Sponsor"; Tab(35); "What They Sponsor"
    picresults.Print "***************************************************************************"
Do While Not EOF(1)
    count = count + 1
    Input #1, sponsornumber, sponsor, sponsorofwhat 'telling the program what is in the text file
    If sponsorofwhat = "UFC" Then 'only the sponsors fro the ufc are written to the text file
    Write #2, sponsornumber, sponsor, sponsorofwhat 'writing data to text file
        picresults.Print sponsornumber, sponsor, Tab(45); sponsorofwhat
    End If
Loop
Close #1
Close #2
End Sub

Private Sub cmdopensponsorlist_Click()
Dim sponsor(1 To 100) As String, sponsornumber(1 To 100) As Integer, sponsorofwhat(1 To 100) As String, count As Integer
Dim count2 As Integer
count = 0
Open App.Path & "\AllSponsors.txt" For Input As #1 'telling program where to get data
        picresults.Print "Sponsor#"; Tab(20); "Sponsor"; Tab(35); "What They Sponsor"
        picresults.Print "************************************************************************"
Do While Not EOF(1)
       count = count + 1
    Input #1, sponsornumber(count), sponsor(count), sponsorofwhat(count)
        'telling the program what data is what
        picresults.Print sponsornumber(count), sponsor(count), Tab(45); sponsorofwhat(count)
Loop
Close #1


    MsgBox "There are " & count & " sponsors in the list of sponsors."
End Sub
