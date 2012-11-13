VERSION 5.00
Begin VB.Form frmAvailabilityForm 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "frmAvailabilityForm.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   8715
      TabIndex        =   5
      Top             =   2400
      Width           =   8775
      Begin VB.CommandButton cmdTotal 
         BackColor       =   &H0000FF00&
         Caption         =   "Click to view expected total"
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   2895
      End
   End
   Begin VB.PictureBox picResults 
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   7155
      TabIndex        =   3
      Top             =   1440
      Width           =   7215
   End
   Begin VB.CommandButton cmdcheckavailability 
      BackColor       =   &H000000FF&
      Caption         =   "Check Availability"
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtdates 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Call: (987)-654-3210 to book your reservation today!!!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAvailabilityForm.frx":BCAA
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmAvailabilityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmAvailabilityForm
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
Option Explicit
'This command goes back to previous form
Private Sub cmdBack_Click()
    frmActivitiesForm.Show
    frmAvailabilityForm.Hide
End Sub

Private Sub cmdcheckavailability_Click()

'Declaration of variables
Dim Dates(1 To 80) As String, Found As Boolean
Dim K As Integer, CTR As Integer, EntryDate As String

'Here we open the file from which to get the dates from
Open App.Path & "\Dates.txt" For Input As #1

'Load the file into an array
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Dates(CTR)
Loop

'Assigning the variables values
EntryDate = txtdates.Text
K = 0
Found = False

    
'This will search the array and find out if the date entered by the user
'has locations that are still available.  If there is nothing avaialable
'for that date the proper message will appear
    Do While ((Not Found) And (K < CTR))
        K = K + 1
        If EntryDate = Dates(K) Then
            Found = True
            picResults.Cls
            picResults.Print "The date of "; EntryDate; " is available.  Please call us today to book your dream vacation!!!"
        End If
    Loop

    If (Not Found) Then
        picResults.Cls
        picResults.Print "The date of "; EntryDate; " is unavailable.  Also note: We are only open June through August."
    End If

'Closes file so it can be done again
Close #1

End Sub

'Closes current form and shows new one
Private Sub cmdTotal_Click()
    frmEndForm.Show
    frmAvailabilityForm.Hide
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

