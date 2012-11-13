VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H0080FFFF&
   Caption         =   "Fish Finder"
   ClientHeight    =   6480
   ClientLeft      =   3315
   ClientTop       =   1755
   ClientWidth     =   8250
   LinkTopic       =   "Form6"
   ScaleHeight     =   6480
   ScaleWidth      =   8250
   Visible         =   0   'False
   Begin VB.CommandButton cmdinput 
      BackColor       =   &H0080FF80&
      Caption         =   "Input data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdsizefish 
      BackColor       =   &H0080FF80&
      Caption         =   "Sort Fish By Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdfish 
      BackColor       =   &H0080FF80&
      Caption         =   "List Fish in Alpabetical Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H0080FF80&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5880
      Picture         =   "fish_finder.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtfish 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   3
      Top             =   1200
      Width           =   5295
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back to Main Page"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Push input button to list possible fish. Type the fish you would like to catch and see what lakes are the best!"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Fish Finder"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Minnesota Fisher
'Fish Finder
'Eric Glorvigen
'Date= March 5
' the purpose of this form is to access to arrays, perform a search using the INSTR function
'and display differnt types of sorts and searches
Dim fish(1 To 100) As String, size(1 To 100) As Single
Dim ctr As Integer

Private Sub cmdexit_Click()
    'return back to main menu
        form1.Show
        Form6.Hide
End Sub

Private Sub cmdfind_Click()
'this program searchs the lakes.text which has different lakes, countys, rank and fish, then it
'shows the location of the fish the user types, this uses the instr function
'because there are multiple fish in each string it searches
        Dim n As Integer
        Dim a As Integer
        Dim fishtype As String
        Dim k As Integer
        Dim lake(1 To 100) As String
        Dim county(1 To 100) As String
        Dim rank(1 To 100) As Single
        Dim fish(1 To 100) As String
        Dim ctrtwo As Integer
        

        picoutput.Cls

            ctr = 0
        Open App.Path & "\lakes.txt" For Input As #1
            Do Until EOF(1)
                ctrtwo = ctrtwo + 1
                Input #1, lake(ctrtwo), county(ctrtwo), rank(ctrtwo), fish(ctrtwo)
            Loop
        Close #1
    
        fishtype = txtfish.Text
        fishtype = Trim(fishtype)
        
        picoutput.Print ; UCase(fishtype); "'s are located in the following lakes:"
        picoutput.Print "*****************"
        
            For n = 1 To ctrtwo
                a = InStr(LCase(Trim(fish(n))), LCase(fishtype))
                    If a <> 0 Then
                        picoutput.Print lake(n)
                        k = k + 1
                    End If
            Next n
    
            If k = 0 Then
                picoutput.Print "Sorry "; UCase(fishtype); " is not in any of these lakes."
            End If
            
    
End Sub

Private Sub cmdfish_Click()
'this button sorts the fish in alpahbet order to show the user what
'possible fish he can look for
  
    Dim temp As String, temptwo As Single
    Dim pass As Integer
    Dim pos As Integer
    Dim c As Integer
    
    picoutput.Cls

        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If fish(pos) > fish(pos + 1) Then
                    temp = fish(pos)
                    fish(pos) = fish(pos + 1)
                    fish(pos + 1) = temp
                    temptwo = size(pos)
                    size(pos) = size(pos + 1)
                    size(pos + 1) = temptwo
                End If
            Next pos
        Next pass
        
          picoutput.Print "Fish"
          picoutput.Print "****************"
          
            For c = 1 To ctr
                picoutput.Print UCase(fish(c))
            Next c
End Sub

Private Sub cmdinput_Click()
' this buttons inputs the array and shows two other buttons, the ones that
'are performing searchs

        cmdfish.Visible = True
        cmdsizefish.Visible = True
        ctr = 0
    
        Open App.Path & "\fish.txt" For Input As #1
            Do Until EOF(1)
                ctr = ctr + 1
                Input #1, fish(ctr), size(ctr)
            Loop
        Close #1
    
End Sub

Private Sub cmdquit_Click()
    'exit program
        End
End Sub

Private Sub cmdsizefish_Click()
'ths button sorts the fish in rank of size, a component of the array
'from biggest to smallest

    Dim temp As String, temptwo As Single
    Dim pass As Integer
    Dim pos As Integer
    Dim c As Integer
    
    picoutput.Cls
  
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If size(pos) > size(pos + 1) Then
                    temp = fish(pos)
                    fish(pos) = fish(pos + 1)
                    fish(pos + 1) = temp
                    temptwo = size(pos)
                    size(pos) = size(pos + 1)
                    size(pos + 1) = temptwo
                End If
            Next pos
        Next pass
        
          picoutput.Print "Fish- largest to smallest"
          picoutput.Print "****************"
          
            For c = 1 To ctr
                picoutput.Print UCase(fish(c))
            Next c
End Sub


