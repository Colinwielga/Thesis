VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Are you the next State Record Holder?"
   ClientHeight    =   7320
   ClientLeft      =   4485
   ClientTop       =   1755
   ClientWidth     =   6660
   LinkTopic       =   "Form5"
   ScaleHeight     =   7320
   ScaleWidth      =   6660
   Visible         =   0   'False
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H000080FF&
      Caption         =   "GO TO Weight Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdleave 
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
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdperch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Perch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdsteelhead 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Steelhead"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      Picture         =   "State_records.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdnorthern 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Northern Pike"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      Picture         =   "State_records.frx":05EE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdbluegill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bluegill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      Picture         =   "State_records.frx":0C1A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdsmall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Smallmouth Bass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      Picture         =   "State_records.frx":12E0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdwalleye 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Walleye"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      Picture         =   "State_records.frx":196A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdmusky 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Musky"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Picture         =   "State_records.frx":1F8E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdcrappie 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Crappie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Picture         =   "State_records.frx":2585
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
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
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click To See If your fish was a record!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'State Records
'Eric Glorvigen
'Date= March 5
'This form uses the public (global) variables that are stored in the weight calculation
'form, then they are compared with the state records of each fish
'and if they are smaller it will show by how many pounds
'to simplify things, if they are larger the same number is used but the abs() function
'converts the number to positive, and shows how much they broke the record by

Private Sub cmdbluegill_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '
        Dim missed As Single
          missed = 3 - weightbluegill
                If weightbluegill <= 3 Then
                    MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
                Else
                    MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
                End If
            
End Sub

Private Sub cmdcrappie_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '

    Dim missed As Single
      missed = 5 - weightcrappie
            If weightcrappie <= 5 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
            
End Sub

Private Sub cmdexit_Click()
    'return back to main page
        form1.Show
        Form5.Hide
End Sub


Private Sub cmdleave_Click()
    'leave the program
        End
End Sub

Private Sub cmdmove_Click()
    'brings user back to weight calculator page if they have not found weight
        Form5.Hide
        Form8.Show
End Sub

Private Sub cmdmusky_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '
        Dim missed As Single
          missed = 54 - weightmusky
            If weightmusky <= 54 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
    
End Sub

Private Sub cmdnorthern_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '

        Dim missed As Single
         missed = 45 - weightnorthern
            If weightnorthern <= 45 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
        
End Sub

Private Sub cmdperch_Click()
    'this displays the ongoing joke of trying to find the weight of the perch with a msgbox
    
        MsgBox "Come on " & inputname & ", you're still trying to find the weight of that perch!", , "Come On Now!!!"
        
End Sub

Private Sub cmdsmall_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '
        Dim missed As Single
          missed = 8 - weightsmall
            If weightsmall <= 8 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
            
End Sub

Private Sub cmdsteelhead_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '
        Dim missed As Single
          missed = 17 - weightsteelhead
            If weightsteelhead <= 17 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
        
End Sub

Private Sub cmdwalleye_Click()
    'this button takes what the user calculated on the weight finder page and displays with a msgbox
    'how much they missed the record by, or how much the broke the record by by using the absolute function
    '
        Dim missed As Single
           missed = 17 - weightwalleye
            If weightwalleye <= 17 Then
                MsgBox "You Missed the State Record by " & FormatNumber(missed, 2) & " pounds!", , "Rats!"
            Else
                MsgBox inputname & ", You just broke the state record by " & Abs(FormatNumber(missed, 2)) & " pounds!!!", , "Attention!!!!"
            End If
            
End Sub



Private Sub Form_Load()
    'msgbox appear as form loading to alert user to find their fishes weight first
        MsgBox "If you have not calculated the weight of your fish, Go to The Weight Calculation Page", , "ATTENTION!!!"
End Sub

