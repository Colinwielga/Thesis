VERSION 5.00
Begin VB.Form FormHotel 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   13275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19890
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   13275
   ScaleWidth      =   19890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10680
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   "OK"
      Height          =   1335
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10680
      Width           =   5295
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H80000003&
      Caption         =   "    1 Star"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   4
      Top             =   9120
      Width           =   5295
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H80000003&
      Caption         =   "    2 Star"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   3
      Top             =   7560
      Width           =   5295
   End
   Begin VB.OptionButton opt3 
      BackColor       =   &H80000003&
      Caption         =   "    3 Star"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   2
      Top             =   6120
      Width           =   5295
   End
   Begin VB.OptionButton opt4 
      BackColor       =   &H80000003&
      Caption         =   "    4 Star"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   1
      Top             =   4560
      Width           =   5295
   End
   Begin VB.OptionButton opt5 
      BackColor       =   &H80000003&
      Caption         =   "    5 Star"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   0
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Choose What Style of Hotel You Would Like"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   12375
   End
End
Attribute VB_Name = "FormHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP
'HOTELS
'HOLLIS FRITTS
'8-18
'THIS IS INFORMATION ON HOTELS IN THE AREA



Option Explicit
'create a set of options using the option button tool and set them to correspond with each form'
'format an If statement for option buttons to direct user to other forms
Private Sub Command1_Click()
    If opt1 = True Then
        frm1star.Show
        FormHotel.Hide
    End If
    If opt2 = True Then
        frm2star.Show
        FormHotel.Hide
    End If
    If opt3 = True Then
        frm3star.Show
        FormHotel.Hide
    End If
    If opt4 = True Then
        frm4star.Show
        FormHotel.Hide
    End If
    If opt5 = True Then
        frm5Star.Show
        FormHotel.Hide
    End If

End Sub

Private Sub Command2_Click()
FormHotel.Hide
Title.Show
End Sub

