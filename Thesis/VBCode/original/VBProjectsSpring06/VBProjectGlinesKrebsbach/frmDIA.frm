VERSION 5.00
Begin VB.Form frmDIA 
   Caption         =   "Denver International Airport"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picUS 
      BackColor       =   &H00FFFFFF&
      Height          =   11055
      Left            =   0
      Picture         =   "frmDIA.frx":0000
      ScaleHeight     =   10995
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton cmdNWA 
         Height          =   1455
         Left            =   8040
         Picture         =   "frmDIA.frx":3156
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   9360
         Width           =   3135
      End
      Begin VB.CommandButton cmdDelta 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   13200
         Picture         =   "frmDIA.frx":B109
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton cmdUnited 
         Height          =   855
         Left            =   9960
         Picture         =   "frmDIA.frx":B56A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H000080FF&
         Caption         =   "Back to Airline page"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   8640
         Width           =   1335
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "By: Levi Glines and John Krebsbach"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   10680
         Width           =   2775
      End
      Begin VB.Label lblNWA 
         BackStyle       =   0  'Transparent
         Caption         =   "Click for NWA rates"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   10
         Top             =   9000
         Width           =   2415
      End
      Begin VB.Label lblDelta 
         BackStyle       =   0  'Transparent
         Caption         =   "Click for Sun Country Airlines rates"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   12480
         TabIndex        =   8
         Top             =   8160
         Width           =   2895
      End
      Begin VB.Label lblUnited 
         BackStyle       =   0  'Transparent
         Caption         =   "Click for United Airlines Rates"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   6
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblDestination 
         BackStyle       =   0  'Transparent
         Caption         =   "Desination: Denver International Airport"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   3
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label lblDIA 
         BackStyle       =   0  'Transparent
         Caption         =   "DIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblMSP 
         BackStyle       =   0  'Transparent
         Caption         =   "MSP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   1
         Top             =   2640
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   5400
         X2              =   8040
         Y1              =   4800
         Y2              =   3120
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   255
         Left            =   5160
         Top             =   4800
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Label Label4 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmDIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmAirline(frmAirline.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to look up the cheapest flights from
'Minneapolis,MN to Denver,Colorado. we chose the three main airlines from Msp,MN to Denver,Colorado and
'researched the lowest prices for a one week stay during CSB/SJU's respective spring
'break dates.

Private Sub cmdback_Click()
frmDIA.Hide
frmAirline.Show

End Sub



Private Sub cmdDelta_Click()
    MsgBox "For your spring break trip the lowest Sun Country Airlines flight from Minneapolis,MN to Denver,CO is $350.60. This flight departs form Minneapolis at 6:10 A.M. and arrives in Denver at 7:10 A.M. MST(mountain standard time).This flight is round trip and will depart a week later from Denver Int. Airport at 8:25 P.M.(MST) and arrives at Msp. Int. at 11:20 P.M.(CST).This nonstop flight is of coach class.", , "NWA Flight"
End Sub

Private Sub cmdNWA_Click()
    MsgBox "For your spring break trip the lowest NWA flight from Minneapolis,MN to Denver,CO is $501.60. This flight departs form Minneapolis at 9:45 A.M. and arrives in Denver at 10:43 A.M. MST(mountain standard time).This flight is round trip and will depart a week later from Denver Int. Airport at 11:30 A.M.(MST) and arrives at Msp. Int. at 2:20 P.M.(CST).This nonstop flight is of coach class.", , "NWA Flight"
End Sub

Private Sub cmdUnited_Click()
    MsgBox "For your spring break trip the lowest United Airlines flight from Minneapolis,MN to Denver,CO is $531.60. This flight departs form Minneapolis at 6:39 A.M. and arrives in Denver at 7:43 A.M. MST(mountain standard time). This flight is round trip and will depart a week later from Denver Int. Airport at 3:10 P.M.(MST) and arrives at Msp. Int. at 6:05 P.M.(CST). This nonstop flight is of coach class.", , "United Flight"
End Sub
