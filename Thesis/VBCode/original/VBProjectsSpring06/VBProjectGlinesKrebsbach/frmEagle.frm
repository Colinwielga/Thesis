VERSION 5.00
Begin VB.Form frmEagle 
   Caption         =   "Eagle County Regional Airport"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmEagle.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicUS 
      BackColor       =   &H00FFFFFF&
      Height          =   11055
      Left            =   0
      Picture         =   "frmEagle.frx":3156
      ScaleHeight     =   10995
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton cmdNWA 
         Caption         =   "Click for NWA flight"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7800
         Picture         =   "frmEagle.frx":62AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   9000
         Width           =   3015
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H000080FF&
         Caption         =   "Back to Airline Page"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9000
         Width           =   1695
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "By: Levi Glines and John Krebsbach"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   10560
         Width           =   2775
      End
      Begin VB.Label lblDestination 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination:  Eagle County Regional Airport"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   4
         Top             =   120
         Width           =   8175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         Top             =   4560
         Width           =   255
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
         TabIndex        =   2
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblEAG 
         BackStyle       =   0  'Transparent
         Caption         =   "Eagle"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   4920
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   4680
         X2              =   8040
         Y1              =   4560
         Y2              =   3120
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   255
         Left            =   4440
         Top             =   4560
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   255
         Left            =   8040
         Top             =   2880
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmEagle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmAirline(frmAirline.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to look up the cheapest flights from
'Minneapolis,MN to Colorado.
'we researched the lowest prices for a one week stay during CSB/SJU's respective spring
'break dates. we also researched the lowest flight available from MN to Eagle Regional
'Airport.
Private Sub cmdback_Click()
    frmEagle.Hide
    frmAirline.Show
End Sub

Private Sub cmdNWA_Click()
    MsgBox "For your spring break trip the lowest NWA flight from Minneapolis,MN to Denver and ending in Vail/Eagle,CO is $1091.20. This flight departs form Minneapolis at 8:53 A.M. and arrives in Vail/Eagle at 12:04 P.M. MST(mountain standard time).This flight is round trip and will depart a week later from Vail/Eagle at 8:09 A.M.(MST) and arrives at MSP. Int. at 3:38 P.M.(CST). This flight stops once in Denver which will continue on to MSP.  This flight is of coach class.", , "NWA Flight"
End Sub

