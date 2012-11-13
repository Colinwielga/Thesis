VERSION 5.00
Begin VB.Form frmKeystone 
   Caption         =   "Keystone"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form4"
   Picture         =   "frmKeystone.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtickets 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdLodge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lodging"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Airfare"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quick Facts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1095
   End
   Begin VB.PictureBox piclogo 
      Height          =   1215
      Left            =   240
      Picture         =   "frmKeystone.frx":1D71E
      ScaleHeight     =   1155
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmKeystone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to navigate all the features of the
'beaver creek resort. it allows the user to access forms that search for ticket prices,
'resorts, and airfair. this form also allows the user to read up on quick facts about the
'ski resort

Private Sub cmdAir_Click()
    frmKeystone.Hide
    frmAirline.Show

End Sub

Private Sub cmdback_Click()
    frmKeystone.Visible = False
    frmContents.Visible = True
End Sub

Private Sub cmdFacts_Click()
    MsgBox "Keystone Resort is comprised of three incredible and unique mountains - Dercum Mountain, North Peak and The Outback.  Each mountain offers some of the best terrain Colorado has to offer steeps, bowls, bumps, glades, rails, hits, lights - you name it - Keystone Has it.", , "Keystone Facts"

End Sub

Private Sub cmdLodge_Click()
    frmKeystone.Hide
    frmKeystoneLodge.Show

End Sub

Private Sub cmdtickets_Click()
    frmKeystone.Hide
    frmKeystonetix.Show
End Sub

