VERSION 5.00
Begin VB.Form frmKeystoneresorts 
   BackColor       =   &H00000000&
   Caption         =   "Keystone Resorts"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmKeystoneresorts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      Height          =   1695
      Left            =   240
      Picture         =   "frmKeystoneresorts.frx":1CBB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Height          =   1695
      Left            =   240
      Picture         =   "frmKeystoneresorts.frx":3BBE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmd4 
      Height          =   1695
      Left            =   240
      Picture         =   "frmKeystoneresorts.frx":5042
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmd5 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmKeystoneresorts.frx":6FE6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   5520
      ScaleHeight     =   7155
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton cmbBack 
      Caption         =   "Back to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmKeystoneresorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmKeystoneresorts(frmKeystoneresortsresorts.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  this form allows the user to get a greater sense of what each
'resort is about. it explains each resort in greater detail than just the price.

Private Sub cmbBack_Click()
    frmKeystoneresorts.Hide 'hides this form
    frmKeystoneLodge.Show 'brings you back to keystone lodge form
End Sub
Private Sub cmd1_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Lakeside Village Condominiums"
    picResults.Print 'prints a blank line
    picResults.Print "These spacious, comfortable condos are located in Keystone's"
    picResults.Print "original village right on the Lake. With an atmosphere that is"
    picResults.Print "festive year-round, Lakeside Village offers plenty of choices"
    picResults.Print "with shops, galleries and restaurants right out the door."
End Sub

Private Sub cmd2_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Keystone Lodge, A RockResort"
    picResults.Print ''prints a blank line
    picResults.Print "Keystone Lodge, A RockResort, located in Lakeside Village, is a"
    picResults.Print "Preferred Hotel, member of RockResorts Hotels and a AAA Four-Diamond"
    picResults.Print "rated property. It provides guests with an abundance of amenities as"
    picResults.Print "well as access to shopping, activities and events located around"
    picResults.Print "Keystone Lake. Direct shuttle to slopes."
End Sub

Private Sub cmd3_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Riverbank Condominiums"
    picResults.Print 'prints a blank line
    picResults.Print "If you want a guaranteed good view, Riverbank is for you. River Run"
    picResults.Print "Village with its great shopping, dining, and nightlife is right"
    picResults.Print "there. These units are a good value for their location and feature"
    picResults.Print "jetted Jacuzzi tubs and fireplaces in every unit. 300 yds to slopes."
End Sub

Private Sub cmd4_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "The Springs"
    picResults.Print 'prints a blank line
    picResults.Print "Built adjacent to the River Run Village and close enough to walk to"
    picResults.Print "the Gondola, it's clear why these are some of the most popular"
    picResults.Print "lodging units at Keystone. Less than 200 yds to slopes."
End Sub

Private Sub cmd5_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Lone Eagle Condominiums"
    picResults.Print 'prints a blank line
    picResults.Print "These ski in/ski out up-scale condominiums are the closest lodging"
    picResults.Print "units to the ski slopes and are located on the eastern slopes of"
    picResults.Print "Keystone Mountain. Ski in/ski out."
End Sub
