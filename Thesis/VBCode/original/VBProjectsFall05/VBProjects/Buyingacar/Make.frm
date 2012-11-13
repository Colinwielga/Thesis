VERSION 5.00
Begin VB.Form MakeForm 
   BackColor       =   &H80000012&
   Caption         =   "Make of car"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToyota 
      BackColor       =   &H80000003&
      Caption         =   "Toyota"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      MaskColor       =   &H008080FF&
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolkswagon 
      BackColor       =   &H8000000D&
      Caption         =   "Volkswagon"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      MaskColor       =   &H008080FF&
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdFord 
      BackColor       =   &H80000012&
      Caption         =   "Ford"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaskColor       =   &H008080FF&
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   4920
      Picture         =   "Make.frx":0000
      Top             =   1920
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   8040
      Picture         =   "Make.frx":0C04
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Image FordLogo 
      Height          =   675
      Left            =   1200
      Picture         =   "Make.frx":14E8
      Top             =   2040
      Width           =   1560
   End
End
Attribute VB_Name = "MakeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Buying A Car(VB-project.vbp)
'Form Name : MakeForm(Make.frm)
'Author : Katie Lee
'Purpose of the form: User selects a specific company
Option Explicit
Private Sub cmdFord_Click()
FordForm.Show ' Brings user to FordForm
MakeForm.Hide
End Sub

Private Sub cmdVolkswagon_Click()
VolkswagonForm.Show 'Brings user to VolkswagonForm
MakeForm.Hide
End Sub

Private Sub cmdToyota_Click()
ToyotaForm.Show 'Brings user to ToyotaForm
MakeForm.Hide
End Sub

