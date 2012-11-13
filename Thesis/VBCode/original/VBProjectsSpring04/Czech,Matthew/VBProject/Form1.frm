VERSION 5.00
Begin VB.Form frmsjuhockey 
   BackColor       =   &H00FF0000&
   Caption         =   "SJU HOCKEY"
   ClientHeight    =   11010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdloadarrays 
      BackColor       =   &H000000FF&
      Caption         =   "2003-2004 University of St. John's:  Hockey Statistics Click To Enter"
      Height          =   1215
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblcreator 
      Caption         =   "By: Matthew Czech"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Image imgteampic 
      Height          =   9000
      Left            =   -3120
      Picture         =   "Form1.frx":000C
      Top             =   360
      Width           =   22500
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   360
      Top             =   3600
      Width           =   4215
   End
End
Attribute VB_Name = "frmsjuhockey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : SJU HOCKEY (Matthew Czech's VB Project.vbp)
'Form Name : frmsjuhockey (Form1.frm)
'Author: Matthew Czech
'Date Written: March 12, 2003
'Purpose of Project: To inform people about how good of a hockey tean that you are.
                 'To see the complied data for my self and anyalyze what I can do to get better
                 'It is something that I am interested in and hopefully someone cane learn a
                 'little bit abought hockey from doing this."
'Purpose of this form: To load the arrays into the  program so they can be used.

Private Sub cmdloadarrays_Click()
Path = "N:\CS130\handin\Czech, Matthew\New Folder\"
Open Path & "sjustats.txt" For Input As #1

    For CTR = 1 To 24 'fills arrays
        Input #1, numbers(CTR), names(CTR), gp(CTR), goals(CTR), assists(CTR), shots(CTR), plusmin(CTR), penmin(CTR), pp(CTR), sh(CTR), gw(CTR)
    Next CTR
Close #1 'close file
frmsjuhockey.Hide
frmstatscalsulations.Show
End Sub
