VERSION 5.00
Begin VB.Form frmCorRep 
   Caption         =   "Coroner's Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicBio 
      Height          =   7455
      Left            =   240
      ScaleHeight     =   7395
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmCorRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

PicBio.Print "Feb. 3, 1967"
PicBio.Print
PicBio.Print "SW1/4 Section 18, Lincoln Twp."
PicBio.Print
PicBio.Print "Queens, New York City, New York"
PicBio.Print
PicBio.Print "The deceased is an eighty-four year old male who has no known major medical history."
PicBio.Print "Mr. Brown was discovered dead 2 February 1967 by Mrs. Reginald Butterfax, living in"
PicBio.Print "neighboring apartment, 134B. On 36 December, Mrs. Butterfax noticed a bad odor"
PicBio.Print "emanating from Mr. Brown's closet room on the first floor. The smell strengthened"
PicBio.Print "over several days, and Mrs. Butterfux tried a number of times to contact Mr. Brown,"
PicBio.Print "who would not respond to knocking, did not possess a phone, and was not was seen"
PicBio.Print "performing his regular custodial duties. Mrs. Reginald called the fire department"
PicBio.Print "at 3:42 5 February. No key could be found which would opened the closet's lock,"
PicBio.Print "so the door was forcibly brought down by Off. Samuel Addux and Lt. Grey Forster."
PicBio.Print
PicBio.Print "Autopsy has shown Mr. Brown to have been deceased for one week at minimum upon"
PicBio.Print "the body's discovery. Cause of death: suffocation. A reptile of the genus-species"
PicBio.Print "Iguana iguana found forcibly wedged inside the victim's throat. Lesions on the"
PicBio.Print "throat's surface lining reveal the lizard to be living upon entry. Unusual levels"
PicBio.Print "of mercury were found on the victim's bloodstream, though this is not thought"
PicBio.Print "to have contributed to death."
PicBio.Print
PicBio.Print "I, Robert S. Protkin, M.D., Acting Coroner of Queens, New York City, New York"
PicBio.Print "on the 4th day of February, 1967 hereby certify that the above facts are made"
PicBio.Print "of record after diligent investigation and I believe them to be correct."

End Sub

Private Sub Label1_Click()

End Sub
