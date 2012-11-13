Attribute VB_Name = "Module1"
'FinalProject:Travel Europe
'Module1
'Jessica Florek
'Written: 3/6/09-3/20/09
'Objective: This form declares all global variables and makes the
'usable throughout the entire project.




Option Explicit

Public flightcost As Single, budget As Single, budget2 As Single, duration As Integer, foodcost As Single, euros As Single
Public citycounter As Integer, Londonhotelcost As Single, londonattractioncost As Single
Public Parishotelcost As Single, parisattractioncost As Single
Public Madridhotelcost As Single, madridattractioncost As Single
Public Venicehotelcost As Single, veniceattractioncost As Single
Public venice As Boolean, paris As Boolean, london As Boolean, madrid As Boolean


'all of these variables are dimmed as public in order to keep them as running totals for the final budget summary
'the cities are dimmed as booleans so only the relevent information is displayed in the final budget


