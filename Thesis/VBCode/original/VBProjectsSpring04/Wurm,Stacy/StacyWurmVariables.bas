Attribute VB_Name = "StacyWurmVariables"
'Project Name: Date Chooser (Wurm, Stacy - VB Project)
' Form Name: This is the module
' Author: Stacy Wurm
' Date Written: Sunday, March 14th, 2004
' Purpose of this Form: ' this module declares variables so they can be used from
                        ' one form to the next for displays and such
    ' Variables for the arrays
        Public Item(1 To 15) As String
        Public Price(1 To 15) As Single
    ' variables for my array sorting
        Public PASS As Integer, COMP As Integer, J As Integer
        Public tempPrice As Single, tempItem As String
    ' Variables for Cost and Price calculations
        Public CTR As Single
        Public Cost As Single
        Public Budget As Single
        Public TotalCost As Single
        Public Choice As String
        Public MovieCost As Single
    ' name and decision variables
        Public UserName As String
        Public Decision1 As String, Decision2 As String
        Public Decision3 As String, Decision4 As String
