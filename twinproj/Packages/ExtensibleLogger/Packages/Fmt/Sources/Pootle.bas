Attribute VB_Name = "Pootle"
Option Explicit

Sub SimpleTest()
    
    Dim myApples As String
    myApples = "Bramleys and Cox"
    Dim myPi As Double
    myPi = 3.142
    
    Debug.Print StringFormat("I'm eating a very nice {0} {1}", myApples, myPi)
    Debug.Print StringFormat("The current price is {0:C2} per ounce.", 17.63245)
    Debug.Print StringFormat("It is now {0:d} at {0:t}", Now())
    Debug.Print StringFormat("LHS{0,8} {1,15}\n\nTwo lines later", "Year", "Population")
    
End Sub
