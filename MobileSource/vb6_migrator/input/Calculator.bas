Attribute VB_Name = "Calculator"
Option Explicit

' Simple calculator module for testing migration

Public Function Add(ByVal a As Double, ByVal b As Double) As Double
    Add = a + b
End Function

Public Function Subtract(ByVal a As Double, ByVal b As Double) As Double
    Subtract = a - b
End Function

Public Function Multiply(ByVal a As Double, ByVal b As Double) As Double
    Multiply = a * b
End Function

Public Function Divide(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then
        Err.Raise 11, "Calculator", "Division by zero"
    End If
    Divide = a / b
End Function

Public Function Factorial(ByVal n As Long) As Long
    Dim result As Long
    Dim i As Long

    If n < 0 Then
        Err.Raise 5, "Calculator", "Negative number"
    End If

    result = 1
    For i = 2 To n
        result = result * i
    Next i

    Factorial = result
End Function

Public Function IsPrime(ByVal n As Long) As Boolean
    Dim i As Long

    If n <= 1 Then
        IsPrime = False
        Exit Function
    End If

    If n <= 3 Then
        IsPrime = True
        Exit Function
    End If

    If n Mod 2 = 0 Or n Mod 3 = 0 Then
        IsPrime = False
        Exit Function
    End If

    i = 5
    Do While i * i <= n
        If n Mod i = 0 Or n Mod (i + 2) = 0 Then
            IsPrime = False
            Exit Function
        End If
        i = i + 6
    Loop

    IsPrime = True
End Function
