Attribute VB_Name = "modHelpers"
Option Explicit

' ==============================================================================
' Module: modHelpers
' Description: Core mathematical and data conversion utilities for EBCT models.
' ==============================================================================

' Converts an Excel Range object into a 2-D Double array.
Function RangeToDouble(rng As Range) As Double()
    Dim n1 As Long, n2 As Long
    n1 = rng.Rows.Count
    n2 = rng.Columns.Count
    
    Dim temp As Variant
    temp = rng.Value2
    
    Dim out() As Double
    ReDim out(1 To n1, 1 To n2)
    
    Dim row As Long, col As Long
    For row = 1 To n1
        For col = 1 To n2
            out(row, col) = CDbl(temp(row, col))
        Next col
    Next row
    
    RangeToDouble = out
End Function

' Calculates the arithmetic mean of a 2-D Double array.
Function mean(X() As Double) As Double
    Dim total As Double
    Dim i As Variant
    
    For Each i In X
        total = total + i
    Next i
    
    Dim totalElements As Long
    totalElements = (UBound(X, 1) - LBound(X, 1) + 1) * (UBound(X, 2) - LBound(X, 2) + 1)
    
    mean = total / totalElements
End Function

' Calculates the sample variance (n-1 degrees of freedom) of a 2-D Double array.
Function var_s(X() As Double) As Double
    Dim sumSq As Double
    Dim i As Variant
    Dim avgValX As Double
    
    avgValX = mean(X)
    
    For Each i In X
        sumSq = sumSq + (CDbl(i) - avgValX) ^ 2
    Next i
    
    Dim totalElements As Long
    totalElements = (UBound(X, 1) - LBound(X, 1) + 1) * (UBound(X, 2) - LBound(X, 2) + 1)
    
    var_s = sumSq / (totalElements - 1)
End Function

' Calculates row sums for a given Range, returning an n x 1 column vector.
Function rowsums(X As Range) As Double()
    Dim nrow As Long, ncol As Long
    nrow = X.Rows.Count
    ncol = X.Columns.Count
    
    Dim out() As Double
    ReDim out(1 To nrow, 1 To 1)
    
    Dim iterrow As Long, itercol As Long
    Dim rsum As Double
    
    For iterrow = 1 To nrow
        rsum = 0
        For itercol = 1 To ncol
            rsum = rsum + X(iterrow, itercol).Value2
        Next itercol
        out(iterrow, 1) = rsum
    Next iterrow
    
    rowsums = out
End Function

' Performs element-wise division between two matrices (Numerator / Denominator).
' Note: Parameter order in signature is (Denominator, Numerator) to match your existing calls.
Function MatrixDivision(denominator() As Double, Numerator() As Double) As Double()
    Dim nrow As Long, ncol As Long
    nrow = UBound(denominator, 1)
    ncol = UBound(denominator, 2)
    
    Dim out() As Double
    ReDim out(1 To nrow, 1 To ncol)
    
    Dim rowiter As Long, coliter As Long
    
    For rowiter = 1 To nrow
        For coliter = 1 To ncol
            If denominator(rowiter, coliter) = 0 Then
                out(rowiter, coliter) = 0
            Else
                out(rowiter, coliter) = Numerator(rowiter, coliter) / denominator(rowiter, coliter)
            End If
        Next coliter
    Next rowiter
    
    MatrixDivision = out
End Function
