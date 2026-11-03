Attribute VB_Name = "modEBCT2Helpers"
Option Explicit

' ==============================================================================
' Module: modEBCT2Helpers
' Description: Calculation functions for Empirical Bayes Credibility Theory Model 2
' ==============================================================================

Function CalculatePbar(P As Range) As Double
    Dim Pbar As Double
    Pbar = basic.mean(RangeToDouble(P)) * (P.Rows.Count * P.Columns.Count)
    CalculatePbar = Pbar
End Function

Function CalculatePibar(P As Range) As Double()
    Dim Pibar() As Double
    Pibar = basic.rowsums(P)
    CalculatePibar = Pibar
End Function

Function CalculatePstar(P As Range) As Double
    Dim Pibar() As Double
    Pibar = CalculatePibar(P)
    
    Dim Pbar As Double
    Pbar = CalculatePbar(P)
    
    Dim nrow As Long, ncol As Long
    nrow = P.Rows.Count
    ncol = P.Columns.Count
    
    Dim pstarout As Double
    Dim iter As Long
    
    For iter = 1 To nrow
        If Pbar <> 0 Then
            pstarout = pstarout + Pibar(iter, 1) * (1 - Pibar(iter, 1) / Pbar)
        End If
    Next iter
    
    Dim denominator As Double
    denominator = (nrow * ncol - 1)
    
    If denominator <> 0 Then
        CalculatePstar = pstarout / denominator
    Else
        CalculatePstar = 0
    End If
End Function

Function CalculateEstheta(P As Range, X() As Double, Xibar() As Double) As Double
    Dim nrow As Long, ncol As Long
    Dim rowiter As Long, coliter As Long
    Dim row_variance() As Double
    Dim row_out As Double

    nrow = P.Rows.Count
    ncol = P.Columns.Count
    ReDim row_variance(1 To nrow, 1 To 1)

    For rowiter = 1 To nrow
        row_out = 0
        For coliter = 1 To ncol
            ' Using .Value2 for fast, format-agnostic numeric access '
            row_out = row_out + P(rowiter, coliter).Value2 * (X(rowiter, coliter) - Xibar(rowiter, 1)) ^ 2
        Next coliter
        
        If ncol > 1 Then
            row_variance(rowiter, 1) = row_out / (ncol - 1)
        Else
            row_variance(rowiter, 1) = 0
        End If
    Next rowiter

    CalculateEstheta = basic.mean(row_variance)
End Function

Function CalculateGrandDifference(P As Range, X() As Double, Xbar As Double) As Double
    Dim nrow As Long, ncol As Long
    Dim rowiter As Long, coliter As Long
    Dim out As Double
    
    nrow = P.Rows.Count
    ncol = P.Columns.Count
    
    For rowiter = 1 To nrow
        For coliter = 1 To ncol
            out = out + P(rowiter, coliter).Value2 * ((X(rowiter, coliter) - Xbar) ^ 2)
        Next coliter
    Next rowiter
    
    Dim denominator As Double
    denominator = (nrow * ncol - 1)
    
    If denominator <> 0 Then
        CalculateGrandDifference = out / denominator
    Else
        CalculateGrandDifference = 0
    End If
End Function

Function CalculateXibar(Y As Range, Pbar As Double, Pibar() As Double) As Double()
    Dim nrow As Long
    Dim Xibar() As Double
    Dim iter As Long
    
    nrow = Y.Rows.Count
    
    ' rowsums already dimensions out as nrow x 1 '
    Xibar = basic.rowsums(Y)
    
    For iter = 1 To nrow
        If Pibar(iter, 1) <> 0 Then
            Xibar(iter, 1) = Xibar(iter, 1) / Pibar(iter, 1)
        Else
            Xibar(iter, 1) = 0
        End If
    Next iter
    
    CalculateXibar = Xibar
End Function
