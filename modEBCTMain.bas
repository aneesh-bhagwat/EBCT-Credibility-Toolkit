Attribute VB_Name = "modEBCTMain"
Option Explicit

' ==============================================================================
' Module: modEBCTMain
' Description: Primary calculation engine for Empirical Bayes Credibility
'              Theory (EBCT) Model 1 and Model 2 frameworks.
' ==============================================================================

' Complete Function for EBCT Model 1 Quantities
Function EBCT(rng As Range, Quantity As String) As Variant
    Dim iter As Long
    Dim nrow As Long
    Dim ncol As Long
    
    nrow = rng.Rows.Count
    ncol = rng.Columns.Count
    
    Dim Xibar() As Double
    Dim sqtheta() As Double
    Dim s As Double
    Dim q1 As Double
    Dim q2 As Double
    Dim v As Double
    
    Select Case UCase$(Trim$(Quantity))
        Case "M", "MEAN", "EM"
            ReDim Xibar(1 To 1, 1 To nrow)
            For iter = 1 To nrow
                Xibar(1, iter) = basic.mean(basic.RangeToDouble(rng.Rows(iter)))
            Next iter
            EBCT = basic.mean(Xibar)
            
        Case "S", "SQUARE", "PROCESSVAR", "SQR"
            ReDim sqtheta(1 To 1, 1 To nrow)
            For iter = 1 To nrow
                sqtheta(1, iter) = basic.var_s(basic.RangeToDouble(rng.Rows(iter)))
            Next iter
            EBCT = basic.mean(sqtheta)
            
        Case "V", "VAR", "BETWEENVAR"
            s = EBCT(rng, "S")
            ReDim Xibar(1 To 1, 1 To nrow)
            For iter = 1 To nrow
                Xibar(1, iter) = basic.mean(basic.RangeToDouble(rng.Rows(iter)))
            Next iter
            q1 = basic.var_s(Xibar)
            q2 = (1 / ncol) * s
            EBCT = q1 - q2
            
        Case "Z", "WEIGHT", "CREDIBILITY"
            s = EBCT(rng, "S")
            v = EBCT(rng, "V")
            If (s + v) <> 0 Then
                EBCT = ncol / (ncol + (s / v))
            Else
                EBCT = 0
            End If
            
        Case Else
            EBCT = CVErr(xlErrValue)
            
    End Select
End Function


' Complete Function for EBCT Model 2 Quantities
Function EBCT2(Y As Range, P As Range, Quantity As String) As Variant
    Dim nrow As Long
    Dim ncol As Long
    Dim Pbar As Double
    Dim Pibar() As Double
    Dim Pstar As Double
    Dim X() As Double
    Dim Xibar() As Double
    Dim Xbar As Double
    Dim yibar() As Double
    Dim sumY As Double
    Dim estheta As Double
    Dim emtheta As Double
    Dim pdouble() As Double
    Dim ydouble() As Double
    Dim varmtheta As Double
    Dim z_req As Double
    Dim Zi() As Double
    Dim out() As Variant
    Dim i As Long

    nrow = Y.Rows.Count
    ncol = Y.Columns.Count
    
    ReDim pdouble(1 To nrow, 1 To ncol)
    ReDim ydouble(1 To nrow, 1 To ncol)
    ReDim Pibar(1 To nrow, 1 To 1)
    ReDim yibar(1 To nrow, 1 To 1)
    ReDim Xibar(1 To nrow, 1 To 1)
    ReDim X(1 To nrow, 1 To ncol)
    
    pdouble = RangeToDouble(P)
    ydouble = RangeToDouble(Y)

    ' Calculate P Quantities
    Pbar = ebct2supplementary.CalculatePbar(P)
    Pibar = ebct2supplementary.CalculatePibar(P)
    Pstar = ebct2supplementary.CalculatePstar(P)

    ' Calculate X Quantities
    yibar = basic.rowsums(Y)
    X = basic.MatrixDivision(pdouble, ydouble)
    Xibar = ebct2supplementary.CalculateXibar(Y, Pbar, Pibar)
    
    sumY = mean(ydouble) * (nrow * ncol)
    If Pbar <> 0 Then
        Xbar = sumY / Pbar
    Else
        Xbar = 0
    End If

    ' EBCT 2 Structural Quantities
    emtheta = Xbar
    estheta = ebct2supplementary.CalculateEstheta(P, basic.MatrixDivision(pdouble, ydouble), Xibar)
    
    If Pstar <> 0 Then
        varmtheta = (ebct2supplementary.CalculateGrandDifference(P, X, Xbar) - estheta) / Pstar
    Else
        varmtheta = 0
    End If

    ' Dispatch on Requested Quantity Flag
    Select Case UCase$(Trim$(Quantity))
        Case "VARMTETA", "VAR"
            EBCT2 = varmtheta

        Case "ESQTHETA", "ESQ", "ESTHETA", "ES"
            EBCT2 = estheta

        Case "EMTHETA", "EM"
            EBCT2 = emtheta

        Case "Z", "ZI"
            ReDim Zi(1 To 1, 1 To nrow)
            For i = 1 To nrow
                If varmtheta <= 0 Then
                    Zi(1, i) = 1
                Else
                    z_req = Pibar(i, 1) / (Pibar(i, 1) + estheta / varmtheta)
                    Zi(1, i) = z_req
                End If
            Next i
            EBCT2 = Zi

        Case "ALL"
            ReDim out(1 To 1, 1 To 3 + nrow)
            out(1, 1) = estheta
            out(1, 2) = varmtheta
            out(1, 3) = emtheta
            For i = 1 To nrow
                If varmtheta <= 0 Then
                    out(1, 3 + i) = 1
                Else
                    z_req = Pibar(i, 1) / (Pibar(i, 1) + estheta / varmtheta)
                    out(1, 3 + i) = z_req
                End If
            Next i
            EBCT2 = out

        Case Else
            EBCT2 = CVErr(xlErrValue)

    End Select
End Function
