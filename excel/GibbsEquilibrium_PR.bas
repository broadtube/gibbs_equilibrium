Attribute VB_Name = "GibbsEquilibrium_PR"
'=============================================================================
' Gibbs Equilibrium Calculator for Excel — Non-ideal gas version (PR / RK)
'
' Extends the ideal-gas module (GibbsEquilibrium.bas) with Peng-Robinson and
' Redlich-Kwong equations of state via a fixed-point "Picard" loop:
'
'   μᵢ/RT = gᵢ⁰(T)/RT + ln(yᵢ·P/P°) + ln(φᵢ)    (gas)
'   μᵢ/RT = gᵢ⁰(T)/RT                          (solid, activity = 1)
'
' Strategy:
'   Stage 1: Solve with φᵢ = 1 (ideal). Get composition y⁽¹⁾.
'   Stage 2: Compute φᵢ(T, P, y⁽¹⁾) via PR or RK, write to phi column.
'   Stage 3: Re-solve. Get y⁽²⁾.
'   Iterate until max|Δφᵢ| < 1e-5 or MAX_PASSES reached.
'
' This avoids solving the cubic EOS inside every Solver iteration (which is
' too slow via UDF and numerically unstable because the vapor/liquid root
' selection creates non-smooth behaviour that GRG cannot follow).
'
' Module coexists with the ideal-gas "GibbsEquilibrium" module: public
' Subs are suffixed _PR. Sheet names (Data / Input / Solve / Output) are
' shared, so running SetupWorkbook_PR rebuilds sheets with the extra
' Tc / Pc / omega / phi columns.
'
' Data sources are identical to the ideal version plus the critical-
' parameter set (NIST WebBook / IAPWS-95 / Reid-Prausnitz-Poling).
'
' Usage:
'   1. In VBE: Tools → References → check "Solver" (or keep the Solver
'      add-in loaded; Application.Run late binding also works).
'   2. Run SetupWorkbook_PR once.
'   3. Edit Input sheet: T [°C], P [atm], EOS, inlet moles.
'   4. Run RunEquilibrium_PR.
'=============================================================================
Option Explicit

Private Const R_GAS As Double = 8.31446261815324   ' J/(mol·K)
Private Const P_REF As Double = 101325#             ' Pa

Private Const SHEET_DATA As String = "Data"
Private Const SHEET_INPUT As String = "Input"
Private Const SHEET_SOLVE As String = "Solve"
Private Const SHEET_OUTPUT As String = "Output"

Private Const N_SPECIES As Long = 17
Private Const ELEMENTS As String = "C,H,O,N"
Private Const N_ELEMENTS As Long = 4

Private Const PHI_COL As Long = 11                 ' column K on Solve sheet
Private Const MAX_PASSES As Long = 10               ' Picard loop cap
Private Const PHI_TOL As Double = 0.00001           ' 1e-5 convergence tolerance

'=============================================================================
' Public entry points
'=============================================================================

Public Sub SetupWorkbook_PR()
    Application.ScreenUpdating = False
    On Error GoTo fail

    Call EnsureSheet(SHEET_DATA)
    Call EnsureSheet(SHEET_INPUT)
    Call EnsureSheet(SHEET_SOLVE)
    Call EnsureSheet(SHEET_OUTPUT)

    Call WriteDataSheet
    Call WriteInputSheet
    Call WriteSolveSheet
    Call ClearOutputSheet

    Worksheets(SHEET_INPUT).Activate
    Application.ScreenUpdating = True
    MsgBox "Setup complete (PR/RK version)." & vbCrLf & _
           "1) Edit the Input sheet (T, P, EOS, inlet moles)." & vbCrLf & _
           "2) Run macro RunEquilibrium_PR.", vbInformation, "Gibbs Equilibrium PR"
    Exit Sub

fail:
    Application.ScreenUpdating = True
    MsgBox "SetupWorkbook_PR failed: " & Err.Description, vbCritical
End Sub

Public Sub RunEquilibrium_PR()
    Application.ScreenUpdating = False
    On Error GoTo fail

    Dim wsInput As Worksheet, wsSolve As Worksheet
    Set wsInput = Worksheets(SHEET_INPUT)
    Set wsSolve = Worksheets(SHEET_SOLVE)

    Dim eos As String
    eos = CStr(wsInput.Range("B5").Value)
    If eos = "" Then eos = "Peng-Robinson"

    ' Reset n_i seed and phi=1 for the first pass
    Call PrepareSolveSheet
    Call SetAllPhi(1#)

    Dim passIdx As Long
    Dim maxDeltaPhi As Double
    Dim lastSolveRes As Variant
    Dim passesUsed As Long: passesUsed = 0

    For passIdx = 1 To MAX_PASSES
        passesUsed = passIdx
        lastSolveRes = InvokeSolver()

        If eos = "ideal-gas" Then
            Exit For
        End If

        ' Snapshot old phi, compute new phi from current composition
        Dim oldPhi() As Double
        Call ReadPhi(oldPhi)
        Call ComputeFugacities(eos)
        Dim newPhi() As Double
        Call ReadPhi(newPhi)

        maxDeltaPhi = 0#
        Dim k As Long
        For k = 1 To N_SPECIES
            Dim d As Double: d = Abs(newPhi(k) - oldPhi(k))
            If d > maxDeltaPhi Then maxDeltaPhi = d
        Next k
        If maxDeltaPhi < PHI_TOL Then Exit For
    Next passIdx

    Call WriteOutputSheet(lastSolveRes, passesUsed, eos, maxDeltaPhi)
    Worksheets(SHEET_OUTPUT).Activate
    Application.ScreenUpdating = True
    Exit Sub

fail:
    Application.ScreenUpdating = True
    MsgBox "RunEquilibrium_PR failed: " & Err.Description & vbCrLf & _
           "Check that the Solver add-in is loaded " & _
           "(File → Options → Add-ins → Solver).", vbCritical
End Sub

'=============================================================================
' Solver invocation
'=============================================================================
Private Function InvokeSolver() As Variant
    Dim wsSolve As Worksheet: Set wsSolve = Worksheets(SHEET_SOLVE)
    Dim setCellAddr As String, byChangeAddr As String
    Dim elemActAddr As String, elemTgtAddr As String
    setCellAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("G_total").Address
    byChangeAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("n_vars").Address
    elemActAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("elem_actual").Address
    elemTgtAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("elem_target").Address

    Application.Run "SolverReset"
    Application.Run "SolverOk", setCellAddr, 2, 0, byChangeAddr, 1   ' Min, GRG
    Application.Run "SolverAdd", byChangeAddr, 3, "1E-20"            ' n_i >= 1e-20
    Application.Run "SolverAdd", elemActAddr, 2, elemTgtAddr         ' element balance
    Application.Run "SolverOptions", 120, 2000, 0.0000001            ' MaxTime, Iter, Prec
    InvokeSolver = Application.Run("SolverSolve", True)
    Application.Run "SolverFinish", 1
End Function

'=============================================================================
' Sheet construction
'=============================================================================
Private Sub WriteDataSheet()
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_DATA)
    ws.Cells.Clear

    Dim hdr As Variant
    hdr = Array("Species", "Phase", "Tlow", "Tmid", "Thigh", _
                "a1_lo", "a2_lo", "a3_lo", "a4_lo", "a5_lo", "a6_lo", "a7_lo", _
                "a1_hi", "a2_hi", "a3_hi", "a4_hi", "a5_hi", "a6_hi", "a7_hi", _
                "C", "H", "O", "N", _
                "Tc[K]", "Pc[Pa]", "omega", "Source")
    Dim j As Long
    For j = 0 To UBound(hdr)
        ws.Cells(1, j + 1).Value = hdr(j)
    Next j
    ws.Rows(1).Font.Bold = True

    Dim r As Long: r = 2
    ' -- GRI-Mech 3.0 (gri30.yaml) species -------------------------------------
    Call WriteSpeciesRow(ws, r, "H2", "gas", 200#, 1000#, 3500#, _
        Array(2.34433112, 0.00798052075, -1.9478151E-05, 2.01572094E-08, -7.37611761E-12, -917.935173, 0.683010238), _
        Array(3.3372792, -4.94024731E-05, 4.99456778E-07, -1.79566394E-10, 2.00255376E-14, -950.158922, -3.20502331), _
        0, 2, 0, 0, 33.19, 1.313E+06, -0.216, "gri30.yaml (GRI-Mech 3.0)")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CO2", "gas", 200#, 1000#, 3500#, _
        Array(2.35677352, 0.00898459677, -7.12356269E-06, 2.45919022E-09, -1.43699548E-13, -48371.9697, 9.90105222), _
        Array(3.85746029, 0.00441437026, -2.21481404E-06, 5.23490188E-10, -4.72084164E-14, -48759.166, 2.27163806), _
        1, 0, 2, 0, 304.21, 7.383E+06, 0.224, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CO", "gas", 200#, 1000#, 3500#, _
        Array(3.57953347, -0.00061035368, 1.01681433E-06, 9.07005884E-10, -9.04424499E-13, -14344.086, 3.50840928), _
        Array(2.71518561, 0.00206252743, -9.98825771E-07, 2.30053008E-10, -2.03647716E-14, -14151.8724, 7.81868772), _
        1, 0, 1, 0, 132.92, 3.499E+06, 0.048, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "H2O", "gas", 200#, 1000#, 3500#, _
        Array(4.19864056, -0.0020364341, 6.52040211E-06, -5.48797062E-09, 1.77197817E-12, -30293.7267, -0.849032208), _
        Array(3.03399249, 0.00217691804, -1.64072518E-07, -9.7041987E-11, 1.68200992E-14, -30004.2971, 4.9667701), _
        0, 2, 1, 0, 647.096, 22.064E+06, 0.3443, "gri30.yaml; Tc/Pc IAPWS-95")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH4", "gas", 200#, 1000#, 3500#, _
        Array(5.14987613, -0.0136709788, 4.91800599E-05, -4.84743026E-08, 1.66693956E-11, -10246.6476, -4.64130376), _
        Array(0.074851495, 0.0133909467, -5.73285809E-06, 1.22292535E-09, -1.0181523E-13, -9468.34459, 18.437318), _
        1, 4, 0, 0, 190.56, 4.599E+06, 0.011, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3OH", "gas", 200#, 1000#, 3500#, _
        Array(5.71539582, -0.0152309129, 6.52441155E-05, -7.10806889E-08, 2.61352698E-11, -25642.7656, -1.50409823), _
        Array(1.78970791, 0.0140938292, -6.36500835E-06, 1.38171085E-09, -1.1706022E-13, -25374.8747, 14.5023623), _
        1, 4, 1, 0, 512.6, 8.084E+06, 0.565, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3CHO", "gas", 200#, 1000#, 6000#, _
        Array(4.7294595, -0.0031932858, 4.7534921E-05, -5.7458611E-08, 2.1931112E-11, -21572.878, 4.1030159), _
        Array(5.4041108, 0.011723059, -4.2263137E-06, 6.8372451E-10, -4.0984863E-14, -22593.122, -3.4807917), _
        2, 4, 1, 0, 466#, 5.57E+06, 0.291, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "HCHO", "gas", 200#, 1000#, 3500#, _
        Array(4.79372315, -0.00990833369, 3.73220008E-05, -3.79285261E-08, 1.31772652E-11, -14308.9567, 0.6028129), _
        Array(1.76069008, 0.00920000082, -4.42258813E-06, 1.00641212E-09, -8.8385564E-14, -13995.8323, 13.656323), _
        1, 2, 1, 0, 408#, 6.59E+06, 0.282, "gri30.yaml (CH2O)")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H6", "gas", 200#, 1000#, 3500#, _
        Array(4.29142492, -0.0055015427, 5.99438288E-05, -7.08466285E-08, 2.68685771E-11, -11522.2055, 2.66682316), _
        Array(1.0718815, 0.0216852677, -1.00256067E-05, 2.21412001E-09, -1.9000289E-13, -11426.3932, 15.1156107), _
        2, 6, 0, 0, 305.32, 4.872E+06, 0.099, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H4", "gas", 200#, 1000#, 3500#, _
        Array(3.95920148, -0.00757052247, 5.70990292E-05, -6.91588753E-08, 2.69884373E-11, 5089.77593, 4.09733096), _
        Array(2.03611116, 0.0146454151, -6.71077915E-06, 1.47222923E-09, -1.25706061E-13, 4939.88614, 10.3053693), _
        2, 4, 0, 0, 282.34, 5.041E+06, 0.087, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "N2", "gas", 300#, 1000#, 5000#, _
        Array(3.298677, 0.0014082404, -3.963222E-06, 5.641515E-09, -2.444854E-12, -1020.8999, 3.950372), _
        Array(2.92664, 0.0014879768, -5.68476E-07, 1.0097038E-10, -6.753351E-15, -922.7977, 5.980528), _
        0, 0, 0, 2, 126.2, 3.39E+06, 0.037, "gri30.yaml")
    r = r + 1

    ' -- NASA Glenn (nasa_gas.yaml) species ------------------------------------
    Call WriteSpeciesRow(ws, r, "CH3OCH3", "gas", 200#, 1000#, 6000#, _
        Array(5.30562279, -0.00214254272, 5.30873244E-05, -6.23147136E-08, 2.30731036E-11, -23986.6295, 0.713264209), _
        Array(5.64844183, 0.0163381899, -5.86802367E-06, 9.46836869E-10, -5.66504738E-14, -25107.469, -5.96264939), _
        2, 6, 1, 0, 400.1, 5.37E+06, 0.2, "nasa_gas.yaml L12/92")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3COOH", "gas", 200#, 1000#, 6000#, _
        Array(2.78936844, 0.0100001016, 3.42557978E-05, -5.09017919E-08, 2.06217504E-11, -53475.2292, 14.1059504), _
        Array(7.67083678, 0.0135152695, -5.25874688E-06, 8.93185062E-10, -5.53180891E-14, -55756.0971, -15.467659), _
        2, 4, 2, 0, 591.95, 5.786E+06, 0.467, "nasa_gas.yaml L 8/88")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H5OH", "gas", 200#, 1000#, 6000#, _
        Array(4.85868178, -0.0037400674, 6.95550267E-05, -8.86541147E-08, 3.5168443E-11, -29996.1309, 4.80192294), _
        Array(6.5628977, 0.0152034264, -5.38922247E-06, 8.62150224E-10, -5.12824683E-14, -31525.7984, -9.47557644), _
        2, 6, 1, 0, 513.92, 6.148E+06, 0.649, "nasa_gas.yaml L 8/88")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "HCOOH", "gas", 200#, 1000#, 6000#, _
        Array(3.23262453, 0.00281129582, 2.44034975E-05, -3.17501066E-08, 1.2063166E-11, -46778.5606, 9.86205647), _
        Array(5.69579404, 0.00772237361, -3.18037808E-06, 5.57949466E-10, -3.52618226E-14, -48159.9723, -6.0168008), _
        1, 2, 2, 0, 588#, 5.81E+06, 0.473, "nasa_gas.yaml L 8/88")
    r = r + 1

    ' -- Burcat T10/07 --------------------------------------------------------
    Call WriteSpeciesRow(ws, r, "CH3COOCH3", "gas", 200#, 1000#, 6000#, _
        Array(7.18744749, -0.00629221513, 8.17059377E-05, -9.82940778E-08, 3.73744521E-11, -52341.7155, -3.24161798), _
        Array(8.38776809, 0.0190836514, -6.8219732E-06, 1.09765423E-09, -6.55561842E-14, -54080.5971, -16.4156253), _
        3, 6, 2, 0, 506.5, 4.75E+06, 0.32, "Burcat T10/07 Meacetate")
    r = r + 1

    ' -- JANAF X 4/83 solid ---------------------------------------------------
    ' Tc/Pc/omega blank for solid (handled as activity = 1)
    Call WriteSpeciesRow(ws, r, "C(gr)", "solid", 200#, 1000#, 5000#, _
        Array(-0.310872072, 0.00440353686, 1.90394118E-06, -6.38546966E-09, 2.98964248E-12, -108.650794, 1.11382953), _
        Array(1.45571829, 0.00171702216, -6.97562786E-07, 1.35277032E-10, -9.67590652E-15, -695.138814, -8.52583033), _
        1, 0, 0, 0, 0#, 0#, 0#, "nasa_condensed.yaml X 4/83 (solid, no EOS)")

    ws.Columns("A:AA").AutoFit
End Sub

Private Sub WriteSpeciesRow(ws As Worksheet, r As Long, name As String, phase As String, _
                            Tl As Double, Tm As Double, Th As Double, _
                            aLo As Variant, aHi As Variant, _
                            nC As Long, nH As Long, nO As Long, nN As Long, _
                            Tc As Double, Pc As Double, omega As Double, _
                            src As String)
    ws.Cells(r, 1).Value = name
    ws.Cells(r, 2).Value = phase
    ws.Cells(r, 3).Value = Tl
    ws.Cells(r, 4).Value = Tm
    ws.Cells(r, 5).Value = Th
    Dim i As Long
    For i = 0 To 6
        ws.Cells(r, 6 + i).Value = aLo(i)
        ws.Cells(r, 13 + i).Value = aHi(i)
    Next i
    ws.Cells(r, 20).Value = nC
    ws.Cells(r, 21).Value = nH
    ws.Cells(r, 22).Value = nO
    ws.Cells(r, 23).Value = nN
    ws.Cells(r, 24).Value = Tc       ' X
    ws.Cells(r, 25).Value = Pc       ' Y
    ws.Cells(r, 26).Value = omega    ' Z
    ws.Cells(r, 27).Value = src      ' AA
End Sub

Private Sub WriteInputSheet()
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_INPUT)
    ws.Cells.Clear

    ws.Range("A1").Value = "Gibbs Equilibrium — Input (PR / RK)"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13

    ws.Range("A3").Value = "Temperature [°C]:"
    ws.Range("B3").Value = 250
    ws.Range("A4").Value = "Pressure [atm]:"
    ws.Range("B4").Value = 50
    ws.Range("A5").Value = "EOS:"
    ws.Range("B5").Value = "Peng-Robinson"
    With ws.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="ideal-gas,Peng-Robinson,Redlich-Kwong"
    End With

    ws.Range("A7").Value = "Species"
    ws.Range("B7").Value = "Inlet moles"
    ws.Range("C7").Value = "Enabled (1/0)"
    ws.Range("D7").Value = "Phase"
    ws.Range("A7:D7").Font.Bold = True

    Dim r As Long, i As Long
    r = 8
    For i = 2 To N_SPECIES + 1
        ws.Cells(r, 1).Value = Worksheets(SHEET_DATA).Cells(i, 1).Value
        ws.Cells(r, 2).Value = 0
        ws.Cells(r, 3).Value = 1
        ws.Cells(r, 4).Value = Worksheets(SHEET_DATA).Cells(i, 2).Value
        r = r + 1
    Next i

    ' Default preset: DME synthesis H2:CO = 2:1
    Worksheets(SHEET_INPUT).Range("B8").Value = 2   ' H2
    Worksheets(SHEET_INPUT).Range("B10").Value = 1  ' CO

    ws.Columns("A:D").AutoFit
    ws.Range("B3:B5").Interior.Color = RGB(255, 255, 204)
    ws.Range("B8:C" & (7 + N_SPECIES)).Interior.Color = RGB(255, 255, 204)
End Sub

Private Sub WriteSolveSheet()
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_SOLVE)
    ws.Cells.Clear

    ws.Range("A1").Value = "T [K]:"
    ws.Range("B1").Formula = "=" & SHEET_INPUT & "!B3+273.15"
    ws.Range("A2").Value = "P [Pa]:"
    ws.Range("B2").Formula = "=" & SHEET_INPUT & "!B4*" & P_REF
    ws.Range("A3").Value = "ln(P/Pref):"
    ws.Range("B3").Formula = "=LN(B2/" & P_REF & ")"
    ws.Range("A4").Value = "EOS:"
    ws.Range("B4").Formula = "=" & SHEET_INPUT & "!B5"

    ws.Range("A5").Value = "Species"
    ws.Range("B5").Value = "Phase"
    ws.Range("C5").Value = "Enabled"
    ws.Range("D5").Value = "n_i"
    ws.Range("E5").Value = "g0/RT"
    ws.Range("F5").Value = "n·μ/RT"
    ws.Range("G5").Value = "n·aC"
    ws.Range("H5").Value = "n·aH"
    ws.Range("I5").Value = "n·aO"
    ws.Range("J5").Value = "n·aN"
    ws.Range("K5").Value = "phi_i"
    ws.Range("A5:K5").Font.Bold = True

    Dim firstRow As Long, lastRow As Long
    firstRow = 6
    lastRow = firstRow + N_SPECIES - 1

    Dim i As Long, sr As Long
    For i = 1 To N_SPECIES
        sr = firstRow + i - 1
        Dim dataRow As Long: dataRow = i + 1

        ws.Cells(sr, 1).Formula = "=" & SHEET_DATA & "!A" & dataRow
        ws.Cells(sr, 2).Formula = "=" & SHEET_DATA & "!B" & dataRow
        ws.Cells(sr, 3).Formula = "=" & SHEET_INPUT & "!C" & (i + 7)   ' enabled (row 8..)
        ws.Cells(sr, 4).Value = 0.000000001                           ' seed n
        ws.Cells(sr, PHI_COL).Value = 1#                              ' phi init = 1

        ' g0/RT
        ws.Cells(sr, 5).Formula = _
            "=IF(C" & sr & "=0, 1E30, " & _
            "IF(B1<=" & SHEET_DATA & "!D" & dataRow & ", " & _
                NasaGoverRTFormula(dataRow, 6) & ", " & _
                NasaGoverRTFormula(dataRow, 13) & "))"

        ' n·μ/RT  — gas now includes +LN(phi_i); solid is unchanged
        ws.Cells(sr, 6).Formula = _
            "=IF(C" & sr & "=0, 0, " & _
            "IF(B" & sr & "=""gas"", " & _
                "D" & sr & "*(E" & sr & "+LN(MAX(D" & sr & "/nTot_gas,1E-300))+B3+LN(MAX(K" & sr & ",1E-30))), " & _
                "D" & sr & "*E" & sr & "))"

        ' Element contributions (n · a_jk · enabled)
        ws.Cells(sr, 7).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!T" & dataRow
        ws.Cells(sr, 8).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!U" & dataRow
        ws.Cells(sr, 9).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!V" & dataRow
        ws.Cells(sr, 10).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!W" & dataRow
    Next i

    ' nTot_gas
    Dim nTotRow As Long: nTotRow = lastRow + 2
    ws.Cells(nTotRow, 1).Value = "nTot_gas:"
    ws.Cells(nTotRow, 2).Formula = "=SUMPRODUCT((B" & firstRow & ":B" & lastRow & _
                                   "=""gas"")*C" & firstRow & ":C" & lastRow & _
                                   "*D" & firstRow & ":D" & lastRow & ")"
    ws.Cells(nTotRow, 1).Font.Bold = True

    ' G_total
    Dim gRow As Long: gRow = lastRow + 3
    ws.Cells(gRow, 1).Value = "G_total/RT:"
    ws.Cells(gRow, 2).Formula = "=SUM(F" & firstRow & ":F" & lastRow & ")"
    ws.Cells(gRow, 1).Font.Bold = True

    ' Element balance block
    Dim eRow As Long: eRow = lastRow + 5
    ws.Cells(eRow, 1).Value = "Element"
    ws.Cells(eRow, 2).Value = "Target (from inlet)"
    ws.Cells(eRow, 3).Value = "Actual (from n)"
    ws.Range(ws.Cells(eRow, 1), ws.Cells(eRow, 3)).Font.Bold = True

    Dim elems As Variant: elems = Split(ELEMENTS, ",")
    Dim dataCols As Variant: dataCols = Array("T", "U", "V", "W")
    Dim solveCols As Variant: solveCols = Array("G", "H", "I", "J")
    Dim eOff As Long
    For eOff = 0 To N_ELEMENTS - 1
        Dim rr As Long: rr = eRow + 1 + eOff
        ws.Cells(rr, 1).Value = elems(eOff)
        ws.Cells(rr, 2).Formula = "=SUMPRODUCT(" & SHEET_INPUT & "!B8:B" & (7 + N_SPECIES) & _
                                  "," & SHEET_INPUT & "!C8:C" & (7 + N_SPECIES) & _
                                  "," & SHEET_DATA & "!" & dataCols(eOff) & "2:" & dataCols(eOff) & (N_SPECIES + 1) & ")"
        ws.Cells(rr, 3).Formula = "=SUM(" & solveCols(eOff) & firstRow & ":" & solveCols(eOff) & lastRow & ")"
    Next eOff

    ' Named ranges
    Call SetName("n_vars", ws.Range("D" & firstRow & ":D" & lastRow))
    Call SetName("G_total", ws.Cells(gRow, 2))
    Call SetName("nTot_gas", ws.Cells(nTotRow, 2))
    Call SetName("elem_target", ws.Range("B" & (eRow + 1) & ":B" & (eRow + N_ELEMENTS)))
    Call SetName("elem_actual", ws.Range("C" & (eRow + 1) & ":C" & (eRow + N_ELEMENTS)))
    Call SetName("phi_range", ws.Range("K" & firstRow & ":K" & lastRow))

    ws.Columns("A:K").AutoFit
End Sub

Private Function NasaGoverRTFormula(dataRow As Long, firstCoefCol As Long) As String
    ' G/RT = a1(1-lnT) - a2·T/2 - a3·T²/6 - a4·T³/12 - a5·T⁴/20 + a6/T - a7
    Dim c1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String, c7 As String
    c1 = SHEET_DATA & "!" & ColLetter(firstCoefCol) & dataRow
    c2 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 1) & dataRow
    c3 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 2) & dataRow
    c4 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 3) & dataRow
    c5 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 4) & dataRow
    c6 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 5) & dataRow
    c7 = SHEET_DATA & "!" & ColLetter(firstCoefCol + 6) & dataRow
    NasaGoverRTFormula = _
        c1 & "*(1-LN(B1))-" & c2 & "*B1/2-" & c3 & "*B1^2/6-" & _
        c4 & "*B1^3/12-" & c5 & "*B1^4/20+" & c6 & "/B1-" & c7
End Function

Private Function ColLetter(col As Long) As String
    Dim s As String: s = ""
    Dim n As Long: n = col
    Do While n > 0
        Dim k As Long: k = (n - 1) Mod 26
        s = Chr(65 + k) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Sub PrepareSolveSheet()
    Dim wsIn As Worksheet, wsS As Worksheet
    Set wsIn = Worksheets(SHEET_INPUT)
    Set wsS = Worksheets(SHEET_SOLVE)

    Dim totalInlet As Double: totalInlet = 0
    Dim i As Long
    For i = 1 To N_SPECIES
        If wsIn.Cells(i + 7, 3).Value <> 0 Then
            totalInlet = totalInlet + CDbl(wsIn.Cells(i + 7, 2).Value)
        End If
    Next i
    If totalInlet <= 0 Then
        MsgBox "Total inlet moles is 0. Check Input sheet.", vbExclamation
        Err.Raise 999
    End If

    Dim seedMin As Double: seedMin = totalInlet * 0.001
    For i = 1 To N_SPECIES
        Dim enabled As Double: enabled = CDbl(wsIn.Cells(i + 7, 3).Value)
        Dim inletV As Double: inletV = CDbl(wsIn.Cells(i + 7, 2).Value)
        Dim seedV As Double
        If enabled = 0 Then
            seedV = 0.00000000000000000001
        ElseIf inletV > 0 Then
            seedV = inletV
        Else
            seedV = seedMin
        End If
        wsS.Cells(5 + i, 4).Value = seedV
    Next i
End Sub

'=============================================================================
' Phi read / write helpers
'=============================================================================
Private Sub SetAllPhi(val As Double)
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_SOLVE)
    Dim i As Long
    For i = 1 To N_SPECIES
        ws.Cells(5 + i, PHI_COL).Value = val
    Next i
End Sub

Private Sub ReadPhi(ByRef arr() As Double)
    ReDim arr(1 To N_SPECIES)
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_SOLVE)
    Dim i As Long
    For i = 1 To N_SPECIES
        arr(i) = CDbl(ws.Cells(5 + i, PHI_COL).Value)
    Next i
End Sub

'=============================================================================
' Fugacity computation (PR / RK)
'=============================================================================
Private Sub ComputeFugacities(eos As String)
    Dim wsS As Worksheet: Set wsS = Worksheets(SHEET_SOLVE)
    Dim wsD As Worksheet: Set wsD = Worksheets(SHEET_DATA)

    Dim T As Double: T = wsS.Range("B1").Value
    Dim P As Double: P = wsS.Range("B2").Value

    ' Gather per-species data
    Dim n_i() As Double, Tc_i() As Double, Pc_i() As Double, om_i() As Double
    Dim isGas() As Boolean, enabled() As Boolean
    ReDim n_i(1 To N_SPECIES), Tc_i(1 To N_SPECIES), Pc_i(1 To N_SPECIES), om_i(1 To N_SPECIES)
    ReDim isGas(1 To N_SPECIES), enabled(1 To N_SPECIES)

    Dim nTotGas As Double: nTotGas = 0#
    Dim i As Long
    For i = 1 To N_SPECIES
        Dim sr As Long: sr = 5 + i
        Dim dr As Long: dr = 1 + i
        enabled(i) = (CDbl(wsS.Cells(sr, 3).Value) <> 0)
        isGas(i) = (CStr(wsS.Cells(sr, 2).Value) = "gas")
        n_i(i) = CDbl(wsS.Cells(sr, 4).Value)
        Tc_i(i) = CDbl(wsD.Cells(dr, 24).Value)
        Pc_i(i) = CDbl(wsD.Cells(dr, 25).Value)
        om_i(i) = CDbl(wsD.Cells(dr, 26).Value)
        If isGas(i) And enabled(i) Then nTotGas = nTotGas + n_i(i)
    Next i

    ' Mole fractions (gas-enabled only)
    Dim y() As Double: ReDim y(1 To N_SPECIES)
    If nTotGas > 0 Then
        For i = 1 To N_SPECIES
            If isGas(i) And enabled(i) Then y(i) = n_i(i) / nTotGas
        Next i
    End If

    ' Per-species a_i, b_i
    Dim a_sp() As Double, b_sp() As Double
    ReDim a_sp(1 To N_SPECIES), b_sp(1 To N_SPECIES)

    Dim usePR As Boolean: usePR = (eos = "Peng-Robinson")

    For i = 1 To N_SPECIES
        If isGas(i) And enabled(i) And Tc_i(i) > 0 Then
            If usePR Then
                ' PR with Soave α(T)
                Dim kappa As Double
                kappa = 0.37464 + 1.54226 * om_i(i) - 0.26992 * om_i(i) ^ 2
                Dim sqrtTr As Double: sqrtTr = Sqr(T / Tc_i(i))
                Dim alpha_ As Double: alpha_ = (1 + kappa * (1 - sqrtTr)) ^ 2
                a_sp(i) = 0.45724 * (R_GAS * Tc_i(i)) ^ 2 / Pc_i(i) * alpha_
                b_sp(i) = 0.0778 * R_GAS * Tc_i(i) / Pc_i(i)
            Else
                ' Basic RK
                a_sp(i) = 0.42748 * (R_GAS ^ 2) * (Tc_i(i) ^ 2.5) / Pc_i(i)
                b_sp(i) = 0.08664 * R_GAS * Tc_i(i) / Pc_i(i)
            End If
        End If
    Next i

    ' Mixing (k_ij = 0):   a_mix = ΣΣ yᵢyⱼ √(aᵢaⱼ),   b_mix = Σ yᵢ bᵢ
    Dim sqrt_a() As Double: ReDim sqrt_a(1 To N_SPECIES)
    Dim sum_y_sqrtA As Double: sum_y_sqrtA = 0#
    Dim b_mix As Double: b_mix = 0#
    For i = 1 To N_SPECIES
        If isGas(i) And enabled(i) And a_sp(i) > 0 Then
            sqrt_a(i) = Sqr(a_sp(i))
            sum_y_sqrtA = sum_y_sqrtA + y(i) * sqrt_a(i)
            b_mix = b_mix + y(i) * b_sp(i)
        End If
    Next i
    Dim a_mix As Double: a_mix = sum_y_sqrtA ^ 2   ' Quadratic mixing reduces to square

    ' A, B dimensionless
    Dim A_dim As Double, B_dim As Double
    If usePR Then
        A_dim = a_mix * P / (R_GAS * T) ^ 2
    Else
        ' RK: A has T^2.5 in denominator
        A_dim = a_mix * P / ((R_GAS ^ 2) * (T ^ 2.5))
    End If
    B_dim = b_mix * P / (R_GAS * T)

    ' Solve cubic for vapor Z
    Dim Z As Double
    If usePR Then
        Z = SolveCubicVaporZ_PR(A_dim, B_dim)
    Else
        Z = SolveCubicVaporZ_RK(A_dim, B_dim)
    End If

    ' Guard Z > B
    If Z <= B_dim Then Z = B_dim * 1.01

    ' Fugacity per species
    Dim phi_out() As Double: ReDim phi_out(1 To N_SPECIES)
    Dim sqrt2 As Double: sqrt2 = Sqr(2#)

    For i = 1 To N_SPECIES
        If Not (isGas(i) And enabled(i) And a_sp(i) > 0) Then
            phi_out(i) = 1#
        Else
            ' Σ_j yⱼ √(aᵢaⱼ) = √aᵢ · Σ_j yⱼ √aⱼ  (uses quadratic mixing)
            Dim sum_j As Double: sum_j = sqrt_a(i) * sum_y_sqrtA

            Dim term1 As Double, term2 As Double, term3 As Double, term4 As Double
            term1 = (b_sp(i) / b_mix) * (Z - 1)
            term2 = Log(Z - B_dim)

            If usePR Then
                term3 = A_dim / (2# * sqrt2 * B_dim) * (2# * sum_j / a_mix - b_sp(i) / b_mix)
                term4 = Log((Z + (1 + sqrt2) * B_dim) / (Z + (1 - sqrt2) * B_dim))
            Else ' RK
                term3 = (A_dim / B_dim) * (2# * sum_j / a_mix - b_sp(i) / b_mix)
                term4 = Log(1 + B_dim / Z)
            End If

            Dim lnPhi As Double: lnPhi = term1 - term2 - term3 * term4
            If lnPhi > 50 Then lnPhi = 50
            If lnPhi < -50 Then lnPhi = -50
            phi_out(i) = Exp(lnPhi)
        End If
    Next i

    ' Write to Solve sheet K column
    For i = 1 To N_SPECIES
        wsS.Cells(5 + i, PHI_COL).Value = phi_out(i)
    Next i
End Sub

'=============================================================================
' Cubic root solvers (Cardano)
'   Return the vapor-phase root (largest real root > B)
'=============================================================================
Private Function SolveCubicVaporZ_PR(A As Double, B As Double) As Double
    ' PR cubic:  Z³ - (1-B)Z² + (A - 3B² - 2B)Z - (AB - B² - B³) = 0
    Dim p As Double, q As Double, r As Double
    p = -(1# - B)
    q = A - 3# * B ^ 2 - 2# * B
    r = -(A * B - B ^ 2 - B ^ 3)
    SolveCubicVaporZ_PR = SolveCubicMaxReal(p, q, r, B)
End Function

Private Function SolveCubicVaporZ_RK(A As Double, B As Double) As Double
    ' RK cubic:  Z³ - Z² + (A - B - B²)Z - AB = 0
    Dim p As Double, q As Double, r As Double
    p = -1#
    q = A - B - B ^ 2
    r = -A * B
    SolveCubicVaporZ_RK = SolveCubicMaxReal(p, q, r, B)
End Function

Private Function SolveCubicMaxReal(p As Double, q As Double, r As Double, Bparam As Double) As Double
    ' Cardano's method for Z³ + p·Z² + q·Z + r = 0
    ' Depressed substitution: Z = t - p/3 →  t³ + pp·t + qq = 0
    Dim pp As Double, qq As Double
    pp = (3# * q - p ^ 2) / 3#
    qq = (2# * p ^ 3 - 9# * p * q + 27# * r) / 27#

    Dim disc As Double: disc = (qq / 2#) ^ 2 + (pp / 3#) ^ 3

    Dim z1 As Double, z2 As Double, z3 As Double
    Dim nReal As Long

    If disc > 0 Then
        ' One real root, two complex conjugates
        Dim sqD As Double: sqD = Sqr(disc)
        Dim u As Double: u = -qq / 2# + sqD
        Dim v As Double: v = -qq / 2# - sqD
        Dim t As Double
        t = CubeRootSigned(u) + CubeRootSigned(v)
        z1 = t - p / 3#
        nReal = 1
    Else
        ' Three real roots (trigonometric form)
        Dim m As Double: m = 2# * Sqr(-pp / 3#)
        Dim cosArg As Double: cosArg = 3# * qq / (pp * m)
        ' Clamp for numerical safety
        If cosArg > 1# Then cosArg = 1#
        If cosArg < -1# Then cosArg = -1#
        Dim theta As Double: theta = Application.Acos(cosArg) / 3#
        Dim PI As Double: PI = 3.14159265358979
        z1 = m * Cos(theta) - p / 3#
        z2 = m * Cos(theta - 2# * PI / 3#) - p / 3#
        z3 = m * Cos(theta - 4# * PI / 3#) - p / 3#
        nReal = 3
    End If

    ' Pick largest real root above B (vapor root)
    Dim best As Double: best = -1E+30
    If z1 > Bparam And z1 > best Then best = z1
    If nReal = 3 Then
        If z2 > Bparam And z2 > best Then best = z2
        If z3 > Bparam And z3 > best Then best = z3
    End If
    If best < -1E+29 Then
        ' Fallback: pick largest even if below B
        best = z1
        If nReal = 3 Then
            If z2 > best Then best = z2
            If z3 > best Then best = z3
        End If
    End If
    SolveCubicMaxReal = best
End Function

Private Function CubeRootSigned(x As Double) As Double
    If x >= 0# Then
        CubeRootSigned = x ^ (1# / 3#)
    Else
        CubeRootSigned = -((-x) ^ (1# / 3#))
    End If
End Function

'=============================================================================
' Output
'=============================================================================
Private Sub WriteOutputSheet(solRes As Variant, passes As Long, eos As String, lastDeltaPhi As Double)
    Dim wsOut As Worksheet, wsS As Worksheet, wsIn As Worksheet
    Set wsOut = Worksheets(SHEET_OUTPUT)
    Set wsS = Worksheets(SHEET_SOLVE)
    Set wsIn = Worksheets(SHEET_INPUT)

    wsOut.Cells.Clear
    wsOut.Range("A1").Value = "Gibbs Equilibrium — Results (PR / RK)"
    wsOut.Range("A1").Font.Bold = True: wsOut.Range("A1").Font.Size = 13

    wsOut.Range("A3").Value = "Temperature [°C]:": wsOut.Range("B3").Value = wsIn.Range("B3").Value
    wsOut.Range("A4").Value = "Pressure [atm]:":   wsOut.Range("B4").Value = wsIn.Range("B4").Value
    wsOut.Range("A5").Value = "EOS:":              wsOut.Range("B5").Value = eos
    wsOut.Range("A6").Value = "G_total/RT:":       wsOut.Range("B6").Value = wsS.Range("G_total").Value
    wsOut.Range("A7").Value = "Solver status:":    wsOut.Range("B7").Value = SolverStatusText(solRes)
    wsOut.Range("A8").Value = "Picard passes:":    wsOut.Range("B8").Value = passes
    wsOut.Range("A9").Value = "max |Δφ| last:":    wsOut.Range("B9").Value = lastDeltaPhi

    ' Compute Z at final composition for reporting
    Dim Z As Double: Z = ComputeZFromCurrent(eos)
    wsOut.Range("A10").Value = "Z (gas):":         wsOut.Range("B10").Value = Z

    wsOut.Range("A12").Value = "Species"
    wsOut.Range("B12").Value = "Phase"
    wsOut.Range("C12").Value = "Inlet [mol]"
    wsOut.Range("D12").Value = "Outlet [mol]"
    wsOut.Range("E12").Value = "Outlet mol%"
    wsOut.Range("F12").Value = "phi_i"
    wsOut.Range("A12:F12").Font.Bold = True

    Dim nTotOut As Double: nTotOut = 0
    Dim i As Long
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 Then nTotOut = nTotOut + wsS.Cells(5 + i, 4).Value
    Next i

    Dim outRow As Long: outRow = 13
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 Then
            wsOut.Cells(outRow, 1).Value = wsS.Cells(5 + i, 1).Value
            wsOut.Cells(outRow, 2).Value = wsS.Cells(5 + i, 2).Value
            wsOut.Cells(outRow, 3).Value = wsIn.Cells(i + 7, 2).Value
            wsOut.Cells(outRow, 4).Value = wsS.Cells(5 + i, 4).Value
            If nTotOut > 0 Then
                wsOut.Cells(outRow, 5).Value = wsS.Cells(5 + i, 4).Value / nTotOut * 100
            End If
            wsOut.Cells(outRow, 6).Value = wsS.Cells(5 + i, PHI_COL).Value
            outRow = outRow + 1
        End If
    Next i

    wsOut.Range("D13:D" & (outRow - 1)).NumberFormat = "0.000000"
    wsOut.Range("E13:E" & (outRow - 1)).NumberFormat = "0.0000"
    wsOut.Range("F13:F" & (outRow - 1)).NumberFormat = "0.0000"
    wsOut.Columns("A:F").AutoFit
End Sub

Private Function ComputeZFromCurrent(eos As String) As Double
    ' Evaluate Z at the current converged composition for reporting only.
    If eos = "ideal-gas" Then
        ComputeZFromCurrent = 1#
        Exit Function
    End If

    Dim wsS As Worksheet: Set wsS = Worksheets(SHEET_SOLVE)
    Dim wsD As Worksheet: Set wsD = Worksheets(SHEET_DATA)
    Dim T As Double: T = wsS.Range("B1").Value
    Dim P As Double: P = wsS.Range("B2").Value

    Dim nTotGas As Double: nTotGas = 0
    Dim a_mix As Double: a_mix = 0
    Dim b_mix As Double: b_mix = 0
    Dim sum_y_sqrtA As Double: sum_y_sqrtA = 0

    Dim usePR As Boolean: usePR = (eos = "Peng-Robinson")

    Dim i As Long
    ' first pass: nTotGas
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 And wsS.Cells(5 + i, 2).Value = "gas" Then
            nTotGas = nTotGas + CDbl(wsS.Cells(5 + i, 4).Value)
        End If
    Next i
    If nTotGas <= 0 Then
        ComputeZFromCurrent = 1#
        Exit Function
    End If

    ' second pass: a_i, b_i, mixing
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 And wsS.Cells(5 + i, 2).Value = "gas" Then
            Dim Tc As Double: Tc = CDbl(wsD.Cells(1 + i, 24).Value)
            Dim Pc As Double: Pc = CDbl(wsD.Cells(1 + i, 25).Value)
            Dim om As Double: om = CDbl(wsD.Cells(1 + i, 26).Value)
            If Tc > 0 Then
                Dim y As Double: y = CDbl(wsS.Cells(5 + i, 4).Value) / nTotGas
                Dim a_sp As Double, b_sp As Double
                If usePR Then
                    Dim kappa As Double
                    kappa = 0.37464 + 1.54226 * om - 0.26992 * om ^ 2
                    Dim alpha_ As Double
                    alpha_ = (1 + kappa * (1 - Sqr(T / Tc))) ^ 2
                    a_sp = 0.45724 * (R_GAS * Tc) ^ 2 / Pc * alpha_
                    b_sp = 0.0778 * R_GAS * Tc / Pc
                Else
                    a_sp = 0.42748 * (R_GAS ^ 2) * (Tc ^ 2.5) / Pc
                    b_sp = 0.08664 * R_GAS * Tc / Pc
                End If
                sum_y_sqrtA = sum_y_sqrtA + y * Sqr(a_sp)
                b_mix = b_mix + y * b_sp
            End If
        End If
    Next i
    a_mix = sum_y_sqrtA ^ 2

    Dim A_dim As Double, B_dim As Double
    If usePR Then
        A_dim = a_mix * P / (R_GAS * T) ^ 2
    Else
        A_dim = a_mix * P / ((R_GAS ^ 2) * (T ^ 2.5))
    End If
    B_dim = b_mix * P / (R_GAS * T)

    If usePR Then
        ComputeZFromCurrent = SolveCubicVaporZ_PR(A_dim, B_dim)
    Else
        ComputeZFromCurrent = SolveCubicVaporZ_RK(A_dim, B_dim)
    End If
End Function

Private Function SolverStatusText(res As Variant) As String
    Select Case res
        Case 0: SolverStatusText = "Converged (optimum found)"
        Case 1: SolverStatusText = "Converged to current solution"
        Case 2: SolverStatusText = "No improvement possible"
        Case 3: SolverStatusText = "Max iterations reached"
        Case 4: SolverStatusText = "Objective does not converge"
        Case 5: SolverStatusText = "No feasible solution found"
        Case Else: SolverStatusText = "Status code " & res
    End Select
End Function

Private Sub ClearOutputSheet()
    Worksheets(SHEET_OUTPUT).Cells.Clear
    Worksheets(SHEET_OUTPUT).Range("A1").Value = "(Run RunEquilibrium_PR to populate.)"
End Sub

'=============================================================================
' Utility helpers
'=============================================================================
Private Sub EnsureSheet(n As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(n)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = n
    End If
End Sub

Private Sub SetName(nm As String, rng As Range)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add name:=nm, RefersTo:=rng
End Sub
