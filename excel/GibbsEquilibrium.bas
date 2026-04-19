Attribute VB_Name = "GibbsEquilibrium"
'=============================================================================
' Gibbs Equilibrium Calculator for Excel (VBA + Solver)
'
' Minimizes total Gibbs free energy G = Σ n_i · μ_i subject to element mass
' balance Σ a_ij · n_i = b_j using Excel Solver (GRG Nonlinear).
'
' μ_i/RT = g0_i(T)/RT + ln(y_i · P/P°)   (ideal gas)
' μ_i/RT = g0_i(T)/RT                    (solid, activity = 1)
'
' Data sources (all public):
'   - GRI-Mech 3.0 (gri30.yaml): H2, CO, CO2, H2O, CH4, CH3OH, CH3CHO, HCHO,
'     C2H6, C2H4, N2
'   - NASA Glenn (nasa_gas.yaml, McBride et al. NASA/TP-2002-211556):
'     CH3OCH3 (L12/92), CH3COOH, C2H5OH, HCOOH (L 8/88)
'   - Burcat T10/07 "Meacetate": CH3COOCH3
'   - JANAF X 4/83 (nasa_condensed.yaml): C(gr)
'
' Usage:
'   1. In VBE, Tools -> References -> check "Solver"
'      (or keep Solver add-in loaded; late-binding via Application.Run also OK)
'   2. Run macro SetupWorkbook  -> creates sheets Data/Input/Solve/Output
'   3. Edit Input sheet: T [°C], P [atm], inlet moles per species
'   4. Run macro RunEquilibrium -> Solver finds equilibrium, writes Output
'=============================================================================
Option Explicit

Private Const R_GAS As Double = 8.31446261815324  ' J/(mol·K)
Private Const P_REF As Double = 101325#            ' Pa (1 atm)

Private Const SHEET_DATA As String = "Data"
Private Const SHEET_INPUT As String = "Input"
Private Const SHEET_SOLVE As String = "Solve"
Private Const SHEET_OUTPUT As String = "Output"

Private Const N_SPECIES As Long = 17
Private Const ELEMENTS As String = "C,H,O,N"  ' fixed element order
Private Const N_ELEMENTS As Long = 4

'=============================================================================
' Public entry points
'=============================================================================

Public Sub SetupWorkbook()
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
    MsgBox "Setup complete." & vbCrLf & _
           "1) Edit the Input sheet (T, P, inlet moles)." & vbCrLf & _
           "2) Run macro RunEquilibrium.", vbInformation, "Gibbs Equilibrium"
    Exit Sub

fail:
    Application.ScreenUpdating = True
    MsgBox "SetupWorkbook failed: " & Err.Description, vbCritical
End Sub

Public Sub RunEquilibrium()
    Application.ScreenUpdating = False
    On Error GoTo fail

    Dim wsInput As Worksheet, wsSolve As Worksheet, wsOut As Worksheet
    Set wsInput = Worksheets(SHEET_INPUT)
    Set wsSolve = Worksheets(SHEET_SOLVE)
    Set wsOut = Worksheets(SHEET_OUTPUT)

    ' Copy inputs → Solve sheet and seed initial n_i
    Call PrepareSolveSheet

    ' Invoke Solver via late binding (works whether or not reference is set)
    Dim setCellAddr As String, byChangeAddr As String
    Dim lbAddr As String, elemActAddr As String, elemTgtAddr As String
    setCellAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("G_total").Address
    byChangeAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("n_vars").Address
    lbAddr = byChangeAddr
    elemActAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("elem_actual").Address
    elemTgtAddr = "'" & SHEET_SOLVE & "'!" & wsSolve.Range("elem_target").Address

    Application.Run "SolverReset"
    Application.Run "SolverOk", setCellAddr, 2, 0, byChangeAddr, 1  ' MinVal=2, Engine 1=GRG
    Application.Run "SolverAdd", lbAddr, 3, "1E-20"                 ' n_i >= 1e-20
    Application.Run "SolverAdd", elemActAddr, 2, elemTgtAddr        ' element balance
    Application.Run "SolverOptions", 120, 2000, 0.0000001           ' MaxTime, Iterations, Precision
    Dim solRes As Variant
    solRes = Application.Run("SolverSolve", True)                   ' UserFinish=True
    Application.Run "SolverFinish", 1                                ' KeepFinal=1

    Call WriteOutputSheet(solRes)
    Worksheets(SHEET_OUTPUT).Activate
    Application.ScreenUpdating = True
    Exit Sub

fail:
    Application.ScreenUpdating = True
    MsgBox "RunEquilibrium failed: " & Err.Description & vbCrLf & _
           "Check that the Solver add-in is loaded " & _
           "(File → Options → Add-ins → Solver).", vbCritical
End Sub

'=============================================================================
' Sheet construction
'=============================================================================

Private Sub WriteDataSheet()
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_DATA)
    ws.Cells.Clear

    ' Header
    Dim hdr As Variant
    hdr = Array("Species", "Phase", "Tlow", "Tmid", "Thigh", _
                "a1_lo", "a2_lo", "a3_lo", "a4_lo", "a5_lo", "a6_lo", "a7_lo", _
                "a1_hi", "a2_hi", "a3_hi", "a4_hi", "a5_hi", "a6_hi", "a7_hi", _
                "C", "H", "O", "N", "Source")
    Dim j As Long
    For j = 0 To UBound(hdr)
        ws.Cells(1, j + 1).Value = hdr(j)
    Next j
    ws.Rows(1).Font.Bold = True

    Dim r As Long: r = 2
    ' -- GRI-Mech 3.0 (gri30.yaml) species --------------------------------
    Call WriteSpeciesRow(ws, r, "H2", "gas", 200#, 1000#, 3500#, _
        Array(2.34433112, 0.00798052075, -1.9478151E-05, 2.01572094E-08, -7.37611761E-12, -917.935173, 0.683010238), _
        Array(3.3372792, -4.94024731E-05, 4.99456778E-07, -1.79566394E-10, 2.00255376E-14, -950.158922, -3.20502331), _
        0, 2, 0, 0, "gri30.yaml (GRI-Mech 3.0)")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CO2", "gas", 200#, 1000#, 3500#, _
        Array(2.35677352, 0.00898459677, -7.12356269E-06, 2.45919022E-09, -1.43699548E-13, -48371.9697, 9.90105222), _
        Array(3.85746029, 0.00441437026, -2.21481404E-06, 5.23490188E-10, -4.72084164E-14, -48759.166, 2.27163806), _
        1, 0, 2, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CO", "gas", 200#, 1000#, 3500#, _
        Array(3.57953347, -0.00061035368, 1.01681433E-06, 9.07005884E-10, -9.04424499E-13, -14344.086, 3.50840928), _
        Array(2.71518561, 0.00206252743, -9.98825771E-07, 2.30053008E-10, -2.03647716E-14, -14151.8724, 7.81868772), _
        1, 0, 1, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "H2O", "gas", 200#, 1000#, 3500#, _
        Array(4.19864056, -0.0020364341, 6.52040211E-06, -5.48797062E-09, 1.77197817E-12, -30293.7267, -0.849032208), _
        Array(3.03399249, 0.00217691804, -1.64072518E-07, -9.7041987E-11, 1.68200992E-14, -30004.2971, 4.9667701), _
        0, 2, 1, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH4", "gas", 200#, 1000#, 3500#, _
        Array(5.14987613, -0.0136709788, 4.91800599E-05, -4.84743026E-08, 1.66693956E-11, -10246.6476, -4.64130376), _
        Array(0.074851495, 0.0133909467, -5.73285809E-06, 1.22292535E-09, -1.0181523E-13, -9468.34459, 18.437318), _
        1, 4, 0, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3OH", "gas", 200#, 1000#, 3500#, _
        Array(5.71539582, -0.0152309129, 6.52441155E-05, -7.10806889E-08, 2.61352698E-11, -25642.7656, -1.50409823), _
        Array(1.78970791, 0.0140938292, -6.36500835E-06, 1.38171085E-09, -1.1706022E-13, -25374.8747, 14.5023623), _
        1, 4, 1, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3CHO", "gas", 200#, 1000#, 6000#, _
        Array(4.7294595, -0.0031932858, 4.7534921E-05, -5.7458611E-08, 2.1931112E-11, -21572.878, 4.1030159), _
        Array(5.4041108, 0.011723059, -4.2263137E-06, 6.8372451E-10, -4.0984863E-14, -22593.122, -3.4807917), _
        2, 4, 1, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "HCHO", "gas", 200#, 1000#, 3500#, _
        Array(4.79372315, -0.00990833369, 3.73220008E-05, -3.79285261E-08, 1.31772652E-11, -14308.9567, 0.6028129), _
        Array(1.76069008, 0.00920000082, -4.42258813E-06, 1.00641212E-09, -8.8385564E-14, -13995.8323, 13.656323), _
        1, 2, 1, 0, "gri30.yaml (CH2O)")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H6", "gas", 200#, 1000#, 3500#, _
        Array(4.29142492, -0.0055015427, 5.99438288E-05, -7.08466285E-08, 2.68685771E-11, -11522.2055, 2.66682316), _
        Array(1.0718815, 0.0216852677, -1.00256067E-05, 2.21412001E-09, -1.9000289E-13, -11426.3932, 15.1156107), _
        2, 6, 0, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H4", "gas", 200#, 1000#, 3500#, _
        Array(3.95920148, -0.00757052247, 5.70990292E-05, -6.91588753E-08, 2.69884373E-11, 5089.77593, 4.09733096), _
        Array(2.03611116, 0.0146454151, -6.71077915E-06, 1.47222923E-09, -1.25706061E-13, 4939.88614, 10.3053693), _
        2, 4, 0, 0, "gri30.yaml")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "N2", "gas", 300#, 1000#, 5000#, _
        Array(3.298677, 0.0014082404, -3.963222E-06, 5.641515E-09, -2.444854E-12, -1020.8999, 3.950372), _
        Array(2.92664, 0.0014879768, -5.68476E-07, 1.0097038E-10, -6.753351E-15, -922.7977, 5.980528), _
        0, 0, 0, 2, "gri30.yaml")
    r = r + 1

    ' -- NASA Glenn (nasa_gas.yaml) species --------------------------------
    Call WriteSpeciesRow(ws, r, "CH3OCH3", "gas", 200#, 1000#, 6000#, _
        Array(5.30562279, -0.00214254272, 5.30873244E-05, -6.23147136E-08, 2.30731036E-11, -23986.6295, 0.713264209), _
        Array(5.64844183, 0.0163381899, -5.86802367E-06, 9.46836869E-10, -5.66504738E-14, -25107.469, -5.96264939), _
        2, 6, 1, 0, "nasa_gas.yaml L12/92")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "CH3COOH", "gas", 200#, 1000#, 6000#, _
        Array(2.78936844, 0.0100001016, 3.42557978E-05, -5.09017919E-08, 2.06217504E-11, -53475.2292, 14.1059504), _
        Array(7.67083678, 0.0135152695, -5.25874688E-06, 8.93185062E-10, -5.53180891E-14, -55756.0971, -15.467659), _
        2, 4, 2, 0, "nasa_gas.yaml L 8/88")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "C2H5OH", "gas", 200#, 1000#, 6000#, _
        Array(4.85868178, -0.0037400674, 6.95550267E-05, -8.86541147E-08, 3.5168443E-11, -29996.1309, 4.80192294), _
        Array(6.5628977, 0.0152034264, -5.38922247E-06, 8.62150224E-10, -5.12824683E-14, -31525.7984, -9.47557644), _
        2, 6, 1, 0, "nasa_gas.yaml L 8/88")
    r = r + 1
    Call WriteSpeciesRow(ws, r, "HCOOH", "gas", 200#, 1000#, 6000#, _
        Array(3.23262453, 0.00281129582, 2.44034975E-05, -3.17501066E-08, 1.2063166E-11, -46778.5606, 9.86205647), _
        Array(5.69579404, 0.00772237361, -3.18037808E-06, 5.57949466E-10, -3.52618226E-14, -48159.9723, -6.0168008), _
        1, 2, 2, 0, "nasa_gas.yaml L 8/88")
    r = r + 1

    ' -- Burcat T10/07 (not in Cantera YAMLs) ------------------------------
    Call WriteSpeciesRow(ws, r, "CH3COOCH3", "gas", 200#, 1000#, 6000#, _
        Array(7.18744749, -0.00629221513, 8.17059377E-05, -9.82940778E-08, 3.73744521E-11, -52341.7155, -3.24161798), _
        Array(8.38776809, 0.0190836514, -6.8219732E-06, 1.09765423E-09, -6.55561842E-14, -54080.5971, -16.4156253), _
        3, 6, 2, 0, "Burcat T10/07 Meacetate")
    r = r + 1

    ' -- NASA Glenn condensed (graphite.yaml) solid ------------------------
    Call WriteSpeciesRow(ws, r, "C(gr)", "solid", 200#, 1000#, 5000#, _
        Array(-0.310872072, 0.00440353686, 1.90394118E-06, -6.38546966E-09, 2.98964248E-12, -108.650794, 1.11382953), _
        Array(1.45571829, 0.00171702216, -6.97562786E-07, 1.35277032E-10, -9.67590652E-15, -695.138814, -8.52583033), _
        1, 0, 0, 0, "nasa_condensed.yaml X 4/83")

    ws.Columns("A:X").AutoFit
End Sub

Private Sub WriteSpeciesRow(ws As Worksheet, r As Long, name As String, phase As String, _
                            Tl As Double, Tm As Double, Th As Double, _
                            aLo As Variant, aHi As Variant, _
                            nC As Long, nH As Long, nO As Long, nN As Long, _
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
    ws.Cells(r, 24).Value = src
End Sub

Private Sub WriteInputSheet()
    Dim ws As Worksheet: Set ws = Worksheets(SHEET_INPUT)
    ws.Cells.Clear

    ws.Range("A1").Value = "Gibbs Equilibrium — Input"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13

    ws.Range("A3").Value = "Temperature [°C]:"
    ws.Range("B3").Value = 250
    ws.Range("A4").Value = "Pressure [atm]:"
    ws.Range("B4").Value = 50

    ws.Range("A6").Value = "Species"
    ws.Range("B6").Value = "Inlet moles"
    ws.Range("C6").Value = "Enabled (1/0)"
    ws.Range("D6").Value = "Phase"
    ws.Range("A6:D6").Font.Bold = True

    Dim r As Long, i As Long
    r = 7
    For i = 2 To N_SPECIES + 1
        ws.Cells(r, 1).Value = Worksheets(SHEET_DATA).Cells(i, 1).Value
        ws.Cells(r, 2).Value = 0
        ws.Cells(r, 3).Value = 1
        ws.Cells(r, 4).Value = Worksheets(SHEET_DATA).Cells(i, 2).Value
        r = r + 1
    Next i

    ' Default preset: DME synthesis H2:CO = 2:1
    Worksheets(SHEET_INPUT).Range("B7").Value = 2   ' H2
    Worksheets(SHEET_INPUT).Range("B9").Value = 1   ' CO

    ws.Columns("A:D").AutoFit
    ws.Range("B3:B4").Interior.Color = RGB(255, 255, 204)
    ws.Range("B7:C" & (6 + N_SPECIES)).Interior.Color = RGB(255, 255, 204)
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
    ws.Range("A5:J5").Font.Bold = True

    Dim firstRow As Long, lastRow As Long
    firstRow = 6
    lastRow = firstRow + N_SPECIES - 1

    Dim i As Long, sr As Long
    For i = 1 To N_SPECIES
        sr = firstRow + i - 1
        Dim dataRow As Long: dataRow = i + 1

        ws.Cells(sr, 1).Formula = "=" & SHEET_DATA & "!A" & dataRow
        ws.Cells(sr, 2).Formula = "=" & SHEET_DATA & "!B" & dataRow
        ws.Cells(sr, 3).Formula = "=" & SHEET_INPUT & "!C" & (i + 6)  ' enabled
        ws.Cells(sr, 4).Value = 0.000000001                          ' n_i seed (replaced later)

        ' g0/RT (NASA7): picks low/high based on T vs Tmid, returns 1e30 if disabled
        ws.Cells(sr, 5).Formula = _
            "=IF(C" & sr & "=0, 1E30, " & _
            "IF(B1<=" & SHEET_DATA & "!D" & dataRow & ", " & _
                NasaGoverRTFormula(dataRow, 6) & ", " & _
                NasaGoverRTFormula(dataRow, 13) & "))"

        ' n·μ/RT: gas uses ln(y·P/Pref); solid is just n·g0/RT
        ws.Cells(sr, 6).Formula = _
            "=IF(C" & sr & "=0, 0, " & _
            "IF(B" & sr & "=""gas"", " & _
                "D" & sr & "*(E" & sr & "+LN(MAX(D" & sr & "/nTot_gas,1E-300))+B3), " & _
                "D" & sr & "*E" & sr & "))"

        ' Element contributions (n · a_jk · enabled). Disabled species are excluded.
        ws.Cells(sr, 7).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!T" & dataRow   ' C
        ws.Cells(sr, 8).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!U" & dataRow   ' H
        ws.Cells(sr, 9).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!V" & dataRow   ' O
        ws.Cells(sr, 10).Formula = "=C" & sr & "*D" & sr & "*" & SHEET_DATA & "!W" & dataRow  ' N
    Next i

    ' Gas total (enabled gas species only)
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
    Dim dataCols As Variant: dataCols = Array("T", "U", "V", "W")    ' Data sheet comp cols
    Dim solveCols As Variant: solveCols = Array("G", "H", "I", "J")   ' Solve n·a cols
    Dim eOff As Long
    For eOff = 0 To N_ELEMENTS - 1
        Dim rr As Long: rr = eRow + 1 + eOff
        ws.Cells(rr, 1).Value = elems(eOff)
        ws.Cells(rr, 2).Formula = "=SUMPRODUCT(" & SHEET_INPUT & "!B7:B" & (6 + N_SPECIES) & _
                                  "," & SHEET_INPUT & "!C7:C" & (6 + N_SPECIES) & _
                                  "," & SHEET_DATA & "!" & dataCols(eOff) & "2:" & dataCols(eOff) & (N_SPECIES + 1) & ")"
        ws.Cells(rr, 3).Formula = "=SUM(" & solveCols(eOff) & firstRow & ":" & solveCols(eOff) & lastRow & ")"
    Next eOff

    ' Named ranges for Solver
    Call SetName("n_vars", ws.Range("D" & firstRow & ":D" & lastRow))
    Call SetName("G_total", ws.Cells(gRow, 2))
    Call SetName("nTot_gas", ws.Cells(nTotRow, 2))
    Call SetName("elem_target", ws.Range("B" & (eRow + 1) & ":B" & (eRow + N_ELEMENTS)))
    Call SetName("elem_actual", ws.Range("C" & (eRow + 1) & ":C" & (eRow + N_ELEMENTS)))

    ws.Columns("A:J").AutoFit
End Sub

Private Function NasaGoverRTFormula(dataRow As Long, firstCoefCol As Long) As String
    ' Returns NASA7 G/RT formula string for one temperature range.
    ' G/RT = a1(1-ln T) - a2·T/2 - a3·T²/6 - a4·T³/12 - a5·T⁴/20 + a6/T - a7
    ' coefficient columns start at `firstCoefCol` (6 for low, 13 for high) on Data sheet.
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
    ' Seed n_i with inlet moles (+ small epsilon for species with zero inlet but enabled)
    Dim wsIn As Worksheet, wsS As Worksheet
    Set wsIn = Worksheets(SHEET_INPUT)
    Set wsS = Worksheets(SHEET_SOLVE)

    Dim totalInlet As Double: totalInlet = 0
    Dim i As Long
    For i = 1 To N_SPECIES
        If wsIn.Cells(i + 6, 3).Value <> 0 Then
            totalInlet = totalInlet + CDbl(wsIn.Cells(i + 6, 2).Value)
        End If
    Next i
    If totalInlet <= 0 Then
        MsgBox "Total inlet moles is 0. Check Input sheet.", vbExclamation
        Err.Raise 999
    End If

    Dim seedMin As Double: seedMin = totalInlet * 0.001
    For i = 1 To N_SPECIES
        Dim enabled As Double: enabled = CDbl(wsIn.Cells(i + 6, 3).Value)
        Dim inletV As Double: inletV = CDbl(wsIn.Cells(i + 6, 2).Value)
        Dim seedV As Double
        If enabled = 0 Then
            seedV = 0.00000000000000000001  ' 1e-20, same as Solver lower bound
        ElseIf inletV > 0 Then
            seedV = inletV
        Else
            seedV = seedMin
        End If
        wsS.Cells(5 + i, 4).Value = seedV
    Next i
End Sub

Private Sub WriteOutputSheet(solRes As Variant)
    Dim wsOut As Worksheet, wsS As Worksheet, wsIn As Worksheet
    Set wsOut = Worksheets(SHEET_OUTPUT)
    Set wsS = Worksheets(SHEET_SOLVE)
    Set wsIn = Worksheets(SHEET_INPUT)

    wsOut.Cells.Clear

    wsOut.Range("A1").Value = "Gibbs Equilibrium — Results"
    wsOut.Range("A1").Font.Bold = True: wsOut.Range("A1").Font.Size = 13

    wsOut.Range("A3").Value = "Temperature [°C]:"
    wsOut.Range("B3").Value = wsIn.Range("B3").Value
    wsOut.Range("A4").Value = "Pressure [atm]:"
    wsOut.Range("B4").Value = wsIn.Range("B4").Value
    wsOut.Range("A5").Value = "G_total/RT:"
    wsOut.Range("B5").Value = wsS.Range("G_total").Value
    wsOut.Range("A6").Value = "Solver status:"
    wsOut.Range("B6").Value = SolverStatusText(solRes)

    wsOut.Range("A8").Value = "Species"
    wsOut.Range("B8").Value = "Phase"
    wsOut.Range("C8").Value = "Inlet [mol]"
    wsOut.Range("D8").Value = "Outlet [mol]"
    wsOut.Range("E8").Value = "Outlet mol%"
    wsOut.Range("A8:E8").Font.Bold = True

    Dim nTotOut As Double: nTotOut = 0
    Dim i As Long
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 Then nTotOut = nTotOut + wsS.Cells(5 + i, 4).Value
    Next i

    Dim outRow As Long: outRow = 9
    For i = 1 To N_SPECIES
        If wsS.Cells(5 + i, 3).Value <> 0 Then
            wsOut.Cells(outRow, 1).Value = wsS.Cells(5 + i, 1).Value
            wsOut.Cells(outRow, 2).Value = wsS.Cells(5 + i, 2).Value
            wsOut.Cells(outRow, 3).Value = wsIn.Cells(i + 6, 2).Value
            wsOut.Cells(outRow, 4).Value = wsS.Cells(5 + i, 4).Value
            If nTotOut > 0 Then
                wsOut.Cells(outRow, 5).Value = wsS.Cells(5 + i, 4).Value / nTotOut * 100
            End If
            outRow = outRow + 1
        End If
    Next i

    wsOut.Range("D" & 9 & ":D" & (outRow - 1)).NumberFormat = "0.000000"
    wsOut.Range("E" & 9 & ":E" & (outRow - 1)).NumberFormat = "0.0000"
    wsOut.Columns("A:E").AutoFit
End Sub

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
    Worksheets(SHEET_OUTPUT).Range("A1").Value = "(Run RunEquilibrium to populate.)"
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
