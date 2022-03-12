Attribute VB_Name = "a_input"
Sub input_points()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'DEBE HACERLOS UNICOS
    'INPUT
'    listmtx = Hoja1.Range("a1:a15").Value 'matrix with the points the route will go by
'    org = Hoja1.Range("a1").Value 'origin
'    fdest = Hoja1.Range("a2").Value 'final destination

    Dim listmtx(1 To 15, 1 To 1)
    listmtx(1, 1) = "20.689513, -103.417709"
    listmtx(2, 1) = "20.727851, -103.517074"
    listmtx(3, 1) = "20.715408, -103.521108"
    listmtx(4, 1) = "20.721036, -103.477931"
    listmtx(5, 1) = "20.722844, -103.528213"
    listmtx(6, 1) = "20.721186, -103.478820"
    listmtx(7, 1) = "20.720940, -103.477582"
    listmtx(8, 1) = "20.718016, -103.474090"
    listmtx(9, 1) = "20.726329, -103.526101"
    listmtx(10, 1) = "20.723788, -103.442798"
    listmtx(11, 1) = "20.717920, -103.528979"
    listmtx(12, 1) = "20.721865, -103.525193"
    listmtx(13, 1) = "20.726526, -103.528367"
    listmtx(14, 1) = "20.721736, -103.528383"
    listmtx(15, 1) = "20.711300, -103.447315"
    
    org = listmtx(1, 1) 'origin
    fdest = listmtx(2, 1) 'final destination
    
    'PROCESS
    tdmtx = td_matrix(listmtx)
    mtx_final = rut_calculate(listmtx, tdmtx, org, fdest)
    
    'PRINT IN SHEET
    Set outp = Range("b1") 'i will spill the resutls matrix in the array
    For i = 1 To UBound(mtx_final, 2)
        For ii = 1 To UBound(mtx_final)
            outp.Offset(ii - 1, i - 1).Value = mtx_final(ii, i)
        Next
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
