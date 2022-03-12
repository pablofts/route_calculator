Attribute VB_Name = "f_output_subprocesses"
Function URLtdLookup(mtx_final, tdmtx)
    sz = UBound(mtx_final)
    
    'fills route time and distance info
    For i = 1 To sz - 1
        s1 = mtx_final(i, 2) + 1 'position of the current point in tdmtx
        s2 = mtx_final(i + 1, 2) + 1 'position of the next point in tdmtx
        mtx_final(i, 3) = tdmtx(s2, s1, 2) 'distance to next point
        dtot = dtot + tdmtx(s2, s1, 2) 'cumulative distance for total route distance
        mtx_final(i, 4) = tdmtx(s2, s1, 1) 'time to next point
        ttot = ttot + tdmtx(s2, s1, 1) 'cumulative time for total route time
    Next
    
    mtx_final(1, 5) = dtot 'total route distance
    mtx_final(1, 6) = ttot 'total route time
    
    'builds google maps url
    url = "https://www.google.com.mx/maps/dir"
    For i = 1 To UBound(mtx_final)
        url = url & "/" & mtx_final(i, 1)
    Next
    
    mtx_final(1, 7) = url 'google maps url
    
    URLtdLookup = mtx_final
End Function
