Attribute VB_Name = "e_route_calculate_subprocesses"
Function dispersion(mn)

'finds the exterior rectangle corners (where I live, y axis coordinates are positive -smaller than 999-; x axis coordinates are negative -bigger than 999-
ymin = 999
ymax = 0
xmin = -999
xmax = 0

'finds biggest values
    For i = 1 To 3
        If Not mn(i, 1) = 999 Then
            yy = CVar(Mid(mn(i, 1), 1, 9)) * 1
            xx = CVar(Mid(mn(i, 1), 12, 11)) * 1
            If ymax < yy Then
                ymax = yy
            End If
            If xmax > xx Then
                xmax = xx
            End If
        End If
    Next i
'finds smallest values
    For i = 1 To 3
        If Not mn(i, 1) = 999 Then
            yy = CVar(Mid(mn(i, 1), 1, 9)) * 1
            xx = CVar(Mid(mn(i, 1), 12, 11)) * 1
            If xmin < xx Then
                xmin = xx
            End If
            If ymin > yy Then
                ymin = yy
            End If
        End If
    Next i

'in km, two sides of the exterior rectange -the ones making a corner |_
dispersion = ((xmin - xmax + ymax - ymin) * 100)

End Function

Function maxmin(mn)
    
    'first, finds shortest distance from current origin to point and longest from point to final destination
    min_d_pt = 999
    max_d_dst = 0
    
    'will contain both found distances
    'row 1: closest point to current origin
    'row 2: furthest point to final destination
    Dim arr(1 To 2)
    
    For i = 1 To 3
        If mn(i, 2) < min_d_pt Then 'And Not IsEmpty(mn(i, 1))
            min_d_pt = mn(i, 2)
        End If
        If mn(i, 3) > max_d_dst And mn(i, 3) <> 999 Then
            max_d_dst = mn(i, 3)
        End If
    Next i
    
    'first, it will convert the distances to a fraction of the greatest
    arr(1) = min_d_pt
    arr(2) = max_d_dst
    
    maxmin = arr
    
End Function

Function findsmin(mtx_info, tdmtx)
        Dim mn(1 To 3, 1 To 5) 'this array will contain the three closest points to the current origin
        'column 1: point's coordinates
        'column 2: distance from current org to point
        'column 3: the distance from point to final destination
        'column 4: later, the parameter considered to chose the next destination
        'column 5: the time from the current point to the final destination -will be used just to keep track of the time implied in the route
        
        sz = UBound(mtx_info)
        
        'starts mn array
        For i = 1 To 3
            mn(i, 1) = 999
            mn(i, 2) = 999
            mn(i, 3) = 999
            mn(i, 4) = 999
            mn(i, 5) = 999
        Next
        
        'finds the three closest points to the current origin
        'picks each point and determines if it should go in first, second, third or out the mn array
        For i = 1 To 3
            If i = 1 Then
                For ii = 1 To sz
                    If mtx_info(ii, 2) < mn(i, 2) And mtx_info(ii, 1) <> 999 Then  'if the distance from org to point is the smaller than the smallest in mn
                        mn(i, 2) = mtx_info(ii, 2) '|
                        mn(i, 1) = mtx_info(ii, 1) '|- fills mn with all the info of mtx_info for pont
                        mn(i, 3) = mtx_info(ii, 3) '|
                        mn(i, 5) = ii              '|
                    End If
                Next
            Else
                For ii = 1 To sz
                    If mtx_info(ii, 2) > mn(i - 1, 2) And mtx_info(ii, 2) < mn(i, 2) And mtx_info(ii, 1) <> 999 Then 'if the distance from org to point is the smaller than the second or third in mn
                        mn(i, 2) = mtx_info(ii, 2) '|
                        mn(i, 1) = mtx_info(ii, 1) '|- fills mn with all the info of mtx_info for point
                        mn(i, 3) = mtx_info(ii, 3) '|
                        mn(i, 5) = ii              '|
                    End If
                Next
            End If
        Next
        findsmin = mn 'output
End Function

Function removes(listmtx, mtx_final)
    sz = UBound(listmtx)
    'loops through mtx final and removes already taken points -puts 999-
    For i = 1 To sz
        For ii = 1 To sz
            If listmtx(i, 1) = mtx_final(ii, 1) Then
                listmtx(i, 1) = 999
            End If
        Next
    Next
    removes = listmtx
End Function

Function fillsinfo(mtx_info, listmtx, tdmtx, j, orgL, fdestL)
    sz = UBound(listmtx)
    'puts all the needed info in mtx_info
    'col 1: coordinates
    'col 2: distance from current origin to point
    'col 3: distance from point to final destination
    
    'only in the first cicle, the names of the points
    If j = 1 Then
        For i = 1 To sz
            mtx_info(i, 1) = listmtx(i, 1)
        Next
    End If
    
    'fills distance info
    For i = 1 To sz
        If listmtx(i, 1) = 999 Then
            mtx_info(i, 2) = 9999
            mtx_info(i, 3) = 0
        Else
            mtx_info(i, 2) = tdmtx(i + 1, orgL + 1, 2) 'distance from current origin to point
            mtx_info(i, 3) = tdmtx(fdestL + 1, i + 1, 2) 'distance from point to final destination
        End If
    Next
    fillsinfo = mtx_info 'output
End Function

