Attribute VB_Name = "d_route_calculate"
Function rut_calculate(listmtx, tdmtx, org, fdest)

    Application.ScreenUpdating = False
    
        'the heuristic starts in the specified origin and picks a place -point- to go next; it choses as next the one _
    that is both, closest to the current origin and furthes to the final destination, between the three closest _
    points to the origin. The chosed point -previously the next inmediate (partial) destination- then becomes the _
    origin, and the loops repets over and over until it goes through every point
    
        'listmtx contains the list of places -points- the route will go by
        'tdmtx is a three demension array that contains all de travel times and distances between the points in the route
        'org is the firts origin; this variable will be recicled as the patrial destinations -points in the route- become _
    the current orrigin
    
    'AMO A ANA PAU
    
    'arrays
    sz = UBound(listmtx) 'number of points the route goes by
    ReDim mtx_info(1 To sz, 1 To 3) 'this array will hold the needed info to perform the heuristic -for each cicle-
        'col 1: coordinates
        'col 2: distance from current origin to point
        'col 3: distance from point to final destination
    ReDim mtx_final(1 To sz, 1 To 7) 'this array will hold the final version of the route and other important info
        'col 1: coordinates
        'col 2: point position in listmtx -to perform a lookup in of times and distances tdmtx once the route order is determined
        'the next columns will only be filled after the route is calculated:
        'col 3: distance to next point
        'col 4: time to next point
        'col 5: total route distance, just in first row
        'col 6: total route time, just in first row
        'col 7: url, just in first row
    
    'final destination Loction in tdmtx
    For i = 1 To sz
        If fdest = listmtx(i, 1) Then
            fdestL = i
        End If
        If org = listmtx(i, 1) Then
            orgL = i
        End If
    Next
    
    'Puts origin in top of the array, puts destination at the end of the array
    mtx_final(1, 1) = org 'punto de partida def por usuario
    mtx_final(1, 2) = orgL 'org position in tdmtx
    mtx_final(sz, 1) = fdest 'dest f def por usuario
    mtx_final(sz, 2) = fdestL 'fdest position in tdmtx
    
    'for each cicle, a point is placed in its final order in mtx_final
    For j = 1 To sz - 2
        'current origin Loction in tdmtx
        For i = 1 To sz
            If org = listmtx(i, 1) Then
                orgL = i
                Exit For
            End If
        Next
        listmtx = removes(listmtx, mtx_final) 'loops through mtx final and removes already taken points -puts 999-
        mtx_info = fillsinfo(mtx_info, listmtx, tdmtx, j, orgL, fdestL) 'loops through mtx_info and puts nedded info for the heuristic
            'col 1: coordinates
            'col 2: distance from current origin to point
            'col 3: distance from point to final destination
        mn = findsmin(mtx_info, tdmtx) 'finds the three cloasest points to the current origin; it will chose from those thre
        disp = dispersion(mn) 'in km, two sides of the exterior rectange, as a measure of how far a part points are from each other, in relation to the final destnation or current origin
        mxmn = maxmin(mn) 'longes and closest distance from the three
            min_d_pt = mxmn(1)  'min distance to point from origin -closest to current origin-
            max_d_dst = mxmn(2) 'max distance from point to final destination -furthest to final destination-
        
        'builds the decision parameter
        For i = 1 To 3
            If Not disp = 0 And Not mn(i, 3) = 999 Then 'if any of those, then only one point remaining, then no adjustment would be needed
                'pt_ot = mn(i, 2) / disp 'used as a power, will adjust current org to point distances according to distances/dispersion ratio
                'pt_df = mn(i, 3) / disp 'used as a power, will adjust point to fdest distances according to distances/dispersion ratio
                pr = ((min_d_pt + max_d_dst) / 2) / disp 'nuevo par
                'if smaller than one, the distances/dispersion ratio is not such that an adjustment needs to be done
                If pt_df < 1 Then
                    pt_df = 1
                End If
                If pt_ot < 1 Then
                    pt_ot = 1
                End If
            Else 'no adjustment needed
                pt_df = 1
                pt_ot = 1
            End If
            If mn(i, 2) <> 999 Or mn(i, 3) <> 999 Then
                mn(i, 2) = mn(i, 2) / min_d_pt 'each of the three distances from current origin to ponit divided by the smallest of them
                mn(i, 3) = mn(i, 3) / max_d_dst 'each of the three distances from ponit to the final dest divided by the biggest of them
                mn(i, 4) = (mn(i, 2) ^ pr) - (mn(i, 3) ^ pr) '(mn(i, 2) ^ pt_ot) - (mn(i, 3) ^ pt_df) 'each of the new "unitary" (?) distances raised to the adjuster value
            End If
        Next i
        
        'choses next partial destination if it has the smallest parameter value -later, in the next cicle, will be the current origin
        min_prom = 999
        For i = 1 To 3
            If min_prom > mn(i, 4) And mn(i, 1) <> 999 Then
                min_prom = mn(i, 4)
                nxt_dest = i
            End If
        Next
        
        mtx_final(j + 1, 1) = mn(nxt_dest, 1) 'coordinates of chosed next point
        mtx_final(j + 1, 2) = mn(nxt_dest, 5) 'point position in tdmtx
        
        org = mn(nxt_dest, 1) 'now, current origin
    Next
    
    'fills pending values in mtx_final
    mtx_final = URLtdLookup(mtx_final, tdmtx)
    
    'function output
    rut_calculate = mtx_final
    
    Application.ScreenUpdating = True
End Function

