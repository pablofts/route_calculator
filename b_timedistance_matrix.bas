Attribute VB_Name = "b_timedistance_matrix"
Function td_matrix(points_list) 'builds times/distance matrix

    'YOUR INFO HERE:
    apikey = "AIzaSyCABUvQS3CpgyHURIVnD7mqfXazt3nO16Y"   
    modo = "driving"
    region = "mx"
    

    sz = UBound(points_list) 'number of points in the route
    
    ReDim tdmtx(1 To sz + 1, 1 To sz + 1, 1 To 2)   'matrix with all possible travel _
                                                    times (layer 1) and distances (layer 2) between two points in the route. _
                                                    !!!!Rows as origin; columns as destiny.
    ' tdmtx representation
    '  __ __ __
    ' /  /  /  /|
    '/__/__/__/||
    '|__|O1|O2|||
    '|D1|__|__||/ distance layer
    '|D2|__|__|/ time layer
    
    'D1 same place as O1; D2 same place as O2.
    'Both would be all the places the route is suposed to go by
                                                        
    'fills matrix axes
    For i = 1 To sz
        tdmtx(1, i + 1, 1) = points_list(i, 1)
        tdmtx(i + 1, 1, 1) = points_list(i, 1)
    Next
    
    'calls function fill_tdmtx to fill matrix with times and distances
    tdmtx = fills_tdmtx(tdmtx)
    
    'function td_matrix output
    td_matrix = tdmtx

End Function

Function fills_tdmtx(tdmtx)
    
    sz = UBound(tdmtx) - 1
    For i = 1 To sz 'loops through origin axis (columns)
    
        origen = tdmtx(1, i + 1, 1)
        origen = WorksheetFunction.EncodeURL(origen)
    
        salir = "no"
        ii = 2 'keeps record of the position in the destination axis (rows) as the 15 destinations blocks go by
        While salir = "no"
            'the variables get renewed each while loop exit
            destino = ""
            For ii = ii To 15 + ii 'it keeps row position
                If Not ii = sz + 2 Then 'will be false when it just finished going through the matrix
                    destino = destino & tdmtx(ii, 1, 1) & "|"
                Else
                    salir = "ya"
                    Exit For
                End If
            Next
            
            destino = WorksheetFunction.EncodeURL(destino)
            mtxrequest = gDist(destino, origen, apikey, modo, region) 'calls gDist function (coded by Matthew Moran)
            
            filrec = ii - UBound(mtxrequest) - 1
            If mtxrequest(0, 0) <> "NO DATA" Then
                For iii = 0 To UBound(mtxrequest)
                    tdmtx(i + 1, filrec + iii, 1) = CDec(mtxrequest(iii, 0)) 'time
                    tdmtx(i + 1, filrec + iii, 2) = CDec(mtxrequest(iii, 1)) 'distance
                Next
            End If
        Wend
    Next
    
    'fills_tdmtx output
    fills_tdmtx = tdmtx
    
End Function



