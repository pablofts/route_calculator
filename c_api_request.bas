Attribute VB_Name = "c_api_request"
Function gDist(strOrig, strDest, strAPI, strMode, strReg) 'makes the api request; matrix with time and distance between points as output

'Following code author: Matthew Moran; it was made public by himself. I made mild modifications _
Available at https://pulseinfomatics.com/new-use-vba-to-retrieve-distances-between-multiple-addresses-in-excel/
    
' because you can run multiple destinations (you can also run multiple origins), I return an array.
' in this example, I limit it to 10 but I believe Google allows 100 at a time. ** i actually changed it to go through 15

'strOrig = origin
'strDest = destinations
'strAPI = your api key
'strMode = walking or driving
'strReg = 2 digit region code

Dim strURL As String
Dim objHttp As MSXML2.XMLHTTP60
Dim objDom As DOMDocument60
Dim aryDest() As String

' in case region code is not included from cheet
If Len(strReg) > 0 Then
    strRegURL = "&region=" & strReg
End If

' I put all the variables below on separate lines just to show how I'm building the URL
'strURL = "https://maps.googleapis.com/maps/api/distancematrix/xml?units=imperial" & _
'        "&origins=" & strOrig & _
'        "&destinations=" & strDest & _
'        "&mode=" & strMode & _
'        strRegURL & _
'        "&key=" & strAPI
        
perro:
' swap destination and origin
strURL = "https://maps.googleapis.com/maps/api/distancematrix/xml?units=imperial" & _
        "&origins=" & strOrig & _
        "&destinations=" & strDest & _
        "&traffic_mode1=optimistic" & _
        "&mode=" & strMode & _
        strRegURL & _
        "&key=" & strAPI
Set objHttp = New MSXML2.XMLHTTP60

With objHttp
    .Open "GET", strURL, False
    .setRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
    .send
End With
Set objDom = New DOMDocument60
objDom.LoadXML (objHttp.responseText)

'objDom.LoadXML objXHTTP.responseText
Dim strStatus As String
strStatus = objDom.SelectSingleNode("//status").Text

If strStatus = "OK" Then 'we have data to parse
    errores = 0
    numrows = objDom.SelectNodes("//row/element").Length
    ReDim aryDest(numrows - 1, 1)
    'get the rows
    For x = 1 To numrows
        Dim datanode As MSXML2.IXMLDOMNode
        Set datanode = objDom.SelectNodes("//row/element")(x - 1)
        
        If datanode.SelectNodes("status")(0).Text = "OK" Then
            strDur = datanode.ChildNodes(1).ChildNodes(0).Text
            strDur = Str(Round(Val(strDur) / 60, 1))  'convert seconds to minutes
            strDist = datanode.ChildNodes(2).ChildNodes(0).Text
            strDist = Round(Val(strDist) / 1000, 3) 'convert from meters to kilometers
            aryDest(x - 1, 0) = strDur  'durations in seconds, converted to minutes
            aryDest(x - 1, 1) = strDist  'distance in meters converted to miles
        Else
            aryDest(x - 1, 0) = datanode.SelectNodes("status")(0).Text
            aryDest(x - 1, 1) = datanode.SelectNodes("status")(0).Text
        End If
    Next
Else
    errores = errores + 1
    If errores <= 3 Then
        Application.Wait (Now + TimeValue("00:00:0" & errores))
        GoTo perro
    End If
    ReDim aryDest(0, 0)
    aryDest(0, 0) = "NO DATA"
End If


Set objDom = Nothing
Set objHttp = Nothing
gDist = aryDest

End Function



