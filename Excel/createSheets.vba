Public Sub createSheets()
    Dim domain As String
    Dim lastLine As Long
    Dim line As Long
    Dim pos As Integer, cell As Integer, index As Integer
    Dim element As Variant
    Dim ending() As Variant
    Dim cache() As Variant
    Dim check As Boolean, isInit As Boolean

    lastLine = Range("A" & Rows.Count).End(xlUp).Row
    isInit = True

    For line = lastLine To 1 Step -1
        domain = Trim(tblAll.Range("A" & line).Value)
        domain = Replace(Replace(domain, Chr(10), ""), Chr(13), "")
        
        pos = InStrRev(domain, ".")
        
        If Not pos = 0 Then
            domain = Right(domain, Len(domain) - pos)
            
            On Error GoTo Continue1
            found = Filter(ending, domain)
            
            check = False
            
            If Len(found(0)) > 0 Then
                check = True
            End If
            
Continue1:
            On Error GoTo -1
            On Error GoTo Continue2
            index = UBound(ending)
            
            If Not check Then
                ReDim cache(UBound(ending))
                cache = ending
                ReDim ending(UBound(ending) + 1)
                
                index = 0
                
                For Each element In cache
                    ending(index) = cache(index)
                    index = index + 1
                Next element
                
                ending(UBound(ending)) = domain
            End If
            
Continue2:
            If isInit Then
                ReDim ending(0)
                ending(0) = domain
                isInit = False
            End If
            
            On Error GoTo -1
        End If
    Next

    For Each element In ending
        Dim first As Variant, second As Variant, third As Variant
        Dim ws As Worksheet
        
        Set ws = mapAddresses.Sheets.Add()
        ws.Name = element
        
        cell = 1
        
        For line = lastLine To 1 Step -1
            domain = tblAll.Range("A" & line)
            pos = InStrRev(domain, ".")
            domain = Right(domain, Len(domain) - pos)
            
            If element = domain Then
                ws.Range("A" & cell).Value = tblAll.Range("A" & line)
                
                cell = cell + 1
            End If
        Next
    Next element
End Sub