Sub payments()
    ' Get FORN and FT when category = to payments
    Dim codascii As Byte
    Dim currentCell, caracter, forn, nf As String
    Dim index, limit, i, j As Integer
    
    
    currentCell = Range("G6").Value
   
    forn = ""
    nf = ""
    limit = 1
    For index = 1 To Len(currentCell)
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
      'find first serie of numbers
      
        'find third ":"
        If codascii = 58 Then
            If limit = 3 Then
                k = index
                limit = limit + 1
            End If
            'find second ":"
            If limit = 2 Then
                j = index
                limit = limit + 1
            End If
            'find first ":"
            If limit <= 1 Then
                i = index
                limit = limit + 1
            End If
        End If
    Next index
    'get forn number
    index = i + 1
    caracter = Mid(currentCell, index, 1)
    codascii = Asc(caracter)
    If codascii = 32 Then
        index = index + 1
    End If
    Do Until index > j
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        
        If codascii >= 48 And codascii <= 57 Then
            forn = forn & caracter
        End If
        index = index + 1
    Loop
    
    'get ft number
    index = k + 1
    caracter = Mid(currentCell, index, 1)
    codascii = Asc(caracter)
    If codascii = 32 Then
        index = index + 1
    End If
    Do Until index > Len(currentCell)
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        If codascii >= 48 And codascii <= 57 Then
            nf = nf & caracter
        End If
        If codascii <= 47 Or codascii > 57 Then
           Exit Do
        End If
        index = index + 1
    Loop
     MsgBox ("Forn: " & forn & "  " _
            & "NF: " & nf)
End Sub

Sub globo()
' Get key value when category = GLOBO
    Dim codascii As Byte
    Dim currentCell, caracter, forn As String
    Dim index, limit, i, j As Integer
        
    currentCell = Range("G10").Value
    forn = ""
    For index = 1 To Len(currentCell)
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        'find PC word
        If codascii = 80 Then
            limit = index + 1
        End If
        If index = limit And codascii = 67 Then
            i = index
        End If
    Next index
    'after PC word I get numbers and quit until I get non number value
    For index = i + 1 To Len(currentCell)
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        
        If codascii <= 47 Or codascii > 57 Then
           Exit For
        End If
        If codascii >= 48 And codascii <= 57 Then
            forn = forn & caracter
        End If
    Next index
End Sub

Sub rec()
' Get FORN and FT when category = to REC
 Dim codascii As Byte
    Dim currentCell, caracter, forn, nf As String
    Dim index, limit, i, j, k As Integer
    
    
    currentCell = Range("G14").Value
   
    forn = ""
    nf = ""
    limit = 1
    For index = 1 To Len(currentCell)
        
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
      'find first serie of numbers
      
      'find third ":"
        If codascii = 58 Then
            If limit = 3 Then
                k = index
                limit = limit + 1
                 
            End If
              'find second ":"
            If limit = 2 Then
                j = index
                limit = limit + 1
                 
            End If
              'find first ":"
            If limit <= 1 Then
                i = index
                limit = limit + 1
                
            End If
        End If
    Next index
    
    'get forn number
    index = i + 1
    caracter = Mid(currentCell, index, 1)
    codascii = Asc(caracter)
    If codascii = 32 Then
        index = index + 1
    End If
    Do Until index > j
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        
        If codascii >= 48 And codascii <= 57 Then
            forn = forn & caracter
        End If
        If codascii <= 47 Or codascii > 57 Then
           Exit Do
        End If
        index = index + 1
    Loop
    
    'get ft number
    index = j + 1
    caracter = Mid(currentCell, index, 1)
    codascii = Asc(caracter)
    If codascii = 32 Then
        index = index + 1
    End If
    Do Until index > k
        caracter = Mid(currentCell, index, 1)
        codascii = Asc(caracter)
        If codascii >= 48 And codascii <= 57 Then
            nf = nf & caracter
        End If
        If codascii <= 47 Or codascii > 57 Then
           Exit Do
        End If
        index = index + 1
    Loop
   
    MsgBox ("Forn: " & forn & "  " _
            & "NF: " & nf)
End Sub




