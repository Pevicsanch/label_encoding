Function CATCODE(R1 As Range) As Variant
    Dim dataArr As Variant
    Dim resultArr As Variant
    Dim uniqueValues As Collection
    Dim valueCount As Integer
    Dim i As Integer, j As Integer
    Dim k As Integer
    
    ' Obtener los datos de la hoja de cálculo en un arreglo
    dataArr = R1.value
    
    ' Inicializar la colección de valores únicos
    Set uniqueValues = New Collection
    
    ' Recorrer el arreglo para encontrar los valores únicos
    For i = 1 To UBound(dataArr, 1)
        For j = 1 To UBound(dataArr, 2)
            On Error Resume Next
            uniqueValues.Add dataArr(i, j), CStr(dataArr(i, j))
            On Error GoTo 0
        Next j
    Next i
    
    ' Calcular el número de valores únicos
    valueCount = uniqueValues.Count
    
    ' Inicializar el arreglo de resultados
    ReDim resultArr(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))
    
    ' Recorrer el arreglo original y asignar los códigos correspondientes
    k = 0
    For i = 1 To UBound(dataArr, 1)
        For j = 1 To UBound(dataArr, 2)
            k = k + 1
            resultArr(i, j) = GetCode(dataArr(i, j), uniqueValues)
        Next j
    Next i
    
    ' Asignar el resultado a la función CATCODE
    CATCODE = resultArr
End Function

Function GetCode(value As Variant, uniqueValues As Collection) As Integer
    Dim i As Integer
    
    ' Buscar el índice del valor en la colección de valores únicos
    For i = 1 To uniqueValues.Count
        If value = uniqueValues.Item(i) Then
            GetCode = i - 1 ' Restar 1 para obtener el código en el rango de 0 a k-1
            Exit Function
        End If
    Next i
    
    ' Si el valor no se encuentra en la colección, se devuelve un código negativo (-1)
    GetCode = -1
End Function

