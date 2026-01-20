Option Explicit

' =====================================================
' MÓDULO 1 - FUNCIONES AUXILIARES Y NAVEGACIÓN
' =====================================================

Public Sub InterceptarEnter_Nuevo()
    ' Deshabilitar el movimiento automático de Excel
    Dim moveDirection As XlDirection
    Dim moveAfterReturn As Boolean
    
    moveAfterReturn = Application.MoveAfterReturn
    If moveAfterReturn Then
        moveDirection = Application.MoveAfterReturnDirection
        Application.MoveAfterReturn = False
    End If
    
    Application.EnableEvents = False
    
    ' Verificar si estamos en Hoja Concentrado
    If ActiveSheet.Name <> "Concentrado" Then
        Selection.Offset(0, 1).Select
        Application.EnableEvents = True
        If moveAfterReturn Then Application.MoveAfterReturn = True
        Exit Sub
    End If
    
    Dim col As Long
    col = Selection.Column
    
    ' =====================================================
    ' NAVEGACIÓN OPTIMIZADA PARA CONTROL DE INVENTARIO
    ' Estructura de columnas (cada 4 columnas):
    ' - Código Material: Q(17), U(21), Y(25), AC(29)... hasta EC(135)
    ' - Referencia (auto): R(18), V(22), Z(26), AD(30)... hasta ED(136)
    ' - Cantidad: S(19), W(23), AA(27), AE(31)... hasta EE(137)
    ' - Lote: T(20), X(24), AB(28), AF(32)... hasta EF(138)
    ' =====================================================
    
    ' LOTE → CANTIDAD del mismo grupo (-1)
    ' Columnas de LOTE: T=20, X=24, AB=28, AF=32... hasta EF=138
    If col >= 20 And col <= 138 Then
        If (col - 20) Mod 4 = 0 Then  ' Es columna de LOTE
            Selection.Offset(0, -1).Select  ' Ir a CANTIDAD (-1)
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    ' CANTIDAD → CÓDIGO MATERIAL del siguiente grupo (+2)
    ' Columnas de CANTIDAD: S=19, W=23, AA=27, AE=31... hasta EE=137
    If col >= 19 And col <= 137 Then
        If (col - 19) Mod 4 = 0 Then  ' Es columna de CANTIDAD
            ' Solo saltar si hay cantidad ingresada
            If Trim(Selection.Value) <> "" And Selection.Value <> 0 Then
                Selection.Offset(0, 2).Select  ' Saltar +2 al siguiente CÓDIGO MATERIAL
            Else
                Selection.Offset(0, 1).Select  ' Salto normal a la derecha
            End If
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    ' Salto normal a la derecha para todas las demás celdas
    Selection.Offset(0, 1).Select
    
    Application.EnableEvents = True
    
    ' Restaurar configuración de Excel
    If moveAfterReturn Then Application.MoveAfterReturn = True
End Sub


Public Sub InterceptarEnterTarget_Nuevo(ByVal rng As Range)
    Application.EnableEvents = False
    
    ' Simple salto a la derecha
    rng.Offset(0, 1).Select
    
    Application.EnableEvents = True
End Sub


' =====================================================
' FUNCIÓN: BUSCAR REFERENCIA EN CATÁLOGO
' Devuelve la descripción correspondiente
' =====================================================
Public Function BuscarDescripcionCatalogo(ByVal referencia As String) As String
    On Error GoTo ErrorHandler
    
    BuscarDescripcionCatalogo = ""
    
    If Trim(referencia) = "" Then Exit Function
    
    Dim wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("Catálogo")
    
    Dim ultimaFila As Long, i As Long
    ultimaFila = wsCat.Cells(wsCat.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To ultimaFila
        If Trim(CStr(wsCat.Cells(i, "C").Value)) = Trim(CStr(referencia)) Then
            BuscarDescripcionCatalogo = wsCat.Cells(i, "B").Value
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    BuscarDescripcionCatalogo = ""
End Function


' =====================================================
' FUNCIÓN: VERIFICAR SI EXISTE LOTE EN CATÁLOGO
' =====================================================
Public Function ExisteLoteEnCatalogo(ByVal referencia As String, ByVal lote As String) As Boolean
    On Error GoTo ErrorHandler
    
    ExisteLoteEnCatalogo = False
    
    If Trim(referencia) = "" Or Trim(lote) = "" Then Exit Function
    
    Dim wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("Catálogo")
    
    Dim ultimaFila As Long, i As Long
    ultimaFila = wsCat.Cells(wsCat.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To ultimaFila
        If Trim(CStr(wsCat.Cells(i, "C").Value)) = Trim(CStr(referencia)) And _
           Trim(CStr(wsCat.Cells(i, "D").Value)) = Trim(CStr(lote)) Then
            ExisteLoteEnCatalogo = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    ExisteLoteEnCatalogo = False
End Function
