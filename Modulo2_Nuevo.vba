Option Explicit

' =====================================================
' MÓDULO 2 - GESTIÓN DE INVENTARIO
' =====================================================

' =====================================================
' FUNCIÓN: VERIFICAR SI EXISTE LOTE EN INVENTARIO
' Verifica que la combinación código + lote exista
' =====================================================
Public Function ExisteLoteEnInventario(ByVal codigo As String, ByVal lote As String) As Boolean
    On Error GoTo ErrorHandler
    
    ExisteLoteEnInventario = False
    
    If Trim(codigo) = "" Or Trim(lote) = "" Then Exit Function

    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")

    Dim ultimaFila As Long, i As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = Trim(CStr(codigo)) _
           And Trim(CStr(wsInv.Cells(i, "C").Value)) = Trim(CStr(lote)) Then
            ExisteLoteEnInventario = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    ExisteLoteEnInventario = False
End Function


' =====================================================
' FUNCIÓN: DEVOLVER INVENTARIO DIRECTO
' Suma la cantidad al inventario existente
' =====================================================
Public Sub DevolverInventarioDirecto_Nuevo( _
        ByVal codigo As String, _
        ByVal lote As String, _
        ByVal cantidad As Double)

    If Trim(codigo) = "" Or Trim(lote) = "" Or cantidad = 0 Then Exit Sub

    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")

    Dim ultimaFila As Long, i As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    ' Buscar el registro en inventario
    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = Trim(CStr(codigo)) _
           And Trim(CStr(wsInv.Cells(i, "C").Value)) = Trim(CStr(lote)) Then

            ' Actualizar cantidad
            wsInv.Cells(i, "D").Value = Val(wsInv.Cells(i, "D").Value) + cantidad

            ' Mensaje opcional (comentado para evitar exceso de notificaciones)
            ' MsgBox "Se devolvieron " & cantidad & " piezas al inventario." & vbCrLf & _
            '        "Referencia: " & codigo & vbCrLf & _
            '        "Lote: " & lote, vbInformation, "Inventario Actualizado"
            Exit Sub
        End If
    Next i

    ' Si no existe, crearlo solo si cantidad es positiva
    If cantidad > 0 Then
        ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row + 1
        wsInv.Cells(ultimaFila, "B").Value = codigo
        wsInv.Cells(ultimaFila, "C").Value = lote
        wsInv.Cells(ultimaFila, "D").Value = cantidad
        
        ' Mensaje opcional (comentado)
        ' MsgBox "Nuevo registro creado en inventario." & vbCrLf & _
        '        "Referencia: " & codigo & vbCrLf & _
        '        "Lote: " & lote & vbCrLf & _
        '        "Cantidad: " & cantidad, vbInformation, "Inventario Actualizado"
    Else
        MsgBox "No se encontró el producto para devolver:" & vbCrLf & _
               codigo & " / Lote " & lote, vbCritical, "Error"
    End If
End Sub


' =====================================================
' FUNCIÓN: DESCONTAR INVENTARIO
' Resta la cantidad del inventario
' =====================================================
Public Function DescontarInventario( _
        ByVal codigo As String, _
        ByVal lote As String, _
        ByVal cantidad As Double) As Boolean
    
    On Error GoTo ErrorHandler
    DescontarInventario = False
    
    If Trim(codigo) = "" Or Trim(lote) = "" Or cantidad <= 0 Then Exit Function

    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")

    Dim ultimaFila As Long, i As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    ' Buscar el registro en inventario
    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = Trim(CStr(codigo)) _
           And Trim(CStr(wsInv.Cells(i, "C").Value)) = Trim(CStr(lote)) Then

            Dim cantidadActual As Double
            cantidadActual = Val(wsInv.Cells(i, "D").Value)
            
            ' Verificar si hay suficiente inventario
            If cantidadActual < cantidad Then
                MsgBox "Inventario insuficiente." & vbCrLf & _
                       "Código: " & codigo & vbCrLf & _
                       "Lote: " & lote, vbExclamation, "Inventario Insuficiente"
                Exit Function
            End If
            
            ' Descontar cantidad
            wsInv.Cells(i, "D").Value = cantidadActual - cantidad
            
            ' Si llega a cero, eliminar el registro
            If wsInv.Cells(i, "D").Value = 0 Then
                wsInv.Rows(i).Delete
            End If
            
            ' Mensaje de éxito
            MsgBox "Piezas descontadas del inventario." & vbCrLf & _
                   "Código: " & codigo & vbCrLf & _
                   "Lote: " & lote, vbInformation, "Inventario Actualizado"
            
            DescontarInventario = True
            Exit Function
        End If
    Next i

    ' No se encontró el producto
    MsgBox "Producto no encontrado en inventario." & vbCrLf & _
           "Código: " & codigo & vbCrLf & _
           "Lote: " & lote, vbCritical, "Error"
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error al descontar inventario: " & Err.Description, vbCritical
    DescontarInventario = False
End Function


' =====================================================
' FUNCIÓN: OBTENER CANTIDAD DISPONIBLE EN INVENTARIO
' =====================================================
Public Function ObtenerCantidadInventario( _
        ByVal codigo As String, _
        ByVal lote As String) As Double
    
    On Error GoTo ErrorHandler
    ObtenerCantidadInventario = 0
    
    If Trim(codigo) = "" Or Trim(lote) = "" Then Exit Function

    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")

    Dim ultimaFila As Long, i As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = Trim(CStr(codigo)) _
           And Trim(CStr(wsInv.Cells(i, "C").Value)) = Trim(CStr(lote)) Then
            ObtenerCantidadInventario = Val(wsInv.Cells(i, "D").Value)
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    ObtenerCantidadInventario = 0
End Function


' =====================================================
' SUB: MOSTRAR/OCULTAR INVENTARIO CON CONTRASEÑA
' =====================================================
Sub MostrarOcultarInventario_Nuevo()
    Dim ws As Worksheet
    Dim pass As String

    pass = InputBox("Ingresa la contraseña:")

    If pass <> "2024comerlat" Then
        MsgBox "Contraseña incorrecta", vbCritical, "Acceso Denegado"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("Inventario")

    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetVeryHidden
        MsgBox "Hoja de Inventario oculta", vbInformation
    Else
        ws.Visible = xlSheetVisible
        ws.Activate
        MsgBox "Hoja de Inventario visible", vbInformation
    End If
End Sub
