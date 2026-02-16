Option Explicit

' =====================================================
' HOJA: ENTRADAS
' Valida que estén completos: Referencia (B), Lote (D), Remisión (G)
' Luego inserta en Catálogo e Inventario
' =====================================================

Dim oldRef As String, oldLote As String, oldRem As Variant
Dim isHandlingEnt As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    ' Guardar valores anteriores para Referencia, Lote y Remisión
    If Target.Column = 2 Or Target.Column = 4 Or Target.Column = 7 Then
        oldRef = Me.Cells(Target.Row, "B").Value
        oldLote = Me.Cells(Target.Row, "D").Value
        oldRem = Me.Cells(Target.Row, "G").Value
    End If
End Sub


' =====================================================
' FUNCIÓN: CREAR O ACTUALIZAR INVENTARIO
' +1 por cada entrada con remisión
' =====================================================
Private Function CrearOActualizarInventario(ByVal referencia As String, ByVal lote As String, ByVal cantidad As Long) As Boolean
    CrearOActualizarInventario = False
    
    If Trim(referencia) = "" Or Trim(lote) = "" Then Exit Function
    
    Dim shInv As Worksheet
    Set shInv = ThisWorkbook.Sheets("Inventario")

    Dim fila As Long
    fila = FindInventarioRow(referencia, lote)
    If fila > 0 Then
        shInv.Cells(fila, "D").Value = Val(shInv.Cells(fila, "D").Value) + cantidad
        If Val(shInv.Cells(fila, "D").Value) <= 0 Then
            shInv.Rows(fila).Delete
        End If
        CrearOActualizarInventario = True
        Exit Function
    End If
    
    ' No existe y cantidad es positiva, crear nuevo registro
    If cantidad > 0 Then
        fila = shInv.Cells(shInv.Rows.Count, "B").End(xlUp).Row + 1
        shInv.Cells(fila, "B").Value = referencia
        shInv.Cells(fila, "C").Value = lote
        shInv.Cells(fila, "D").Value = cantidad
        CrearOActualizarInventario = True
    End If
End Function


' =====================================================
' FUNCIÓN: INSERTAR O ACTUALIZAR EN CATÁLOGO
' =====================================================
Private Function InsertarEnCatalogo(ByVal referencia As String, ByVal lote As String) As Boolean
    InsertarEnCatalogo = False
    
    If Trim(referencia) = "" Or Trim(lote) = "" Then Exit Function
    
    Dim wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("Catálogo")

    Dim fila As Long
    fila = FindCatalogoRow(referencia, lote)
    If fila > 0 Then
        wsCat.Cells(fila, "E").Value = Val(wsCat.Cells(fila, "E").Value) + 1
        InsertarEnCatalogo = True
        Exit Function
    End If
    
    ' No existe, crear nuevo registro
    ' Buscar descripción (podría venir de otro sistema o dejarse vacía)
    Dim descripcion As String
    descripcion = ""  ' Aquí podrías buscar la descripción de otra fuente

    fila = wsCat.Cells(wsCat.Rows.Count, "C").End(xlUp).Row + 1
    wsCat.Cells(fila, "B").Value = descripcion
    wsCat.Cells(fila, "C").Value = referencia
    wsCat.Cells(fila, "D").Value = lote
    wsCat.Cells(fila, "E").Value = 1  ' Primera vez que aparece este lote
    
    InsertarEnCatalogo = True
End Function


Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    If isHandlingEnt Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Row < 7 Then Exit Sub ' No procesar encabezados
    
    ' Solo procesar columnas B, D, G (Referencia, Lote, Remisión)
    If Target.Column <> 2 And Target.Column <> 4 And Target.Column <> 7 Then Exit Sub

    Dim newRef As String, newLote As String, newRem As Variant
    
    newRef = Trim(Me.Cells(Target.Row, "B").Value)
    newLote = Trim(Me.Cells(Target.Row, "D").Value)
    newRem = Trim(Me.Cells(Target.Row, "G").Value)
    
    isHandlingEnt = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '===========================================================
    ' BLOQUE: IMPEDIR BORRADO DE LOTE SI HAY REMISIÓN
    '===========================================================
    If Target.Column = 4 Then ' Columna D = Lote
        If newRem <> "" And newLote = "" Then
            Target.Value = oldLote  ' restaurar el lote
            MsgBox "No puedes borrar el Lote porque la Remisión tiene dato.", vbExclamation, "OPERACIÓN NO PERMITIDA"
            GoTo Salida
        End If
    End If
    
    '===========================================================
    ' BLOQUE: IMPEDIR BORRADO DE REFERENCIA SI HAY REMISIÓN
    '===========================================================
    If Target.Column = 2 Then ' Columna B = Referencia
        If newRem <> "" And newRef = "" Then
            Target.Value = oldRef  ' restaurar la referencia
            MsgBox "No puedes borrar la Referencia porque la Remisión tiene dato.", vbExclamation, "OPERACIÓN NO PERMITIDA"
            GoTo Salida
        End If
    End If

    '===========================================================
    ' CASO A: CAMBIO EN REFERENCIA
    '===========================================================
    If Target.Column = 2 Then ' Columna B = Referencia
        
        ' Si había remisión, restar del inventario anterior
        If oldRem <> "" And oldRef <> "" And oldLote <> "" Then
            Call CrearOActualizarInventario(oldRef, oldLote, -1)
        End If
        
        ' Si hay nueva referencia con remisión, sumar al inventario nuevo
        If newRef <> "" And newLote <> "" And newRem <> "" Then
            Call CrearOActualizarInventario(newRef, newLote, 1)
        End If
        
        GoTo Salida
    End If

    '===========================================================
    ' CASO B: CAMBIO EN LOTE
    '===========================================================
    If Target.Column = 4 Then ' Columna D = Lote
        
        ' Si había remisión, restar del inventario anterior
        If oldRem <> "" And oldRef <> "" And oldLote <> "" Then
            Call CrearOActualizarInventario(oldRef, oldLote, -1)
        End If
        
        ' Si hay nuevo lote con remisión, sumar al inventario nuevo
        If newRef <> "" And newLote <> "" And newRem <> "" Then
            Call CrearOActualizarInventario(newRef, newLote, 1)
        End If
        
        GoTo Salida
    End If

    '===========================================================
    ' CASO C: CAMBIO EN REMISIÓN (EL MÁS IMPORTANTE)
    '===========================================================
    If Target.Column = 7 Then ' Columna G = Remisión
        
        ' VALIDACIÓN: Referencia y Lote deben existir
        If newRem <> "" Then
            If newRef = "" Or newLote = "" Then
                Target.Value = oldRem
                MsgBox "No se puede registrar la remisión: falta Referencia o Lote.", vbCritical, "VALIDACIÓN FALLIDA"
                GoTo Salida
            End If
        End If
        
        ' CASO: Se borra remisión (había dato, ahora está vacío)
        If oldRem <> "" And newRem = "" Then
            If oldRef <> "" And oldLote <> "" Then
                Call CrearOActualizarInventario(oldRef, oldLote, -1)
            End If
            GoTo Salida
        End If
        
        ' CASO: Se cambia remisión (había dato, ahora hay otro dato)
        If oldRem <> "" And newRem <> "" And oldRem <> newRem Then
            ' No hay cambio en inventario, solo se actualiza la remisión
            GoTo Salida
        End If
        
        ' CASO: Se agrega remisión nueva (no había dato, ahora sí)
        If oldRem = "" And newRem <> "" Then
            If newRef <> "" And newLote <> "" Then
                Call CrearOActualizarInventario(newRef, newLote, 1)
            End If
            GoTo Salida
        End If
        
    End If

Salida:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isHandlingEnt = False
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    isHandlingEnt = False
    MsgBox "Error en Worksheet_Change: " & Err.Description, vbCritical, "Error"
End Sub