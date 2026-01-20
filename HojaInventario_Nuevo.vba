Option Explicit

' =====================================================
' HOJA: INVENTARIO
' Columnas: Referencia (B), Lote (C), Cantidad (D)
' Control principal de existencias
' =====================================================

Dim oldRefInv As String, oldLoteInv As String, oldCantInv As Variant

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    ' Guardar valores anteriores
    If Target.Column = 2 Or Target.Column = 3 Or Target.Column = 4 Then
        oldRefInv = Me.Cells(Target.Row, "B").Value
        oldLoteInv = Me.Cells(Target.Row, "C").Value
        oldCantInv = Me.Cells(Target.Row, "D").Value
    End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub ' No procesar encabezados
    
    ' Solo procesar columnas B, C, D (Referencia, Lote, Cantidad)
    If Target.Column <> 2 And Target.Column <> 3 And Target.Column <> 4 Then Exit Sub

    Dim newRef As String, newLote As String, newCant As Variant
    
    newRef = Trim(Me.Cells(Target.Row, "B").Value)
    newLote = Trim(Me.Cells(Target.Row, "C").Value)
    newCant = Me.Cells(Target.Row, "D").Value
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '===========================================================
    ' VALIDACIÓN: NO PERMITIR CANTIDAD ANTES QUE LOTE
    '===========================================================
    If Target.Column = 4 Then ' Columna D = Cantidad
        Dim refActual As String, loteActual As String
        refActual = Trim(Me.Cells(Target.Row, "B").Value)
        loteActual = Trim(Me.Cells(Target.Row, "C").Value)
        
        If loteActual = "" And newCant <> "" Then
            Target.Value = oldCantInv  ' Restaurar valor anterior
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "Debe ingresar primero el LOTE antes de establecer la CANTIDAD.", _
                   vbExclamation, "VALIDACIÓN FALLIDA"
            Exit Sub
        End If
        
        If refActual = "" And newCant <> "" Then
            Target.Value = oldCantInv  ' Restaurar valor anterior
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "Debe ingresar primero la REFERENCIA antes de establecer la CANTIDAD.", _
                   vbExclamation, "VALIDACIÓN FALLIDA"
            Exit Sub
        End If
    End If
    
    '===========================================================
    ' VALIDACIÓN: NO PERMITIR LOTE SIN REFERENCIA
    '===========================================================
    If Target.Column = 3 Then ' Columna C = Lote
        Dim refCheck As String
        refCheck = Trim(Me.Cells(Target.Row, "B").Value)
        
        If refCheck = "" And newLote <> "" Then
            Target.Value = oldLoteInv  ' Restaurar valor anterior
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "Debe ingresar primero la REFERENCIA antes del LOTE.", _
                   vbExclamation, "VALIDACIÓN FALLIDA"
            Exit Sub
        End If
    End If

    '===========================================================
    ' CASO: CAMBIO EN CANTIDAD
    ' Actualizar en Hoja Concentrado si es necesario
    '===========================================================
    If Target.Column = 4 Then ' Columna D = Cantidad
        Dim cantAnterior As Double, cantNueva As Double
        cantAnterior = Val(oldCantInv)
        cantNueva = Val(newCant)
        
        ' Validar que la cantidad no sea negativa
        If cantNueva < 0 Then
            Target.Value = oldCantInv
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "La cantidad no puede ser negativa.", vbExclamation, "VALIDACIÓN FALLIDA"
            Exit Sub
        End If
        
        ' Si la cantidad llega a cero, preguntar si desea eliminar el registro
        If cantNueva = 0 And cantAnterior > 0 Then
            Dim respuesta As VbMsgBoxResult
            respuesta = MsgBox("La cantidad es cero. ¿Desea eliminar este registro?", _
                              vbQuestion + vbYesNo, "Confirmar Eliminación")
            If respuesta = vbYes Then
                Me.Rows(Target.Row).Delete
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Exit Sub
            End If
        End If
    End If
    
    '===========================================================
    ' CASO: CAMBIO EN REFERENCIA O LOTE
    ' Verificar que no exista duplicado y manejar borrados
    '===========================================================
    If Target.Column = 2 Or Target.Column = 3 Then
        Dim filaActual As Long
        filaActual = Target.Row
        
        Dim refVerif As String, loteVerif As String
        refVerif = Trim(Me.Cells(filaActual, "B").Value)
        loteVerif = Trim(Me.Cells(filaActual, "C").Value)
        
        ' DETECCIÓN DE BORRADO: Si se borra referencia o lote
        If Target.Column = 2 And newRef = "" And oldRefInv <> "" Then
            ' Se borró la referencia, limpiar lote y cantidad
            Me.Cells(filaActual, "C").Value = ""
            Me.Cells(filaActual, "D").Value = ""
            MsgBox "Se limpió el registro completo (Lote y Cantidad)." & vbCrLf & _
                   "Referencia borrada: " & oldRefInv, vbInformation, "Registro Limpiado"
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        If Target.Column = 3 And newLote = "" And oldLoteInv <> "" Then
            ' Se borró el lote, limpiar cantidad
            Me.Cells(filaActual, "D").Value = ""
            MsgBox "Se limpió la Cantidad del registro." & vbCrLf & _
                   "Lote borrado: " & oldLoteInv, vbInformation, "Cantidad Limpiada"
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        ' Solo verificar si ambos campos están completos
        If refVerif <> "" And loteVerif <> "" Then
            Dim i As Long, ultima As Long
            ultima = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
            
            For i = 2 To ultima
                If i <> filaActual Then
                    If Trim(Me.Cells(i, "B").Value) = refVerif And _
                       Trim(Me.Cells(i, "C").Value) = loteVerif Then
                        
                        ' Duplicado encontrado
                        If Target.Column = 2 Then
                            Target.Value = oldRefInv
                        Else
                            Target.Value = oldLoteInv
                        End If
                        
                        Application.EnableEvents = True
                        Application.ScreenUpdating = True
                        MsgBox "Ya existe un registro con esta combinación de Referencia y Lote.", _
                               vbExclamation, "REGISTRO DUPLICADO"
                        Exit Sub
                    End If
                End If
            Next i
        End If
    End If

Salida:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en Inventario: " & Err.Description, vbCritical, "Error"
End Sub


' =====================================================
' FUNCIÓN: OBTENER CANTIDAD DISPONIBLE
' Devuelve la cantidad disponible para una referencia y lote
' =====================================================
Public Function ObtenerDisponible(ByVal referencia As String, ByVal lote As String) As Double
    On Error GoTo ErrorHandler
    
    ObtenerDisponible = 0
    
    If Trim(referencia) = "" Or Trim(lote) = "" Then Exit Function
    
    Dim i As Long, ultima As Long
    ultima = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To ultima
        If Trim(Me.Cells(i, "B").Value) = Trim(referencia) And _
           Trim(Me.Cells(i, "C").Value) = Trim(lote) Then
            ObtenerDisponible = Val(Me.Cells(i, "D").Value)
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    ObtenerDisponible = 0
End Function
