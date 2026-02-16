Option Explicit

Private oldRef As String
Private oldLote As String
Private oldRem As Variant

' =====================================================
' EVENTO: CAMBIO DE SELECCIÓN
' Guarda valores previos de Referencia, Lote y Remision
' =====================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Row < 7 Then Exit Sub

    If Target.Column = 2 Or Target.Column = 4 Or Target.Column = 7 Then
        oldRef = CStr(Me.Cells(Target.Row, "B").Value)
        oldLote = CStr(Me.Cells(Target.Row, "D").Value)
        oldRem = Me.Cells(Target.Row, "G").Value
    End If
End Sub

' =====================================================
' CREAR O ACTUALIZAR INVENTARIO
' Retorna True si se procesó correctamente
' =====================================================
Private Function CrearOActualizarInventario(ByVal referencia As String, _
                                             ByVal lote As String, _
                                             ByVal cantidad As Long) As Boolean
    CrearOActualizarInventario = False
    If Trim(referencia) = "" Or Trim(lote) = "" Then Exit Function

    Dim shInv As Worksheet
    On Error Resume Next
    Set shInv = ThisWorkbook.Sheets("Inventario")
    On Error GoTo 0
    If shInv Is Nothing Then Exit Function

    Dim ultima As Long, i As Long
    ultima = shInv.Cells(shInv.Rows.Count, "B").End(xlUp).Row

    ' Buscar si ya existe la combinacion referencia + lote
    For i = 2 To ultima
        If Trim(CStr(shInv.Cells(i, "B").Value)) = Trim(referencia) And _
           Trim(CStr(shInv.Cells(i, "C").Value)) = Trim(lote) Then

            shInv.Cells(i, "D").Value = Val(shInv.Cells(i, "D").Value) + cantidad

            ' Eliminar registro si cantidad llega a 0 o menos
            If Val(shInv.Cells(i, "D").Value) <= 0 Then
                shInv.Rows(i).Delete
            End If

            CrearOActualizarInventario = True
            Exit Function
        End If
    Next i

    ' No existe: crear solo si cantidad es positiva
    If cantidad > 0 Then
        shInv.Cells(ultima + 1, "B").Value = referencia
        shInv.Cells(ultima + 1, "C").Value = lote
        shInv.Cells(ultima + 1, "D").Value = cantidad
        CrearOActualizarInventario = True
    End If
End Function

' =====================================================
' EVENTO: CAMBIO DE CELDA
' Gestiona inventario segun Referencia / Lote / Remision
' =====================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Row < 7 Then Exit Sub
    If Target.Column <> 2 And Target.Column <> 4 And Target.Column <> 7 Then Exit Sub

    Dim newRef As String:  newRef = Trim(CStr(Me.Cells(Target.Row, "B").Value))
    Dim newLote As String: newLote = Trim(CStr(Me.Cells(Target.Row, "D").Value))
    Dim newRem As String:  newRem = Trim(CStr(Me.Cells(Target.Row, "G").Value))

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' --- Impedir borrado de lote si hay remision ---
    If Target.Column = 4 And newRem <> "" And newLote = "" Then
        Target.Value = oldLote
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No puedes borrar el Lote porque la Remision tiene dato.", _
               vbExclamation, "OPERACION NO PERMITIDA"
        Exit Sub
    End If

    ' --- CASO A: CAMBIO EN REFERENCIA ---
    If Target.Column = 2 Then
        If CStr(oldRem) <> "" And Trim(oldRef) <> "" And Trim(oldLote) <> "" Then
            Call CrearOActualizarInventario(oldRef, oldLote, -1)
        End If
        If newRef <> "" And newLote <> "" And newRem <> "" Then
            Call CrearOActualizarInventario(newRef, newLote, 1)
        End If
        GoTo Salida
    End If

    ' --- CASO B: CAMBIO EN LOTE ---
    If Target.Column = 4 Then
        If CStr(oldRem) <> "" And Trim(oldRef) <> "" And Trim(oldLote) <> "" Then
            Call CrearOActualizarInventario(oldRef, oldLote, -1)
        End If
        If newRef <> "" And newLote <> "" And newRem <> "" Then
            Call CrearOActualizarInventario(newRef, newLote, 1)
        End If
        GoTo Salida
    End If

    ' --- CASO C: CAMBIO EN REMISION ---
    If Target.Column = 7 Then
        ' Validar prerrequisitos
        If newRem <> "" And (newRef = "" Or newLote = "") Then
            Target.Value = oldRem
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "No se puede registrar la remision: falta Referencia o Lote.", vbCritical
            Exit Sub
        End If

        Dim strOldRem As String: strOldRem = Trim(CStr(oldRem))

        If strOldRem <> "" And newRem = "" Then
            ' Se borra remision
            If Trim(oldRef) <> "" And Trim(oldLote) <> "" Then
                Call CrearOActualizarInventario(oldRef, oldLote, -1)
            End If

        ElseIf strOldRem <> "" And newRem <> "" And strOldRem <> newRem Then
            ' Se cambia remision
            If Trim(oldRef) <> "" And Trim(oldLote) <> "" Then
                Call CrearOActualizarInventario(oldRef, oldLote, -1)
            End If
            If newRef <> "" And newLote <> "" Then
                Call CrearOActualizarInventario(newRef, newLote, 1)
            End If

        ElseIf strOldRem = "" And newRem <> "" Then
            ' Se agrega remision nueva
            If newRef <> "" And newLote <> "" Then
                Call CrearOActualizarInventario(newRef, newLote, 1)
            End If
        End If
    End If

Salida:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en Worksheet_Change: " & Err.Description, vbCritical
End Sub

' =====================================================
' FUNCION PUBLICA: FORMULA DIFERENCIA
' Equivale a: =SI(D="","",SI(D>=E,0,E-D))
' =====================================================
Public Function AplicarFormulaDiferencia(ByVal valorD As Variant, ByVal valorE As Variant) As Variant
    If IsEmpty(valorD) Or Trim(valorD & "") = "" Then
        AplicarFormulaDiferencia = ""
    ElseIf Val(valorD) >= Val(valorE) Then
        AplicarFormulaDiferencia = 0
    Else
        AplicarFormulaDiferencia = Val(valorE) - Val(valorD)
    End If
End Function
