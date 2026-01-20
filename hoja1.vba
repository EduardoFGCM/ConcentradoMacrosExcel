
Option Explicit
Dim cols, resetCols As Variant
Dim UltimaReferencia As Object

Private Sub Worksheet_Activate()
If UltimaReferencia Is Nothing Then
    Set UltimaReferencia = CreateObject("Scripting.Dictionary")
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.Count <> 1 Then Exit Sub

    If UltimaReferencia Is Nothing Then
        Set UltimaReferencia = CreateObject("Scripting.Dictionary")
    End If

    ' Guardar valor previo de referencia
    If Not Intersect(Target, Me.Range("R:R,W:W,AB:AB,AE:AE,AH:AH,AK:AK,AN:AN,AQ:AQ,AW:AW,AZ:AZ,BC:BC,BF:BF,BI:BI")) Is Nothing Then
        UltimaReferencia(Target.Address) = Target.Value
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim ValorNuevo As Variant, ValorAnterior As Variant
Dim celdaQty As Range
Dim codigoR As String

    On Error GoTo ErrHandler
    
    If Target.Cells.CountLarge <> 1 Then Exit Sub

    ' =====================================================
    ' CAMBIO REAL DE REFERENCIA (DEVUELVE Y LIMPIA)
    ' =====================================================
    resetCols = Array("R", "W", "AB", "AE", "AH", "AK", "AN", "AQ", "AW", "AZ", "BC", "BF", "BI")
    
    Dim refCol As Variant
    For Each refCol In resetCols
    
        If Target.Column = Me.Range(refCol & "1").Column Then
    
            Application.EnableEvents = False
    
            Dim clave As String
            clave = Target.Address
    
            Dim refAnterior As String
            If UltimaReferencia.Exists(clave) Then
                refAnterior = UltimaReferencia(clave)
            Else
                refAnterior = ""
            End If
    
            Dim refNueva As String
            refNueva = Target.Value
    
            ' Guardar referencia actual
            UltimaReferencia(clave) = refNueva
    
            ' Si no hubo cambio real, salir
            If Trim(refAnterior) = Trim(refNueva) Then GoTo SalidaRef
    
            ' Ubicar cantidad y lote
            Set celdaQty = Target.Offset(0, 1)
            Dim cantidadAnterior As Double
            cantidadAnterior = Val(celdaQty.Value)
    
            Dim loteAnt As String
            loteAnt = Target.Offset(0, 2).Value
    
            ' DEVOLVER INVENTARIO
            If cantidadAnterior <> 0 And loteAnt <> "" And refAnterior <> "" Then
                Call DevolverInventarioDirecto(refAnterior, loteAnt, cantidadAnterior)
            End If
    
            ' LIMPIEZA VISUAL
            celdaQty.Value = ""
            Target.Offset(0, 2).Value = ""
    
SalidaRef:
            Application.EnableEvents = True
            GoTo Cleanup
        End If
    Next refCol
    
'=========================================================================='

    ' =====================================================
    ' VALIDAR: NO PERITIR CANTIDAD SIN LOTE
    ' =====================================================
    
    ' ===== BLOQUE AE ? BK =====
    If Target.Column >= 31 And Target.Column <= 63 Then   ' AE a BK
    
        ' Columnas de CANTIDAD (AF, AI, AL, ...)
        If (Target.Column - 31) Mod 3 = 1 Then
    
            ' Lote está a la derecha
            If Trim(Target.Offset(0, 1).Value) = "" Then
                MsgBox "Debe ingresar primero el LOTE y después la CANTIDAD.", vbExclamation
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                Exit Sub
            End If
    
        End If
    End If
    
    ' ===== COLUMNAS AISLADAS =====
    Select Case Target.Column
    
        ' CANTIDAD sin lote
        Case Columns("S").Column, Columns("X").Column, Columns("AC").Column
    
            ' Lote está a la derecha
            If Trim(Target.Offset(0, 1).Value) = "" Then
                MsgBox "Debe ingresar primero el LOTE y después la CANTIDAD.", vbExclamation
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = True
                Exit Sub
            End If
    
    End Select

    ' ---------------------------------------------------
    '  DETECTAR Y MANEJAR CAMBIO DE “LOTE”
    ' ---------------------------------------------------
    Dim ultimaCol As Long, colTitulo As Long
    ultimaCol = Me.Cells(4, Me.Columns.Count).End(xlToLeft).Column

    For colTitulo = 1 To ultimaCol
        If Trim(LCase(Me.Cells(4, colTitulo).Value)) = "lote" Then

            If Target.Column = colTitulo Then
                
                Application.EnableEvents = False

                Dim loteNuevo As String, loteAnterior As String
                Dim CodigoLot As String
                
                loteNuevo = Target.Value
                Application.Undo
                loteAnterior = Target.Value
                Target.Value = loteNuevo

                ' Si realmente cambió el lote ? resetear cantidad asociada
                If Trim(loteNuevo) <> Trim(loteAnterior) Then
                    ValorNuevo = 0
                    ValorAnterior = 0
                    
                    codigoR = Target.Offset(0, -2)
                    
                    Set celdaQty = Target.Offset(0, -1)
                    ValorAnterior = celdaQty.Value
                    celdaQty.Value = 0
                    Target.Offset(0, -1).Value = ""
                    
                    ValorNuevo = celdaQty.Value
                    
                    Call Inventario(ValorNuevo, ValorAnterior, celdaQty, codigoR, loteAnterior)
                End If

                Application.EnableEvents = True
                Exit Sub
            End If
        End If
Next colTitulo

    ' ---------------------------------------------------
    ' COLUMNAS QUE DESCUENTAN INVENTARIO
    ' ---------------------------------------------------
    cols = Array("S", "X", "AC", "AF", "AI", "AL", "AO", "AR", "AU", "AX", "BA", "BD", "BG", "BJ")

    Dim col As Variant, necesitaProcesar As Boolean: necesitaProcesar = False
    For Each col In cols
        If Target.Column = Me.Range(col & "1").Column Then
            necesitaProcesar = True
            Exit For
        End If
    Next col
    If Not necesitaProcesar Then Exit Sub

    Application.EnableEvents = False

    ValorNuevo = Target.Value

    Application.Undo
    ValorAnterior = Target.Value

    Target.Value = ValorNuevo
    
    Call Inventario(ValorNuevo, ValorAnterior, Target)
    
    ' Restaurar eventos
    Application.EnableEvents = True
    Exit Sub


Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Cleanup
    

End Sub

' ===============================================================
' ============   FUNCIÓN AJUSTADA CON INVENTARIO   ==============
' ===============================================================
Private Sub Inventario(ValorNuevo As Variant, ValorAnterior As Variant, Target As Range, _
                        Optional codigoAnt As String = "", Optional loteAnt As String)

    On Error GoTo ErrorInv
    
    Dim fila As Long: fila = Target.Row
    Dim col As Variant: col = ColLetter(Target.Column)

    Dim codigo As String
    If codigoAnt = "" Then
        codigo = Target.Offset(0, -1).Value
    Else
        codigo = codigoAnt
    End If
    
    Dim lote As String
    If loteAnt = "" Then
        lote = Target.Offset(0, 1).Value
    Else
        lote = loteAnt
    End If

    ' ================================
    '  NUEVA BÚSQUEDA EN INVENTARIO
    ' ================================
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")

    Dim ultimaFila As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row

    Dim filaInv As Long: filaInv = 0
    Dim i As Long
    

    If Trim(codigo) = "" Or Trim(lote) = "" Then
        Application.EnableEvents = True
        Exit Sub
    End If

    For i = 2 To ultimaFila
        If wsInv.Cells(i, "B").Value = codigo And wsInv.Cells(i, "C").Value = lote Then
            filaInv = i
            Exit For
        End If
    Next i

    If filaInv = 0 Then
        Application.EnableEvents = True
        MsgBox "No existe en INVENTARIO el código: " & codigo & vbCrLf & _
               "con lote: " & lote, vbCritical
        Application.EnableEvents = False
        Target.Value = ValorAnterior
        Application.EnableEvents = True
        Exit Sub
    End If

    Dim existencia As Double
    existencia = Val(wsInv.Cells(filaInv, "D").Value)

    ' =====================================
    ' Cálculo de total anterior y nuevo
    ' =====================================
    Dim TotalAnterior As Double: TotalAnterior = 0
    For Each col In cols
        If Target.Column = Me.Range(col & "1").Column Then
            TotalAnterior = TotalAnterior + Val(ValorAnterior)
        Else
            TotalAnterior = TotalAnterior + Val(Me.Range(col & fila).Value)
        End If
    Next col

    Dim TotalNuevo As Double: TotalNuevo = 0
    For Each col In cols
        TotalNuevo = TotalNuevo + Val(Me.Range(col & fila).Value)
    Next col

    Dim diferenciaTotal As Double
    diferenciaTotal = TotalNuevo - TotalAnterior

    ' Validación
    If diferenciaTotal > 0 Then
        If diferenciaTotal > existencia Then
            Application.EnableEvents = True
            MsgBox "La suma excede la existencia.", vbCritical
            Application.EnableEvents = False
            Target.Value = ValorAnterior
            Application.EnableEvents = True
            Exit Sub
        End If
    End If

    ' ============================
    ' Descuento / devolución
    ' ============================
    wsInv.Cells(filaInv, "D").Value = existencia - diferenciaTotal

    Application.EnableEvents = True
    
    If diferenciaTotal > 0 Then
        MsgBox "Se descontaron las piezas del Inventario.", vbInformation
    ElseIf diferenciaTotal < 0 Then
        MsgBox "Se devolvieron las piezas al Inventario.", vbInformation
    End If
    
    Application.EnableEvents = False
    Exit Sub

ErrorInv:
    Application.EnableEvents = True
    MsgBox "Error en Inventario: " & Err.Description, vbCritical

End Sub

' Obtener letra de columna
Function ColLetter(colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function
