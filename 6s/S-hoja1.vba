Option Explicit

' =====================================================
' DICCIONARIO PARA RASTREAR VALORES PREVIOS DE REFERENCIA
' =====================================================
Private UltimaReferencia As Object

' =====================================================
' HELPERS: ARRAYS DE COLUMNAS (REFERENCIA / CANTIDAD / LOTE)
' Centralizados para evitar inconsistencias
' =====================================================
Private Function RefCols() As Variant
    RefCols = Array("R", "W", "AB", "AE", "AH", "AK", "AN", "AQ", "AW", "AZ", "BC", "BF", "BI")
End Function

Private Function QtyCols() As Variant
    QtyCols = Array("S", "X", "AC", "AF", "AI", "AL", "AO", "AR", "AU", "AX", "BA", "BD", "BG", "BJ")
End Function

Private Function LoteCols() As Variant
    LoteCols = Array("T", "Y", "AD", "AG", "AJ", "AM", "AP", "AS", "AV", "AY", "BB", "BE", "BH", "BK")
End Function

' Verifica si una columna pertenece a un array de letras de columna
Private Function EsColumna(ByVal colNum As Long, ByRef arrCols As Variant) As Boolean
    Dim c As Variant
    For Each c In arrCols
        If colNum = Me.Range(CStr(c) & "1").Column Then
            EsColumna = True
            Exit Function
        End If
    Next c
    EsColumna = False
End Function

' Inicializa el diccionario de referencias si es necesario
Private Sub InicializarDiccionario()
    If UltimaReferencia Is Nothing Then
        Set UltimaReferencia = CreateObject("Scripting.Dictionary")
    End If
End Sub

' =====================================================
' EVENTO: ACTIVAR HOJA
' =====================================================
Private Sub Worksheet_Activate()
    InicializarDiccionario
End Sub

' =====================================================
' EVENTO: CAMBIO DE SELECCIÓN
' Guarda el valor previo de las celdas de referencia
' =====================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count <> 1 Then Exit Sub
    InicializarDiccionario

    Dim refs As Variant: refs = RefCols()
    If EsColumna(Target.Column, refs) Then
        UltimaReferencia(Target.Address) = Target.Value
    End If
End Sub

' =====================================================
' EVENTO PRINCIPAL: CAMBIO DE CELDA
' Despacha a sub-rutinas según tipo de columna
' =====================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    If Target.Row <= 4 Then Exit Sub  ' No procesar encabezados

    InicializarDiccionario

    Dim refs As Variant:  refs = RefCols()
    Dim qtys As Variant:  qtys = QtyCols()
    Dim lotes As Variant: lotes = LoteCols()

    ' --- CASO 1: CAMBIO DE REFERENCIA ---
    If EsColumna(Target.Column, refs) Then
        Call ProcesarCambioReferencia(Target)
        Exit Sub
    End If

    ' --- CASO 2: CAMBIO DE LOTE ---
    If EsColumna(Target.Column, lotes) Then
        Call ProcesarCambioLote(Target)
        Exit Sub
    End If

    ' --- CASO 3: CAMBIO DE CANTIDAD ---
    If EsColumna(Target.Column, qtys) Then
        ' Validar: debe existir lote (1 col a la derecha)
        If Trim(Target.Offset(0, 1).Value) = "" Then
            MsgBox "Debe ingresar primero el LOTE y despues la CANTIDAD.", vbExclamation
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            Exit Sub
        End If

        ' Capturar valor anterior con Undo
        Application.EnableEvents = False
        Dim ValorNuevo As Variant:    ValorNuevo = Target.Value
        Application.Undo
        Dim ValorAnterior As Variant: ValorAnterior = Target.Value
        Target.Value = ValorNuevo

        Call ProcesarInventario(ValorNuevo, ValorAnterior, Target)

        Application.EnableEvents = True
    End If

    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' =====================================================
' PROCESAR CAMBIO DE REFERENCIA
' Devuelve inventario de la referencia anterior y limpia
' cantidad + lote adyacentes
' =====================================================
Private Sub ProcesarCambioReferencia(ByVal Target As Range)
    On Error GoTo ErrRef
    Application.EnableEvents = False

    Dim clave As String: clave = Target.Address
    Dim refAnterior As String
    If UltimaReferencia.Exists(clave) Then
        refAnterior = CStr(UltimaReferencia(clave))
    Else
        refAnterior = ""
    End If

    Dim refNueva As String: refNueva = CStr(Target.Value)
    UltimaReferencia(clave) = refNueva

    ' Sin cambio real -> salir
    If Trim(refAnterior) = Trim(refNueva) Then GoTo FinRef

    ' Cantidad (+1) y lote (+2) adyacentes
    Dim cantAnterior As Double: cantAnterior = Val(Target.Offset(0, 1).Value)
    Dim loteAnt As String:      loteAnt = Trim(CStr(Target.Offset(0, 2).Value))

    ' Devolver inventario de la referencia anterior
    If cantAnterior <> 0 And loteAnt <> "" And refAnterior <> "" Then
        Call DevolverInventarioDirecto(refAnterior, loteAnt, cantAnterior)
    End If

    ' Limpiar cantidad y lote
    Target.Offset(0, 1).Value = ""
    Target.Offset(0, 2).Value = ""

FinRef:
    Application.EnableEvents = True
    Exit Sub

ErrRef:
    Application.EnableEvents = True
    MsgBox "Error al procesar referencia: " & Err.Description, vbExclamation
End Sub

' =====================================================
' PROCESAR CAMBIO DE LOTE
' Devuelve inventario del lote anterior y limpia cantidad
' =====================================================
Private Sub ProcesarCambioLote(ByVal Target As Range)
    On Error GoTo ErrLote
    Application.EnableEvents = False

    Dim loteNuevo As String: loteNuevo = CStr(Target.Value)
    Application.Undo
    Dim loteAnterior As String: loteAnterior = CStr(Target.Value)
    Target.Value = loteNuevo

    ' Sin cambio real -> salir
    If Trim(loteNuevo) = Trim(loteAnterior) Then GoTo FinLote

    ' Referencia (-2) y cantidad (-1) respecto al lote
    Dim codigoR As String:     codigoR = Trim(CStr(Target.Offset(0, -2).Value))
    Dim cantAnterior As Double: cantAnterior = Val(Target.Offset(0, -1).Value)

    ' Devolver inventario del lote anterior
    If cantAnterior <> 0 And loteAnterior <> "" And codigoR <> "" Then
        Call DevolverInventarioDirecto(codigoR, loteAnterior, cantAnterior)
    End If

    ' Limpiar cantidad
    Target.Offset(0, -1).Value = ""

FinLote:
    Application.EnableEvents = True
    Exit Sub

ErrLote:
    Application.EnableEvents = True
    MsgBox "Error al procesar lote: " & Err.Description, vbExclamation
End Sub

' =====================================================
' PROCESAR DESCUENTO / DEVOLUCIÓN DE INVENTARIO
' Valida existencia y aplica el delta correspondiente
' =====================================================
Private Sub ProcesarInventario(ByVal ValorNuevo As Variant, ByVal ValorAnterior As Variant, _
                                ByVal Target As Range)
    On Error GoTo ErrorInv

    ' Código (-1) y lote (+1) respecto a la cantidad
    Dim codigo As String: codigo = Trim(CStr(Target.Offset(0, -1).Value))
    Dim lote As String:   lote = Trim(CStr(Target.Offset(0, 1).Value))

    If codigo = "" Or lote = "" Then Exit Sub

    ' Obtener hoja de Inventario
    Dim wsInv As Worksheet
    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("Inventario")
    On Error GoTo ErrorInv
    If wsInv Is Nothing Then
        MsgBox "No se encontro la hoja 'Inventario'.", vbCritical
        Target.Value = ValorAnterior
        Exit Sub
    End If

    ' Buscar producto + lote
    Dim ultimaFila As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    Dim filaInv As Long: filaInv = 0
    Dim i As Long
    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = codigo And _
           Trim(CStr(wsInv.Cells(i, "C").Value)) = lote Then
            filaInv = i
            Exit For
        End If
    Next i

    If filaInv = 0 Then
        MsgBox "No existe en INVENTARIO el codigo: " & codigo & vbCrLf & _
               "con lote: " & lote, vbCritical
        Target.Value = ValorAnterior
        Exit Sub
    End If

    Dim existencia As Double
    existencia = Val(wsInv.Cells(filaInv, "D").Value)

    ' Calcular delta (positivo = descuento, negativo = devolucion)
    Dim delta As Double
    delta = Val(ValorNuevo) - Val(ValorAnterior)

    ' Validar: no puede exceder existencia
    If delta > 0 And delta > existencia Then
        MsgBox "La cantidad excede la existencia disponible (" & existencia & ").", vbCritical
        Target.Value = ValorAnterior
        Exit Sub
    End If

    ' Aplicar cambio
    wsInv.Cells(filaInv, "D").Value = existencia - delta

    If delta > 0 Then
        MsgBox "Se descontaron " & delta & " piezas del Inventario.", vbInformation
    ElseIf delta < 0 Then
        MsgBox "Se devolvieron " & Abs(delta) & " piezas al Inventario.", vbInformation
    End If

    Exit Sub

ErrorInv:
    Application.EnableEvents = True
    MsgBox "Error en Inventario: " & Err.Description, vbCritical
End Sub
