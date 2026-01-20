Option Explicit

Public Sub InterceptarEnter()
    ' Deshabilitar el movimiento automático de Excel
    Dim moveDirection As XlDirection
    Dim moveAfterReturn As Boolean
    
    moveAfterReturn = Application.MoveAfterReturn
    If moveAfterReturn Then
        moveDirection = Application.MoveAfterReturnDirection
        Application.MoveAfterReturn = False
    End If
    
    Application.EnableEvents = False
    
    ' Verificar si estamos en Hoja1
    If ActiveSheet.CodeName <> "Hoja1" Then
        Selection.Offset(0, 1).Select
        Application.EnableEvents = True
        ' Restaurar configuración
        If moveAfterReturn Then Application.MoveAfterReturn = True
        Exit Sub
    End If
    
    Dim col As Long
    col = Selection.Column
    
    ' =====================================================
    ' SALTO DOBLE: CANTIDAD → SIGUIENTE REFERENCIA (+2)
    ' (Solo si la celda YA tiene valor, es decir, segunda pasada)
    ' =====================================================
    ' Columnas de CANTIDAD: S, X, AC, AF, AI, AL, AO, AR, AU, AX, BA, BD, BG, BJ
    Select Case col
        Case Columns("S").Column    ' S → U (+2) o T (+1)
            If Trim(Selection.Value) <> "" Then
                Selection.Offset(0, 2).Select  ' Ya tiene valor, saltar +2
            Else
                Selection.Offset(0, 1).Select  ' Sin valor, ir a LOTE
            End If
            Application.EnableEvents = True
            Exit Sub
        Case Columns("X").Column    ' X → Z (+2) o Y (+1)
            If Trim(Selection.Value) <> "" Then
                Selection.Offset(0, 2).Select
            Else
                Selection.Offset(0, 1).Select
            End If
            Application.EnableEvents = True
            Exit Sub
        Case Columns("AC").Column   ' AC → AE (+2) o AD (+1)
            If Trim(Selection.Value) <> "" Then
                Selection.Offset(0, 2).Select
            Else
                Selection.Offset(0, 1).Select
            End If
            Application.EnableEvents = True
            Exit Sub
    End Select
    
    ' Bloque AE-BK: Columnas de CANTIDAD (AF, AI, AL, AO, AR, AU, AX, BA, BD, BG, BJ)
    If col >= 32 And col <= 62 Then   ' AF a BJ
        If (col - 31) Mod 3 = 1 Then  ' Es columna de CANTIDAD
            If Trim(Selection.Value) <> "" Then
                Selection.Offset(0, 2).Select  ' Ya tiene valor, saltar +2
            Else
                Selection.Offset(0, 1).Select  ' Sin valor, ir a LOTE
            End If
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    ' =====================================================
    ' SALTO REGRESIVO: LOTE → CANTIDAD (una celda atrás)
    ' =====================================================
    ' Columnas de LOTE: T, Y, AD, AG, AJ, AM, AP, AS, AV, AY, BB, BE, BH, BK
    If col = Columns("T").Column Or _
       col = Columns("Y").Column Or _
       col = Columns("AD").Column Or _
       (col >= 33 And col <= 63 And (col - 31) Mod 3 = 2) Then
       
        Selection.Offset(0, -1).Select  ' Regresar a CANTIDAD
        Application.EnableEvents = True
        Exit Sub
    End If
    
    ' Salto normal a la derecha para todas las demás celdas
    Selection.Offset(0, 1).Select
    
    Application.EnableEvents = True
    
    ' Restaurar configuración de Excel
    If moveAfterReturn Then Application.MoveAfterReturn = True
End Sub

Public Sub InterceptarEnterTarget(ByVal rng As Range)
    Application.EnableEvents = False
    
    ' Simple salto a la derecha
    rng.Offset(0, 1).Select
    
    Application.EnableEvents = True
End Sub
