Option Explicit

' =====================================================
' INTERCEPTAR ENTER PARA SALTOS PERSONALIZADOS
' Logica: Cantidad vacia   -> ir a Lote (+1)
'         Cantidad con valor -> saltar a siguiente Ref (+2)
'         Lote              -> regresar a Cantidad (-1)
'         Otras celdas      -> salto normal a la derecha
' =====================================================
Public Sub InterceptarEnter()
    Dim moveWasActive As Boolean
    moveWasActive = Application.MoveAfterReturn

    If moveWasActive Then Application.MoveAfterReturn = False

    Application.EnableEvents = False

    ' Fuera de Hoja1: salto normal a la derecha
    If ActiveSheet.CodeName <> "Hoja1" Then
        Selection.Offset(0, 1).Select
        GoTo Restaurar
    End If

    Dim col As Long: col = Selection.Column

    ' --- Columnas de CANTIDAD ---
    If EsColumnaQty(col) Then
        If Trim(Selection.Value) <> "" Then
            Selection.Offset(0, 2).Select   ' con valor -> +2 (siguiente ref)
        Else
            Selection.Offset(0, 1).Select   ' vacia -> +1 (lote)
        End If
        GoTo Restaurar
    End If

    ' --- Columnas de LOTE -> regresar a CANTIDAD ---
    If EsColumnaLote(col) Then
        Selection.Offset(0, -1).Select
        GoTo Restaurar
    End If

    ' --- Resto -> derecha ---
    Selection.Offset(0, 1).Select

Restaurar:
    Application.EnableEvents = True
    If moveWasActive Then Application.MoveAfterReturn = True
End Sub

' =====================================================
' HELPERS: IDENTIFICAR COLUMNAS POR NUMERO
' =====================================================

' Columnas de CANTIDAD: S(19), X(24), AC(29),
' y bloque AE-BK donde (col-31) Mod 3 = 1
Private Function EsColumnaQty(ByVal col As Long) As Boolean
    Select Case col
        Case 19, 24, 29
            EsColumnaQty = True
        Case Else
            EsColumnaQty = (col >= 32 And col <= 62 And (col - 31) Mod 3 = 1)
    End Select
End Function

' Columnas de LOTE: T(20), Y(25), AD(30),
' y bloque AE-BK donde (col-31) Mod 3 = 2
Private Function EsColumnaLote(ByVal col As Long) As Boolean
    Select Case col
        Case 20, 25, 30
            EsColumnaLote = True
        Case Else
            EsColumnaLote = (col >= 33 And col <= 63 And (col - 31) Mod 3 = 2)
    End Select
End Function
