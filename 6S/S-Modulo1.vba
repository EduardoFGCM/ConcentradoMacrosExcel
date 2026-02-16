Option Explicit

' =====================================================
' MÓDULO 1 - FUNCIONES AUXILIARES Y NAVEGACIÓN
' =====================================================

Public Const APP_PASSWORD As String = "cOMERLAT2025"

Public Sub InterceptarEnter_Nuevo()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge <> 1 Then Exit Sub

    Dim moveDirection As XlDirection
    Dim moveAfterReturn As Boolean
    Dim col As Long

    On Error GoTo CleanUp

    moveAfterReturn = Application.MoveAfterReturn
    If moveAfterReturn Then
        moveDirection = Application.MoveAfterReturnDirection
        Application.MoveAfterReturn = False
    End If

    Application.EnableEvents = False

    ' Verificar si estamos en Hoja Concentrado
    If ActiveSheet.Name <> "Concentrado" Then
        Selection.Offset(0, 1).Select
        GoTo CleanUp
    End If

    col = Selection.Column
    
    ' =====================================================
    ' NAVEGACIÓN OPTIMIZADA PARA CONTROL DE INVENTARIO
    ' Estructura de columnas (cada 4 columnas):
    ' - Código Material: Q(17), U(21), Y(25), AC(29)... hasta EC(135)
    ' - Referencia (auto): R(18), V(22), Z(26), AD(30)... hasta ED(136)
    ' - Cantidad: S(19), W(23), AA(27), AE(31)... hasta EE(137)
    ' - Lote: T(20), X(24), AB(28), AF(32)... hasta EF(138)
    ' =====================================================
    
    ' LOTE ? CANTIDAD del mismo grupo (-1)
    ' Columnas de LOTE: T=20, X=24, AB=28, AF=32... hasta EF=138
    If col >= 20 And col <= 138 Then
        If (col - 20) Mod 4 = 0 Then  ' Es columna de LOTE
            Selection.Offset(0, -1).Select  ' Ir a CANTIDAD (-1)
            GoTo CleanUp
        End If
    End If
    
    ' CANTIDAD ? CÓDIGO MATERIAL del siguiente grupo (+2)
    ' Columnas de CANTIDAD: S=19, W=23, AA=27, AE=31... hasta EE=137
    If col >= 19 And col <= 137 Then
        If (col - 19) Mod 4 = 0 Then  ' Es columna de CANTIDAD
            ' Solo saltar si hay cantidad ingresada
            If Trim(Selection.Value) <> "" And Selection.Value <> 0 Then
                Selection.Offset(0, 2).Select  ' Saltar +2 al siguiente CÓDIGO MATERIAL
            Else
                Selection.Offset(0, 1).Select  ' Salto normal a la derecha
            End If
            GoTo CleanUp
        End If
    End If
    
    ' Salto normal a la derecha para todas las demás celdas
    Selection.Offset(0, 1).Select
    
CleanUp:
    Application.EnableEvents = True

    ' Restaurar configuración de Excel
    If moveAfterReturn Then
        Application.MoveAfterReturn = True
        Application.MoveAfterReturnDirection = moveDirection
    End If
End Sub


Public Sub InterceptarEnterTarget_Nuevo(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub
    If rng.Cells.CountLarge <> 1 Then Exit Sub

    Application.EnableEvents = False
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
    referencia = Trim(CStr(referencia))
    If referencia = "" Then Exit Function

    Dim fila As Long
    fila = FindCatalogoRow(referencia, "")
    If fila > 0 Then
        BuscarDescripcionCatalogo = ThisWorkbook.Sheets("Catálogo").Cells(fila, "B").Value
    End If
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
    referencia = Trim(CStr(referencia))
    lote = Trim(CStr(lote))
    If referencia = "" Or lote = "" Then Exit Function

    ExisteLoteEnCatalogo = (FindCatalogoRow(referencia, lote) > 0)
    Exit Function

ErrorHandler:
    ExisteLoteEnCatalogo = False
End Function

Public Function FindCatalogoRow(ByVal referencia As String, ByVal lote As String) As Long
    On Error GoTo ErrorHandler

    FindCatalogoRow = 0
    referencia = Trim(CStr(referencia))
    lote = Trim(CStr(lote))
    If referencia = "" Then Exit Function

    Dim wsCat As Worksheet
    Dim ultimaFila As Long
    Dim rng As Range
    Dim cell As Range
    Dim firstAddr As String

    Set wsCat = ThisWorkbook.Sheets("Catálogo")
    ultimaFila = wsCat.Cells(wsCat.Rows.Count, "C").End(xlUp).Row
    If ultimaFila < 2 Then Exit Function

    Set rng = wsCat.Range("C2:C" & ultimaFila)
    Set cell = rng.Find(What:=referencia, LookIn:=xlValues, LookAt:=xlWhole)
    If cell Is Nothing Then Exit Function

    firstAddr = cell.Address
    Do
        If lote = "" Or Trim(CStr(wsCat.Cells(cell.Row, "D").Value)) = lote Then
            FindCatalogoRow = cell.Row
            Exit Function
        End If
        Set cell = rng.FindNext(cell)
    Loop While Not cell Is Nothing And cell.Address <> firstAddr

    Exit Function

ErrorHandler:
    FindCatalogoRow = 0
End Function