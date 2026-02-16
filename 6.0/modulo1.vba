Option Explicit

' =====================================================
' DEVOLVER INVENTARIO DIRECTAMENTE
' Suma la cantidad de vuelta a la existencia
' =====================================================
Public Sub DevolverInventarioDirecto( _
        ByVal codigo As String, _
        ByVal lote As String, _
        ByVal cantidad As Double)

    If Trim(codigo) = "" Or Trim(lote) = "" Or cantidad = 0 Then Exit Sub

    Dim wsInv As Worksheet
    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("Inventario")
    On Error GoTo 0

    If wsInv Is Nothing Then
        MsgBox "No se encontro la hoja 'Inventario'.", vbCritical
        Exit Sub
    End If

    Dim ultimaFila As Long, i As Long
    ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row

    For i = 2 To ultimaFila
        If Trim(CStr(wsInv.Cells(i, "B").Value)) = Trim(codigo) _
           And Trim(CStr(wsInv.Cells(i, "C").Value)) = Trim(lote) Then

            wsInv.Cells(i, "D").Value = Val(wsInv.Cells(i, "D").Value) + cantidad

            MsgBox "Se devolvieron " & cantidad & _
                   " piezas al inventario.", vbInformation
            Exit Sub
        End If
    Next i

    MsgBox "No se encontro el producto para devolver:" & vbCrLf & _
           codigo & " / Lote " & lote, vbCritical
End Sub

' =====================================================
' MOSTRAR / OCULTAR HOJA DE INVENTARIO (CON CONTRASENA)
' =====================================================
Public Sub MostrarOcultarInventario()
    Dim pass As String
    pass = InputBox("Ingresa la contrasena:", "Acceso a Inventario")
    If Len(pass) = 0 Then Exit Sub

    If pass <> "2024comerlat" Then
        MsgBox "Contrasena incorrecta.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Inventario")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No se encontro la hoja 'Inventario'.", vbCritical
        Exit Sub
    End If

    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetVeryHidden
    Else
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
End Sub

