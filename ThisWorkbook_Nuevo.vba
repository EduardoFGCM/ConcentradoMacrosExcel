Option Explicit

' =====================================================
' THISWORKBOOK - CONFIGURACIÓN INICIAL Y EVENTOS
' =====================================================

' Variable para guardar la dirección original del Enter
Private direccionOriginal As Long

Private Sub Workbook_Open()
    ' Guardar la configuración original de dirección del Enter
    direccionOriginal = Application.MoveAfterReturn
    
    ' Cambiar la dirección del Enter a la derecha (xlToRight)
    ' xlDown = 1, xlToRight = 2, xlToLeft = 3, xlUp = 4
    Application.MoveAfterReturn = True
    Application.MoveAfterReturnDirection = xlToRight
    
    ' Restaurar intercepción del Enter para saltos personalizados
    On Error Resume Next
    Application.OnKey "~"
    Application.OnKey "~", "InterceptarEnter_Nuevo"
    On Error GoTo 0
    
    ' Ocultar automáticamente la hoja de Inventario
    On Error Resume Next
    ThisWorkbook.Sheets("Inventario").Visible = xlSheetVeryHidden
    On Error GoTo 0
    
    ' Mensaje de bienvenida (opcional)
    ' MsgBox "Sistema de Control de Inventario iniciado.", vbInformation, "Bienvenido"
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Restaurar la configuración original del Enter
    Application.MoveAfterReturnDirection = xlDown
    
    ' Limpiar intercepción del Enter
    On Error Resume Next
    Application.OnKey "~"
    On Error GoTo 0
    
    ' Guardar cambios automáticamente (opcional)
    ' Me.Save
End Sub


' =====================================================
' SUB: CONFIGURACIÓN INICIAL DE HOJAS
' Ejecutar una vez para configurar las hojas
' =====================================================
Sub ConfigurarHojasIniciales()
    On Error Resume Next
    
    ' Verificar que existan las hojas necesarias
    Dim hojas As Variant
    hojas = Array("ENTRADAS Y SALIDAS", "Catálogo", "Inventario", "Concentrado")
    
    Dim hoja As Variant
    Dim existe As Boolean
    
    For Each hoja In hojas
        existe = False
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = hoja Then
                existe = True
                Exit For
            End If
        Next ws
        
        If Not existe Then
            MsgBox "Advertencia: No existe la hoja '" & hoja & "'", vbExclamation
        End If
    Next hoja
    
    MsgBox "Verificación de hojas completada.", vbInformation
End Sub


' =====================================================
' SUB: PROTEGER/DESPROTEGER HOJAS
' =====================================================
Sub ProtegerHojas()
    Dim pass As String
    pass = "2024comerlat"
    
    On Error Resume Next
    
    ' Proteger hoja de Inventario
    ThisWorkbook.Sheets("Inventario").Protect Password:=pass, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    ' Proteger hoja de Catálogo (solo lectura para usuarios)
    ThisWorkbook.Sheets("Catálogo").Protect Password:=pass, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    MsgBox "Hojas protegidas correctamente.", vbInformation
End Sub


Sub DesprotegerHojas()
    Dim pass As String
    pass = InputBox("Ingresa la contraseña:")
    
    If pass <> "2024comerlat" Then
        MsgBox "Contraseña incorrecta", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    
    ThisWorkbook.Sheets("Inventario").Unprotect Password:="2024comerlat"
    ThisWorkbook.Sheets("Catálogo").Unprotect Password:="2024comerlat"
    
    MsgBox "Hojas desprotegidas correctamente.", vbInformation
End Sub


' =====================================================
' SUB: LIMPIAR DATOS DE PRUEBA
' =====================================================
Sub LimpiarDatosPrueba()
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro de que desea limpiar todos los datos?" & vbCrLf & _
                       "Esta acción no se puede deshacer.", vbQuestion + vbYesNo, "Confirmar Limpieza")
    
    If respuesta = vbNo Then Exit Sub
    
    Dim pass As String
    pass = InputBox("Ingresa la contraseña para confirmar:")
    
    If pass <> "2024comerlat" Then
        MsgBox "Contraseña incorrecta. Operación cancelada.", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Limpiar Entradas (desde fila 7)
    Dim wsEnt As Worksheet
    Set wsEnt = ThisWorkbook.Sheets("ENTRADAS Y SALIDAS")
    If wsEnt.Cells(7, 1).Value <> "" Then
        wsEnt.Rows("7:" & wsEnt.Rows.Count).ClearContents
    End If
    
    ' Limpiar Catálogo (desde fila 2)
    Dim wsCat As Worksheet
    Set wsCat = ThisWorkbook.Sheets("Catálogo")
    If wsCat.Cells(2, 1).Value <> "" Then
        wsCat.Rows("2:" & wsCat.Rows.Count).ClearContents
    End If
    
    ' Limpiar Inventario (desde fila 2)
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")
    If wsInv.Cells(2, 1).Value <> "" Then
        wsInv.Rows("2:" & wsInv.Rows.Count).ClearContents
    End If
    
    ' Limpiar Concentrado (desde fila especificada)
    Dim wsConc As Worksheet
    Set wsConc = ThisWorkbook.Sheets("Concentrado")
    ' Nota: Ajustar según la estructura de tu hoja Concentrado
    ' wsConc.Range("Q5:EF1000").ClearContents
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Datos de prueba limpiados correctamente.", vbInformation
End Sub


' =====================================================
' SUB: GENERAR REPORTE DE INVENTARIO
' =====================================================
Sub GenerarReporteInventario()
    On Error GoTo ErrorHandler
    
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Sheets("Inventario")
    
    Dim ultima As Long
    ultima = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row
    
    If ultima < 2 Then
        MsgBox "No hay datos en el inventario.", vbInformation
        Exit Sub
    End If
    
    Dim reporte As String
    reporte = "REPORTE DE INVENTARIO" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = 2 To ultima
        reporte = reporte & "Referencia: " & wsInv.Cells(i, "B").Value & vbCrLf
        reporte = reporte & "Lote: " & wsInv.Cells(i, "C").Value & vbCrLf
        reporte = reporte & "Cantidad: " & wsInv.Cells(i, "D").Value & vbCrLf
        reporte = reporte & String(50, "-") & vbCrLf
    Next i
    
    ' Mostrar en cuadro de mensaje (para inventarios pequeños)
    ' Para inventarios grandes, considerar exportar a archivo
    MsgBox reporte, vbInformation, "Reporte de Inventario"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al generar reporte: " & Err.Description, vbCritical
End Sub
