Option Explicit

' =====================================================
' THISWORKBOOK - CONFIGURACIÓN INICIAL Y EVENTOS
' =====================================================

Private direccionOriginal As XlDirection
Private moveAfterReturnOriginal As Boolean

Private Sub Workbook_Open()
    On Error GoTo CleanUp

    ' Guardar configuración original
    moveAfterReturnOriginal = Application.MoveAfterReturn
    direccionOriginal = Application.MoveAfterReturnDirection

    ' Cambiar la dirección del Enter a la derecha (xlToRight)
    ' xlDown = 1, xlToRight = 2, xlToLeft = 3, xlUp = 4
    Application.MoveAfterReturn = True
    Application.MoveAfterReturnDirection = xlToRight

    ' Interceptar Enter para saltos personalizados
    Application.OnKey "~"
    Application.OnKey "~", "InterceptarEnter_Nuevo"

    ' Ocultar automáticamente la hoja de Inventario
    ThisWorkbook.Sheets("Inventario").Visible = xlSheetVeryHidden

CleanUp:
    On Error GoTo 0
    
    ' Mensaje de bienvenida (opcional)
    ' MsgBox "Sistema de Control de Inventario iniciado.", vbInformation, "Bienvenido"
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo CleanUp

    ' Restaurar la configuración original del Enter
    Application.MoveAfterReturn = moveAfterReturnOriginal
    Application.MoveAfterReturnDirection = direccionOriginal

    ' Limpiar intercepción del Enter
    Application.OnKey "~"

CleanUp:
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
    On Error Resume Next
    
    ' Proteger hoja de Inventario
    ThisWorkbook.Sheets("Inventario").Protect PASSWORD:=APP_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    ' Proteger hoja de Catálogo (solo lectura para usuarios)
    ThisWorkbook.Sheets("Catálogo").Protect PASSWORD:=APP_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    MsgBox "Hojas protegidas correctamente.", vbInformation
End Sub


Sub DesprotegerHojas()
    Dim pass As String
    pass = InputBox("Ingresa la contraseña:")

    If pass <> APP_PASSWORD Then
        MsgBox "Contraseña incorrecta", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    
    ThisWorkbook.Sheets("Inventario").Unprotect PASSWORD:=APP_PASSWORD
    ThisWorkbook.Sheets("Catálogo").Unprotect PASSWORD:=APP_PASSWORD
    
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
    
    If pass <> APP_PASSWORD Then
        MsgBox "Contraseña incorrecta. Operación cancelada.", vbCritical
        Exit Sub
    End If
    
    On Error GoTo CleanUp
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
    
CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number = 0 Then
        MsgBox "Datos de prueba limpiados correctamente.", vbInformation
    Else
        MsgBox "Error al limpiar datos: " & Err.Description, vbCritical
    End If
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