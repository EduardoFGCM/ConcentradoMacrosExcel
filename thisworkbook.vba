Option Explicit

' =====================================================
' CONFIGURACIÓN PARA SALTO A LA DERECHA
' Solo afecta este archivo, restaura configuración al salir
' =====================================================

' Variables para guardar configuración original del usuario
Private ConfigOriginal_MoveAfterReturn As Boolean
Private ConfigOriginal_Direction As XlDirection

Private Sub Workbook_Open()
    ' Guardar configuración original y aplicar salto a la derecha
    ConfigurarSaltosDerecha
    
    ' Activar intercepción del Enter para saltos personalizados
    ' NOTA: Asegúrate de que el Módulo1 esté insertado en el proyecto VBA
    On Error Resume Next
    Application.OnKey "~", "InterceptarEnter"
    On Error GoTo 0
End Sub

Private Sub Workbook_Activate()
    ' Al volver a este archivo, aplicar salto a la derecha
    ConfigurarSaltosDerecha
    
    ' Reactivar intercepción del Enter
    On Error Resume Next
    Application.OnKey "~", "InterceptarEnter"
    On Error GoTo 0
End Sub

Private Sub Workbook_Deactivate()
    ' Al cambiar a otro archivo, restaurar configuración original
    RestaurarConfiguracionOriginal
    
    ' Limpiar intercepción del Enter
    On Error Resume Next
    Application.OnKey "~"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Al cerrar, restaurar configuración original de Excel
    RestaurarConfiguracionOriginal
    
    ' Limpiar intercepción del Enter
    On Error Resume Next
    Application.OnKey "~"
    On Error GoTo 0
End Sub


' =====================================================
' CONFIGURAR SALTO A LA DERECHA (SOLO ESTE ARCHIVO)
' =====================================================
Private Sub ConfigurarSaltosDerecha()
    ' Guardar configuración actual del usuario (solo una vez)
    Static yaGuardado As Boolean
    If Not yaGuardado Then
        ConfigOriginal_MoveAfterReturn = Application.MoveAfterReturn
        ConfigOriginal_Direction = Application.MoveAfterReturnDirection
        yaGuardado = True
    End If
    
    ' Aplicar configuración: salto a la DERECHA después de Enter
    Application.MoveAfterReturn = True
    Application.MoveAfterReturnDirection = xlToRight
End Sub


' =====================================================
' RESTAURAR CONFIGURACIÓN ORIGINAL DEL USUARIO
' =====================================================
Private Sub RestaurarConfiguracionOriginal()
    ' Restaurar la configuración que tenía el usuario antes de abrir este archivo
    Application.MoveAfterReturn = ConfigOriginal_MoveAfterReturn
    Application.MoveAfterReturnDirection = ConfigOriginal_Direction
End Sub
