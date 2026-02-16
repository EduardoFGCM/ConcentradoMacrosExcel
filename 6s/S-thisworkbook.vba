Option Explicit

' =====================================================
' CONFIGURACION DE SALTO A LA DERECHA
' Guarda y restaura la configuracion original del usuario
' =====================================================
Private ConfigOriginal_MoveAfterReturn As Boolean
Private ConfigOriginal_Direction As XlDirection
Private ConfigGuardada As Boolean

Private Sub Workbook_Open()
    GuardarConfigOriginal
    AplicarSaltosDerecha
    RegistrarInterceptor
End Sub

Private Sub Workbook_Activate()
    AplicarSaltosDerecha
    RegistrarInterceptor
End Sub

Private Sub Workbook_Deactivate()
    RestaurarConfigOriginal
    LimpiarInterceptor
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RestaurarConfigOriginal
    LimpiarInterceptor
End Sub

' --- Guardar config original (solo una vez) ---
Private Sub GuardarConfigOriginal()
    If Not ConfigGuardada Then
        ConfigOriginal_MoveAfterReturn = Application.MoveAfterReturn
        ConfigOriginal_Direction = Application.MoveAfterReturnDirection
        ConfigGuardada = True
    End If
End Sub

' --- Aplicar salto a la derecha ---
Private Sub AplicarSaltosDerecha()
    Application.MoveAfterReturn = True
    Application.MoveAfterReturnDirection = xlToRight
End Sub

' --- Restaurar config del usuario ---
Private Sub RestaurarConfigOriginal()
    If ConfigGuardada Then
        Application.MoveAfterReturn = ConfigOriginal_MoveAfterReturn
        Application.MoveAfterReturnDirection = ConfigOriginal_Direction
    End If
End Sub

' --- Registrar interceptor de Enter ---
Private Sub RegistrarInterceptor()
    On Error Resume Next
    Application.OnKey "~", "InterceptarEnter"
    On Error GoTo 0
End Sub

' --- Limpiar interceptor de Enter ---
Private Sub LimpiarInterceptor()
    On Error Resume Next
    Application.OnKey "~"
    On Error GoTo 0
End Sub
