Option Explicit

' Version 6.0 - Sin logica de saltos
' Limpia cualquier binding de OnKey residual de versiones anteriores

Private Sub Workbook_Open()
    LimpiarSaltos
End Sub

Private Sub Workbook_Activate()
    LimpiarSaltos
End Sub

Private Sub Workbook_Deactivate()
    LimpiarSaltos
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    LimpiarSaltos
End Sub

Private Sub LimpiarSaltos()
    On Error Resume Next
    Application.OnKey "~"       ' Restaura Enter a comportamiento default
    Application.OnKey "{ENTER}" ' Restaura Enter numerico a default
    On Error GoTo 0
End Sub
