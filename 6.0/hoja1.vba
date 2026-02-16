Option Explicit

' =====================================================
' HOJA: CONCENTRADO
' Columnas repetitivas cada 4 columnas (Q-T, V-Y, Z-AC...)
' Material (auto), Referencia (auto), Cantidad (manual), Lote (manual)
' NOTA: Material y Referencia se llenan automáticamente con fórmulas/listas
' ENFOQUE: Control correcto de inventario mediante Cantidad y Lote
' =====================================================

Dim oldCantidadConc As Variant
Dim oldLoteConc As String
Dim oldCodigoConc As String
Dim isHandlingConc As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
	On Error Resume Next
	If Target.Cells.Count <> 1 Then Exit Sub

	' Guardar valores anteriores de Cantidad, Lote y Código
	If Target.Column >= 19 And Target.Column <= 136 Then
		If (Target.Column - 19) Mod 4 = 0 Then  ' Columna de Cantidad
			oldCantidadConc = Target.Value
			oldLoteConc = Target.Offset(0, 1).Value
			oldCodigoConc = Target.Offset(0, -2).Value
		ElseIf (Target.Column - 20) Mod 4 = 0 Then  ' Columna de Lote
			oldCantidadConc = Target.Offset(0, -1).Value
			oldLoteConc = Target.Value
			oldCodigoConc = Target.Offset(0, -3).Value
		End If
	End If

	' Guardar valores para columnas de Código de Material
	If Target.Column >= 17 And Target.Column <= 134 Then
		If (Target.Column - 17) Mod 4 = 0 Then  ' Columna de Código Material
			oldCodigoConc = Target.Value
			oldCantidadConc = Target.Offset(0, 2).Value
			oldLoteConc = Target.Offset(0, 3).Value
		End If
	End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
	Dim valorNuevo As Variant, valorAnterior As Variant
	Dim referencia As String, lote As String
	Dim refAnterior As String, loteAnterior As String
	Dim loteNuevo As String, cantidadActual As Double
	Dim colCodigo As Long, colDescripcion As Long, colLote As Long

	On Error GoTo ErrHandler

	If isHandlingConc Then Exit Sub
	If Target.Cells.CountLarge <> 1 Then Exit Sub

	isHandlingConc = True
	Application.EnableEvents = False

	' =====================================================
	' MANEJO DE BORRADO/CAMBIO DE CÓDIGO DE MATERIAL
	' =====================================================
	If Target.Column >= 17 And Target.Column <= 134 Then
		If (Target.Column - 17) Mod 4 = 0 Then  ' Es columna de Código Material

			Dim codigoNuevo As String
			codigoNuevo = Trim(Target.Value)

			' Si se borró o cambió el código y había cantidad, devolver al inventario
			If codigoNuevo <> oldCodigoConc And oldCodigoConc <> "" And Val(oldCantidadConc) > 0 Then
				If oldLoteConc <> "" Then
					Call DevolverInventarioDirecto_Nuevo(oldCodigoConc, oldLoteConc, Val(oldCantidadConc))

					' Limpiar cantidad y lote
					Target.Offset(0, 2).Value = ""  ' Cantidad
					Target.Offset(0, 3).Value = ""  ' Lote
				End If
			End If

			GoTo Salida
		End If
	End If

	' =====================================================
	' MANEJO DE CANTIDAD
	' =====================================================
	If Target.Column >= 19 And Target.Column <= 136 Then
		If (Target.Column - 19) Mod 4 = 0 Then  ' Es columna de Cantidad

			' Obtener CÓDIGO de Material (2 columnas a la izquierda) y Lote (1 columna a la derecha)
			colCodigo = Target.Column - 2
			colDescripcion = Target.Column - 1
			colLote = Target.Column + 1

			referencia = Trim(Target.Offset(0, -2).Value)  ' CÓDIGO de material
			lote = Trim(Target.Offset(0, 1).Value)         ' LOTE

			' DETECCIÓN DE BORRADO: Si se borra la cantidad y había cantidad anterior
			If Trim(Target.Value) = "" And oldCantidadConc <> "" And Val(oldCantidadConc) > 0 Then
				refAnterior = referencia
				loteAnterior = lote

				If refAnterior <> "" And loteAnterior <> "" Then
					Call DevolverInventarioDirecto_Nuevo(refAnterior, loteAnterior, Val(oldCantidadConc))
				End If
				GoTo Salida
			End If

			' VALIDACIÓN 1: Verificar que haya Referencia
			If referencia = "" And Trim(Target.Value) <> "" Then
				Target.Value = oldCantidadConc
				MsgBox "Falta Código de Material." & vbCrLf & _
					   "Columna: " & Split(Cells(1, colCodigo).Address, "$")(1), _
					   vbExclamation, "Validación"
				GoTo Salida
			End If

			' VALIDACIÓN 2: No permitir Cantidad sin Lote
			If lote = "" And Trim(Target.Value) <> "" Then
				Target.Value = oldCantidadConc
				MsgBox "Falta Lote. Ingrese primero el Lote.", vbExclamation, "Validación"
				GoTo Salida
			End If

			' VALIDACIÓN 3: Verificar que el lote corresponda al código de material
			If referencia <> "" And lote <> "" And Trim(Target.Value) <> "" Then
				If Not ExisteLoteEnInventario(referencia, lote) Then
					Target.Value = ""
					Target.Offset(0, 1).Value = ""  ' Limpiar lote también
					MsgBox "El Lote NO corresponde al Código." & vbCrLf & _
						   "Código: " & referencia & vbCrLf & _
						   "Lote: " & lote, vbExclamation, "Lote Incorrecto"
					GoTo Salida
				End If
			End If

			' Procesar cambio de Cantidad
			valorNuevo = Target.Value
			valorAnterior = oldCantidadConc

			Call InventarioConcentrado(valorNuevo, valorAnterior, Target, referencia, lote)

			GoTo Salida
		End If
	End If

	' =====================================================
	' MANEJO DE LOTE
	' =====================================================
	If Target.Column >= 20 And Target.Column <= 137 Then
		If (Target.Column - 20) Mod 4 = 0 Then  ' Es columna de Lote

			loteNuevo = Trim(Target.Value)
			loteAnterior = Trim(oldLoteConc)
			referencia = Trim(Target.Offset(0, -3).Value)  ' CÓDIGO de material
			cantidadActual = Val(Target.Offset(0, -1).Value)

			' Si cambió el lote y había cantidad, devolver inventario y limpiar cantidad
			If loteNuevo <> loteAnterior And cantidadActual > 0 And loteAnterior <> "" And referencia <> "" Then
				Call DevolverInventarioDirecto_Nuevo(referencia, loteAnterior, cantidadActual)

				' Limpiar cantidad
				Target.Offset(0, -1).Value = ""
			End If

			GoTo Salida
		End If
	End If

Salida:
	Application.EnableEvents = True
	isHandlingConc = False
	Exit Sub

ErrHandler:
	Application.EnableEvents = True
	isHandlingConc = False
	MsgBox "Error en Worksheet_Change: " & Err.Description, vbCritical, "Error"
End Sub


' =====================================================
' SUB: GESTIÓN DE INVENTARIO PARA CONCENTRADO
' Controla el descuento y devolución de inventario
' =====================================================
Private Sub InventarioConcentrado(ByVal valorNuevo As Variant, ByVal valorAnterior As Variant, _
								  ByVal celdaCantidad As Range, ByVal referencia As String, _
								  ByVal lote As String)

	Dim cantNueva As Double, cantAnterior As Double

	cantNueva = Val(valorNuevo)
	cantAnterior = Val(valorAnterior)

	' Si no hay referencia o lote, salir
	If referencia = "" Or lote = "" Then Exit Sub

	' Calcular diferencia
	Dim diferencia As Double
	diferencia = cantNueva - cantAnterior

	If diferencia > 0 Then
		' Se incrementó la cantidad consumida, descontar del inventario
		If Not DescontarInventario(referencia, lote, diferencia) Then
			' Si falla el descuento, revertir
			Application.EnableEvents = False
			celdaCantidad.Value = valorAnterior
			Application.EnableEvents = True
			MsgBox "No se pudo descontar del inventario." & vbCrLf & _
				   "Código: " & referencia & vbCrLf & _
				   "Lote: " & lote, vbExclamation, "Error"
		End If

	ElseIf diferencia < 0 Then
		' Se redujo la cantidad consumida, devolver al inventario
		Call DevolverInventarioDirecto_Nuevo(referencia, lote, Abs(diferencia))

	ElseIf cantNueva = 0 And cantAnterior > 0 Then
		' Se limpió la cantidad, devolver todo al inventario
		Call DevolverInventarioDirecto_Nuevo(referencia, lote, cantAnterior)
	End If

End Sub
