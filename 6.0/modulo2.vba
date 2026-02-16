Option Explicit

' =====================================================
' MÓDULO 2 - GESTIÓN DE INVENTARIO
' =====================================================

' =====================================================
' FUNCIÓN: VERIFICAR SI EXISTE LOTE EN INVENTARIO
' Verifica que la combinación código + lote exista
' =====================================================
Public Function ExisteLoteEnInventario(ByVal codigo As String, ByVal lote As String) As Boolean
	On Error GoTo ErrorHandler

	ExisteLoteEnInventario = False

	codigo = Trim(CStr(codigo))
	lote = Trim(CStr(lote))
	If codigo = "" Or lote = "" Then Exit Function

	ExisteLoteEnInventario = (FindInventarioRow(codigo, lote) > 0)

	Exit Function

ErrorHandler:
	ExisteLoteEnInventario = False
End Function


' =====================================================
' FUNCIÓN: DEVOLVER INVENTARIO DIRECTO
' Suma la cantidad al inventario existente
' =====================================================
Public Sub DevolverInventarioDirecto_Nuevo( _
		ByVal codigo As String, _
		ByVal lote As String, _
		ByVal cantidad As Double)

	codigo = Trim(CStr(codigo))
	lote = Trim(CStr(lote))
	If codigo = "" Or lote = "" Or cantidad = 0 Then Exit Sub

	Dim wsInv As Worksheet
	Set wsInv = ThisWorkbook.Sheets("Inventario")

	Dim fila As Long
	fila = FindInventarioRow(codigo, lote)

	If fila > 0 Then
		wsInv.Cells(fila, "D").Value = Val(wsInv.Cells(fila, "D").Value) + cantidad
		Exit Sub
	End If

	' Si no existe, crearlo solo si cantidad es positiva
	If cantidad > 0 Then
		fila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row + 1
		wsInv.Cells(fila, "B").Value = codigo
		wsInv.Cells(fila, "C").Value = lote
		wsInv.Cells(fila, "D").Value = cantidad
	Else
		MsgBox "No se encontró el producto para devolver:" & vbCrLf & _
			   codigo & " / Lote " & lote, vbCritical, "Error"
	End If
End Sub


' =====================================================
' FUNCIÓN: DESCONTAR INVENTARIO
' Resta la cantidad del inventario
' =====================================================
Public Function DescontarInventario( _
		ByVal codigo As String, _
		ByVal lote As String, _
		ByVal cantidad As Double) As Boolean

	On Error GoTo ErrorHandler
	DescontarInventario = False

	codigo = Trim(CStr(codigo))
	lote = Trim(CStr(lote))
	If codigo = "" Or lote = "" Or cantidad <= 0 Then Exit Function

	Dim wsInv As Worksheet
	Set wsInv = ThisWorkbook.Sheets("Inventario")

	Dim fila As Long
	fila = FindInventarioRow(codigo, lote)
	If fila > 0 Then
		Dim cantidadActual As Double
		cantidadActual = Val(wsInv.Cells(fila, "D").Value)

		If cantidadActual < cantidad Then
			MsgBox "Inventario insuficiente." & vbCrLf & _
				   "Código: " & codigo & vbCrLf & _
				   "Lote: " & lote, vbExclamation, "Inventario Insuficiente"
			Exit Function
		End If

		wsInv.Cells(fila, "D").Value = cantidadActual - cantidad

		If wsInv.Cells(fila, "D").Value = 0 Then
			wsInv.Rows(fila).Delete
		End If

		DescontarInventario = True
		Exit Function
	End If

	MsgBox "Producto no encontrado en inventario." & vbCrLf & _
		   "Código: " & codigo & vbCrLf & _
		   "Lote: " & lote, vbCritical, "Error"

	Exit Function

ErrorHandler:
	MsgBox "Error al descontar inventario: " & Err.Description, vbCritical
	DescontarInventario = False
End Function


' =====================================================
' FUNCIÓN: OBTENER CANTIDAD DISPONIBLE EN INVENTARIO
' =====================================================
Public Function ObtenerCantidadInventario( _
		ByVal codigo As String, _
		ByVal lote As String) As Double

	On Error GoTo ErrorHandler
	ObtenerCantidadInventario = 0

	codigo = Trim(CStr(codigo))
	lote = Trim(CStr(lote))
	If codigo = "" Or lote = "" Then Exit Function

	Dim wsInv As Worksheet
	Set wsInv = ThisWorkbook.Sheets("Inventario")

	Dim fila As Long
	fila = FindInventarioRow(codigo, lote)
	If fila > 0 Then
		ObtenerCantidadInventario = Val(wsInv.Cells(fila, "D").Value)
		Exit Function
	End If

	Exit Function

ErrorHandler:
	ObtenerCantidadInventario = 0
End Function


' =====================================================
' SUB: MOSTRAR/OCULTAR INVENTARIO CON CONTRASEÑA
' =====================================================
Sub MostrarOcultarInventario_Nuevo()
	Dim ws As Worksheet
	Dim pass As String

	pass = InputBox("Ingresa la contraseña:")

	If pass <> APP_PASSWORD Then
		MsgBox "Contraseña incorrecta", vbCritical, "Acceso Denegado"
		Exit Sub
	End If

	Set ws = ThisWorkbook.Sheets("Inventario")

	If ws.Visible = xlSheetVisible Then
		ws.Visible = xlSheetVeryHidden
		MsgBox "Hoja de Inventario oculta", vbInformation
	Else
		ws.Visible = xlSheetVisible
		ws.Activate
		MsgBox "Hoja de Inventario visible", vbInformation
	End If
End Sub


Public Function FindInventarioRow(ByVal codigo As String, ByVal lote As String) As Long
	On Error GoTo ErrorHandler

	FindInventarioRow = 0
	codigo = Trim(CStr(codigo))
	lote = Trim(CStr(lote))
	If codigo = "" Or lote = "" Then Exit Function

	Dim wsInv As Worksheet
	Dim ultimaFila As Long
	Dim rng As Range
	Dim cell As Range
	Dim firstAddr As String

	Set wsInv = ThisWorkbook.Sheets("Inventario")
	ultimaFila = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row
	If ultimaFila < 2 Then Exit Function

	Set rng = wsInv.Range("B2:B" & ultimaFila)
	Set cell = rng.Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole)
	If cell Is Nothing Then Exit Function

	firstAddr = cell.Address
	Do
		If Trim(CStr(wsInv.Cells(cell.Row, "C").Value)) = lote Then
			FindInventarioRow = cell.Row
			Exit Function
		End If
		Set cell = rng.FindNext(cell)
	Loop While Not cell Is Nothing And cell.Address <> firstAddr

	Exit Function

ErrorHandler:
	FindInventarioRow = 0
End Function
