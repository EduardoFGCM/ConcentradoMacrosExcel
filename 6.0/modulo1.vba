Option Explicit

' =====================================================
' MÓDULO 1 - FUNCIONES AUXILIARES
' =====================================================

Public Const APP_PASSWORD As String = "cOMERLAT2025"

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
