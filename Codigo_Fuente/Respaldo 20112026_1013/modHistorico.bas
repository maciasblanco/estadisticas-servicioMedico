Attribute VB_Name = "modHistorico"
Option Explicit

' ═══════════════════════════════════════════════════════════════════
'  MÓDULO: modHistorico
'  Registro de auditoría de todas las acciones
' ═══════════════════════════════════════════════════════════════════

Public Sub RegistrarHistorico(ByVal accion As String, _
                               ByVal hoja As String, _
                               ByVal idRegistro As Variant, _
                               ByVal campo As String, _
                               ByVal valorAnt As Variant, _
                               ByVal valorNuevo As Variant, _
                               Optional ByVal detalle As String = "")
    On Error GoTo Errores
    
    Dim ws As Worksheet
    Dim fila As Long
    Set ws = ThisWorkbook.Sheets(HOJA_HISTORICO)
    fila = ws.Cells(ws.Rows.count, HCOL_FECHA_HORA).End(xlUp).Row + 1
    
    With ws
        .Cells(fila, HCOL_FECHA_HORA).Value = Format(Now, FORMATO_DATETIME)
        .Cells(fila, HCOL_USUARIO).Value = gUsuarioActual
        .Cells(fila, HCOL_ACCION).Value = accion
        .Cells(fila, HCOL_HOJA).Value = hoja
        .Cells(fila, HCOL_ID_REGISTRO).Value = CStr(idRegistro)
        .Cells(fila, HCOL_CAMPO).Value = campo
        .Cells(fila, HCOL_VALOR_ANT).Value = CStr(valorAnt)
        .Cells(fila, HCOL_VALOR_NUEVO).Value = CStr(valorNuevo)
        .Cells(fila, HCOL_DETALLE).Value = detalle
        .Cells(fila, HCOL_EQUIPO).Value = gEquipoActual
    End With
    Exit Sub
Errores:
    ' Silencioso para no interrumpir la operación principal
End Sub
