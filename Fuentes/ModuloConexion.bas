Attribute VB_Name = "ModuloConexion"
'**************************************************************************************************************************************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************************************************************************************************************************************
' Módulo   : ModuloConexion
' Fecha    : 20/11/2006 13:51
' Autor    : Miguel
' Propósito:MODULO PARA CONEXION A LA BD Y VARIAS VARIABLES
'**************************************************************************************************************************************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************************************************************************************************************************************
Option Explicit

'CONEXION PARA ORACLE
Global Conexion As ADODB.Connection
Global GUsuario As String
Global GNumSol As Long

Global StrSql As String

'VARIABLE PARA CONECTARME AL TIPO DE PROVEEDOR
Global GConectionString As String
Global GConeccion As String
Global GPassword As String
Global GOrigenDeDatos As String

'VARIABLES PARA REPORTE
Global GNombreReporte As String
Global GCriteriaReporte As String
Global GTituloReporte As String
Global GComentarioReporte As String
Global GSubtituloReporte As String

'VARIABLES PARA FECHA  /////////////
Global GvFechaHoy As String
Global GvDescMes As String
Global GvAño As String



Public Sub Abrir_Recordset(Recordset As ADODB.Recordset, StrSql As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Abrir_Recordset-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Abrir_Recordset
' Procedimiento : Abrir_Recordset
' Fecha         : 20/11/2006 13:51
' Autor         : Miguel
' Propósito     :ABRIR RECORDSET
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Abrir_Recordset
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Abrir_Recordset------------------------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo Abrir_Recordset_Error

On Error Resume Next
    Recordset.ActiveConnection = Conexion
    Recordset.LockType = adLockOptimistic
    Recordset.CursorLocation = adUseClient
    Recordset.CursorType = adOpenDynamic
    Recordset.Open StrSql

    If Err <> 0 Then
        'MsgBox Err.Description
    End If

On Error GoTo 0
    Exit Sub
Abrir_Recordset_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Abrir_Recordset de Módulo ModuloConexion"
    
End Sub
 
Public Sub Desconectar()
    Conexion.Close
    Set Conexion = Nothing
End Sub

