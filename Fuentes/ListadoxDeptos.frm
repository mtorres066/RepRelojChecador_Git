VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FrmImpresionDeptos 
   Caption         =   "Impresion"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 Reporte 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      lastProp        =   500
      _cx             =   12303
      _cy             =   10610
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FrmImpresionDeptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Crystal As New CRAXDRT.Application
'Private Reporte As New CRAXDRT.Report
Private BasedeDatos As CRAXDRT.Database
Private Tablas As CRAXDRT.DatabaseTables
Private Tabla As CRAXDRT.DatabaseTable
Dim SubReport As CRAXDRT.Report
Dim Sections As CRAXDRT.Sections
Dim Section As CRAXDRT.Section
Dim RepObjs As CRAXDRT.ReportObjects
Dim SubReportObj As CRAXDRT.SubreportObject

Dim i As Integer
Dim Cont As Integer
Dim n As Integer

Private Sub Form_Load()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load-------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load
' Procedimiento : Form_Load
' Fecha         : 20/11/2006 13:50
' Autor         : Miguel
' Propósito     :CARGA PRINCIPAL DEL FORMULARIO
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------Form_Load------------------------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo Form_Load_Error
On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
        'ASIGNA A LA VARIABLE REPORTE EL NOMBRE Y RUTA DEL REPORTE.
        Set Reporte = Crystal.OpenReport(App.Path & "\ListadoxDeptos.rpt", 1)
        Reporte.DiscardSavedData
        Reporte.ReportTitle = GTituloReporte
        Reporte.ReportComments = GComentarioReporte
        Reporte.ReportAuthor = GSubtituloReporte
                        
            Set BasedeDatos = Reporte.Database
                Set Tablas = BasedeDatos.Tables
                Cont = 1
                
                    For Each Tabla In Tablas
                    
                         Tabla.SetDataSource "\\Respaldosfepsa\relojchecadorhuella\fepsa2.mdb"
                         
                            With Tabla.ConnectionProperties
                    
                                .Item("Jet Database Password") = ""
                            
                            End With
                            
                        Cont = Cont + 1
                           
                           If Err <> 0 Then
                                MsgBox Err.Description
                                Err.Clear
                            End If
                            
                    Next Tabla
                
                Set Sections = Reporte.Sections
                For n = 1 To Sections.Count
                        Set Section = Sections.Item(n)
                        Set RepObjs = Section.ReportObjects
                        
                            For i = 1 To RepObjs.Count
                                  
                                  If RepObjs.Item(i).Kind = crSubreportObject Then
                                  
                                        Set SubReportObj = RepObjs.Item(i)
                                        Set SubReport = SubReportObj.OpenSubreport
                                        Set BasedeDatos = SubReport.Database
                                        Set Tablas = BasedeDatos.Tables
                                        Cont = 1
                                      
                                            For Each Tabla In Tablas
                                                
                                                 Tabla.SetDataSource "\\Respaldosfepsa\relojchecadorhuella\fepsa2.mdb"
                                                                                                    
                                                    With Tabla.ConnectionProperties
                                                    
                                                        .Item("Jet Database Password") = ""
                                                        
                                                    End With
                                                    
                                                Cont = Cont + 1
                                                   
                                                   If Err <> 0 Then
                                                        MsgBox Err.Description
                                                        Err.Clear
                                                    End If
                                                    
                                            Next Tabla
                                              
                                  End If
                              
                            Next i
                Next n
               
        
        Sacar_Filtro_Imprimir
        
        
        'SELECCIONA LOS DATOS DEL REPORTE
        Reporte.RecordSelectionFormula = GCriteriaReporte
        'ASIGNA EL REPORTE AL CRViewer
        crViewer.ReportSource = Reporte
        crViewer.ViewReport
        crViewer.Zoom (85)
                    
        If Err <> 0 Then
            MsgBox "err" & Err.Number & Err.Description
            Err.Clear
        End If

    Screen.MousePointer = vbDefault

On Error GoTo 0
    Exit Sub
Form_Load_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario Reporte"
End Sub

Private Sub Form_Resize()
    crViewer.Top = 0
    crViewer.Left = 0
    crViewer.Height = ScaleHeight
    crViewer.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
     
     Set Reporte = Nothing
     Set Crystal = Nothing
     If Err <> 0 Then
     End If

End Sub

Private Sub PrinterBotton_Click()
   Reporte.PrinterSetupEx Me.hWnd
    Reporte.PrintOut True, 1

End Sub


Public Sub Sacar_Filtro_Imprimir()
   On Error GoTo Sacar_Filtro_Imprimir_Error

   
               VDia = Day(DTPSolComFecIni.Value)
            VMes = Month(DTPSolComFecIni.Value)
            VAño = Year(DTPSolComFecIni.Value)
            VDia2 = Day(DTPSolComFecFin.Value)
            VMes2 = Month(DTPSolComFecFin.Value)
            VAño2 = Year(DTPSolComFecFin.Value)
            
                    'FECHAS __________________________________________________________________
                    If OptSolCom.Item(0).Value = True Then
                                GTituloReporte = "Desde " & DTPSolComFecIni.Value & " Hasta " & DTPSolComFecFin.Value
                                GCriteriaReporte = "{M_RepuestosEncSolicitudCompra.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    'REPUESTO
                    ElseIf OptSolCom.Item(1).Value = True Then
                                GTituloReporte = "Desde " & DTPSolComFecIni.Value & " Hasta " & DTPSolComFecFin.Value & " Producto " & LblSolCom.Caption
                                GCriteriaReporte = "{M_RepuestosEncSolicitudCompra.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {M_RepuestosDetSolicitudCompra.Codigo} Like '" & TxtSolCom.Text & "*'"
                                
                    'PROVEEDOR
                    ElseIf OptSolCom.Item(2).Value = True Then
                                GTituloReporte = "Desde " & DTPSolComFecIni.Value & " Hasta " & DTPSolComFecFin.Value & " Producto " & LblSolCom.Caption & " Proveedor " & LblSolComPro.Caption
                                GCriteriaReporte = "{M_RepuestosEncSolicitudCompra.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {M_RepuestosDetSolicitudCompra.Proveedor} Like '" & TxtSolComPro.Text & "*'"
                    'ORDEN
                    ElseIf OptSolCom.Item(3).Value = True Then
                                If Not IsNumeric(TxtSolCom.Text) Then
                                    MsgBox "# De Orden De Compra Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                                    Exit Sub
                                Else
                                    GTituloReporte = ""
                                    GCriteriaReporte = "{M_RepuestosEncSolicitudCompra.Documento} = " & TxtSolCom.Text
                                End If
                    End If
                    
                    If OptSolCom.Item(3).Value = True Then
                    Else
                            If OptSolComEst.Item(0).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & " And {M_RepuestosEncSolicitudCompra.Estado} = 'PENDIENTE'"
                            ElseIf OptSolComEst.Item(1).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & " And {M_RepuestosEncSolicitudCompra.Estado} = 'AUTORIZADA'"
                            ElseIf OptSolComEst.Item(2).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & " And {M_RepuestosEncSolicitudCompra.Estado} = 'ENTREGADA'"
                            ElseIf OptSolComEst.Item(3).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & " And {M_RepuestosEncSolicitudCompra.Estado} = 'PARCIAL'"
                            ElseIf OptSolComEst.Item(4).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & " And {M_RepuestosEncSolicitudCompra.Estado} = 'ANULADA'"
                            ElseIf OptSolComEst.Item(5).Value = True Then
                                GCriteriaReporte = GCriteriaReporte & "And {M_RepuestosEncSolicitudCompra.Estado} = 'DENEGADA'"
                            End If
                    End If
                
                   'ELIGE REPORTE DE ACUERDO A LA OPCION
                    'DETALLE
                    If OptSolComTipRep.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "SolicitudCompraDetalle.rpt"
                        Else
                            GNombreReporte = "SolicitudCompraDetalleO.rpt"
                        End If
                    'FROMATO
                    ElseIf OptSolComTipRep.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "SolicitudCompraFormato.rpt"
                        Else
                            GNombreReporte = "SolicitudCompraFormatoO.rpt"
                        End If
                    'RESUMEN
                    ElseIf OptSolComTipRep.Item(2).Value = True Then
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "SolicitudCompraResumen.rpt"
                        Else
                            GNombreReporte = "SolicitudCompraResumenO.rpt"
                        End If
                                                      
                    End If
   
   
   
   
   
   

   On Error GoTo 0

   Exit Sub
Sacar_Filtro_Imprimir_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Sacar_Filtro_Imprimir of Formulario FrmImpresionDeptos"
End Sub
