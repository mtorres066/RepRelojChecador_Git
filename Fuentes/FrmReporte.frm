VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form FrmReporte 
   Caption         =   "Impresion"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporte.frx":0000
      PICN            =   "FrmReporte.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      lastProp        =   500
      _cx             =   22251
      _cy             =   11456
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
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Aplicacion As CRAXDRT.Application
Dim Reporte As CRAXDRT.Report
Dim BasedeDatos As CRAXDRT.Database
Dim Tablas As CRAXDRT.DatabaseTables
Dim Tabla As CRAXDRT.DatabaseTable

Private Sub Command1_Click()
    Reporte.PrinterSetupEx Me.hWnd
    Reporte.PrintOut True, 1

End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Screen.MousePointer = vbHourglass
                    
            Set Aplicacion = New CRAXDRT.Application
            Set Reporte = Aplicacion.OpenReport(App.Path & "\" & GNombreReporte)
                            
                Set BasedeDatos = Reporte.Database
                Set Tablas = BasedeDatos.Tables
                
                For Each Tabla In Tablas
                                                                    
                    'Tabla.ConnectionProperties("Database Name").Value = App.Path & "\metalenvases.mdb"
                    Tabla.ConnectionProperties("Jet OLEDB:Database Password").Value = ""
                    
                    'Use next line, if you are using Native connection to SQL database
                    'crxDatabaseTable.SetLogOnInfo "servername", "databasename", "userid", "password"
                        
                    'Use next line, if you are using ODBC connection to a PC or SQL database
                    'crxDatabaseTable.SetLogOnInfo "ODBC_DSN", "databasename", "userid", "password"
                    
                       If Err <> 0 Then
                            MsgBox Err.Description
                            Err.Clear
                        End If
                                                                
                Next Tabla
                                                
        'SELECCIONA LOS DATOS DEL REPORTE
        Reporte.RecordSelectionFormula = GCriteriaReporte
        
        GTituloReporte = "Desde " & Format(FrmMenuReportes.DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(FrmMenuReportes.DTFechaFinal.Value, "dd/mm/yyyy")
        Reporte.ReportTitle = GTituloReporte
        
        'Reporte.ReportComments = "Prueba"
        
        'ASIGNA EL REPORTE AL CRVIEWER
        CRViewer.ReportSource = Reporte
        CRViewer.ViewReport
        'CRViewer.Zoom (100)
                    
        If Err <> 0 Then
            MsgBox "err" & Err.Number & Err.Description
            Err.Clear
        End If

    Screen.MousePointer = vbDefault
End Sub
    

Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
          
     Set BasedeDatos = Nothing
     Set Tabla = Nothing
     Set Tablas = Nothing
     Set Reporte = Nothing
     Set Aplicacion = Nothing
                         
     If Err <> 0 Then
     End If


End Sub
