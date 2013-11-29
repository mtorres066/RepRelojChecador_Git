VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMenuReportes 
   BackColor       =   &H80000016&
   Caption         =   "Menu de Reportes Reloj Checador"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   Icon            =   "FrmMenuReportes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   7335
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3615
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6376
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin KewlButtonz.KewlButtons cmdBuscar 
         Height          =   615
         Left            =   3840
         TabIndex        =   18
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Buscar"
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
         BCOL            =   16777152
         BCOLO           =   16777152
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMenuReportes.frx":0442
         PICN            =   "FrmMenuReportes.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtBusqueda2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   3255
      End
      Begin KewlButtonz.KewlButtons cmdSalirBuscar2 
         Height          =   615
         Left            =   5640
         TabIndex        =   16
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Salir"
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
         BCOL            =   12640511
         BCOLO           =   12640511
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMenuReportes.frx":66F8
         PICN            =   "FrmMenuReportes.frx":6714
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2400
      TabIndex        =   8
      Top             =   360
      Width           =   3135
      Begin VB.OptionButton OptionFiltro 
         Caption         =   "x Empleados (Todos)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   20
         Top             =   2760
         Width           =   2175
      End
      Begin VB.OptionButton OptionFiltro 
         Caption         =   "x Empleados (Seleccion)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton OptionFiltro 
         Caption         =   "x Empleado (Individual)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton OptionFiltro 
         Caption         =   "x Deptos. (Seleccion)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   3600
         Width           =   2535
      End
      Begin VB.OptionButton OptionFiltro 
         Caption         =   "x Deptos. (Todos)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Tarjeta de Asistencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   2880
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccion:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Rango de Fechas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      Begin MSComCtl2.DTPicker DTFechaFinal 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17170433
         CurrentDate     =   39761
      End
      Begin MSComCtl2.DTPicker DTFechaInicio 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17170433
         CurrentDate     =   39761
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Inicio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Rango de Fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
   Begin KewlButtonz.KewlButtons CmdSalir 
      Height          =   2055
      Left            =   5880
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3625
      BTYPE           =   3
      TX              =   "Salir"
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
      BCOL            =   12640511
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMenuReportes.frx":C9AE
      PICN            =   "FrmMenuReportes.frx":C9CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons CmdImprimir 
      Height          =   2055
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3625
      BTYPE           =   3
      TX              =   "Imprimir"
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
      BCOL            =   13684430
      BCOLO           =   13684430
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmMenuReportes.frx":13ECC
      PICN            =   "FrmMenuReportes.frx":13EE8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Equipo:"
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reloj Chiapas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reloj SLP"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reloj Culiacan"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMenuReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaSeleccionDeptos   As New ADODB.Recordset

Dim RBuscaSeleccionEmpleado As New ADODB.Recordset

Dim GTituloReporte          As String

Dim VDia                    As String

Dim VMes                    As String

Dim VAño As String

Dim VDia2 As String

Dim VMes2 As String

Dim VAño2 As String

Dim BBuscandoDeptosTodos        As Boolean

Dim BBuscandoDeptosSeleccion    As Boolean

Dim BBuscandoEmpleadosSeleccion As Boolean

Dim vEquipo                     As String
Dim vEquipoCul                     As String

Dim BCuliacan As Boolean
Dim BSLP As Boolean
Dim BChiapas As Boolean


'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Form_Load
' Descripción   : CARGA PRINCIPAL DEL FORM
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:02:54
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
            
BCuliacan = False
BSLP = False
BChiapas = False


100     DTFechaInicio.Value = Date
102     DTFechaFinal.Value = Date
        
104     OptionFiltro.Item(0).Value = True
                
106     VDia = ""
108     VMes = ""
110     VAño = ""
112     VDia2 = ""
114     VMes2 = ""
116     VAño2 = ""
        
118     FrameBusqueda.Visible = False
                          
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.Form_Load " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : CmdImprimir_Click
' Descripción   : EJECUTA LA IMPRESION
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:03:45
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CmdImprimir_Click()
        '<EhHeader>
        On Error GoTo CmdImprimir_Click_Err
        '</EhHeader>
    
100     Screen.MousePointer = vbHourglass
    
102     GCriteriaReporte = ""
104     GTituloReporte = ""
106     GComentarioReporte = ""
        
108     If OptionFiltro.Item(0).Value = True Then
        
110         Imprime_Deptos
            
112     ElseIf OptionFiltro.Item(1).Value = True Then
        
114         If TxtBusqueda.Text = "" Then   'No eligio nada el usuario.
116             Call MsgBox("No eligió ningún departamento. Favor de elegir uno.", vbExclamation Or vbSystemModal, "Atencion !")
118             Screen.MousePointer = vbDefault

                Exit Sub

            Else
120             Imprime_Seleccion_Deptos
            End If
            
122     ElseIf OptionFiltro.Item(2).Value = True Then
        
124         If TxtBusqueda.Text = "" Then   'No eligio nada el usuario.
126             Call MsgBox("No eligió ningún empleado. Favor de elegir uno.", vbExclamation Or vbSystemModal, "Atencion !")
128             Screen.MousePointer = vbDefault

                Exit Sub

            Else
130             Imprime_Seleccion_Empleado
            End If
            
132     ElseIf OptionFiltro.Item(3).Value = True Then   'IMPRIME LAS TARJETAS DE ASISTENCIAS POR SELECCION EMPLEADO.//////
    
134         If TxtBusqueda.Text = "" Then   'No eligio nada el usuario.
136             Call MsgBox("No eligió ningún empleado. Favor de elegir uno.", vbExclamation Or vbSystemModal, "Atencion !")
138             Screen.MousePointer = vbDefault

                Exit Sub

            Else
140             Imprime_Seleccion_Empleado
            End If
        
142     ElseIf OptionFiltro.Item(4).Value = True Then   'IMPRIME LAS TARJETAS DE ASISTENCIAS TODOS LOS EMPLEADOS.//////
    
            '        If TxtBusqueda.Text = "" Then   'No eligio nada el usuario.
            '            Call MsgBox("No eligió ningún empleado. Favor de elegir uno.", vbExclamation Or vbSystemModal, "Atencion !")
            '            Screen.MousePointer = vbDefault
            '            Exit Sub
            '        Else
144         Imprime_Seleccion_Empleado
            
        End If
    
        'DESPLIEGA EL REPORTE
146     FrmReporte.Show 1
        'CrReportes.DiscardSavedData = True
148     MousePointer = 0

150     If Err <> 0 Then
152         MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"

            Exit Sub

        End If
             
154     TxtBusqueda.Text = ""
156     TxtBusqueda2.Text = ""
    
158     TxtBusqueda.SetFocus
    
160     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

CmdImprimir_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.CmdImprimir_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid1_Click
' Descripción   : CUANDO DOY UN CLICK DENTRO DEL DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:04:08
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid1_Click()
        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err
        '</EhHeader>
    
        'Buscando Todos los Deptos.
    
100     If BBuscandoDeptosTodos = True And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = False Then
            
            'imprime el reporte de todos los deptos.
102         FrameBusqueda.Visible = False
104         CmdImprimir.SetFocus
            
106     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = True And BBuscandoEmpleadosSeleccion = False Then
        
108         TxtBusqueda.Text = DataGrid1.Columns(0).Text
110         FrameBusqueda.Visible = False
112         CmdImprimir.SetFocus
            
114     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = True Then
116         TxtBusqueda.Text = DataGrid1.Columns(1).Text
118         FrameBusqueda.Visible = False
120         CmdImprimir.SetFocus
        End If
        
122     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

DataGrid1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.DataGrid1_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid1_DblClick
' Descripción   : CUANDO DOY DOBLE CLICK DENTRO DEL DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:04:33
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid1_DblClick()
        '<EhHeader>
        On Error GoTo DataGrid1_DblClick_Err
        '</EhHeader>

        'Buscando Todos los Deptos.
    
100     If BBuscandoDeptosTodos = True And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = False Then
            
            'imprime el reporte de todos los deptos.
102         FrameBusqueda.Visible = False
104         CmdImprimir.SetFocus
            
106     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = True And BBuscandoEmpleadosSeleccion = False Then
        
108         TxtBusqueda.Text = DataGrid1.Columns(0).Text
110         FrameBusqueda.Visible = False
112         CmdImprimir.SetFocus
            
114     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = True Then
116         TxtBusqueda.Text = DataGrid1.Columns(1).Text
118         FrameBusqueda.Visible = False
120         CmdImprimir.SetFocus
        End If
        
122     Screen.MousePointer = vbDefault
 
        '<EhFooter>
        Exit Sub

DataGrid1_DblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.DataGrid1_DblClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid1_KeyPress
' Descripción   : CUANDO PRESIONO CUALQUIER TECLA DENTRO DEL DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:05:19
'
' Parámetros    : KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err
        '</EhHeader>
    
        'Buscando Todos los Deptos.
    
100     If BBuscandoDeptosTodos = True And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = False Then
            
            'imprime el reporte de todos los deptos.
102         FrameBusqueda.Visible = False
104         CmdImprimir.SetFocus
            
106     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = True And BBuscandoEmpleadosSeleccion = False Then
        
108         TxtBusqueda.Text = DataGrid1.Columns(0).Text
110         FrameBusqueda.Visible = False
112         CmdImprimir.SetFocus
            
114     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = True Then
116         TxtBusqueda.Text = DataGrid1.Columns(1).Text
118         FrameBusqueda.Visible = False
120         CmdImprimir.SetFocus
        End If
        
122     Screen.MousePointer = vbDefault
    
        '<EhFooter>
        Exit Sub

DataGrid1_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.DataGrid1_KeyPress " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : cmdBuscar_Click
' Descripción   : CUANDO EJECUTO LA ORDEN BUSCAR
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:05:59
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdBuscar_Click()
        '<EhHeader>
        On Error GoTo cmdBuscar_Click_Err
        '</EhHeader>
   
        'Buscando Todos los Deptos.
    
100     If BBuscandoDeptosTodos = True And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = False Then
            
            'imprime el reporte de todos los deptos.
102         FrameBusqueda.Visible = False
104         CmdImprimir.SetFocus
            
106     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = True And BBuscandoEmpleadosSeleccion = False Then
                
            'Abre el Recordset y busca la seleccion del Depto
108         Set RBuscaSeleccionDeptos = New ADODB.Recordset
110         Call Abrir_Recordset(RBuscaSeleccionDeptos, "SELECT deptname FROM DEPARTMENTS WHERE deptname LIKE '%" & TxtBusqueda2.Text & "%'")
          
112         If RBuscaSeleccionDeptos.RecordCount > 0 Then
114             FrameBusqueda.Visible = True
116             Set DataGrid1.DataSource = RBuscaSeleccionDeptos
118             DataGrid1.Columns(0).Width = "2800"
            Else
120             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
            End If
            
122     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = True Then
            
            'Abre el Recordset y busca la seleccion del empleado.
124         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
126         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO WHERE street LIKE '%" & TxtBusqueda2.Text & "%'")

128         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
130             FrameBusqueda.Visible = True
132             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
134             DataGrid1.Columns(0).Width = "700"
136             DataGrid1.Columns(1).Width = "2700"
            Else
138             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
            End If
            
        End If
        
140     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

cmdBuscar_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.cmdBuscar_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptionFiltro_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptionFiltro_Click_Err
        '</EhHeader>
  
100     TxtBusqueda.Text = ""
102     TxtBusqueda2.Text = ""
 
        '<EhFooter>
        Exit Sub

OptionFiltro_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.OptionFiltro_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub TxtBusqueda_DblClick()
        '<EhHeader>
        On Error GoTo TxtBusqueda_DblClick_Err
        '</EhHeader>
    
100     If TxtBusqueda.Text = "" And OptionFiltro.Item(0).Value = True Then   'Esta buscando Todos los deptos.
        
102         Call MsgBox("Imprimirá todos los deptos. presione el boton de imprimir directamente.", vbInformation Or vbSystemModal, "Impresion")
104         FrameBusqueda.Visible = False
106         CmdImprimir.SetFocus
    
108     ElseIf TxtBusqueda.Text <> "" And OptionFiltro.Item(1).Value = True Then   'Esta buscando por una seleccion depto.
    
110         BBuscandoDeptosTodos = False
112         BBuscandoDeptosSeleccion = True
114         BBuscandoEmpleadosSeleccion = False
        
116         FrameBusqueda.Visible = True
        
118         Set RBuscaSeleccionDeptos = New ADODB.Recordset
            'Abre el Recordset y busca la seleccion del Depto
120         Set RBuscaSeleccionDeptos = New ADODB.Recordset
122         Call Abrir_Recordset(RBuscaSeleccionDeptos, "SELECT deptname FROM DEPARTMENTS WHERE deptname LIKE '%" & TxtBusqueda.Text & "%'")
          
124         If RBuscaSeleccionDeptos.RecordCount > 0 Then
126             FrameBusqueda.Visible = True
128             Set DataGrid1.DataSource = RBuscaSeleccionDeptos
130             DataGrid1.Columns(0).Width = "2800"
            Else
132             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
134             Screen.MousePointer = vbDefault
            End If

136         DataGrid1.SetFocus
          
138     ElseIf TxtBusqueda.Text = "" And OptionFiltro.Item(1).Value = True Then   'Esta buscando por una seleccion depto.
    
140         BBuscandoDeptosTodos = False
142         BBuscandoDeptosSeleccion = True
144         BBuscandoEmpleadosSeleccion = False
        
146         FrameBusqueda.Visible = True
                
            'Abre el Recordset y busca la seleccion del Depto
148         Set RBuscaSeleccionDeptos = New ADODB.Recordset
150         Call Abrir_Recordset(RBuscaSeleccionDeptos, "SELECT deptname FROM DEPARTMENTS ")
          
152         If RBuscaSeleccionDeptos.RecordCount > 0 Then
154             FrameBusqueda.Visible = True
156             Set DataGrid1.DataSource = RBuscaSeleccionDeptos
158             DataGrid1.Columns(0).Width = "2800"
            Else
160             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
162             Screen.MousePointer = vbDefault
            End If

164         DataGrid1.SetFocus
    
166     ElseIf TxtBusqueda.Text <> "" And OptionFiltro.Item(2).Value = True Then   'Esta buscando por una seleccion Empleado.
            
168         BBuscandoDeptosTodos = False
170         BBuscandoDeptosSeleccion = False
172         BBuscandoEmpleadosSeleccion = True
        
174         FrameBusqueda.Visible = True
        
            'Abre el Recordset y busca la seleccion del empleado.
176         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
178         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO WHERE street LIKE '%" & TxtBusqueda.Text & "%'")

180         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
182             FrameBusqueda.Visible = True
184             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
186             DataGrid1.Columns(0).Width = "700"
188             DataGrid1.Columns(1).Width = "2700"
            Else
190             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
192             Screen.MousePointer = vbDefault
            End If
          
194         DataGrid1.SetFocus
          
196     ElseIf TxtBusqueda.Text = "" And OptionFiltro.Item(2).Value = True Then   'Esta buscando por una seleccion Empleado.
    
198         BBuscandoDeptosTodos = False
200         BBuscandoDeptosSeleccion = False
202         BBuscandoEmpleadosSeleccion = True
            
            'Abre el Recordset y busca la seleccion del empleado.
204         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
206         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO ORDER BY badgenumber")

208         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
210             FrameBusqueda.Visible = True
212             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
214             DataGrid1.Columns(0).Width = "700"
216             DataGrid1.Columns(1).Width = "2700"
            Else
218             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
220             Screen.MousePointer = vbDefault
            End If
    
            'PARA LAS TARJETAS DE ASISTENCIAS (SELECCION X EMPLEADO)./////////////
222     ElseIf TxtBusqueda.Text <> "" And OptionFiltro.Item(3).Value = True Then   'Esta buscando por una seleccion Empleado.
            
224         BBuscandoDeptosTodos = False
226         BBuscandoDeptosSeleccion = False
228         BBuscandoEmpleadosSeleccion = True
        
230         FrameBusqueda.Visible = True
        
            'Abre el Recordset y busca la seleccion del empleado.
232         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
234         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO WHERE street LIKE '%" & TxtBusqueda.Text & "%'")

236         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
238             FrameBusqueda.Visible = True
240             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
242             DataGrid1.Columns(0).Width = "700"
244             DataGrid1.Columns(1).Width = "2700"
            Else
246             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
248             Screen.MousePointer = vbDefault
            End If
          
250         DataGrid1.SetFocus
    
            'PARA LAS TARJETAS DE ASISTENCIAS (SELECCION X EMPLEADO)./////////////
252     ElseIf TxtBusqueda.Text = "" And OptionFiltro.Item(3).Value = True Then   'Esta buscando por una seleccion Empleado.
    
254         BBuscandoDeptosTodos = False
256         BBuscandoDeptosSeleccion = False
258         BBuscandoEmpleadosSeleccion = True
            
            'Abre el Recordset y busca la seleccion del empleado.
260         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
262         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO ORDER BY badgenumber")

264         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
266             FrameBusqueda.Visible = True
268             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
270             DataGrid1.Columns(0).Width = "700"
272             DataGrid1.Columns(1).Width = "2700"
            Else
274             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
276             Screen.MousePointer = vbDefault
            End If
          
278         DataGrid1.SetFocus
    
        End If
    
280     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

TxtBusqueda_DblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.TxtBusqueda_DblClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub TxtBusqueda2_Change()
        '<EhHeader>
        On Error GoTo TxtBusqueda2_Change_Err
        '</EhHeader>

        'Buscando Todos los Deptos.
    
100     If BBuscandoDeptosTodos = True And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = False Then
            
            'imprime el reporte de todos los deptos.
102         FrameBusqueda.Visible = False
104         CmdImprimir.SetFocus
            
106     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = True And BBuscandoEmpleadosSeleccion = False Then
                
            'Abre el Recordset y busca la seleccion del Depto
108         Set RBuscaSeleccionDeptos = New ADODB.Recordset
110         Call Abrir_Recordset(RBuscaSeleccionDeptos, "SELECT deptname FROM DEPARTMENTS WHERE deptname LIKE '%" & TxtBusqueda2.Text & "%'")
          
112         If RBuscaSeleccionDeptos.RecordCount > 0 Then
114             FrameBusqueda.Visible = True
116             Set DataGrid1.DataSource = RBuscaSeleccionDeptos
118             DataGrid1.Columns(0).Width = "2800"
            Else
120             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
122             Screen.MousePointer = vbDefault
            End If
            
124     ElseIf BBuscandoDeptosTodos = False And BBuscandoDeptosSeleccion = False And BBuscandoEmpleadosSeleccion = True Then
            
            'Abre el Recordset y busca la seleccion del empleado.
126         Set RBuscaSeleccionEmpleado = New ADODB.Recordset
128         Call Abrir_Recordset(RBuscaSeleccionEmpleado, "SELECT badgenumber, street FROM USERINFO WHERE street LIKE '%" & TxtBusqueda2.Text & "%'")

130         If RBuscaSeleccionEmpleado.RecordCount > 0 Then
132             FrameBusqueda.Visible = True
134             Set DataGrid1.DataSource = RBuscaSeleccionEmpleado
136             DataGrid1.Columns(0).Width = "700"
138             DataGrid1.Columns(1).Width = "2700"
            Else
140             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
142             Screen.MousePointer = vbDefault
            End If
            
        End If
        
144     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

TxtBusqueda2_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.TxtBusqueda2_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CmdSalir_Click()
        '<EhHeader>
        On Error GoTo CmdSalir_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

CmdSalir_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.CmdSalir_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSalirBuscar2_Click()
        '<EhHeader>
        On Error GoTo cmdSalirBuscar2_Click_Err
        '</EhHeader>
100     FrameBusqueda.Visible = False
        '<EhFooter>
        Exit Sub

cmdSalirBuscar2_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.cmdSalirBuscar2_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Imprime_Deptos
' Descripción   : IMPRIME TODOS LOS DEPTOS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:07:20
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Imprime_Deptos()
        '<EhHeader>
        On Error GoTo Imprime_Deptos_Err
        '</EhHeader>
        
        BCuliacan = False
        BSLP = False
        BChiapas = False
        
        'Establece valores para las variables de fechas
100     VDia = Day(DTFechaInicio.Value)
102     VMes = Month(DTFechaInicio.Value)
104     VAño = Year(DTFechaInicio.Value)
106     VDia2 = Day(DTFechaFinal.Value)
108     VMes2 = Month(DTFechaFinal.Value)
110     VAño2 = Year(DTFechaFinal.Value)
    
        'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
112     If Option1(0).Value = True Then     'Reloj Culiacan
114         vEquipo = "1"
            vEquipoCul = "8"
            BCuliacan = True
            BSLP = False
            BChiapas = False
            '         ElseIf Option2(1).Value = True Then     'Baño Culiacan
            '             vEquipo = "2"
            '             vIDEquipo = 3
            '         ElseIf Option2(2).Value = True Then     'Vestidores Culiacan
            '             vEquipo = "3"
            '             vIDEquipo = 4
116     ElseIf Option1(1).Value = True Then     'Reloj SLP
118         vEquipo = "5"
            BCuliacan = False
            BSLP = True
            BChiapas = False
120     ElseIf Option1(2).Value = True Then     'Reloj Chiapas
122         vEquipo = "6"
            BCuliacan = False
            BSLP = False
            BChiapas = True
        End If
    
        'Rango de FECHAS y todos los Departamentos seleccionado __________________________________________________________________
        
124     GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
                
        If BCuliacan = True Then    'Esta consultando todos en culiacan
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "' "
            GCriteriaReporte = GCriteriaReporte & " OR {CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipoCul & "' "
        ElseIf BSLP = True Then 'ESTA CONSULTANDO TODOS EN SLP
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "' "
        Else    'ESTA CONSULTANDO TODOS EN CHIAPAS
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "' "
        End If
        GNombreReporte = "ListadoxDeptos.rpt"
       
130     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

Imprime_Deptos_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.Imprime_Deptos " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Imprime_Seleccion_Deptos
' Descripción   : IMPRESION SELECTIVA POR DEPTOS.
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:07:44
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Imprime_Seleccion_Deptos()
        '<EhHeader>
        On Error GoTo Imprime_Seleccion_Deptos_Err
        '</EhHeader>
        
        BCuliacan = False
        BSLP = False
        BChiapas = False

'BORRA EL CONTENIDO DE LA VARIABLE GLOBAL "GCriteriaReporte"
GCriteriaReporte = ""
GTituloReporte = ""

        'Establece valores para las variables de fechas
100     VDia = Day(DTFechaInicio.Value)
102     VMes = Month(DTFechaInicio.Value)
104     VAño = Year(DTFechaInicio.Value)
106     VDia2 = Day(DTFechaFinal.Value)
108     VMes2 = Month(DTFechaFinal.Value)
110     VAño2 = Year(DTFechaFinal.Value)
    
        'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
112     If Option1(0).Value = True Then     'Reloj Culiacan
114         vEquipo = "1"
            vEquipoCul = "8"
            BCuliacan = True
            BSLP = False
            BChiapas = False
116     ElseIf Option1(1).Value = True Then     'Reloj SLP
118         vEquipo = "5"
            BCuliacan = False
            BSLP = True
            BChiapas = False
120     ElseIf Option1(2).Value = True Then     'Reloj Chiapas
122         vEquipo = "6"
            BCuliacan = False
            BSLP = False
            BChiapas = True
        End If
    
        
        If TxtBusqueda.Text <> "" Then
        
        
            'Rango de FECHAS y todos los Departamentos seleccionado __________________________________________________________________
            If TxtBusqueda.Text <> "Administración" Then
                 'Rango de FECHAS y todos los Departamentos seleccionado __________________________________________________________________
                If TxtBusqueda.Text <> "Ventas Culiacan" Then
                    'LA CONSULTA NO ES ADMINISTRACION NI VENTAS - ES DIFERENTE A ESOS DEPTOS.
                    GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
                    GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "'"
                Else
                    'LA CONSULTA ES VENTAS CULIACAN
                    GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
                    GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "' "
                    GCriteriaReporte = GCriteriaReporte & " OR {CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} LIKE '" & vEquipoCul & "' "
                End If
            Else
                ' LA CONSULTA ES ADMINISTRACION
                GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
                If BCuliacan = True Then    'ES CULIACAN
                    GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} <> '5' "
                ElseIf BSLP = True Then
                    ' ES CONSULTA SLP
                    GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} = '5' "
                Else
                    ' ES CONSULTA CHIAPAS
                    GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    GCriteriaReporte = GCriteriaReporte & " And {DEPARTMENTS.DEPTNAME} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} = '6' "
                End If
            End If
            
        Else
        
            MsgBox "No eligió depto., vuelva a intentar ", vbCritical, "No eligió depto."
            Screen.MousePointer = vbDefault
        
        End If
         
        
130     GNombreReporte = "ListadoxSeleccionDeptos.rpt"
        'GNombreReporte = "UbicacionesCuadricula.rpt"
        
132     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

Imprime_Seleccion_Deptos_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.Imprime_Seleccion_Deptos " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Imprime_Seleccion_Empleado
' Descripción   : IMPRESION SELECTIVA POR EMPLEADO
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-17:08:00
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Imprime_Seleccion_Empleado()
        '<EhHeader>
        On Error GoTo Imprime_Seleccion_Empleado_Err
        '</EhHeader>
        
        BCuliacan = False
        BSLP = False
        BChiapas = False

    
        'Establece valores para las variables de fechas
100     VDia = Day(DTFechaInicio.Value)
102     VMes = Month(DTFechaInicio.Value)
104     VAño = Year(DTFechaInicio.Value)
106     VDia2 = Day(DTFechaFinal.Value)
108     VMes2 = Month(DTFechaFinal.Value)
110     VAño2 = Year(DTFechaFinal.Value)
    
        'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
112     If Option1(0).Value = True Then     'Reloj Culiacan
114         vEquipo = "1"
            vEquipoCul = "8"
            BCuliacan = True
            BSLP = False
            BChiapas = False
116     ElseIf Option1(1).Value = True Then     'Reloj SLP
118         vEquipo = "5"
            BCuliacan = False
            BSLP = True
            BChiapas = False
120     ElseIf Option1(2).Value = True Then     'Reloj Chiapas
122         vEquipo = "6"
            BCuliacan = False
            BSLP = False
            BChiapas = True
        End If
    
        'IMPRESION DE UN SOLO EMPLEADO  __________________________________________________________________
    
124     GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
126
        If BCuliacan = True Then    'ESTA CONSULTANDO EMPLEADO DE CULIACAN
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            If TxtBusqueda.Text <> "Aida Cecilia Santoyo Amaral" Then
                GCriteriaReporte = GCriteriaReporte & " And {USERINFO.STREET} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} <> '5' "
            Else
                GCriteriaReporte = GCriteriaReporte & " And {USERINFO.Badgenumber} = '604' AND {CHECKINOUT.SENSORID} <> '5' "
            End If
        ElseIf BSLP = True Then 'ESTA CONSULTANDO EMPLEADO DE SLP
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            GCriteriaReporte = GCriteriaReporte & " And {USERINFO.STREET} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} LIKE '5' "
        Else    'ESTA CONSULTANDO EMPLEADO DE CHIAPAS
            GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            GCriteriaReporte = GCriteriaReporte & " And {USERINFO.STREET} LIKE '" & TxtBusqueda.Text & "' AND {CHECKINOUT.SENSORID} LIKE '6' "
        End If
        
        
        'SECCION DE IMPRESION DE TARJETAS DE ASISTENCIA ==============================================================================
        
        If OptionFiltro.Item(3).Value = True Then  'TARJETA DE ASISTENCIA POR SELECCION DE EMPLEADOS. ////
        
132         GNombreReporte = "TarjetaAsistenciaxEmpleado.rpt"
        
134     ElseIf OptionFiltro.Item(4).Value = True Then  'TARJETA DE ASISTENCIA TODOS LOS EMPLEADOS. ////
            
136         GTituloReporte = "Desde " & Format(DTFechaInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(DTFechaFinal.Value, "dd/mm/yyyy")
            If BCuliacan = True Then    'ES IMP TARJETA DE ASISTENCIA CULIACAN -TODOS-
                GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipo & "' "
                GCriteriaReporte = GCriteriaReporte & " OR {CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '" & vEquipoCul & "' "
            ElseIf BSLP = True Then 'ES IMP TARJETA DE ASISTENCIA SLP -TODOS-
                GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '5' "
            Else        'ES IMP TARJETA DE ASISTENCIA CHIAPAS -TODOS-
                GCriteriaReporte = "{CHECKINOUT.CHECKTIME} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") AND {CHECKINOUT.SENSORID} LIKE '6' "
            End If

140         GNombreReporte = "TarjetaAsistenciaxEmpleado.rpt"
        
        Else    'NO SON TARJETAS DE ASISTENCIAS.  //////////
        
142         GNombreReporte = "ListadoxSeleccionEmpleado.rpt"
        
        End If

        'GNombreReporte = "UbicacionesCuadricula.rpt"
    
144     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

Imprime_Seleccion_Empleado_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.Imprime_Seleccion_Empleado " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ReportesRelojHuella.FrmMenuReportes.Form_Unload " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

