VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReportesReloj 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes Reloj Checador Huella"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9345
   Icon            =   "FrmReportesReloj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda"
      Height          =   9375
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   8775
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   8535
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   15055
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
   End
   Begin KewlButtonz.KewlButtons Boton 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReportesReloj.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12640511
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
   Begin VB.Frame Frame2 
      Caption         =   "Criterio de Busqueda:"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
      Begin VB.OptionButton Option1 
         Caption         =   "Empleado"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Departamento"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fecha:"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   8775
      Begin MSComCtl2.DTPicker DTFinal 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   39755
      End
      Begin MSComCtl2.DTPicker DTInicio 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   39755
         MinDate         =   39448
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin:"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin KewlButtonz.KewlButtons Boton 
      Height          =   855
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   4
      TX              =   "Otra Consulta"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReportesReloj.frx":0326
      PICN            =   "FrmReportesReloj.frx":0342
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Boton 
      Height          =   855
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   4
      TX              =   "Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReportesReloj.frx":EC64
      PICN            =   "FrmReportesReloj.frx":EC80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Boton 
      Height          =   855
      Index           =   3
      Left            =   6360
      TabIndex        =   14
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   4
      TX              =   "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReportesReloj.frx":16182
      PICN            =   "FrmReportesReloj.frx":1619E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      TabIndex        =   15
      Top             =   7440
      Width           =   8775
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   855
         Left            =   6360
         TabIndex        =   25
         Top             =   240
         Width           =   975
         Begin VB.OptionButton OptionDia 
            Caption         =   "Noche"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton OptionDia 
            Caption         =   "Dia"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin KewlButtonz.KewlButtons cmdCalculaTotHrs 
         Height          =   855
         Left            =   7440
         TabIndex        =   24
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BTYPE           =   4
         TX              =   "Calcula Total Hrs."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmReportesReloj.frx":24AC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Horas Totales (Según rango de fechas seleccionado)"
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
         Left            =   1200
         TabIndex        =   23
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Horas/Dia/Turno"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   3960
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Busqueda:"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   8775
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "# Equipo"
      Height          =   855
      Left            =   3840
      TabIndex        =   28
      Top             =   1920
      Width           =   5175
      Begin VB.OptionButton Option2 
         Caption         =   "Reloj Chiapas"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Reloj SLP"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Reloj Culiacan"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reporte Semanal Checadas Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "FrmReportesReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GvFechaHoy          As String

Dim Msg                 As String

Dim RBuscaDepto         As New ADODB.Recordset

Dim RBuscaEmpleado      As New ADODB.Recordset

Dim RBuscaNada          As New ADODB.Recordset

Dim RBuscaUltimaChecada As New ADODB.Recordset

Dim RBuscaTodo          As New ADODB.Recordset

Dim RBuscaChecadaNoche  As New ADODB.Recordset

Dim BBuscaEmpleado As Boolean

Dim BBuscaDepto As Boolean
Dim BEquipoCulican As Boolean

Dim vEquipo             As String
Dim vIDEquipo           As Integer

Dim vEquipoCul             As String
Dim vIDEquipoCul           As Integer

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Form_Load
' Descripción   : CARGA EL FORM PRINCIPAL
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:54:28
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     FrameBusqueda.Visible = False
102     TxtBusqueda.Text = ""

104     Formato_Fechas
106     Criterio_Busqueda
108     OptionDia(0).Value = True
110     Option2(0).Value = True
112     vEquipo = "1"
    
        'Inicia el proceso de conexion a la BD del reloj checador de huella.
113     GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=\\Respaldosfepsa\relojchecadorhuella\fepsa2.mdb; Jet OLEDB:Database Password="
114     'GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\Desarrollo\ReportesRelojChecador\fepsa2.mdb; Jet OLEDB:Database Password="
    
        'Casa
        'GConectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=c:\fepsa2.mdb; Jet OLEDB:Database Password="
    
        '///////////////////////////////////////   POR SI YA ESTA EJECUTANDOSE EL APLICATIVO
116     If App.PrevInstance Then
118         Msg = App.EXEName & ".EXE" & " ya está en ejecución y no debe haber dos sesiones abiertas"
120         MsgBox Msg, 16, "Aplicación."

122         End

        End If
     
        'INICIALIZA O CREA LA INSTANCIA DE LA CONECCION
124     Set Conexion = New ADODB.Connection
126     Conexion.ConnectionString = GConectionString
128     Conexion.Open
  
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Form_Load " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Formato_Fechas
' Descripción   : DA FORMATO A LAS VARIABLES QUE UTILIZO PARA EL MANEJO DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:55:07
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Formato_Fechas()

        '<EhHeader>
        On Error GoTo Formato_Fechas_Err

        '</EhHeader>

        'RANGO DE FECHAS  //////////////////
        'Formatos de Fechas
100     GvFechaHoy = ""
102     GvFechaHoy = Date
104     DTInicio.Value = GvFechaHoy
106     DTFinal.Value = GvFechaHoy
 
        '<EhFooter>
        Exit Sub

Formato_Fechas_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Formato_Fechas " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Criterio_Busqueda
' Descripción   : DEFINO POR DEFAULT LOS CRITERIOS DE BUSQUEDAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:55:42
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Criterio_Busqueda()

        '<EhHeader>
        On Error GoTo Criterio_Busqueda_Err

        '</EhHeader>

        'CRITERIO DE BUSQUEDA  //////////
        'Por default es por Depto.
100     Option1(0).Value = True
102     Option1(1).Value = False

        '<EhFooter>
        Exit Sub

Criterio_Busqueda_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Criterio_Busqueda " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Boton_Click
' Descripción   : EJECUTAR LA BUSQUEDA DE REGISTROS -CONSULTA-
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:56:38
'
' Parámetros    : Index (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Boton_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo Boton_Click_Err

        '</EhHeader>
    
        Dim vFechaFinal As Date

100     vFechaFinal = DTFinal.Value
102     vFechaFinal = vFechaFinal + 1

104     If Index = 0 Then   'Ejecutar
    
106         Screen.MousePointer = vbHourglass
    
            'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
108         If Option2(0).Value = True Then     'Reloj Culiacan
110             vEquipo = "1"
                vEquipoCul = "8"
112             vIDEquipo = 2
                vIDEquipoCul = 8

                '114         ElseIf Option2(1).Value = True Then     'Baño Culiacan
                '116             vEquipo = "2"
                '118             vIDEquipo = 3
                '120         ElseIf Option2(2).Value = True Then     'Vestidores Culiacan
                '122             vEquipo = "3"
                '124             vIDEquipo = 4
126         ElseIf Option2(3).Value = True Then     'Reloj SLP
128             vEquipo = "5"
130             vIDEquipo = 5
132         ElseIf Option2(4).Value = True Then     'Reloj Chiapas
134             vEquipo = "6"
136             vIDEquipo = 6
            End If
    
138         If Option1(0).Value = True And TxtBusqueda.Text = "" Then 'Busqueda Completa x depto.
            
140             Set RBuscaDepto = New ADODB.Recordset
142             'Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.sensorid LIKE M.MachineNumber " & "AND registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY badgenumber, checktime, deptname ASC")
                Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                    "M.ID = " & vIDEquipo & " " & "AND " & _
                    "registros.sensorid LIKE M.MachineNumber " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "UNION " & _
                    "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                    "M.ID = " & vIDEquipoCul & " " & "AND " & _
                    "registros.sensorid LIKE M.MachineNumber " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "ORDER BY badgenumber, checktime, deptname ASC")
            
144             If RBuscaDepto.RecordCount > 0 Then 'si hay registros.
                
146                 Set DataGrid1.DataSource = RBuscaDepto
                    
148                 DataGrid1.Columns(0).Alignment = dbgLeft
                    'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
150                 DataGrid1.Columns(0).Width = "2000"
152                 DataGrid1.Columns(0).Caption = "Fecha y Hora"
                    
154                 DataGrid1.Columns(1).Alignment = dbgLeft
                    'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
156                 DataGrid1.Columns(1).Width = "1000"
158                 DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                    
160                 DataGrid1.Columns(2).Alignment = dbgLeft
                    'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
162                 DataGrid1.Columns(2).Width = "700"
164                 DataGrid1.Columns(2).Caption = "Cód.Emp"
                    
166                 DataGrid1.Columns(3).Alignment = dbgLeft
                    'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
168                 DataGrid1.Columns(3).Width = "2500"
170                 DataGrid1.Columns(3).Caption = "Nombre"
                    
172                 DataGrid1.Columns(4).Alignment = dbgLeft
                    'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
174                 DataGrid1.Columns(4).Width = "2000"
176                 DataGrid1.Columns(4).Caption = "Depto"
                    
                    'Saca_SumaMinutos
                    
                Else
178                 Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")
180                 TxtBusqueda.Text = ""

182                 Formato_Fechas
184                 Criterio_Busqueda
186                 DTInicio.SetFocus
                    
                    'CIERRA EL RECORDSET xDepto  ///////////////////////////////
188                 Screen.MousePointer = 11
                    'RBuscaDepto.Close
                    'Set RBuscaDepto = Nothing
190                 Screen.MousePointer = 0
                    
                End If
        
                '///////BUSQUEDA SELECTIVA X DEPTOS.
192         ElseIf Option1(0).Value = True And TxtBusqueda.Text <> "" Then 'Busqueda Selectiva x depto.
                
194             Set RBuscaDepto = New ADODB.Recordset
196             'Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND deptos.deptname = '" & TxtBusqueda.Text & "' " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY badgenumber, checktime, deptname ASC")
                Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                    "M.ID = " & vIDEquipo & " " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "deptos.deptname = '" & TxtBusqueda.Text & "' " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "UNION " & _
                    "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                    "M.ID = " & vIDEquipoCul & " " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "deptos.deptname = '" & TxtBusqueda.Text & "' " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "ORDER BY badgenumber, checktime, deptname ASC")
                                
                                
198             If RBuscaDepto.RecordCount > 0 Then 'si hay registros.
200                 Set DataGrid1.DataSource = RBuscaDepto
                        
202                 DataGrid1.Columns(0).Alignment = dbgLeft
                    'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
204                 DataGrid1.Columns(0).Width = "2000"
206                 DataGrid1.Columns(0).Caption = "Fecha y Hora"
                        
208                 DataGrid1.Columns(1).Alignment = dbgLeft
                    'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
210                 DataGrid1.Columns(1).Width = "1000"
212                 DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                        
214                 DataGrid1.Columns(2).Alignment = dbgLeft
                    'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
216                 DataGrid1.Columns(2).Width = "700"
218                 DataGrid1.Columns(2).Caption = "Cód.Emp"
                        
220                 DataGrid1.Columns(3).Alignment = dbgLeft
                    'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
222                 DataGrid1.Columns(3).Width = "2500"
224                 DataGrid1.Columns(3).Caption = "Nombre"
                        
226                 DataGrid1.Columns(4).Alignment = dbgLeft
                    'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
228                 DataGrid1.Columns(4).Width = "2000"
230                 DataGrid1.Columns(4).Caption = "Depto"
                        
                Else
232                 Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")
234                 TxtBusqueda.Text = ""

236                 Formato_Fechas
238                 Criterio_Busqueda
240                 DTInicio.SetFocus
                        
                    'CIERRA EL RECORDSET xDepto  ///////////////////////////////
242                 Screen.MousePointer = 11
                    'RBuscaDepto.Close
                    'Set RBuscaDepto = Nothing
244                 Screen.MousePointer = 0
                        
                End If
                    
                'POR EMPLEADO.   ///////////////////////////////////////////////////////////////
        
246         ElseIf Option1(0).Value = False And TxtBusqueda.Text = "" Then 'Busqueda Completa x Empleado.
            
248             Set RBuscaEmpleado = New ADODB.Recordset
250             Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                        "M.ID = " & vIDEquipo & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "UNION " & _
                        "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                        "M.ID = " & vIDEquipoCul & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "ORDER BY badgenumber, checktime, deptname ASC")
            
252             If RBuscaEmpleado.RecordCount > 0 Then 'si hay registros.
254                 Set DataGrid1.DataSource = RBuscaEmpleado
                    
256                 DataGrid1.Columns(0).Alignment = dbgLeft
                    'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
258                 DataGrid1.Columns(0).Width = "2000"
260                 DataGrid1.Columns(0).Caption = "Fecha y Hora"
                        
262                 DataGrid1.Columns(1).Alignment = dbgLeft
                    'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
264                 DataGrid1.Columns(1).Width = "1000"
266                 DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                        
268                 DataGrid1.Columns(2).Alignment = dbgLeft
                    'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
270                 DataGrid1.Columns(2).Width = "700"
272                 DataGrid1.Columns(2).Caption = "Cód.Emp"
                        
274                 DataGrid1.Columns(3).Alignment = dbgLeft
                    'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
276                 DataGrid1.Columns(3).Width = "2500"
278                 DataGrid1.Columns(3).Caption = "Nombre"
                        
280                 DataGrid1.Columns(4).Alignment = dbgLeft
                    'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
282                 DataGrid1.Columns(4).Width = "2000"
284                 DataGrid1.Columns(4).Caption = "Depto"
                    
                Else
286                 Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")
288                 TxtBusqueda.Text = ""

290                 Formato_Fechas
292                 Criterio_Busqueda
294                 DTInicio.SetFocus
                    
                    'CIERRA EL RECORDSET xDepto  ///////////////////////////////
296                 Screen.MousePointer = 11
                    'RBuscaEmpleado.Close
                    'Set RBuscaEmpleado = Nothing
298                 Screen.MousePointer = 0
                    
                End If
        
                '///////BUSQUEDA SELECTIVA X EMPLEADOS.
300         ElseIf Option1(0).Value = False And TxtBusqueda.Text <> "" Then 'Busqueda Selectiva x empleados.
                
302             Set RBuscaEmpleado = New ADODB.Recordset
304             Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                        "M.ID = " & vIDEquipo & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "usuarios.street = '" & TxtBusqueda.Text & "' " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "UNION " & _
                        "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                        "M.ID = " & vIDEquipoCul & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "usuarios.street = '" & TxtBusqueda.Text & "' " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "ORDER BY checktime, street, deptname ASC")
                
306             If RBuscaEmpleado.RecordCount > 0 Then 'si hay registros.
308                 Set DataGrid1.DataSource = RBuscaEmpleado
                        
310                 DataGrid1.Columns(0).Alignment = dbgLeft
                    'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
312                 DataGrid1.Columns(0).Width = "2000"
314                 DataGrid1.Columns(0).Caption = "Fecha y Hora"
                        
316                 DataGrid1.Columns(1).Alignment = dbgLeft
                    'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
318                 DataGrid1.Columns(1).Width = "1000"
320                 DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                        
322                 DataGrid1.Columns(2).Alignment = dbgLeft
                    'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
324                 DataGrid1.Columns(2).Width = "700"
326                 DataGrid1.Columns(2).Caption = "Cód.Emp"
                        
328                 DataGrid1.Columns(3).Alignment = dbgLeft
                    'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
330                 DataGrid1.Columns(3).Width = "2500"
332                 DataGrid1.Columns(3).Caption = "Nombre"
                        
334                 DataGrid1.Columns(4).Alignment = dbgLeft
                    'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
336                 DataGrid1.Columns(4).Width = "2000"
338                 DataGrid1.Columns(4).Caption = "Depto"
                        
                Else
340                 Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")
342                 TxtBusqueda.Text = ""

344                 Formato_Fechas
346                 Criterio_Busqueda
348                 DTInicio.SetFocus
                        
                    'CIERRA EL RECORDSET xDepto  ///////////////////////////////
350                 Screen.MousePointer = 11
                    'RBuscaEmpleado.Close
                    'Set RBuscaEmpleado = Nothing
352                 Screen.MousePointer = 0
                        
                End If
        
            End If  'para saber que tipo de busqueda hacer.
    
354     ElseIf Index = 1 Then 'Imprimir  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
356         FrmMenuReportes.Visible = True
    
358     ElseIf Index = 2 Then 'Limpiar Campos  ///////////////////////////////////////////////////////
360         TxtBusqueda.Text = ""

362         Formato_Fechas
364         Criterio_Busqueda
366         DTInicio.SetFocus
368         Label3.Caption = ""
370         Label4.Caption = ""
372         OptionDia(0).Value = True
374         Set DataGrid1.DataSource = RBuscaNada
        
376     ElseIf Index = 3 Then 'Salir  ////////////////////////////////////////////////////////////////
378         Unload Me
380         Close

382         End
        
        End If
    
384     Boton(0).BackColor = &H8000000F
386     Boton(1).BackColor = &H8000000F
388     Boton(2).BackColor = &H8000000F
390     Boton(3).BackColor = &H8000000F
    
392     Screen.MousePointer = vbDefault
 
        '<EhFooter>
        Exit Sub

Boton_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Boton_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTInicio_CallbackKeyDown
' Descripción   : LIMPIO LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:57:41
'
' Parámetros    : KeyCode (Integer)
'                 Shift (Integer)
'                 CallbackField (String)
'                 CallbackDate (Date)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTInicio_CallbackKeyDown(ByVal KeyCode As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal CallbackField As String, _
                                     CallbackDate As Date)

        '<EhHeader>
        On Error GoTo DTInicio_CallbackKeyDown_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTInicio_CallbackKeyDown_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTInicio_CallbackKeyDown " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTFinal_CallbackKeyDown
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:58:00
'
' Parámetros    : KeyCode (Integer)
'                 Shift (Integer)
'                 CallbackField (String)
'                 CallbackDate (Date)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTFinal_CallbackKeyDown(ByVal KeyCode As Integer, _
                                    ByVal Shift As Integer, _
                                    ByVal CallbackField As String, _
                                    CallbackDate As Date)

        '<EhHeader>
        On Error GoTo DTFinal_CallbackKeyDown_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTFinal_CallbackKeyDown_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTFinal_CallbackKeyDown " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTInicio_Click
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:58:37
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTInicio_Click()

        '<EhHeader>
        On Error GoTo DTInicio_Click_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTInicio_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTInicio_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTFinal_Click
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:58:48
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTFinal_Click()

        '<EhHeader>
        On Error GoTo DTFinal_Click_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTFinal_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTFinal_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTInicio_Change
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:59:02
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTInicio_Change()

        '<EhHeader>
        On Error GoTo DTInicio_Change_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTInicio_Change_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTInicio_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTFinal_Change
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:59:15
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTFinal_Change()

        '<EhHeader>
        On Error GoTo DTFinal_Change_Err

        '</EhHeader>

100     Label3.Caption = ""
102     Label4.Caption = ""

        '<EhFooter>
        Exit Sub

DTFinal_Change_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTFinal_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DTInicio_KeyPress
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:59:36
'
' Parámetros    : KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DTInicio_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DTInicio_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = 13 Then
102         Option1(0).SetFocus
        End If

        '<EhFooter>
        Exit Sub

DTInicio_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DTInicio_KeyPress " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Option1_KeyPress
' Descripción   : LIMPIO TAMBIEN LOS CAMPOS DE FECHAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-15:59:48
'
' Parámetros    : Index (Integer)
'                 KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Option1_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = 13 Then
102         TxtBusqueda.SetFocus
        End If

        '<EhFooter>
        Exit Sub

Option1_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Option1_KeyPress " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : TxtBusqueda_Change
' Descripción   : BUSQUEDA CUANDO CAMBIA VALOR EL TEXTS DE BUSQUEDAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:00:00
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub TxtBusqueda_Change()

        '<EhHeader>
        On Error GoTo TxtBusqueda_Change_Err

        '</EhHeader>

        Dim vFechaFinal As Date

100     vFechaFinal = DTFinal.Value
102     vFechaFinal = vFechaFinal + 1
            
104     Screen.MousePointer = vbHourglass
            
        'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
106     If Option2(0).Value = True Then     'Reloj Culiacan
108         vEquipo = "1"
110         vIDEquipo = 2
            '112     ElseIf Option2(1).Value = True Then     'Baño Culiacan
            '114         vEquipo = "2"
            '116         vIDEquipo = 3
            '118     ElseIf Option2(2).Value = True Then     'Vestidores Culiacan
            '120         vEquipo = "3"
            '122         vIDEquipo = 4
124     ElseIf Option2(3).Value = True Then     'Reloj SLP
126         vEquipo = "5"
128         vIDEquipo = 5
130     ElseIf Option2(4).Value = True Then     'Reloj Chiapas
132         vEquipo = "6"
134         vIDEquipo = 6
        End If
        
136     If Option1(0).Value = True And TxtBusqueda.Text = "" Then 'Busqueda Completa x depto.
            
138         Set RBuscaDepto = New ADODB.Recordset
140         'Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY badgenumber, checktime, deptname ASC")
            
            Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                    "M.ID = " & vIDEquipo & " " & "AND " & _
                    "registros.sensorid LIKE M.MachineNumber " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "UNION " & _
                    "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                    "M.ID = " & vIDEquipoCul & " " & "AND " & _
                    "registros.sensorid LIKE M.MachineNumber " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "ORDER BY badgenumber, checktime, deptname ASC")
            
142         If RBuscaDepto.RecordCount > 0 Then 'si hay registros.
                            
144             Set DataGrid1.DataSource = RBuscaDepto
                                
146             DataGrid1.Columns(0).Alignment = dbgLeft
                'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
148             DataGrid1.Columns(0).Width = "2000"
150             DataGrid1.Columns(0).Caption = "Fecha y Hora"
                                
152             DataGrid1.Columns(1).Alignment = dbgLeft
                'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
154             DataGrid1.Columns(1).Width = "1000"
156             DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                                
158             DataGrid1.Columns(2).Alignment = dbgLeft
                'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
160             DataGrid1.Columns(2).Width = "700"
162             DataGrid1.Columns(2).Caption = "Cód.Emp"
                                
164             DataGrid1.Columns(3).Alignment = dbgLeft
                'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
166             DataGrid1.Columns(3).Width = "2500"
168             DataGrid1.Columns(3).Caption = "Nombre"
                                
170             DataGrid1.Columns(4).Alignment = dbgLeft
                'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
172             DataGrid1.Columns(4).Width = "2000"
174             DataGrid1.Columns(4).Caption = "Depto"
                                
                'Saca_SumaMinutos
                                
            Else
176             Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")

                'TxtBusqueda.Text = ""
178             Formato_Fechas
180             Criterio_Busqueda
182             DTInicio.SetFocus
                                
                'CIERRA EL RECORDSET xDepto  ///////////////////////////////
184             Screen.MousePointer = 11
                'RBuscaDepto.Close
                'Set RBuscaDepto = Nothing
186             Screen.MousePointer = 0
                                
            End If
            
            '///////BUSQUEDA SELECTIVA X DEPTOS.
188     ElseIf Option1(0).Value = True And TxtBusqueda.Text <> "" Then 'Busqueda Selectiva x depto.
                    
190         Set RBuscaDepto = New ADODB.Recordset
192         'Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND deptos.deptname LIKE '%" & TxtBusqueda.Text & "%' " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY badgenumber, checktime, deptname ASC")
                    
            Call Abrir_Recordset(RBuscaDepto, "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                    "M.ID = " & vIDEquipo & " " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "deptos.deptname = '" & TxtBusqueda.Text & "' " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "UNION " & _
                    "SELECT registros.checktime, registros.checktype, " & _
                    "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                    "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                    "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                    "M.ID = " & vIDEquipoCul & " " & "AND " & _
                    "registros.USERID = usuarios.userid " & "AND " & _
                    "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                    "deptos.deptname = '" & TxtBusqueda.Text & "' " & "AND " & _
                    "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                    "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                    "ORDER BY badgenumber, checktime, deptname ASC")
                    
194         If RBuscaDepto.RecordCount > 0 Then 'si hay registros.
196             Set DataGrid1.DataSource = RBuscaDepto
                            
198             DataGrid1.Columns(0).Alignment = dbgLeft
                'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
200             DataGrid1.Columns(0).Width = "2000"
202             DataGrid1.Columns(0).Caption = "Fecha y Hora"
                            
204             DataGrid1.Columns(1).Alignment = dbgLeft
                'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
206             DataGrid1.Columns(1).Width = "1000"
208             DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                            
210             DataGrid1.Columns(2).Alignment = dbgLeft
                'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
212             DataGrid1.Columns(2).Width = "700"
214             DataGrid1.Columns(2).Caption = "Cód.Emp"
                            
216             DataGrid1.Columns(3).Alignment = dbgLeft
                'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
218             DataGrid1.Columns(3).Width = "2500"
220             DataGrid1.Columns(3).Caption = "Nombre"
                            
222             DataGrid1.Columns(4).Alignment = dbgLeft
                'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
224             DataGrid1.Columns(4).Width = "2000"
226             DataGrid1.Columns(4).Caption = "Depto"
                            
            Else
                        
228             Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")

                'TxtBusqueda.Text = ""
230             Formato_Fechas
232             Criterio_Busqueda
234             DTInicio.SetFocus
                            
                'CIERRA EL RECORDSET xDepto  ///////////////////////////////
236             Screen.MousePointer = 11
                'RBuscaDepto.Close
                'Set RBuscaDepto = Nothing
238             Screen.MousePointer = 0
                             
            End If
                        
            'POR EMPLEADO.   ///////////////////////////////////////////////////////////////
            
240     ElseIf Option1(0).Value = False And TxtBusqueda.Text = "" Then 'Busqueda Completa x Empleado.
                    
242         Set RBuscaEmpleado = New ADODB.Recordset
244         'Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND usuarios.street LIKE '%" & TxtBusqueda.Text & "%' " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY badgenumber, checktime, deptname ASC")
                        
             Set RBuscaEmpleado = New ADODB.Recordset
             Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                        "M.ID = " & vIDEquipo & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "UNION " & _
                        "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                        "M.ID = " & vIDEquipoCul & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "ORDER BY badgenumber, checktime, deptname ASC")
                        
                        
246         If RBuscaEmpleado.RecordCount > 0 Then 'si hay registros.
             Set DataGrid1.DataSource = RBuscaEmpleado
                                
250             DataGrid1.Columns(0).Alignment = dbgLeft
                'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
252             DataGrid1.Columns(0).Width = "2000"
254             DataGrid1.Columns(0).Caption = "Fecha y Hora"
                                    
256             DataGrid1.Columns(1).Alignment = dbgLeft
                'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
258             DataGrid1.Columns(1).Width = "1000"
260             DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                                    
262             DataGrid1.Columns(2).Alignment = dbgLeft
                'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
264             DataGrid1.Columns(2).Width = "700"
266             DataGrid1.Columns(2).Caption = "Cód.Emp"
                                    
268             DataGrid1.Columns(3).Alignment = dbgLeft
                'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
270             DataGrid1.Columns(3).Width = "2500"
272             DataGrid1.Columns(3).Caption = "Nombre"
                                    
274             DataGrid1.Columns(4).Alignment = dbgLeft
                'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
276             DataGrid1.Columns(4).Width = "2000"
278             DataGrid1.Columns(4).Caption = "Depto"
                                
            Else
280             Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")

                'TxtBusqueda.Text = ""
282             Formato_Fechas
284             Criterio_Busqueda
286             DTInicio.SetFocus
                                
                'CIERRA EL RECORDSET xDepto  ///////////////////////////////
288             Screen.MousePointer = 11
                'RBuscaEmpleado.Close
                'Set RBuscaEmpleado = Nothing
290             Screen.MousePointer = 0
                                
            End If
                            
            '///////BUSQUEDA SELECTIVA X EMPLEADOS.
292     ElseIf Option1(0).Value = False And TxtBusqueda.Text <> "" Then 'Busqueda Selectiva x empleados.
                    
294         Set RBuscaEmpleado = New ADODB.Recordset
296         'Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.defaultdeptid = deptos.deptid " & "AND usuarios.street LIKE '%" & TxtBusqueda.Text & "%' " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime, street, deptname ASC")
                    
         Call Abrir_Recordset(RBuscaEmpleado, "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND " & _
                        "M.ID = " & vIDEquipo & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "usuarios.street = '" & TxtBusqueda.Text & "' " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "UNION " & _
                        "SELECT registros.checktime, registros.checktype, " & _
                        "usuarios.badgenumber, " & "usuarios.street, deptos.deptname, registros.sensorid " & " " & _
                        "FROM CHECKINOUT registros, USERINFO usuarios, DEPARTMENTS deptos, Machines M " & " " & _
                        "WHERE registros.sensorid LIKE " & "'" & vEquipoCul & "' " & "AND " & _
                        "M.ID = " & vIDEquipoCul & " " & "AND " & _
                        "registros.USERID = usuarios.userid " & "AND " & _
                        "usuarios.defaultdeptid = deptos.deptid " & "AND " & _
                        "usuarios.street = '" & TxtBusqueda.Text & "' " & "AND " & _
                        "registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND " & _
                        "registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & " " & _
                        "ORDER BY checktime, street, deptname ASC")
                    
298         If RBuscaEmpleado.RecordCount > 0 Then 'si hay registros.
300             Set DataGrid1.DataSource = RBuscaEmpleado
                            
302             DataGrid1.Columns(0).Alignment = dbgLeft
                'DataGrid1.Columns(0).NumberFormat = ("#,###,##0.#0")
304             DataGrid1.Columns(0).Width = "2000"
306             DataGrid1.Columns(0).Caption = "Fecha y Hora"
                            
308             DataGrid1.Columns(1).Alignment = dbgLeft
                'DataGrid1.Columns(1).NumberFormat = ("#,###,##0.#0")
310             DataGrid1.Columns(1).Width = "1000"
312             DataGrid1.Columns(1).Caption = "I=Ent/O=Sal"
                            
314             DataGrid1.Columns(2).Alignment = dbgLeft
                'DataGrid1.Columns(2).NumberFormat = ("#,###,##0.#0")
316             DataGrid1.Columns(2).Width = "700"
318             DataGrid1.Columns(2).Caption = "Cód.Emp"
                            
320             DataGrid1.Columns(3).Alignment = dbgLeft
                'DataGrid1.Columns(3).NumberFormat = ("#,###,##0.#0")
322             DataGrid1.Columns(3).Width = "2500"
324             DataGrid1.Columns(3).Caption = "Nombre"
                            
326             DataGrid1.Columns(4).Alignment = dbgLeft
                'DataGrid1.Columns(4).NumberFormat = ("#,###,##0.#0")
328             DataGrid1.Columns(4).Width = "2000"
330             DataGrid1.Columns(4).Caption = "Depto"
                        
            Else
                        
332             Call MsgBox("No se encontraron registros en esta busqueda. Vuelva a intentarlo.", vbCritical Or vbSystemModal, "No hay información")

                'TxtBusqueda.Text = ""
334             Formato_Fechas
336             Criterio_Busqueda
338             DTInicio.SetFocus
                                
                'CIERRA EL RECORDSET xDepto  ///////////////////////////////
340             Screen.MousePointer = 11
                'RBuscaEmpleado.Close
                'Set RBuscaEmpleado = Nothing
342             Screen.MousePointer = 0
                            
            End If
            
        End If  'para saber que tipo de busqueda hacer.
            
344     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

TxtBusqueda_Change_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.TxtBusqueda_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : TxtBusqueda_DblClick
' Descripción   : BUSQUEDA CON SOLO PRESIONAR DOBLE CLICK DEL MOUSE
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:01:05
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub TxtBusqueda_DblClick()

        '<EhHeader>
        On Error GoTo TxtBusqueda_DblClick_Err

        '</EhHeader>
          
100     If Option1(0).Value = True And TxtBusqueda.Text = "" Then  'xDepto
                      
102         BBuscaEmpleado = False
104         BBuscaDepto = True
            'Abre el Recordset y busca el empleado
106         Set RBuscaDepto = New ADODB.Recordset
108         Call Abrir_Recordset(RBuscaDepto, "SELECT deptname FROM DEPARTMENTS")
          
110         If RBuscaDepto.RecordCount > 0 Then
                    
112             FrameBusqueda.Visible = True
114             Set DataGrid2.DataSource = RBuscaDepto
116             DataGrid2.Columns(0).Width = "2800"
                                  
            Else
118             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
            End If

120         DataGrid2.SetFocus
              
        Else 'Esta buscando Empleado  ////////
      
122         BBuscaEmpleado = True
124         BBuscaDepto = False
            'Abre el Recordset y busca el Depto
126         Set RBuscaEmpleado = New ADODB.Recordset
128         Call Abrir_Recordset(RBuscaEmpleado, "SELECT badgenumber, street FROM USERINFO ORDER BY badgenumber")
          
130         If RBuscaEmpleado.RecordCount > 0 Then
                    
132             FrameBusqueda.Visible = True
134             Set DataGrid2.DataSource = RBuscaEmpleado
136             DataGrid2.Columns(0).Width = "700"
138             DataGrid2.Columns(1).Width = "2700"
                    
            Else
140             MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
            End If

142         DataGrid2.SetFocus
                  
        End If  'Fin de Bandera de Depto y Puesto

144     Screen.MousePointer = vbDefault
  
        '<EhFooter>
        Exit Sub

TxtBusqueda_DblClick_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.TxtBusqueda_DblClick " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : TxtBusqueda_KeyPress
' Descripción   : BUSQUEDA CON CUALQUIER CAMBIO QUE REGISTRE EL VALOR DEL TEXT
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:01:42
'
' Parámetros    : KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo TxtBusqueda_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = 13 And TxtBusqueda.Text = "" Or TxtBusqueda.Text = " " Then
102         Boton(0).SetFocus
        
104         If Option1(0).Value = True And TxtBusqueda.Text <> "" Then  'xDepto
                            
106             BBuscaEmpleado = False
108             BBuscaDepto = True
                'Abre el Recordset y busca el empleado
110             Set RBuscaDepto = New ADODB.Recordset
112             Call Abrir_Recordset(RBuscaDepto, "SELECT * FROM DEPARTMENS")
                
114             If RBuscaDepto.RecordCount > 0 Then
                          
116                 FrameBusqueda.Visible = True
118                 Set DataGrid2.DataSource = RBuscaDepto
120                 DataGrid2.Columns(0).Width = "2800"
                    
                Else
122                 MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
                End If

124             DataGrid2.SetFocus
                    
            Else 'Esta buscando Empleado  ////////
            
126             BBuscaEmpleado = True
128             BBuscaDepto = False
                'Abre el Recordset y busca el Depto
130             Set RBuscaEmpleado = New ADODB.Recordset
132             Call Abrir_Recordset(RBuscaEmpleado, "SELECT badgenumber, street FROM USERINFO ORDER BY badgenumber")
                
134             If RBuscaEmpleado.RecordCount > 0 Then
                          
136                 FrameBusqueda.Visible = True
138                 Set DataGrid2.DataSource = RBuscaEmpleado
140                 DataGrid2.Columns(0).Width = "700"
142                 DataGrid2.Columns(1).Width = "2700"
                          
                Else
144                 MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
                End If

146             DataGrid2.SetFocus
                        
            End If  'Fin de Bandera de Depto y Puesto
        
            '//'/////////////////////////////////////////////    'Busca Codigo Procesos   +   /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
148     ElseIf KeyAscii = 43 Then
    
150         If Option1(0).Value = True And TxtBusqueda.Text = "" Then  'xDepto
                            
152             BBuscaEmpleado = False
154             BBuscaDepto = True
                'Abre el Recordset y busca el empleado
156             Set RBuscaDepto = New ADODB.Recordset
158             Call Abrir_Recordset(RBuscaDepto, "SELECT * FROM DEPARTMENS")
                
160             If RBuscaDepto.RecordCount > 0 Then
                          
162                 FrameBusqueda.Visible = True
164                 Set DataGrid2.DataSource = RBuscaDepto
166                 DataGrid2.Columns(0).Width = "2800"
                    
                Else
168                 MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
                End If

170             DataGrid2.SetFocus
                    
            Else 'Esta buscando Empleado  ////////
            
172             BBuscaEmpleado = True
174             BBuscaDepto = False
                'Abre el Recordset y busca el Depto
176             Set RBuscaEmpleado = New ADODB.Recordset
178             Call Abrir_Recordset(RBuscaEmpleado, "SELECT badgenumber, street FROM USERINFO ORDER BY badgenumber")
                
180             If RBuscaEmpleado.RecordCount > 0 Then
                          
182                 FrameBusqueda.Visible = True
184                 Set DataGrid2.DataSource = RBuscaEmpleado
186                 DataGrid2.Columns(0).Width = "700"
188                 DataGrid2.Columns(1).Width = "2700"
                          
                Else
190                 MsgBox "No puedo encontrarlo, vuelva a intentar ", vbCritical, "No se encontró"
                End If

192             DataGrid2.SetFocus
                        
            End If  'Fin de Bandera de Depto y Puesto
    
        End If

194     Screen.MousePointer = vbDefault

        '<EhFooter>
        Exit Sub

TxtBusqueda_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.TxtBusqueda_KeyPress " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid2_KeyPress
' Descripción   : CUANDO PRESIONA UNA TECLA DENTRO DE ESTE DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:03:10
'
' Parámetros    : KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid2_KeyPress_Err

        '</EhHeader>

100     If KeyAscii = 13 Then
    
102         FrameBusqueda.Visible = False
            
104         If BBuscaDepto = True Then
        
106             TxtBusqueda.Text = DataGrid2.Columns(0).Text
108             Boton(0).SetFocus
                
            Else
            
110             TxtBusqueda.Text = DataGrid2.Columns(1).Text
            
            End If
        
        End If
    
112     Boton(0).SetFocus

        '<EhFooter>
        Exit Sub

DataGrid2_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DataGrid2_KeyPress " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid2_Click
' Descripción   : CUANDO DOY UN CLICK DENTRO DEL DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:03:30
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid2_Click()

        '<EhHeader>
        On Error GoTo DataGrid2_Click_Err

        '</EhHeader>

100     FrameBusqueda.Visible = False

102     If BBuscaDepto = True Then
        
104         TxtBusqueda.Text = DataGrid2.Columns(0).Text
106         Boton(0).SetFocus
        
        Else
    
108         TxtBusqueda.Text = DataGrid2.Columns(1).Text
110         Boton(0).SetFocus
    
        End If

        '<EhFooter>
        Exit Sub

DataGrid2_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DataGrid2_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : DataGrid2_DblClick
' Descripción   : CUANDO PRESIONO DOBLE CLICK DENTRO DEL DATAGRID
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:04:06
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub DataGrid2_DblClick()

        '<EhHeader>
        On Error GoTo DataGrid2_DblClick_Err

        '</EhHeader>

100     FrameBusqueda.Visible = False

102     If BBuscaDepto = True Then
        
104         TxtBusqueda.Text = DataGrid2.Columns(0).Text
106         Boton(0).SetFocus
        
        Else
    
108         TxtBusqueda.Text = DataGrid2.Columns(1).Text
110         Boton(0).SetFocus
    
        End If

        '<EhFooter>
        Exit Sub

DataGrid2_DblClick_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.DataGrid2_DblClick " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Boton_GotFocus
' Descripción   : CUANDO OBTENGO EL ENFOQUE DE ESTE BOTON
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:04:29
'
' Parámetros    : Index (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Boton_GotFocus(Index As Integer)

        '<EhHeader>
        On Error GoTo Boton_GotFocus_Err

        '</EhHeader>

100     If Index = 0 Then
102         Boton(0).BackColor = &H80FF80
104     ElseIf Index = 1 Then
106         Boton(0).BackColor = &H80FF80
        End If
 
        '<EhFooter>
        Exit Sub

Boton_GotFocus_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Boton_GotFocus " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Form_Unload
' Descripción   : DESCARGO EL FORM
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:04:50
'
' Parámetros    : Cancel (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>

100     Close

102     End

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Form_Unload " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Saca_SumaMinutos
' Descripción   : PROCESO PARA SACAR LOS MINUTOS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:05:12
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Saca_SumaMinutos()

        '<EhHeader>
        On Error GoTo Saca_SumaMinutos_Err

        '</EhHeader>

100     Screen.MousePointer = vbHourglass

        Dim vh1 As Variant

        Dim vh2 As Variant

102     vh1 = 0
104     vh2 = 0

        Dim vSumaHoras1  As Variant

        Dim vFechaFinal  As Date

        Dim vCodEmpleado As Integer

        Dim vHoraInicial As Date
    
        'PARA SACAR EL NUMERO DE EQUIPO (POR SI SOLO QUIERE CONSULTAR UNO SOLO)
    
114     If Option2(0).Value = True Then     'Reloj Culiacan
116         vEquipo = 1
            '118     ElseIf Option2(1).Value = True Then     'Baño Culiacan
            '120         vEquipo = 2
            '122     ElseIf Option2(1).Value = True Then     'Vestidores Culiacan
            '124         vEquipo = 3
126     ElseIf Option2(1).Value = True Then     'Reloj SLP
128         vEquipo = 4
130     ElseIf Option2(1).Value = True Then     'Reloj Chiapas
132         vEquipo = 5
        End If
    
134     vCodEmpleado = DataGrid1.Columns(2).Text
               
136     If OptionDia(0).Value = True Then   'es para calcular de dia.
    
            'Pasa valor del DataGrid1 a una variable para saber la fecha inicial de busqueda.
138         vHoraInicial = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
            
140         vFechaFinal = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
142         vFechaFinal = vFechaFinal + 1
               
            'Busca la primera y la ultima checada del dia.
144         Set RBuscaUltimaChecada = New ADODB.Recordset
146         Call Abrir_Recordset(RBuscaUltimaChecada, "SELECT registros.checktime, registros.checktype, " & "usuarios.badgenumber, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.badgenumber LIKE '" & vCodEmpleado & "' " & "AND registros.checktime >= #" & Format(vHoraInicial, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime")
            
148         If RBuscaUltimaChecada.RecordCount > 0 Then 'si hay registros.
150             RBuscaUltimaChecada.MoveFirst
152             vh1 = Format(RBuscaUltimaChecada!checktime, "h:mm")
154             RBuscaUltimaChecada.MoveLast
156             vh2 = Format(RBuscaUltimaChecada!checktime, "h:mm")
            End If

158         vSumaHoras1 = Format(DateDiff("n", vh1, vh2) / 60, "#0")
                
160         If vSumaHoras1 = 0 Then 'Se salió el rango de datos.
162             Call MsgBox("No se encontró informacion en el rango seleccionado de fechas (Fecha Final). Calculo de Horas = 0. Vuelva a intentar.", vbCritical Or vbSystemModal, "Error en rango de Fechas !")
            End If
        
        Else   'Para calcular de noche.
     
            'Pasa valor del DataGrid1 a una variable para saber la fecha inicial de busqueda.
164         vHoraInicial = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
            
166         vFechaFinal = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
168         vFechaFinal = vFechaFinal + 1
            
            'Busca la ultima checada del dia y la primera del siguiente dia.
170         Set RBuscaUltimaChecada = New ADODB.Recordset
172         Call Abrir_Recordset(RBuscaUltimaChecada, "SELECT registros.checktime, registros.checktype, " & "usuarios.badgenumber, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.badgenumber LIKE '" & vCodEmpleado & "' " & "AND registros.checktime >= #" & Format(vHoraInicial, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime")
            
174         If RBuscaUltimaChecada.RecordCount > 0 Then 'si hay registros.
                    
                'Para sacar la ultima checada del dia (entró en la noche).
176             RBuscaUltimaChecada.MoveLast
178             vh1 = Format(RBuscaUltimaChecada!checktime, "h:mm")
                    
                'Para sacar la primera checada del dia siguiente (salió al día siguiente).
180             vHoraInicial = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
182             vHoraInicial = vHoraInicial + 1
184             vFechaFinal = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
186             vFechaFinal = vFechaFinal + 2
                    
188             Set RBuscaChecadaNoche = New ADODB.Recordset
190             Call Abrir_Recordset(RBuscaChecadaNoche, "SELECT registros.checktime, registros.checktype, " & "usuarios.badgenumber, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.badgenumber LIKE '" & vCodEmpleado & "' " & "AND registros.checktime >= #" & Format(vHoraInicial, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime")
                    
192             If RBuscaChecadaNoche.RecordCount > 0 Then 'si hay registros.
                    
194                 RBuscaChecadaNoche.MoveFirst
196                 vh2 = Format(RBuscaChecadaNoche!checktime, "h:mm")
                        
                End If
            End If
                
198         vSumaHoras1 = Format(DateDiff("n", vh1, vh2) / 60, "#0")
                
200         If vSumaHoras1 = 0 Then 'Se salió el rango de datos.
202             Call MsgBox("No se encontró informacion en el rango seleccionado de fechas (Fecha Final). Calculo de Horas = 0. Vuelva a intentar.", vbCritical Or vbSystemModal, "Error en rango de Fechas !")
            End If
                
        End If 'Fin de preguntar si es para dia o Noche.
        
204     Label4.Caption = vSumaHoras1
206     Label3.Caption = DataGrid1.Columns(3).Text
                
        'Saca_TodasHoras
        
208     Screen.MousePointer = vbDefault
 
        '<EhFooter>
        Exit Sub

Saca_SumaMinutos_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Saca_SumaMinutos " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : Saca_TodasHoras
' Descripción   : CALCULA EL TOTAL DE HORAS
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:06:08
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Saca_TodasHoras()

        '<EhHeader>
        On Error GoTo Saca_TodasHoras_Err

        '</EhHeader>

        Dim vh1 As Variant

        Dim vh2 As Variant

100     vh1 = 0
102     vh2 = 0

        Dim vhtodas1 As Variant

        Dim vhtodas2 As Variant

104     vhtodas1 = 0
106     vhtodas2 = 0
            
        Dim vSumaHoras2 As Variant

        Dim vFechaFinal As Date
            
        'Pasa valor del DataGrid1 a una variable para saber la fecha inicial de busqueda.
        'Dim vFechaFinal As Date
108     vFechaFinal = DTFinal.Value
110     vFechaFinal = vFechaFinal + 1
            
        Dim vHoraInicial As Date

112     vHoraInicial = Format(DataGrid1.Columns(0).Text, "dd/mm/yyyy hh:mm:ss")
            
        Dim vCodEmpleado As Integer

114     vCodEmpleado = DataGrid1.Columns(2).Text
            
116     Set RBuscaTodo = New ADODB.Recordset
118     Call Abrir_Recordset(RBuscaTodo, "SELECT registros.checktime, registros.checktype, " & "usuarios.badgenumber, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.badgenumber LIKE '" & vCodEmpleado & "' " & "AND registros.checktime >= #" & Format(DTInicio.Value, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime")
            
120     If RBuscaTodo.RecordCount > 0 Then 'si hay registros.

122         Do Until RBuscaTodo.EOF
                        
124             If vHoraInicial = Format(RBuscaTodo!checktime, "dd/mm/yyyy") Then
126                 RBuscaTodo.MoveNext
                                 
                Else
128                 vFechaFinal = vHoraInicial + 1
                                 
130                 Set RBuscaUltimaChecada = New ADODB.Recordset
132                 Call Abrir_Recordset(RBuscaUltimaChecada, "SELECT registros.checktime, registros.checktype, " & "usuarios.badgenumber, registros.sensorid " & "FROM CHECKINOUT registros, USERINFO usuarios, Machines M " & "WHERE registros.sensorid LIKE " & "'" & vEquipo & "' " & "AND M.ID = " & vIDEquipo & " " & "AND registros.USERID = usuarios.userid " & "AND usuarios.badgenumber LIKE '" & vCodEmpleado & "' " & "AND registros.checktime >= #" & Format(vHoraInicial, "mm/dd/yyyy") & "# " & "AND registros.checktime <= #" & Format(vFechaFinal, "mm/dd/yyyy") & "# " & "ORDER BY checktime")
                                   
134                 If RBuscaUltimaChecada.RecordCount > 0 Then 'si hay registros.
136                     RBuscaUltimaChecada.MoveFirst
138                     vh1 = Format(RBuscaUltimaChecada!checktime, "h:mm")
140                     RBuscaUltimaChecada.MoveLast
142                     vh2 = Format(RBuscaUltimaChecada!checktime, "h:mm")
                    End If
                                       
144                 vSumaHoras2 = vSumaHoras2 + Format(DateDiff("n", vh1, vh2) / 60, "#0")
146                 RBuscaUltimaChecada.MoveNext
                                     
                End If
                             
            Loop
                         
        End If
                 
148     vSumaHoras2 = Format(DateDiff("n", vhtodas1, vhtodas2) / 60, "#0")
        
150     Label6.Caption = vSumaHoras2

        '<EhFooter>
        Exit Sub

Saca_TodasHoras_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.Saca_TodasHoras " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Proyecto      : ReportesRelojHuella
' Procedimiento : cmdCalculaTotHrs_Click
' Descripción   : SUMA LAS HORAS TOTALES
' Creado por    : Miguel Angel
' Fecha-Hora    : 8/12/2011-16:10:38
'
' Parámetros    :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdCalculaTotHrs_Click()

        '<EhHeader>
        On Error GoTo cmdCalculaTotHrs_Click_Err

        '</EhHeader>
    
        'Para sacar las horas del dia segun el click en el datagrid1.
100     Saca_SumaMinutos

        '<EhFooter>
        Exit Sub

cmdCalculaTotHrs_Click_Err:
        MsgBox Err.Description & vbCrLf & "in ReportesRelojHuella.FrmReportesReloj.cmdCalculaTotHrs_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"

        Resume Next

        '</EhFooter>

End Sub

