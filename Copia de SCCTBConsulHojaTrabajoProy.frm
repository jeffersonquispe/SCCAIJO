VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCTBConsulHojaTrabajoProy 
   Caption         =   "SGCaijo-Contabilidad, Balance de Comprobación por Proyecto"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   104
   Icon            =   "SCCTBConsulHojaTrabajoProy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboCtasContables 
      Height          =   315
      ItemData        =   "SCCTBConsulHojaTrabajoProy.frx":08CA
      Left            =   5040
      List            =   "SCCTBConsulHojaTrabajoProy.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton cmdPProy 
      Height          =   255
      Left            =   6060
      Picture         =   "SCCTBConsulHojaTrabajoProy.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   310
      Width           =   220
   End
   Begin VB.TextBox txtProyecto 
      Height          =   315
      Left            =   960
      MaxLength       =   2
      TabIndex        =   0
      Top             =   280
      Width           =   415
   End
   Begin VB.ComboBox cboProyecto 
      Height          =   315
      Left            =   1400
      Style           =   1  'Simple Combo
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   280
      Width           =   4920
   End
   Begin VB.Frame Frame13 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11655
      Begin VB.TextBox TxtFechaInicioProyecto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   280
         Width           =   990
      End
      Begin VB.TextBox txtFinan 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   280
         Width           =   2480
      End
      Begin VB.Label Label2 
         Caption         =   "F.Inicio Proy:"
         Height          =   255
         Left            =   9600
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Proyecto:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lblFinanciera 
         Caption         =   "Financiera:"
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   300
         Width           =   795
      End
   End
   Begin Crystal.CrystalReport rptInformes 
      Left            =   600
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10320
      TabIndex        =   8
      Top             =   8040
      Width           =   1000
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Generar Informe"
      Height          =   400
      Left            =   8280
      TabIndex        =   7
      Top             =   8040
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   11655
      Begin MSMask.MaskEdBox mskNivelCont 
         Height          =   315
         Left            =   11040
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   4005
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   330
         Left            =   6405
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin:"
         Height          =   195
         Left            =   5520
         TabIndex        =   18
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio:"
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Top             =   285
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel contable (2,3,5) :"
         Height          =   255
         Left            =   9360
         TabIndex        =   10
         Top             =   315
         Visible         =   0   'False
         Width           =   1605
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   2
      HighLight       =   0
      FillStyle       =   1
   End
   Begin ComctlLib.ProgressBar prgInforme 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCTBConsulHojaTrabajoProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variables para la carga de la consulta
Dim msAnioMes As String
Dim FechaInicioConsulta As String
Dim FechaFinConsulta As String

' Cursores para la carga de la consulta
Dim mcurCtasContables As New clsBD2
Dim mcurSumMayorActDebe As New clsBD2
Dim mcurSumMayorActHaber As New clsBD2
Dim mcurSumMayorAntDebe As New clsBD2
Dim mcurSumMayorAntHaber As New clsBD2
Dim mcursumMovMesActDebe As New clsBD2
Dim mcursumMovMesActHAber As New clsBD2

' Variables para acumular los totales
Dim dblDebMayorAct As Double
Dim dblHabMayorAct As Double
Dim dblDebSaldoMayorAnt As Double
Dim dblHabSaldoMayorAnt As Double
Dim dblDebMovMesAct As Double
Dim dblHabMovMesAct As Double
Dim dblDebSaldoMes As Double
Dim dblHabSaldoMes As Double
Dim dblDebSaldoAct As Double
Dim dblHabSaldoAct As Double

'Colecciones para la carga del combo Proyectos
Private mcolCodDesProy As New Collection
Private mcolCodProy As New Collection

Dim NumeroCtaContableProyecto As String
Dim NumeroCtaContableBancaria As String
Dim IdCtaBancaria As String
Dim IndiceCtasContables As Integer
Dim NumeroCtaContable As String
Dim DescripcionCtaContable As String
Dim AnioAnteriorFinconsulta As String
Dim MesAnteriorFinconsulta As String
Dim DiaAnteriorFinconsulta As String
Dim FechaAnteriorFinconsulta As String

' Cursores para el manejo de ingresos y egresos Mes Anterior
Dim mcurSumIngresosAnt As New clsBD2
Dim mcurSumGastosAnt As New clsBD2
Dim mcurSumConceptosPlanillaAnt As New clsBD2
Dim mcurGastosConAfectacionAnt As New clsBD2
Dim mcurGastosSinAfectacionAnt As New clsBD2
Dim mcurAsientoManualProyectoAnt As New clsBD2
Dim mcurCTSAnt As New clsBD2
Dim mcurGratificacionAnt As New clsBD2

' Cursores para el manejo de ingresos y egresos Mes Actual
Dim mcurSumIngresosAct As New clsBD2
Dim mcurSumGastosAct As New clsBD2
Dim mcurSumConceptosPlanillaAct As New clsBD2
Dim mcurGastosConAfectacionAct As New clsBD2
Dim mcurGastosSinAfectacionAct As New clsBD2
Dim mcurAsientoManualProyectoAct As New clsBD2
Dim mcurCTSAct As New clsBD2
Dim mcurGratificacionAct As New clsBD2

' Cursores para el manejo de ingresos y egresos hasta la fecha de corte
Dim mcurSumIngresosCorte As New clsBD2
Dim mcurSumGastosCorte As New clsBD2
Dim mcurSumConceptosPlanillaCorte As New clsBD2
Dim mcurGastosConAfectacionCorte As New clsBD2
Dim mcurGastosSinAfectacionCorte As New clsBD2
Dim mcurAsientoManualProyectoCorte As New clsBD2
Dim mcurCTSCorte As New clsBD2
Dim mcurGratificacionCorte As New clsBD2

Dim TotalIngresos As Double
Dim TotalEgresos As Double
Dim TotalRemuneracionesPorPagar As Double
Dim PersonalDelProyectoSinBorrar As String
Dim curTotalPrestamo As New clsBD2
Dim curCanceladoPrestamo As New clsBD2
Dim curAfpEssaludRenta5ta As New clsBD2
Dim curCtasContablesContraCuentas As New clsBD2
Dim FechaInicioPlanilla As String
Dim InstrucPersonal As String
Dim sSQL As String

Dim curNumeroAsientoManual As New clsBD2
Dim AsientosManuales As String
Dim InstrucCtasContablesManuales As String

Dim ExistenAsientosManualesMayor As Boolean
Dim ExistenAsientosManualesMesAnterior As Boolean
Dim ExistenAsientosManualesMesActual As Boolean
  
Dim ExistenPlanillaMesAnterior As Boolean
Dim ExistenPlanillaMesActual As Boolean

Dim ExistenProyectoIngresoMesAnterior As Boolean
Dim ExistenProyectoEgresoMesAnterior As Boolean

Dim TotalCTSCorte As Double
Dim TotalCTSAct As Double
Dim TotalCTSAnt As Double
  
Private Sub cboProyecto_Change()
  ' verifica SI lo ingresado esta en la lista del combo
  If VerificarTextoEnLista(cboProyecto) = True Then SendKeys "{down}"
End Sub

Private Sub cboProyecto_Click()
  ' Verifica SI el evento ha sido activado por el teclado o Mouse
  If VerificarClick(cboProyecto.ListIndex) = False And cboProyecto.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
End Sub

Private Sub cboProyecto_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Verifica SI es enter para salir o flechas para recorrer
  VerificaKeyDowncbo (KeyCode)
End Sub

Private Sub cboProyecto_LostFocus()
  ' sale del combo y acualiza datos enlazados
  If ValidarDatoCbo(cboProyecto, vbWhite) = True Then
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboProyecto.Text, txtProyecto, mcolCodProy, mcolCodDesProy
    
    'actualizar la financiera del proyecto
    ActualizarFinancieraPeriodo
      
    'If cboMes.BackColor = vbWhite And mskAnio.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
    'If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
    If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite Then
      'Muestra los datos de hoja de trabajo
      CargaConsulta
    End If
  Else '  Vaciar Controles enlazados al combo
    txtProyecto.Text = Empty
    txtFinan.Text = Empty
  End If
  
  'Cambia el alto del combo
  cboProyecto.Height = CBONORMAL
End Sub

Private Sub cmdImprimir_Click()
Dim rptHojaTrabajo As New clsBD4

' Deshabilita el botón aceptar
  cmdImprimir.Enabled = False

' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón informe
       cmdImprimir.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla consulta de Hoja de trabajo
  LlenaTablaConsul
  
' Genera el reporte
' Formulario
  Set rptHojaTrabajo.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptHojaTrabajo.AsignarRpt

' Formula/s de Crystal.
  rptHojaTrabajo.Formulas.Add "MesAnio=' DEL " & mskFechaIni.Text & " AL " & mskFechaFin.Text & "'"
                            
' Clausula WHERE de las relaciones del rpt.
  rptHojaTrabajo.FiltroSelectionFormula = ""

' Nombre del fichero
  rptHojaTrabajo.NombreRPT = "rptCtbHojaTrabajo.rpt"

' Presentación preliminar del Informe
  rptHojaTrabajo.PresentancionPreliminar

' Elimina los datos de la tabla
  BorraDatosTablaConsul

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón aceptar
  cmdImprimir.Enabled = True
  
End Sub

Private Sub cmdPProy_Click()
  If cboProyecto.Enabled Then
    ' alto
     cboProyecto.Height = CBOALTO
    ' focus a cbo
    cboProyecto.SetFocus
  End If
End Sub

Private Sub cmdSalir_Click()

' Cierra el formulario de consulta de Nomina del mes
Unload Me

End Sub

Private Sub Form_Load()
Dim aTitulos1 As Variant
Dim aTitulos2 As Variant
Dim iCol As Integer

aTitulos1 = Array("CTA", "DESCRIPCION", _
                      "SUMAS DEL Mayor", "SUMAS DEL Mayor", _
                      "SALDO DEL MES ANTERIOR", "SALDO DEL MES ANTERIOR", _
                      "MOVIMIENTO DEL MES", "MOVIMIENTO DEL MES", _
                      "SALDO DEL MES", "SALDO DEL MES", _
                      "SALDO ACTUAL", "SALDO ACTUAL")
aTitulos2 = Array("CTA", "DESCRIPCION", _
                 "DEBE", "HABER", _
                 "DEBE", "HABER", _
                 "DEBE", "HABER", _
                 "DEBE", "HABER", _
                 "DEBE", "HABER")
                    
' Carga los tamaños del grid
aTamañosColumnas = Array(600, 3500, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200)
For iCol = 0 To grdConsulta.Cols - 1
  ' Carga los títulos y tamaños de las columnas
  grdConsulta.Row = 0
  grdConsulta.Col = iCol
  grdConsulta.ColWidth(iCol) = aTamañosColumnas(iCol)
  grdConsulta.Text = aTitulos1(iCol)
  grdConsulta.Row = 1
  grdConsulta.Text = aTitulos2(iCol)
Next

' Agrupa las filas con el mismo contenido
grdConsulta.MergeCells = 4
grdConsulta.MergeCol(0) = True
grdConsulta.MergeCol(1) = True
grdConsulta.MergeRow(0) = True
grdConsulta.MergeRow(1) = True

' Inhabilita el botón Imprimir
cmdImprimir.Enabled = False

'Establece obligatorios
txtProyecto.BackColor = Obligatorio
mskFechaIni.BackColor = Obligatorio
mskFechaFin.BackColor = Obligatorio
'mskNivelCont.BackColor = Obligatorio

'Se carga el combo de Proyectos
sSQL = ""
sSQL = "SELECT IdProy, IdProy + '   ' + DescProy FROM PROYECTOS WHERE CLASECTA = 'NAC' " & _
       " ORDER BY IdProy + '   ' + DescProy "
CD_CargarColsCbo cboProyecto, sSQL, mcolCodProy, mcolCodDesProy

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Colecciones para la carga del combo Proyectos
  Set mcolCodDesProy = Nothing
  Set mcolCodProy = Nothing
End Sub

Private Sub BorraDatosTablaConsul()
'------------------------------------------------------------
' Propósito: Borra los datos de la tabla RPTCTBHOJATRABAJO
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim modTablaConsul As New clsBD3

' Carga la sentencia
sSQL = "DELETE * FROM RPTCTBHOJATRABAJO"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modTablaConsul.Cerrar

End Sub

Private Sub LlenaTablaConsul()
'------------------------------------------------------------
' Propósito: LLena la tabla RPTCTBHOJATRABAJO
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim modTablaConsul As New clsBD3
Dim i As Long
' Recorre los datos del grid
For i = 2 To grdConsulta.Rows - 1
        ' Carga la sentencia sSQL
        sSQL = "INSERT INTO RPTCTBHOJATRABAJO VALUES('" _
        & grdConsulta.TextMatrix(i, 0) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 1)) & "'," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 2))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 3))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 4))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 5))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 6))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 7))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 8))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 9))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 10))) & "," _
        & Var32(Var37(grdConsulta.TextMatrix(i, 11))) & ")"
        
        ' Ejecuta la sentencia
        modTablaConsul.SQL = sSQL
        If modTablaConsul.Ejecutar = HAY_ERROR Then End
        modTablaConsul.Cerrar
Next i

End Sub

Private Sub CargaConsulta()
' -------------------------------------------------------
' Propósito : Verifica los datos y carga la consulta
' Recibe : Nada
' Entrega : Nada
' -------------------------------------------------------
' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
      grdConsulta.Rows = 2
    ' Deshabilita imprimir
      cmdImprimir.Enabled = False
    Exit Sub
  End If
  ' Inicia progreso
    prgInforme.Max = 6
    prgInforme.Min = 0
    prgInforme.Value = 0
    
  AnioAnteriorFinconsulta = Left(AnioMesAnterior(Val(Mid(FechaFinConsulta, 5, 2)), Val(Left(FechaFinConsulta, 4))), 4)
  MesAnteriorFinconsulta = Right(AnioMesAnterior(Val(Mid(FechaFinConsulta, 5, 2)), Val(Left(FechaFinConsulta, 4))), 2)
  DiaAnteriorFinconsulta = NumeroDiasMes(Val(MesAnteriorFinconsulta), Val(AnioAnteriorFinconsulta))
  FechaAnteriorFinconsulta = AnioAnteriorFinconsulta & MesAnteriorFinconsulta & DiaAnteriorFinconsulta

  prgInforme.Value = 1
' Cargar Ctas contables que tuvieron movimiento
  CargaDatosCtasContables
  prgInforme.Value = prgInforme.Value + 1
' Cargar sumas del Mayor hasta el mes elegido(Debe,Haber)
  CargaSumasMayorMesAct
  prgInforme.Value = prgInforme.Value + 1
' Cargar sumas del Mayor hasta el mes anterior(Debe,Haber)
  CargaSumasMayorMesAnt
  prgInforme.Value = prgInforme.Value + 1
' Cargar sumas del movimiento del mes elegido(Debe,Haber)
  CargaSumasMovActual
  prgInforme.Value = prgInforme.Value + 1
'Calcular CTS
  CalcularCTS
' Carga el grid consulta
  CargarGridConsulta
  prgInforme.Value = prgInforme.Value + 1
' Habilita el botón imprimir
  cmdImprimir.Enabled = True
  prgInforme.Value = 0

End Sub

Private Sub CargaSumasMovActual()
'-------------------------------------------------------------------------
'Propósito: Carga los cursores del debe y haber acumulados _
            en la contabilidad que pertenecen al mes elegido
'Recibe: Nada
'Entrega : Nada
'-------------------------------------------------------------------------

  ParaProyectoMesActual
  
  ParaPlanillaMesActual
  
  ParaGastoMesActual

End Sub

Private Sub CargaSumasMayorMesAnt()
'-------------------------------------------------------------------------
'Propósito: Carga los cursores del debe y haber acumulados _
            en la contabilidad hasta el mes anterior al elegido
'Recibe: Nada
'Entrega : Nada
'-------------------------------------------------------------------------
  
  ParaProyectoMesAnterior
  
  ParaPlanillaMesAnterior
  
  ParaGastoMesAnterior
End Sub

Private Sub CargaSumasMayorMesAct()
'-------------------------------------------------------------------------
'Propósito: Carga los cursores del debe y haber acumulados _
            en la contabilidad hasta el mes elegido
'Recibe: Nada
'Entrega : Nada
'-------------------------------------------------------------------------

  ParaProyectoSumasMayor
  
  ParaPlanillaSumasMayor
  
  ParaGastoSumasMayor
  
  
End Sub

Private Sub CargaDatosCtasContables()
'-------------------------------------------------------------------------
'Propósito: Carga los datos de las Ctas contables usadas
'Recibe: Nada
'Entrega : Nada
'-------------------------------------------------------------------------
' Carga la consulta
'  VADICK MODIFICACION PARA CARGAR LAS CTAS CONTABLES Y EL PERSONAL DEL PROYECTO

  ' Proyecto
  CtasContablesProyecto
  ' Planilla
  CtasContablesPlanilla
  ' GastoProductosServicios
  CtasContablesGasto
  ' carga Contracuentas
  CtasContablesContraCuentas
End Sub

Private Sub CargarGridConsulta()
'-------------------------------------------------------------------------
'Propósito: Carga la consulta de la hoja de trabajo
'Recibe: Nada
'Entrega : Nada
'-------------------------------------------------------------------------
Dim i As Integer

' Inicializa acumuladores
dblDebMayorAct = 0: dblHabMayorAct = 0
dblDebSaldoMayorAnt = 0: dblHabSaldoMayorAnt = 0
dblDebMovMesAct = 0: dblHabMovMesAct = 0
dblDebSaldoMes = 0: dblHabSaldoMes = 0
dblDebSaldoAct = 0: dblHabSaldoAct = 0

' Inicializa el grid
grdConsulta.Rows = 2
grdConsulta.ScrollBars = flexScrollBarNone
grdConsulta.Visible = False

' Recorre las ctas
For IndiceCtasContables = 0 To ComboCtasContables.ListCount - 1
  ' Añade una fila para la cta. contable al grid
  NumeroCtaContable = Left(ComboCtasContables.List(IndiceCtasContables), InStr(1, ComboCtasContables.List(IndiceCtasContables), "@") - 2)
  DescripcionCtaContable = Mid(ComboCtasContables.List(IndiceCtasContables), InStr(1, ComboCtasContables.List(IndiceCtasContables), "@") + 2, Len(ComboCtasContables.List(IndiceCtasContables)))
  grdConsulta.AddItem NumeroCtaContable & vbTab & DescripcionCtaContable
  
  ' Muestra los datos del Mayor
  MuestraSumasMayor
  ' Muestra el saldo hasta el mes anterior
  MuestraMesAnterior
  ' Muestra el movimiento del mes
  MuestraMesActual
  ' Muestra los saldos del mes
  'MuestraSaldoMovMes
  ' Muestra el saldo total hasta el mes
  'MuestraSaldoActual
  ' Carga acumuladores
  'CargaAcumuladores
  ' Mueve a la siguiente cuenta
  'mcurCtasContables.MoverSiguiente
Next
  
  CargarCTS
  
  TotalRemuneraciones 2
  CargaValoresContraCuentas 2
  ObtenerSaldo 2, 10
  
  TotalRemuneraciones 4
  CargaValoresContraCuentas 4
  ObtenerSaldo 4, 4
  
  TotalRemuneraciones 6
  CargaValoresContraCuentas 6
  ObtenerSaldo 6, 8
    
  grdConsulta.AddItem vbTab & "TOTALES BALANCE"
  ' Coloca el color a los totales
  grdConsulta.Row = grdConsulta.Rows - 1
  MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
  
  ObtenerTotales 2
  ObtenerTotales 4
  ObtenerTotales 6
  ObtenerTotales 8
  ObtenerTotales 10
  
' Coloca las barras de desplazamiento
grdConsulta.ScrollBars = flexScrollBarBoth
grdConsulta.Visible = True
End Sub

Private Sub MuestraTotales()
' ----------------------------------------------------
' Propósito: Muestra los totales acumulados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' Añade totales al grid
grdConsulta.AddItem vbTab & "TOTALES" & vbTab & _
        Format(dblDebMayorAct, "###,###,##0.00") & vbTab & _
        Format(dblHabMayorAct, "###,###,##0.00") & vbTab & _
        Format(dblDebSaldoMayorAnt, "###,###,##0.00") & vbTab & _
        Format(dblHabSaldoMayorAnt, "###,###,##0.00") & vbTab & _
        Format(dblDebMovMesAct, "###,###,##0.00") & vbTab & _
        Format(dblHabMovMesAct, "###,###,##0.00") & vbTab & _
        Format(dblDebSaldoMes, "###,###,##0.00") & vbTab & _
        Format(dblHabSaldoMes, "###,###,##0.00") & vbTab & _
        Format(dblDebSaldoAct, "###,###,##0.00") & vbTab & _
        Format(dblHabSaldoAct, "###,###,##0.00")

' Coloca el color a los totales
grdConsulta.Row = grdConsulta.Rows - 1
MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
  
End Sub

Private Sub CargaAcumuladores()
' ----------------------------------------------------
' Propósito: Acumula en las variables los montos calculados _
            para la cta contable
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' Acumula montos
dblDebMayorAct = dblDebMayorAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2)))
dblHabMayorAct = dblHabMayorAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3)))
dblDebSaldoMayorAnt = dblDebSaldoMayorAnt + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4)))
dblHabSaldoMayorAnt = dblHabSaldoMayorAnt + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5)))
dblDebMovMesAct = dblDebMovMesAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6)))
dblHabMovMesAct = dblHabMovMesAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7)))
dblDebSaldoMes = dblDebSaldoMes + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 8)))
dblHabSaldoMes = dblHabSaldoMes + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 9)))
dblDebSaldoAct = dblDebSaldoAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 10)))
dblHabSaldoAct = dblHabSaldoAct + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 11)))

End Sub

Private Sub MuestraSaldoActual()
' ----------------------------------------------------
' Propósito: Coloca en sus columnas respectivas el debe y _
            haber del saldo Actual hasta el mes elegido
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim dblSaldo As Double
' Carga el debe del movimiento para calcular el saldo
  dblSaldo = dblSaldo + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2)))
' Carga el haber del movimiento para calcular el saldo
  dblSaldo = dblSaldo - Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3)))
' Verifica si el saldo esta en el debe o haber del saldo actual
  If dblSaldo > 0 Then
    ' Carga al Debe
    grdConsulta.TextMatrix(grdConsulta.Rows - 1, 10) = _
    Format(dblSaldo, "###,###,##0.00")
  ElseIf dblSaldo < 0 Then
    ' Carga al Haber
    grdConsulta.TextMatrix(grdConsulta.Rows - 1, 11) = _
    Format(dblSaldo * (-1), "###,###,##0.00")
  End If

End Sub

Private Sub MuestraSaldoMovMes()
' ----------------------------------------------------
' Propósito: Coloca en sus columnas respectivas el debe y _
            haber del saldo del movimiento del mes
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim dblSaldo As Double
' Carga el debe del movimiento para calcular el saldo
  dblSaldo = dblSaldo + Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6)))
' Carga el haber del movimiento para calcular el saldo
  dblSaldo = dblSaldo - Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7)))
' Verifica si el saldo esta en el debe o haber del mov del mes
  If dblSaldo > 0 Then
    ' Carga al Debe
    grdConsulta.TextMatrix(grdConsulta.Rows - 1, 8) = _
    Format(dblSaldo, "###,###,##0.00")
  ElseIf dblSaldo < 0 Then
    ' Carga al Haber
    grdConsulta.TextMatrix(grdConsulta.Rows - 1, 9) = _
    Format(dblSaldo * (-1), "###,###,##0.00")
  End If
  
End Sub

Private Sub MuestraMesActual()
' ----------------------------------------------------
' Propósito: Coloca en sus columnas respectivas el debe y _
            haber de el movimiento del mes
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------

  Dim i As Integer
  
  ' Verifica si se terminó de recorrer el cursor
  ' */*/*/*/      44113
  If Not (mcurSumIngresosAct.EOF) Then
    ' Verifica si tiene el Mayor montos al debe
    'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
    If NumeroCtaContable = "44113" Then
      ' Carga al debe
      If mcurSumIngresosAct.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumIngresosAct.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumIngresosAct.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumIngresosAct.MoverSiguiente
      
      If Not (mcurSumIngresosAct.EOF) Then
        ' Carga al Haber
        If mcurSumIngresosAct.campo(1) = "D" Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumIngresosAct.campo(2), "###,###,##0.00")
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumIngresosAct.campo(2), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
        mcurSumIngresosAct.MoverSiguiente
      End If
    End If
  End If
  
  ' Verifica si se terminó de recorrer el cursor
  ' */*/*/*/      10414
  If Not (mcurSumGastosAct.EOF) Then
    ' Verifica si tiene el Mayor montos al Haber
  '  If NumeroCtaContable = mcurSumGastosCorte.campo(0) Then
  '    ' Carga al Haber
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumGastosCorte.campo(1), "###,###,##0.00")
  '    ' Mueve al siguiente elemento
  '        mcurSumGastosCorte.MoverSiguiente
  '  End If
    
    If NumeroCtaContable = "10414" Then
      ' Carga al debe
      If mcurSumGastosAct.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumGastosAct.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumGastosAct.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumGastosAct.MoverSiguiente
      
      If Not (mcurSumGastosAct.EOF) Then
        ' Carga al Haber
        If mcurSumGastosAct.campo(1) = "D" Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumGastosAct.campo(2), "###,###,##0.00")
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumGastosAct.campo(2), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
        mcurSumGastosAct.MoverSiguiente
      End If
    End If
  End If
  
  If ExistenAsientosManualesMesActual = True Then
    If Not (mcurAsientoManualProyectoAct.EOF) Then
      ' Verifica si tiene el Mayor montos al debe
      'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
      If NumeroCtaContable = mcurAsientoManualProyectoAct.campo(0) Then
        ' Carga al debe
        If mcurAsientoManualProyectoAct.campo(1) = "D" Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurAsientoManualProyectoAct.campo(2), "###,###,##0.00")
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurAsientoManualProyectoAct.campo(2), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
        mcurAsientoManualProyectoAct.MoverSiguiente
        
    '    If Not (mcurAsientoManualProyectoCorte.EOF) Then
    '      ' Carga al Haber
    '      If mcurAsientoManualProyectoCorte.campo(1) = "D" Then
    '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
    '      Else
    '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
    '      End If
    '      ' Mueve al siguiente elemento
    '      mcurAsientoManualProyectoCorte.MoverSiguiente
    '    End If
      End If
    End If
  End If
  
  If ExistenPlanillaMesActual = True Then
    ' Verifica si se terminó de recorrer el cursor
    If Not (mcurSumConceptosPlanillaAct.EOF) Then
      ' Verifica si tiene el Mayor montos al Haber
      If NumeroCtaContable = mcurSumConceptosPlanillaAct.campo(0) Then
        ' Carga al Haber
        If InStr(1, NumeroCtaContable, "14111") Then
          CtaPrestamos
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(curTotalPrestamo.campo(0), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(curCanceladoPrestamo.campo(0), "###,###,##0.00")
        
        ElseIf InStr(1, NumeroCtaContable, "38511") Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
        
        ElseIf InStr(1, NumeroCtaContable, "40172") Then
          'CtaAfpEssaludRenta5ta 40172
          'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
    
        ElseIf InStr(1, NumeroCtaContable, "4691") Then
          'CtaAfpEssaludRenta5ta NumeroCtaContable
          'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 7) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
          
        ElseIf InStr(1, NumeroCtaContable, "62711") Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
          For i = 0 To grdConsulta.Rows - 1
            If grdConsulta.TextMatrix(i, 0) = "40311" Then
              'CtaAfpEssaludRenta5ta 62711
              'grdConsulta.TextMatrix(i, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
              grdConsulta.TextMatrix(i, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
              grdConsulta.TextMatrix(i, 7) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
            End If
          Next i
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurSumConceptosPlanillaAct.campo(1), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
          mcurSumConceptosPlanillaAct.MoverSiguiente
      End If
    End If
  End If
  
  ' Verifica si se terminó de recorrer el cursor
  If Not (mcurGastosConAfectacionAct.EOF) Then
    ' Verifica si tiene el Mayor montos al Haber
    If NumeroCtaContable = mcurGastosConAfectacionAct.campo(0) Then
      ' Carga al Haber
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurGastosConAfectacionAct.campo(1), "###,###,##0.00")
      ' Mueve al siguiente elemento
        mcurGastosConAfectacionAct.MoverSiguiente
    End If
  End If
  
  ' Verifica si se terminó de recorrer el cursor
  If Not (mcurGastosSinAfectacionAct.EOF) Then
    ' Verifica si tiene el Mayor montos al Haber
    If NumeroCtaContable = mcurGastosSinAfectacionAct.campo(0) Then
      ' Carga al Haber
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 6) = Format(mcurGastosSinAfectacionAct.campo(1), "###,###,##0.00")
      ' Mueve al siguiente elemento
          mcurGastosSinAfectacionAct.MoverSiguiente
    End If
  End If
End Sub

Private Sub MuestraMesAnterior()
' ----------------------------------------------------
' Propósito: Coloca en sus columnas respectivas el debe y _
            haber de el Saldo hasta el mes anterior
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------

Dim i As Integer

If ExistenProyectoIngresoMesAnterior = True Then
  ' Verifica si se terminó de recorrer el cursor
  ' */*/*/*/      44113
  If Not (mcurSumIngresosAnt.EOF) Then
    ' Verifica si tiene el Mayor montos al debe
    'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
    If NumeroCtaContable = "44113" Then
      ' Carga al debe
      If mcurSumIngresosAnt.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumIngresosAnt.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumIngresosAnt.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumIngresosAnt.MoverSiguiente
      
      If Not (mcurSumIngresosAnt.EOF) Then
        ' Carga al Haber
        If mcurSumIngresosAnt.campo(1) = "D" Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumIngresosAnt.campo(2), "###,###,##0.00")
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumIngresosAnt.campo(2), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
        mcurSumIngresosAnt.MoverSiguiente
      End If
    End If
  End If
End If

If ExistenProyectoEgresoMesAnterior = True Then
  ' Verifica si se terminó de recorrer el cursor
  ' */*/*/*/      10414
  If Not (mcurSumGastosAnt.EOF) Then
    If NumeroCtaContable = "10414" Then
      ' Carga al debe
      If mcurSumGastosAnt.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumGastosAnt.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumGastosAnt.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumGastosAnt.MoverSiguiente
      
      If Not (mcurSumGastosAnt.EOF) Then
        ' Carga al Haber
        If mcurSumGastosAnt.campo(1) = "D" Then
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumGastosAnt.campo(2), "###,###,##0.00")
        Else
          grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumGastosAnt.campo(2), "###,###,##0.00")
        End If
        ' Mueve al siguiente elemento
        mcurSumGastosAnt.MoverSiguiente
      End If
    End If
  End If
End If

If ExistenAsientosManualesMesAnterior = True Then
  If Not (mcurAsientoManualProyectoAnt.EOF) Then
    ' Verifica si tiene el Mayor montos al debe
    'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
    If NumeroCtaContable = mcurAsientoManualProyectoAnt.campo(0) Then
      ' Carga al debe
      If mcurAsientoManualProyectoAnt.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurAsientoManualProyectoAnt.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurAsientoManualProyectoAnt.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurAsientoManualProyectoAnt.MoverSiguiente
      
  '    If Not (mcurAsientoManualProyectoCorte.EOF) Then
  '      ' Carga al Haber
  '      If mcurAsientoManualProyectoCorte.campo(1) = "D" Then
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
  '      Else
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
  '      End If
  '      ' Mueve al siguiente elemento
  '      mcurAsientoManualProyectoCorte.MoverSiguiente
  '    End If
    End If
  End If
End If

If ExistenPlanillaMesAnterior = True Then
  ' Verifica si se terminó de recorrer el cursor
  If Not (mcurSumConceptosPlanillaAnt.EOF) Then
    ' Verifica si tiene el Mayor montos al Haber
    If NumeroCtaContable = mcurSumConceptosPlanillaAnt.campo(0) Then
      ' Carga al Haber
      If InStr(1, NumeroCtaContable, "14111") Then
        CtaPrestamos
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(curTotalPrestamo.campo(0), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(curCanceladoPrestamo.campo(0), "###,###,##0.00")
      
      ElseIf InStr(1, NumeroCtaContable, "38511") Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
      
      ElseIf InStr(1, NumeroCtaContable, "40172") Then
        'CtaAfpEssaludRenta5ta 40172
        'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
  
      ElseIf InStr(1, NumeroCtaContable, "4691") Then
        'CtaAfpEssaludRenta5ta NumeroCtaContable
        'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
        
      ElseIf InStr(1, NumeroCtaContable, "62711") Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
        For i = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(i, 0) = "40311" Then
            'CtaAfpEssaludRenta5ta 62711
            'grdConsulta.TextMatrix(i, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
            grdConsulta.TextMatrix(i, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
            grdConsulta.TextMatrix(i, 5) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
          End If
        Next i
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurSumConceptosPlanillaAnt.campo(1), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
        mcurSumConceptosPlanillaAnt.MoverSiguiente
    End If
  End If
End If

' Verifica si se terminó de recorrer el cursor
If Not (mcurGastosConAfectacionAnt.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
  If NumeroCtaContable = mcurGastosConAfectacionAnt.campo(0) Then
    ' Carga al Haber
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurGastosConAfectacionAnt.campo(1), "###,###,##0.00")
    ' Mueve al siguiente elemento
      mcurGastosConAfectacionAnt.MoverSiguiente
  End If
End If

' Verifica si se terminó de recorrer el cursor
If Not (mcurGastosSinAfectacionAnt.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
  If NumeroCtaContable = mcurGastosSinAfectacionAnt.campo(0) Then
    ' Carga al Haber
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 4) = Format(mcurGastosSinAfectacionAnt.campo(1), "###,###,##0.00")
    ' Mueve al siguiente elemento
        mcurGastosSinAfectacionAnt.MoverSiguiente
  End If
End If
End Sub

Private Sub MuestraSumasMayor()
' ----------------------------------------------------
' Propósito: coloca en sus columnas respectivas el debe y _
            haber de el cursor de sumas del Mayor
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim i As Integer

' Verifica si se terminó de recorrer el cursor
' */*/*/*/      44113
If Not (mcurSumIngresosCorte.EOF) Then
  ' Verifica si tiene el Mayor montos al debe
  'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
  If NumeroCtaContable = "44113" Then
    ' Carga al debe
    If mcurSumIngresosCorte.campo(1) = "D" Then
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
    Else
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
    End If
    ' Mueve al siguiente elemento
    mcurSumIngresosCorte.MoverSiguiente
    
    If Not (mcurSumIngresosCorte.EOF) Then
      ' Carga al Haber
      If mcurSumIngresosCorte.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumIngresosCorte.MoverSiguiente
    End If
  End If
End If

' Verifica si se terminó de recorrer el cursor
' */*/*/*/      10414
If Not (mcurSumGastosCorte.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
'  If NumeroCtaContable = mcurSumGastosCorte.campo(0) Then
'    ' Carga al Haber
'        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumIngresosCorte.campo(2), "###,###,##0.00")
'        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumGastosCorte.campo(1), "###,###,##0.00")
'    ' Mueve al siguiente elemento
'        mcurSumGastosCorte.MoverSiguiente
'  End If
  
  If NumeroCtaContable = "10414" Then
    ' Carga al debe
    If mcurSumGastosCorte.campo(1) = "D" Then
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumGastosCorte.campo(2), "###,###,##0.00")
    Else
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumGastosCorte.campo(2), "###,###,##0.00")
    End If
    ' Mueve al siguiente elemento
    mcurSumGastosCorte.MoverSiguiente
    
    If Not (mcurSumGastosCorte.EOF) Then
      ' Carga al Haber
      If mcurSumGastosCorte.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumGastosCorte.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumGastosCorte.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurSumGastosCorte.MoverSiguiente
    End If
  End If
End If

If ExistenAsientosManualesMayor = True Then
  If Not (mcurAsientoManualProyectoCorte.EOF) Then
    ' Verifica si tiene el Mayor montos al debe
    'If NumeroCtaContable = mcurSumIngresosCorte.campo(0) Then
    If NumeroCtaContable = mcurAsientoManualProyectoCorte.campo(0) Then
      ' Carga al debe
      If mcurAsientoManualProyectoCorte.campo(1) = "D" Then
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
      Else
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
      End If
      ' Mueve al siguiente elemento
      mcurAsientoManualProyectoCorte.MoverSiguiente
      
  '    If Not (mcurAsientoManualProyectoCorte.EOF) Then
  '      ' Carga al Haber
  '      If mcurAsientoManualProyectoCorte.campo(1) = "D" Then
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
  '      Else
  '        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurAsientoManualProyectoCorte.campo(2), "###,###,##0.00")
  '      End If
  '      ' Mueve al siguiente elemento
  '      mcurAsientoManualProyectoCorte.MoverSiguiente
  '    End If
    End If
  End If
End If

' Verifica si se terminó de recorrer el cursor
If Not (mcurSumConceptosPlanillaCorte.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
  If NumeroCtaContable = mcurSumConceptosPlanillaCorte.campo(0) Then
    ' Carga al Haber
    If InStr(1, NumeroCtaContable, "14111") Then
      CtaPrestamos
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curTotalPrestamo.campo(0), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(curCanceladoPrestamo.campo(0), "###,###,##0.00")
    
    ElseIf InStr(1, NumeroCtaContable, "38511") Then
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
    
    ElseIf InStr(1, NumeroCtaContable, "40172") Then
      'CtaAfpEssaludRenta5ta 40172
      'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")

    ElseIf InStr(1, NumeroCtaContable, "4691") Then
      'CtaAfpEssaludRenta5ta NumeroCtaContable
      'grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 3) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
      
    ElseIf InStr(1, NumeroCtaContable, "62711") Then
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
      For i = 0 To grdConsulta.Rows - 1
        If grdConsulta.TextMatrix(i, 0) = "40311" Then
          'CtaAfpEssaludRenta5ta 62711
          'grdConsulta.TextMatrix(i, 2) = Format(curAfpEssaludRenta5ta.campo(0), "###,###,##0.00")
          grdConsulta.TextMatrix(i, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
          grdConsulta.TextMatrix(i, 3) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
        End If
      Next i
    Else
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurSumConceptosPlanillaCorte.campo(1), "###,###,##0.00")
    End If
    ' Mueve al siguiente elemento
      mcurSumConceptosPlanillaCorte.MoverSiguiente
  End If
End If

' Verifica si se terminó de recorrer el cursor
If Not (mcurGastosConAfectacionCorte.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
  If NumeroCtaContable = mcurGastosConAfectacionCorte.campo(0) Then
    ' Carga al Haber
      grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurGastosConAfectacionCorte.campo(1), "###,###,##0.00")
    ' Mueve al siguiente elemento
      mcurGastosConAfectacionCorte.MoverSiguiente
  End If
End If

' Verifica si se terminó de recorrer el cursor
If Not (mcurGastosSinAfectacionCorte.EOF) Then
  ' Verifica si tiene el Mayor montos al Haber
  If NumeroCtaContable = mcurGastosSinAfectacionCorte.campo(0) Then
    ' Carga al Haber
        grdConsulta.TextMatrix(grdConsulta.Rows - 1, 2) = Format(mcurGastosSinAfectacionCorte.campo(1), "###,###,##0.00")
    ' Mueve al siguiente elemento
        mcurGastosSinAfectacionCorte.MoverSiguiente
  End If
End If
End Sub


Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito: Verifica si esta bien los datos para ejecutar _
            la consulta
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim curContabilidad As New clsBD2

'If mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite _
  And txtProyecto.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
If mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite _
  And txtProyecto.BackColor = vbWhite Then
' Verifica que la fecha de inicio sea Menor a la fecha final
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = True Then
        fbOkDatosIntroducidos = False
        Exit Function
    End If
Else ' Alguna fecha es obligatorio
        fbOkDatosIntroducidos = False
        Exit Function
End If

' Carga el código de planilla introducido
msAnioMes = Right(mskFechaFin.Text, 4) & Mid(mskFechaFin.Text, 4, 2)
FechaInicioConsulta = FechaAMD(mskFechaIni.Text)
FechaFinConsulta = FechaAMD(mskFechaFin.Text)

' Verifica si existen asientos para Mes y Año, ' Carga la consulta
sSQL = ""
sSQL = "SELECT COUNT (*) FROM CTB_ASIENTOS " _
    & "WHERE Fecha like '" & msAnioMes & "*' and " _
    & "Anulado='NO'"

' Ejecuta la sentencia
curContabilidad.SQL = sSQL
If curContabilidad.Abrir = HAY_ERROR Then End

' Verifica la cantidad de registros en contabilidad
If curContabilidad.campo(0) = 0 Then
    MsgBox "No se tienen asientos para el mes y año introducidos " _
        , , "SGCcaijo-Boletas de Pago"
    ' Cierra el cursor y devuelve la función
    curContabilidad.Cerrar
    fbOkDatosIntroducidos = False
    'mskAnio.SetFocus
    Exit Function
End If

' Devuelve la función y cierra la consulta
curContabilidad.Cerrar

fbOkDatosIntroducidos = True

End Function

Private Sub mskFechaFin_Change()
  If InStr(1, mskFechaFin.Text, "_") < 1 Then
    If mskFechaFin.Text = UltimoDiaMes(Mid(mskFechaFin.Text, 4, 2), Right(mskFechaFin.Text, 4)) Then
      ' Se valida que la fecha fin de la consulta
      If ValidarFecha(mskFechaFin) Then
        mskFechaFin.BackColor = vbWhite
        'If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
        If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite Then
          ' Carga consulta
          CargaConsulta
        End If
      Else
        mskFechaFin.BackColor = Obligatorio
        grdConsulta.Rows = 2
        ' Deshabilita el botón generar informe
        cmdImprimir.Enabled = False
      End If
    Else
      mskFechaFin.Text = UltimoDiaMes(Mid(mskFechaFin.Text, 4, 2), Right(mskFechaFin.Text, 4))
    End If
  Else
    mskFechaFin.BackColor = Obligatorio
    grdConsulta.Rows = 2
    ' Deshabilita el botón generar informe
    cmdImprimir.Enabled = False
  End If
End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)
' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
End Sub

Private Sub mskFechaIni_Change()
  If InStr(1, mskFechaIni.Text, "_") < 1 Then
    If FechaAMD(mskFechaIni.Text) < FechaAMD(TxtFechaInicioProyecto.Text) Then
      MsgBox "El Inicio de la consulta no puede ser Menor a la Fecha de Inicio del Proyecto.", vbInformation + vbOKOnly, "SGCcaijo-Balance de Comprobaciòn"
      mskFechaIni.Text = "__/__/____"
      mskFechaIni.BackColor = Obligatorio
    Else
      ' Se valida que la fecha fin de la consulta
      If ValidarFecha(mskFechaIni) Then
          mskFechaIni.BackColor = vbWhite
          'If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
          If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite Then
            ' Carga consulta
            CargaConsulta
          End If
      Else
        mskFechaIni.BackColor = Obligatorio
        grdConsulta.Rows = 2
        ' Deshabilita el botón generar informe
        cmdImprimir.Enabled = False
      End If
    End If
  Else
    mskFechaIni.BackColor = Obligatorio
    grdConsulta.Rows = 2
    ' Deshabilita el botón generar informe
    cmdImprimir.Enabled = False
  End If
End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)
  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
End Sub

'Private Sub mskNivelCont_Change()
'
'  If InStr(1, mskNivelCont.Text, "_") < 1 Then
'    ' Muestra la consulta
'    Select Case mskNivelCont
'    Case 2, 3, 5
'        ' Nivel correcto
'        mskNivelCont.BackColor = vbWhite
'        ' Verifica si el combo de mes es <> de vacio
'        If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
'          ' Muestra los datos de planilla
'          CargaConsulta
'        End If
'    Case Else
'        ' Nivel incorrecto
'        mskNivelCont.BackColor = Obligatorio
'
'     ' Limpia el grid
'        grdConsulta.Rows = 2
'        cmdImprimir.Enabled = False
'
'      ' Msg Debe ser 2, 3, 5
'        MsgBox "El nivel contable debe ser 2,3 o 5", vbInformation + vbOKOnly, "SGCcaijo-Valida nivel contable"
'
'    End Select
'  Else
'    mskNivelCont.BackColor = Obligatorio
'    grdConsulta.Rows = 2
'    ' Deshabilita el botón generar informe
'    cmdImprimir.Enabled = False
'  End If
'
'End Sub

'Private Sub mskNivelCont_KeyPress(KeyAscii As Integer)
'
'' Si presiona enter pasa al siguiente control
'If KeyAscii = 13 Then
'    SendKeys vbTab
'End If
'
'End Sub

Private Sub txtProyecto_Change()
  If (txtProyecto.Text <> Empty) And Len(txtProyecto.Text) = 2 Then
    ' SI procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboProyecto, txtProyecto, mcolCodDesProy
  
    ' Verifica SI el campo esta vacio
    If txtProyecto.Text <> Empty And cboProyecto.Text <> "" Then
      ' Los campos coloca a color blanco
      txtProyecto.BackColor = vbWhite
      
      'Actualiza Financiera y periodo del proyecto
      ActualizarFinancieraPeriodo
      
      'If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite And mskNivelCont.BackColor = vbWhite Then
      If txtProyecto.BackColor = vbWhite And mskFechaIni.BackColor = vbWhite And mskFechaFin.BackColor = vbWhite Then
        ' Carga consulta
        CargaConsulta
      End If
    End If
  Else
    'Marca los campos obligatorios, y limpia el combo
    txtProyecto.BackColor = Obligatorio
    cboProyecto.Text = ""
    txtFinan.Text = ""
    TxtFechaInicioProyecto.Text = ""
    
    grdConsulta.Rows = 2
    ' Deshabilita el botón generar informe
    cmdImprimir.Enabled = False
  End If
End Sub

Private Sub txtProyecto_KeyPress(KeyAscii As Integer)
  ' Si presiona enter entonces manda al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
End Sub

Public Sub ActualizarFinancieraPeriodo()
'----------------------------------------------------------------------------
'PROPÓSITO: Actualizar los controles referentes a financiera, despues de ingresar un proyecto
'Recibe:    nada
'Devuelve:  nada
'----------------------------------------------------------------------------
' nota: llamado desde textbox, y combobox de proyecto al ingresar un proyecto
 Dim curFinanPerioProy As New clsBD2

    'Recupera financiera del proyecto seleccionado
    sSQL = ""
    sSQL = "SELECT F.DescFinan, P.PerioProy, P.FECINICIO " & _
           "FROM PROYECTOS P, Tipo_Financieras F " & _
           "WHERE P.IdProy=" & "'" & txtProyecto.Text & "'" & _
           " And P.IdFinan=F.IdFinan"
           
    curFinanPerioProy.SQL = sSQL
    
    ' ejecuta la consulta y asignamos al txt de proyecto
    If curFinanPerioProy.Abrir = HAY_ERROR Then
      Unload Me
      End
    End If
    txtFinan.Text = curFinanPerioProy.campo(0)
    TxtFechaInicioProyecto.Text = FechaDMA(curFinanPerioProy.campo(2))
    
    curFinanPerioProy.Cerrar 'Cierra el cursor

End Sub

Sub CtasContablesProyecto()
  Dim curCtasContProy As New clsBD2
  Dim curCtasContCtaBancaria As New clsBD2
  Dim curAsientosCtaContProy As New clsBD2
  Dim curTodasCtaContProy As New clsBD2
  Dim TotalAsientosContablesProy As String
  Dim InstrucCtasContablesProy As String
  Dim curAsientosCtaContBanc As New clsBD2
  Dim curTodasCtaContBanc As New clsBD2
  Dim TotalAsientosContablesBanc As String
  Dim InstrucCtasContablesBanc As String
  Dim curNumeroAsientoManual As New clsBD2
  Dim curCtasContManuales As New clsBD2
  
  ComboCtasContables.Clear
  
  sSQL = ""
  sSQL = "SELECT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA, PROYECTOS.IDPROY, PROYECTOS.IDCTA " _
      & "FROM PROYECTOS, PLAN_CONTABLE " _
      & "WHERE (PROYECTOS.CTACONTAPROYECTO = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(PROYECTOS.IDPROY = '" & Trim(txtProyecto.Text) & "')"
  
  ' Ejecuta la consulta
  curCtasContProy.SQL = sSQL

  If curCtasContProy.Abrir = HAY_ERROR Then
    End
  End If
  
  If Not curCtasContProy.EOF Then
    IdCtaBancaria = curCtasContProy.campo(3)
    ControlarDuplicados (curCtasContProy.campo(0) & " @ " & curCtasContProy.campo(1))
    NumeroCtaContableProyecto = curCtasContProy.campo(0)
    
    ' AVERIGUAMOS LA CUENTA CONTABLE DE LA CUENTA BANCARIA
    sSQL = ""
    sSQL = "SELECT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA, TIPO_CUENTASBANC.IDCTA, TIPO_CUENTASBANC.DESCCTA " _
      & "FROM TIPO_CUENTASBANC, PLAN_CONTABLE " _
      & "WHERE (TIPO_CUENTASBANC.CODCONT = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(TIPO_CUENTASBANC.IDCTA = '" & IdCtaBancaria & "')"
    
    ' Ejecuta la consulta
    curCtasContCtaBancaria.SQL = sSQL
  
    If curCtasContCtaBancaria.Abrir = HAY_ERROR Then
      End
    End If
    
    If Not curCtasContCtaBancaria.EOF Then
      ControlarDuplicados (curCtasContCtaBancaria.campo(0) & " @ " & curCtasContCtaBancaria.campo(1))
      NumeroCtaContableBancaria = curCtasContCtaBancaria.campo(0)
    End If
  End If
      
  curCtasContProy.Cerrar
  curCtasContCtaBancaria.Cerrar
  
  ' -----*****-----*****-----*****
  ' RECUPERANDO LOS ASIENTOS MANUALES
  ' -----*****-----*****-----*****
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS.NUMASIENTO, CTB_ASIENTOS.PROCORIGEN, CTB_ASIENTOS_DET.CODCONTABLE " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS.PROCORIGEN = 'AM') AND (CTB_ASIENTOS.ANULADO = 'NO') AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAMD(mskFechaFin.Text) & "') "
      
  ' Ejecuta la consulta
  curNumeroAsientoManual.SQL = sSQL

  If curNumeroAsientoManual.Abrir = HAY_ERROR Then
    End
  End If
    
  AsientosManuales = ""
  If Not curNumeroAsientoManual.EOF Then
    Do While Not curNumeroAsientoManual.EOF
      AsientosManuales = AsientosManuales & curNumeroAsientoManual.campo(0) & "@"
      
      curNumeroAsientoManual.MoverSiguiente
    Loop
  End If
  
  InstrucCtasContablesManuales = ""
  If Len(AsientosManuales) > 0 Then
    Do While InStr(1, AsientosManuales, "@")
      InstrucCtasContablesManuales = InstrucCtasContablesManuales & "CTB_ASIENTOS_DET.NUMASIENTO = '" & Left(AsientosManuales, 10) & "' OR "
      AsientosManuales = Mid(AsientosManuales, 12, Len(AsientosManuales))
    Loop
    
    InstrucCtasContablesManuales = Left(InstrucCtasContablesManuales, Len(InstrucCtasContablesManuales) - 4)
  End If
  
  If Len(InstrucCtasContablesManuales) > 26 Then
    ' AVERIGUAMOS LAS CUENTAS CONTABLES DE LOS ASIENTOS MANUALES
    sSQL = ""
    sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
      & "FROM CTB_ASIENTOS_DET, PLAN_CONTABLE " _
      & "WHERE (CTB_ASIENTOS_DET.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(" & InstrucCtasContablesManuales & ") "
      
    ' Ejecuta la consulta
    curCtasContManuales.SQL = sSQL
  
    If curCtasContManuales.Abrir = HAY_ERROR Then
      End
    End If
    
    If Not curCtasContManuales.EOF Then
      Do While Not curCtasContManuales.EOF
        ControlarDuplicados (curCtasContManuales.campo(0) & " @ " & curCtasContManuales.campo(1))
  
        curCtasContManuales.MoverSiguiente
      Loop
    End If
  
    curCtasContManuales.Cerrar
  End If
  
  curNumeroAsientoManual.Cerrar
End Sub

Sub CtasContablesPlanilla()
  Dim curPersonalProy As New clsBD2
  Dim curCtasContPlanilla As New clsBD2
  Dim PersonalDelProyecto As String
     
  sSQL = ""
  sSQL = "SELECT DISTINCT PLN_PERSONAL_HIST.IDPERSONA, MIN(PLN_PERSONAL_HIST.CODPLANILLA) " _
      & "FROM PLN_PERSONAL_HIST " _
      & "WHERE (PLN_PERSONAL_HIST.PROYQUEPAGA = '" & Trim(txtProyecto.Text) & "') " _
      & "GROUP BY PLN_PERSONAL_HIST.IDPERSONA " _
      & "ORDER BY PLN_PERSONAL_HIST.IDPERSONA "

  ' Ejecuta la consulta
  curPersonalProy.SQL = sSQL

  If curPersonalProy.Abrir = HAY_ERROR Then
    End
  End If
  
  PersonalDelProyecto = ""
  If Not curPersonalProy.EOF Then
    Do While Not curPersonalProy.EOF
      PersonalDelProyecto = PersonalDelProyecto & curPersonalProy.campo(0) & "@"
      PersonalDelProyectoSinBorrar = PersonalDelProyecto
      
      FechaInicioPlanilla = curPersonalProy.campo(1)
      
      curPersonalProy.MoverSiguiente
    Loop
  End If
  
  InstrucPersonal = ""
  Do While InStr(1, PersonalDelProyecto, "@")
    InstrucPersonal = InstrucPersonal & "PLN_VALOR_NOMINA.IDPERSONA = '" & Left(PersonalDelProyecto, 4) & "' OR "
    PersonalDelProyecto = Mid(PersonalDelProyecto, 6, Len(PersonalDelProyecto))
  Loop
  
  InstrucPersonal = Left(InstrucPersonal, Len(InstrucPersonal) - 4)
  
  If Left(FechaFinConsulta, 6) >= FechaInicioPlanilla Then
    sSQL = ""
    sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
        & "FROM PLAN_CONTABLE, PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
        & "(" & InstrucPersonal & ") AND " _
        & "(PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & FechaInicioPlanilla & "' AND '" & Left(FechaFinConsulta, 6) & "') "
  
    ' Ejecuta la consulta
    curCtasContPlanilla.SQL = sSQL
  
    If curCtasContPlanilla.Abrir = HAY_ERROR Then
      End
    End If
    
    If Not curCtasContPlanilla.EOF Then
      Do While Not curCtasContPlanilla.EOF
        If (curCtasContPlanilla.campo(0) = "14111") Or (curCtasContPlanilla.campo(0) = "38511") Then
        
        Else
          ControlarDuplicados (curCtasContPlanilla.campo(0) & " @ " & curCtasContPlanilla.campo(1))
          
        End If
        
        curCtasContPlanilla.MoverSiguiente
      Loop
    End If
    curCtasContPlanilla.Cerrar
  End If
  ControlarDuplicados ("40311" & " @ " & "ESSALUD")
  ControlarDuplicados ("41111" & " @ " & "REMUNERACIONES POR PAGAR")
  ControlarDuplicados ("47111" & " @ " & "COMPENSACIÓN POR TIEMPO DE SERVICIOS")
  ControlarDuplicados ("68611" & " @ " & "COMPENSACIÓN TIEMPO DE SERVICIOS")
  
  curPersonalProy.Cerrar
End Sub

Sub CtasContablesGasto()
  GastoConAfectacion
  
  GastoSinAfectacion
End Sub

Sub GastoConAfectacion()
  Dim curProdServEgresoProy As New clsBD2
  Dim curCtasContProdServ As New clsBD2
  
  sSQL = ""
  sSQL = "SELECT EGRESOS.ORDEN, GASTOS.CONCEPTO, GASTOS.CODCONCEPTO " _
      & "FROM EGRESOS, GASTOS " _
      & "WHERE (EGRESOS.ORDEN = GASTOS.ORDEN) AND (EGRESOS.IDPROY = '" & Trim(txtProyecto.Text) & "') AND " _
      & "(EGRESOS.FECMOV BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaFinConsulta & "') " _
      & "ORDER BY EGRESOS.FECMOV, EGRESOS.ORDEN "

  ' Ejecuta la consulta
  curProdServEgresoProy.SQL = sSQL

  If curProdServEgresoProy.Abrir = HAY_ERROR Then
    End
  End If
  
  If Not curProdServEgresoProy.EOF Then
    Do While Not curProdServEgresoProy.EOF
      sSQL = ""
      If curProdServEgresoProy.campo(1) = "P" Then      ' PRODUCTOS
        sSQL = ""
        sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
          & "FROM PLAN_CONTABLE, PRODUCTOS " _
          & "WHERE (PRODUCTOS.CODCONT = PLAN_CONTABLE.CODCONTABLE) AND " _
          & "(PRODUCTOS.IDPROD = '" & curProdServEgresoProy.campo(2) & "') "
      
      Else                                              ' SERVICIOS
        sSQL = ""
        sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
          & "FROM PLAN_CONTABLE, SERVICIOS " _
          & "WHERE (SERVICIOS.CODCONT = PLAN_CONTABLE.CODCONTABLE) AND " _
          & "(SERVICIOS.IDSERV = '" & curProdServEgresoProy.campo(2) & "') "
      End If
      
      ' Ejecuta la consulta
      curCtasContProdServ.SQL = sSQL
    
      If curCtasContProdServ.Abrir = HAY_ERROR Then
        End
      End If
      
      If Not curCtasContProdServ.EOF Then
          ControlarDuplicados (curCtasContProdServ.campo(0) & " @ " & curCtasContProdServ.campo(1))
      End If
      
      curProdServEgresoProy.MoverSiguiente
    Loop
  End If
  
  curProdServEgresoProy.Cerrar
  curCtasContProdServ.Cerrar
End Sub

Sub GastoSinAfectacion()
  Dim curCtasContEgreso As New clsBD2
  
  sSQL = ""
  sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
      & "FROM EGRESOS, PLAN_CONTABLE " _
      & "WHERE (EGRESOS.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(EGRESOS.IDPROY = ' ') AND (EGRESOS.IDCTA = '" & IdCtaBancaria & "') AND " _
      & "(EGRESOS.CODCONTABLE <> ' ') AND " _
      & "(EGRESOS.FECMOV BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaFinConsulta & "') "

  ' Ejecuta la consulta
  curCtasContEgreso.SQL = sSQL

  If curCtasContEgreso.Abrir = HAY_ERROR Then
    End
  End If
  
  If Not curCtasContEgreso.EOF Then
    ControlarDuplicados (curCtasContEgreso.campo(0) & " @ " & curCtasContEgreso.campo(1))
  End If
  
  curCtasContEgreso.Cerrar
End Sub

Sub ControlarDuplicados(CtaContable As String)
  Dim i As Integer
  Dim Existe As Boolean
   
  Existe = False
  ' recorre el grid
  For i = 0 To ComboCtasContables.ListCount - 1
    If ComboCtasContables.List(i) = CtaContable Then
      Existe = True
    End If
  Next i
    
  If Existe Then
    'MsgBox ("La Cta Contable ya fue agregada, Ingrese otro.")
  Else
    ComboCtasContables.AddItem (CtaContable)
  End If
End Sub

Sub CtaPrestamos()
  Dim i As Integer
  Dim Existe As Boolean
  Dim ConsultaPersonal As String
  Dim PersonalDelProyecto As String
  
  PersonalDelProyecto = PersonalDelProyectoSinBorrar
  ConsultaPersonal = ""
  Do While InStr(1, PersonalDelProyecto, "@")
    ConsultaPersonal = ConsultaPersonal & "PRESTAMOS_CUOTAS.IDPERSONA = '" & Left(PersonalDelProyecto, 4) & "' OR "
    PersonalDelProyecto = Mid(PersonalDelProyecto, 6, Len(PersonalDelProyecto))
  Loop
  
  ConsultaPersonal = Left(ConsultaPersonal, Len(ConsultaPersonal) - 4)

  sSQL = ""
  sSQL = "SELECT SUM(PRESTAMOS_CUOTAS.CUOTA) " _
      & "FROM PRESTAMOS_CUOTAS " _
      & "WHERE (" & ConsultaPersonal & ") AND " _
      & "(PRESTAMOS_CUOTAS.ANIOMES >= '" & FechaInicioPlanilla & "') "

  ' Ejecuta la consulta
  curTotalPrestamo.SQL = sSQL

  If curTotalPrestamo.Abrir = HAY_ERROR Then
    End
  End If
  
  sSQL = ""
  sSQL = "SELECT SUM(PRESTAMOS_CUOTAS.CUOTA) " _
      & "FROM PRESTAMOS_CUOTAS " _
      & "WHERE (PRESTAMOS_CUOTAS.CANCELADO = 'SI') AND " _
      & "(" & ConsultaPersonal & ") AND " _
      & "(PRESTAMOS_CUOTAS.ANIOMES >= '" & FechaInicioPlanilla & "') "

  ' Ejecuta la consulta
  curCanceladoPrestamo.SQL = sSQL

  If curCanceladoPrestamo.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Sub CtaAfpEssaludRenta5ta(CtaContableConsultar As String)
  Dim ConsultaPersonal As String
  Dim PersonalDelProyecto As String
  
  PersonalDelProyecto = PersonalDelProyectoSinBorrar
  ConsultaPersonal = ""
  Do While InStr(1, PersonalDelProyecto, "@")
    ConsultaPersonal = ConsultaPersonal & "PLN_VALOR_NOMINA.IDPERSONA = '" & Left(PersonalDelProyecto, 4) & "' OR "
    PersonalDelProyecto = Mid(PersonalDelProyecto, 6, Len(PersonalDelProyecto))
  Loop
  
  ConsultaPersonal = Left(ConsultaPersonal, Len(ConsultaPersonal) - 4)

  sSQL = ""
  sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
      & "FROM PLN_VALOR_NOMINA " _
      & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & FechaInicioPlanilla & "' AND '" & Left(FechaAnteriorFinconsulta, 6) & "') AND " _
      & "(" & ConsultaPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '" & CtaContableConsultar & "') "
      
  ' Ejecuta la consulta
  curAfpEssaludRenta5ta.SQL = sSQL

  If curAfpEssaludRenta5ta.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Public Sub CtasContablesContraCuentas()
'----------------------------------------------------------------------------
'PROPÓSITO: Carga las Contr Cuentas
'Recibe:    nada
'Devuelve:  nada
'----------------------------------------------------------------------------
  Dim curContraCta01 As New clsBD2
  Dim curContraCta02 As New clsBD2
  Dim i As Integer
  
  'Recupera financiera del proyecto seleccionado
  sSQL = ""
  sSQL = "SELECT CE.CodContable, EE.Componente, EE.CodContable " & _
         "FROM (CTB_ENLACE CE LEFT JOIN CTB_ESPECIF_ENLACE EE ON CE.IdEnlace=EE.IdEnlace) " & _
         "ORDER BY CE.CodContable "
         
  curCtasContablesContraCuentas.SQL = sSQL
  
  ' ejecuta la consulta
  If curCtasContablesContraCuentas.Abrir = HAY_ERROR Then
    End
  End If
  
  If Not curCtasContablesContraCuentas.EOF Then
    For i = 0 To ComboCtasContables.ListCount - 1
      curCtasContablesContraCuentas.MoverPrimero
      Do While Not curCtasContablesContraCuentas.EOF
        If Left(ComboCtasContables.List(i), InStr(1, ComboCtasContables.List(i), "@") - 2) = curCtasContablesContraCuentas.campo(0) Then
          sSQL = ""
          sSQL = "SELECT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
            & "FROM PLAN_CONTABLE " _
            & "WHERE (PLAN_CONTABLE.CODCONTABLE = '" & curCtasContablesContraCuentas.campo(1) & "') "
          
          ' Ejecuta la consulta
          curContraCta01.SQL = sSQL
        
          If curContraCta01.Abrir = HAY_ERROR Then
            End
          End If
          
          If Not curContraCta01.EOF Then
              ControlarDuplicados (curContraCta01.campo(0) & " @ " & curContraCta01.campo(1))
          End If
          
          curContraCta01.Cerrar
          
          sSQL = ""
          sSQL = "SELECT PLAN_CONTABLE.CODCONTABLE, PLAN_CONTABLE.DESCCUENTA " _
            & "FROM PLAN_CONTABLE " _
            & "WHERE (PLAN_CONTABLE.CODCONTABLE = '" & curCtasContablesContraCuentas.campo(2) & "') "
          
          ' Ejecuta la consulta
          curContraCta02.SQL = sSQL
        
          If curContraCta02.Abrir = HAY_ERROR Then
            End
          End If
          
          If Not curContraCta02.EOF Then
              ControlarDuplicados (curContraCta02.campo(0) & " @ " & curContraCta02.campo(1))
          End If
          
          curContraCta02.Cerrar
        End If
        curCtasContablesContraCuentas.MoverSiguiente
      Loop
    Next i
  End If
End Sub



Sub ParaProyectoSumasMayor()
'  /*/*/*/*/*/*/*/    I N G R E S O S

  ' OBTENEMOS LA SUMA DE LOS INGRESOS EN EL PROYECTO DEPOSITOS EN LA CUENTA BANCARIA
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
    & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
    & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
    & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
    & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAMD(mskFechaFin.Text) & "') " _
    & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
    & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "

  ' Ejecuta la consulta
  mcurSumIngresosCorte.SQL = sSQL

  If mcurSumIngresosCorte.Abrir = HAY_ERROR Then
    End
  End If

'  /*/*/*/*/*/*/*/  E G R E S O S

  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
    & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
    & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
    & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableBancaria & "') AND " _
    & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAMD(mskFechaFin.Text) & "') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0391') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0906CB0410') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0402') " _
    & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
    & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "

  ' Ejecuta la consulta
  mcurSumGastosCorte.SQL = sSQL

  If mcurSumGastosCorte.Abrir = HAY_ERROR Then
    End
  End If
  
'  /*/*/*/*/*/*/*/  A S I E N T O S   M A N U A L E S    C U E N T A    44113
  
  ExistenAsientosManualesMayor = False
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS.NUMASIENTO, CTB_ASIENTOS.PROCORIGEN, CTB_ASIENTOS_DET.CODCONTABLE " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS.PROCORIGEN = 'AM') AND (CTB_ASIENTOS.ANULADO = 'NO') AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAMD(mskFechaFin.Text) & "') "

  ' Ejecuta la consulta
  curNumeroAsientoManual.SQL = sSQL

  If curNumeroAsientoManual.Abrir = HAY_ERROR Then
    End
  End If
  
  AsientosManuales = ""
  If Not curNumeroAsientoManual.EOF Then
    Do While Not curNumeroAsientoManual.EOF
      AsientosManuales = AsientosManuales & curNumeroAsientoManual.campo(0) & "@"
      
      curNumeroAsientoManual.MoverSiguiente
    Loop
  End If
  
  If Len(AsientosManuales) > 0 Then
    InstrucCtasContablesManuales = ""
    Do While InStr(1, AsientosManuales, "@")
      InstrucCtasContablesManuales = InstrucCtasContablesManuales & "CTB_ASIENTOS_DET.NUMASIENTO = '" & Left(AsientosManuales, 10) & "' OR "
      AsientosManuales = Mid(AsientosManuales, 12, Len(AsientosManuales))
    Loop
    
    InstrucCtasContablesManuales = Left(InstrucCtasContablesManuales, Len(InstrucCtasContablesManuales) - 4)
  End If
  
  If Len(InstrucCtasContablesManuales) > 26 Then
    ExistenAsientosManualesMayor = True
    sSQL = ""
    sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
        & "FROM CTB_ASIENTOS_DET " _
        & "WHERE (" & InstrucCtasContablesManuales & ") AND " _
        & "(CTB_ASIENTOS_DET.CODCONTABLE <> '" & NumeroCtaContableProyecto & "') " _
        & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
        & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "
  
    ' Ejecuta la consulta
    mcurAsientoManualProyectoCorte.SQL = sSQL
  
    If mcurAsientoManualProyectoCorte.Abrir = HAY_ERROR Then
      End
    End If
  End If
    
  curNumeroAsientoManual.Cerrar
End Sub

Sub ParaPlanillaSumasMayor()
  sSQL = ""
  sSQL = "SELECT PLN_VALOR_NOMINA.CODCONTABLE, SUM(PLN_VALOR_NOMINA.VALOR) " _
      & "FROM PLN_VALOR_NOMINA " _
      & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & FechaInicioPlanilla & "' AND '" & Left(FechaFinConsulta, 6) & "') AND " _
      & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') " _
      & "GROUP BY PLN_VALOR_NOMINA.CODCONTABLE " _
      & "ORDER BY PLN_VALOR_NOMINA.CODCONTABLE "
      
  ' Ejecuta la consulta
  mcurSumConceptosPlanillaCorte.SQL = sSQL

  If mcurSumConceptosPlanillaCorte.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Sub ParaGastoSumasMayor()
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS    C O N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  sSQL = ""
  sSQL = "SELECT PR.CODCONT & SV.CODCONT, SUM(G.MONTO) " _
       & "FROM ((EGRESOS E INNER  JOIN TIPO_DOCUM D ON E.IdTipoDoc=D.IdTipoDoc) " _
       & "INNER JOIN PROVEEDORES P ON E.IdProveedor=P.IdProveedor) " _
       & "INNER JOIN ((GASTOS G LEFT JOIN PRODUCTOS PR ON G.CodConcepto=PR.IdProd) " _
       & "LEFT JOIN SERVICIOS SV ON G.CodConcepto=SV.IdServ) ON E.Orden=G.Orden " _
       & "WHERE E.idProy='" & Trim(txtProyecto.Text) & "' and " _
       & "E.Anulado = 'NO' and (E.FecMov BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaFinConsulta & "') " _
       & "GROUP BY PR.CODCONT & SV.CODCONT " _
       & "ORDER BY PR.CODCONT & SV.CODCONT "
  
  ' Ejecuta la sentencia
  mcurGastosConAfectacionCorte.SQL = sSQL
  
  If mcurGastosConAfectacionCorte.Abrir = HAY_ERROR Then
    End
  End If
  
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS     S I N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  
  sSQL = ""
  sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, EGRESOS.MONTOCB " _
      & "FROM EGRESOS, PLAN_CONTABLE " _
      & "WHERE (EGRESOS.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(EGRESOS.IDPROY = ' ') AND (EGRESOS.IDCTA = '" & IdCtaBancaria & "') AND " _
      & "(EGRESOS.CODCONTABLE <> ' ') AND " _
      & "(EGRESOS.FECMOV BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaFinConsulta & "') "

  ' Ejecuta la consulta
  mcurGastosSinAfectacionCorte.SQL = sSQL

  If mcurGastosSinAfectacionCorte.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Sub ParaProyectoMesAnterior()
  'AnioAnteriorFinconsulta = 2009
  'MesAnteriorFinconsulta = 07
  'DiaAnteriorFinconsulta = 31
  'FechaAnteriorFinconsulta = 20090731
  
'  /*/*/*/*/*/*/*/    I N G R E S O S
  
  ExistenProyectoIngresoMesAnterior = False
  If FechaAnteriorFinconsulta >= FechaAMD(TxtFechaInicioProyecto.Text) Then
    ExistenProyectoIngresoMesAnterior = True
    ' OBTENEMOS LA SUMA DE LOS INGRESOS EN EL PROYECTO DEPOSITOS EN LA CUENTA BANCARIA
    sSQL = ""
    sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAnteriorFinconsulta & "') " _
      & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
      & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "
  
    ' Ejecuta la consulta
    mcurSumIngresosAnt.SQL = sSQL
  
    If mcurSumIngresosAnt.Abrir = HAY_ERROR Then
      End
    End If
  End If

'  /*/*/*/*/*/*/*/  E G R E S O S
  ExistenProyectoEgresoMesAnterior = False
  If FechaAnteriorFinconsulta >= FechaAMD(TxtFechaInicioProyecto.Text) Then
    ExistenProyectoEgresoMesAnterior = True
    sSQL = ""
    sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableBancaria & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAnteriorFinconsulta & "') AND " _
      & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0391') AND " _
      & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0906CB0410') AND " _
      & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0402') " _
      & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
      & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "
  
    ' Ejecuta la consulta
    mcurSumGastosAnt.SQL = sSQL
  
    If mcurSumGastosAnt.Abrir = HAY_ERROR Then
      End
    End If
  End If
'  /*/*/*/*/*/*/*/  A S I E N T O S   M A N U A L E S    C U E N T A    44113
  
  ExistenAsientosManualesMesAnterior = False
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS.NUMASIENTO, CTB_ASIENTOS.PROCORIGEN, CTB_ASIENTOS_DET.CODCONTABLE " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS.PROCORIGEN = 'AM') AND (CTB_ASIENTOS.ANULADO = 'NO') AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & FechaAMD(TxtFechaInicioProyecto.Text) & "' AND '" & FechaAnteriorFinconsulta & "') "

  ' Ejecuta la consulta
  curNumeroAsientoManual.SQL = sSQL

  If curNumeroAsientoManual.Abrir = HAY_ERROR Then
    End
  End If
  
  AsientosManuales = ""
  If Not curNumeroAsientoManual.EOF Then
    Do While Not curNumeroAsientoManual.EOF
      AsientosManuales = AsientosManuales & curNumeroAsientoManual.campo(0) & "@"
      
      curNumeroAsientoManual.MoverSiguiente
    Loop
  End If
  
  InstrucCtasContablesManuales = ""
  If Len(AsientosManuales) > 0 Then
    Do While InStr(1, AsientosManuales, "@")
      InstrucCtasContablesManuales = InstrucCtasContablesManuales & "CTB_ASIENTOS_DET.NUMASIENTO = '" & Left(AsientosManuales, 10) & "' OR "
      AsientosManuales = Mid(AsientosManuales, 12, Len(AsientosManuales))
    Loop
    
    InstrucCtasContablesManuales = Left(InstrucCtasContablesManuales, Len(InstrucCtasContablesManuales) - 4)
  End If
  
  If Len(InstrucCtasContablesManuales) > 26 Then
    ExistenAsientosManualesMesAnterior = True
    sSQL = ""
    sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
        & "FROM CTB_ASIENTOS_DET " _
        & "WHERE (" & InstrucCtasContablesManuales & ") AND " _
        & "(CTB_ASIENTOS_DET.CODCONTABLE <> '" & NumeroCtaContableProyecto & "') " _
        & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
        & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "
  
    ' Ejecuta la consulta
    mcurAsientoManualProyectoAnt.SQL = sSQL
  
    If mcurAsientoManualProyectoAnt.Abrir = HAY_ERROR Then
      End
    End If
  End If
    
  curNumeroAsientoManual.Cerrar
End Sub

Sub ParaPlanillaMesAnterior()
  ExistenPlanillaMesAnterior = False
  If Left(FechaAnteriorFinconsulta, 6) >= FechaInicioPlanilla Then
    ExistenPlanillaMesAnterior = True
    sSQL = ""
    sSQL = "SELECT PLN_VALOR_NOMINA.CODCONTABLE, SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & FechaInicioPlanilla & "' AND '" & Left(FechaAnteriorFinconsulta, 6) & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') " _
        & "GROUP BY PLN_VALOR_NOMINA.CODCONTABLE " _
        & "ORDER BY PLN_VALOR_NOMINA.CODCONTABLE "
        
    ' Ejecuta la consulta
    mcurSumConceptosPlanillaAnt.SQL = sSQL
  
    If mcurSumConceptosPlanillaAnt.Abrir = HAY_ERROR Then
      End
    End If
  End If
End Sub

Sub ParaGastoMesAnterior()
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS    C O N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  sSQL = ""
  sSQL = "SELECT PR.CODCONT & SV.CODCONT, SUM(G.MONTO) " _
       & "FROM ((EGRESOS E INNER  JOIN TIPO_DOCUM D ON E.IdTipoDoc=D.IdTipoDoc) " _
       & "INNER JOIN PROVEEDORES P ON E.IdProveedor=P.IdProveedor) " _
       & "INNER JOIN ((GASTOS G LEFT JOIN PRODUCTOS PR ON G.CodConcepto=PR.IdProd) " _
       & "LEFT JOIN SERVICIOS SV ON G.CodConcepto=SV.IdServ) ON E.Orden=G.Orden " _
       & "WHERE E.idProy='" & Trim(txtProyecto.Text) & "' and " _
       & "E.Anulado = 'NO' and (E.FecMov BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaAnteriorFinconsulta & "') " _
       & "GROUP BY PR.CODCONT & SV.CODCONT " _
       & "ORDER BY PR.CODCONT & SV.CODCONT "
  
  ' Ejecuta la sentencia
  mcurGastosConAfectacionAnt.SQL = sSQL
  
  If mcurGastosConAfectacionAnt.Abrir = HAY_ERROR Then
    End
  End If
  
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS     S I N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  
  sSQL = ""
  sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, EGRESOS.MONTOCB " _
      & "FROM EGRESOS, PLAN_CONTABLE " _
      & "WHERE (EGRESOS.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(EGRESOS.IDPROY = ' ') AND (EGRESOS.IDCTA = '" & IdCtaBancaria & "') AND " _
      & "(EGRESOS.CODCONTABLE <> ' ') AND " _
      & "(EGRESOS.FECMOV BETWEEN '" & FechaInicioConsulta & "' AND '" & FechaAnteriorFinconsulta & "') "

  ' Ejecuta la consulta
  mcurGastosSinAfectacionAnt.SQL = sSQL

  If mcurGastosSinAfectacionAnt.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Sub ParaProyectoMesActual()
  'AnioAnteriorFinconsulta = 2009
  'MesAnteriorFinconsulta = 07
  'DiaAnteriorFinconsulta = 31
  'FechaAnteriorFinconsulta = 20090731
  
  'FechaInicioConsulta = FechaAMD(mskFechaIni.Text)   20090521
  'FechaFinConsulta = FechaAMD(mskFechaFin.Text)      20090831
  
'  /*/*/*/*/*/*/*/    I N G R E S O S

  ' OBTENEMOS LA SUMA DE LOS INGRESOS EN EL PROYECTO DEPOSITOS EN LA CUENTA BANCARIA
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
    & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
    & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
    & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
    & "(CTB_ASIENTOS.FECHA BETWEEN '" & Left(FechaFinConsulta, 6) & "01" & "' AND '" & FechaFinConsulta & "') " _
    & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
    & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "

  ' Ejecuta la consulta
  mcurSumIngresosAct.SQL = sSQL

  If mcurSumIngresosAct.Abrir = HAY_ERROR Then
    End
  End If

'  /*/*/*/*/*/*/*/  E G R E S O S

  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
    & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
    & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
    & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableBancaria & "') AND " _
    & "(CTB_ASIENTOS.FECHA BETWEEN '" & Left(FechaFinConsulta, 6) & "01" & "' AND '" & FechaFinConsulta & "') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0391') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0906CB0410') AND " _
    & "(CTB_ASIENTOS_DET.NUMASIENTO <> '0907CB0402') " _
    & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
    & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "

  ' Ejecuta la consulta
  mcurSumGastosAct.SQL = sSQL

  If mcurSumGastosAct.Abrir = HAY_ERROR Then
    End
  End If
  
'  /*/*/*/*/*/*/*/  A S I E N T O S   M A N U A L E S    C U E N T A    44113
  
  ExistenAsientosManualesMesActual = False
  sSQL = ""
  sSQL = "SELECT CTB_ASIENTOS.NUMASIENTO, CTB_ASIENTOS.PROCORIGEN, CTB_ASIENTOS_DET.CODCONTABLE " _
      & "FROM CTB_ASIENTOS, CTB_ASIENTOS_DET " _
      & "WHERE (CTB_ASIENTOS.NUMASIENTO = CTB_ASIENTOS_DET.NUMASIENTO) AND " _
      & "(CTB_ASIENTOS.PROCORIGEN = 'AM') AND (CTB_ASIENTOS.ANULADO = 'NO') AND " _
      & "(CTB_ASIENTOS_DET.CODCONTABLE = '" & NumeroCtaContableProyecto & "') AND " _
      & "(CTB_ASIENTOS.FECHA BETWEEN '" & Left(FechaFinConsulta, 6) & "01" & "' AND '" & FechaFinConsulta & "') "

  ' Ejecuta la consulta
  curNumeroAsientoManual.SQL = sSQL

  If curNumeroAsientoManual.Abrir = HAY_ERROR Then
    End
  End If
  
  AsientosManuales = ""
  If Not curNumeroAsientoManual.EOF Then
    Do While Not curNumeroAsientoManual.EOF
      AsientosManuales = AsientosManuales & curNumeroAsientoManual.campo(0) & "@"
      
      curNumeroAsientoManual.MoverSiguiente
    Loop
  End If
  
  InstrucCtasContablesManuales = ""
  If Len(AsientosManuales) > 0 Then
    Do While InStr(1, AsientosManuales, "@")
      InstrucCtasContablesManuales = InstrucCtasContablesManuales & "CTB_ASIENTOS_DET.NUMASIENTO = '" & Left(AsientosManuales, 10) & "' OR "
      AsientosManuales = Mid(AsientosManuales, 12, Len(AsientosManuales))
    Loop
    
    InstrucCtasContablesManuales = Left(InstrucCtasContablesManuales, Len(InstrucCtasContablesManuales) - 4)
  End If
  
  If Len(InstrucCtasContablesManuales) > 26 Then
    ExistenAsientosManualesMesActual = True
    sSQL = ""
    sSQL = "SELECT CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER, SUM(CTB_ASIENTOS_DET.MONTO) " _
        & "FROM CTB_ASIENTOS_DET " _
        & "WHERE (" & InstrucCtasContablesManuales & ") AND " _
        & "(CTB_ASIENTOS_DET.CODCONTABLE <> '" & NumeroCtaContableProyecto & "') " _
        & "GROUP BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER " _
        & "ORDER BY CTB_ASIENTOS_DET.CODCONTABLE, CTB_ASIENTOS_DET.DEBEHABER "
  
    ' Ejecuta la consulta
    mcurAsientoManualProyectoAct.SQL = sSQL
  
    If mcurAsientoManualProyectoAct.Abrir = HAY_ERROR Then
      End
    End If
  End If
    
  curNumeroAsientoManual.Cerrar
End Sub

Sub ParaPlanillaMesActual()
  ExistenPlanillaMesActual = False
  If FechaInicioPlanilla <= Left(FechaFinConsulta, 6) Then
    ExistenPlanillaMesActual = True
    sSQL = ""
    sSQL = "SELECT PLN_VALOR_NOMINA.CODCONTABLE, SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 6) & "' AND '" & Left(FechaFinConsulta, 6) & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') " _
        & "GROUP BY PLN_VALOR_NOMINA.CODCONTABLE " _
        & "ORDER BY PLN_VALOR_NOMINA.CODCONTABLE "
        
    ' Ejecuta la consulta
    mcurSumConceptosPlanillaAct.SQL = sSQL
  
    If mcurSumConceptosPlanillaAct.Abrir = HAY_ERROR Then
      End
    End If
  End If
End Sub

Sub ParaGastoMesActual()
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS    C O N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  sSQL = ""
  sSQL = "SELECT PR.CODCONT & SV.CODCONT, SUM(G.MONTO) " _
       & "FROM ((EGRESOS E INNER  JOIN TIPO_DOCUM D ON E.IdTipoDoc=D.IdTipoDoc) " _
       & "INNER JOIN PROVEEDORES P ON E.IdProveedor=P.IdProveedor) " _
       & "INNER JOIN ((GASTOS G LEFT JOIN PRODUCTOS PR ON G.CodConcepto=PR.IdProd) " _
       & "LEFT JOIN SERVICIOS SV ON G.CodConcepto=SV.IdServ) ON E.Orden=G.Orden " _
       & "WHERE E.idProy='" & Trim(txtProyecto.Text) & "' and " _
       & "E.Anulado = 'NO' and (E.FecMov BETWEEN '" & Left(FechaFinConsulta, 6) & "01" & "' AND '" & FechaFinConsulta & "') " _
       & "GROUP BY PR.CODCONT & SV.CODCONT " _
       & "ORDER BY PR.CODCONT & SV.CODCONT "
  
  ' Ejecuta la sentencia
  mcurGastosConAfectacionAct.SQL = sSQL
  
  If mcurGastosConAfectacionAct.Abrir = HAY_ERROR Then
    End
  End If
  
  '/*/*/*/*/*/*/*/*/*/*
  ' OBTENEMOS LA SUMA DE LOS GASTOS     S I N     AFECTACION PARA EL PROYECTO
  '/*/*/*/*/*/*/*/*/*/*
  
  sSQL = ""
  sSQL = "SELECT DISTINCT PLAN_CONTABLE.CODCONTABLE, EGRESOS.MONTOCB " _
      & "FROM EGRESOS, PLAN_CONTABLE " _
      & "WHERE (EGRESOS.CODCONTABLE = PLAN_CONTABLE.CODCONTABLE) AND " _
      & "(EGRESOS.IDPROY = ' ') AND (EGRESOS.IDCTA = '" & IdCtaBancaria & "') AND " _
      & "(EGRESOS.CODCONTABLE <> ' ') AND " _
      & "(EGRESOS.FECMOV BETWEEN '" & Left(FechaFinConsulta, 6) & "01" & "' AND '" & FechaFinConsulta & "') "

  ' Ejecuta la consulta
  mcurGastosSinAfectacionAct.SQL = sSQL

  If mcurGastosSinAfectacionAct.Abrir = HAY_ERROR Then
    End
  End If
End Sub

Sub TotalRemuneraciones(ColumnaDatos As Integer)
  Dim i As Integer
  
  TotalIngresos = 0
  TotalEgresos = 0
  
  For i = 0 To grdConsulta.Rows - 1
    'If InStr(1, Left(grdConsulta.TextMatrix(i, 0), 2), "14") Or InStr(1, grdConsulta.TextMatrix(i, 0), "40172") Or InStr(1, Left(grdConsulta.TextMatrix(i, 0), 2), "46") Or InStr(1, grdConsulta.TextMatrix(i, 0), "38511") Then
    If InStr(1, grdConsulta.TextMatrix(i, 0), "40172") Or InStr(1, Left(grdConsulta.TextMatrix(i, 0), 3), "469") Then
      'If InStr(1, Left(grdConsulta.TextMatrix(i, 0), 2), "14") Or InStr(1, grdConsulta.TextMatrix(i, 0), "40172") Or InStr(1, Left(grdConsulta.TextMatrix(i, 0), 2), "46") Then
      If InStr(1, grdConsulta.TextMatrix(i, 0), "40172") Or InStr(1, Left(grdConsulta.TextMatrix(i, 0), 3), "469") Then
        If grdConsulta.TextMatrix(i, ColumnaDatos + 1) <> "" Then
          TotalEgresos = TotalEgresos + grdConsulta.TextMatrix(i, ColumnaDatos + 1)
        End If
      Else
        If grdConsulta.TextMatrix(i, ColumnaDatos) <> "" Then
          TotalEgresos = TotalEgresos + grdConsulta.TextMatrix(i, ColumnaDatos)
        End If
      End If
      
    ElseIf InStr(1, Left(grdConsulta.TextMatrix(i, 0), 2), "62") Then
      If InStr(1, grdConsulta.TextMatrix(i, 0), "62711") Then
      
      Else
        If grdConsulta.TextMatrix(i, ColumnaDatos) <> "" Then
          TotalIngresos = TotalIngresos + grdConsulta.TextMatrix(i, ColumnaDatos)
        End If
      End If
    End If
  Next i
    
  TotalRemuneracionesPorPagar = TotalIngresos - TotalEgresos
  
  For i = 0 To grdConsulta.Rows - 1
    If grdConsulta.TextMatrix(i, 0) = "41111" Then
      grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalRemuneracionesPorPagar, "###,###,##0.00")
      grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalRemuneracionesPorPagar, "###,###,##0.00")
    End If
  Next i
End Sub

Sub CargaValoresContraCuentas(ColumnaDatos As Integer)
  Dim i, j, k As Integer
  Dim TotalManejado As Double
  
  TotalIngresos = 0
  TotalEgresos = 0

  For i = 0 To grdConsulta.Rows - 1
    Select Case grdConsulta.TextMatrix(i, 0)
      Case "26111"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60611" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26119"
        TotalManejado = 0
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60614" Or grdConsulta.TextMatrix(j, 0) = "60615" Or grdConsulta.TextMatrix(j, 0) = "60619" Then
            If grdConsulta.TextMatrix(j, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(j, ColumnaDatos)
            End If
          End If
        Next j
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
        
      Case "26211"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60621" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26212"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60622" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26311"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60631" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26313"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60632" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26312"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60633" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26319"
        TotalManejado = 0
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60634" Or grdConsulta.TextMatrix(j, 0) = "60635" Then
            If grdConsulta.TextMatrix(j, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(j, ColumnaDatos)
            End If
          End If
        Next j
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
      
      Case "26611"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60671" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26619"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60673" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26711"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60691" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26712"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60692" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26713"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60693" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      Case "26719"
        For j = i To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(j, 0) = "60699" Then
            grdConsulta.TextMatrix(i, ColumnaDatos) = grdConsulta.TextMatrix(j, ColumnaDatos)
            grdConsulta.TextMatrix(i, ColumnaDatos + 1) = grdConsulta.TextMatrix(j, ColumnaDatos)
          End If
        Next j
        
      
      
      Case "61611"
        TotalManejado = 0
        For k = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(k, 0) = "60611" Or grdConsulta.TextMatrix(k, 0) = "60614" Or grdConsulta.TextMatrix(k, 0) = "60615" Or grdConsulta.TextMatrix(k, 0) = "60619" Then
            If grdConsulta.TextMatrix(k, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(k, ColumnaDatos)
            End If
          End If
        Next k
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
      
      Case "61621"
        TotalManejado = 0
        For k = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(k, 0) = "60621" Or grdConsulta.TextMatrix(k, 0) = "60622" Then
            If grdConsulta.TextMatrix(k, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(k, ColumnaDatos)
            End If
          End If
        Next k
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
      
      Case "61631"
        TotalManejado = 0
        For k = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(k, 0) = "60631" Or grdConsulta.TextMatrix(k, 0) = "60632" Or grdConsulta.TextMatrix(k, 0) = "60633" Or grdConsulta.TextMatrix(k, 0) = "60634" Or grdConsulta.TextMatrix(k, 0) = "60635" Then
            If grdConsulta.TextMatrix(k, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(k, ColumnaDatos)
            End If
          End If
        Next k
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
    
      Case "61671"
        TotalManejado = 0
        For k = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(k, 0) = "60671" Or grdConsulta.TextMatrix(k, 0) = "60673" Then
            If grdConsulta.TextMatrix(k, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(k, ColumnaDatos)
            End If
          End If
        Next k
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
        
      Case "61691"
        TotalManejado = 0
        For k = 0 To grdConsulta.Rows - 1
          If grdConsulta.TextMatrix(k, 0) = "60691" Or grdConsulta.TextMatrix(k, 0) = "60692" Or grdConsulta.TextMatrix(k, 0) = "60693" Or grdConsulta.TextMatrix(k, 0) = "60699" Then
            If grdConsulta.TextMatrix(k, ColumnaDatos) <> "" Then
              TotalManejado = TotalManejado + grdConsulta.TextMatrix(k, ColumnaDatos)
            End If
          End If
        Next k
        
        grdConsulta.TextMatrix(i, ColumnaDatos) = Format(TotalManejado, "###,###,##0.00")
        grdConsulta.TextMatrix(i, ColumnaDatos + 1) = Format(TotalManejado, "###,###,##0.00")
            
    End Select
  Next i
End Sub

Sub ObtenerSaldo(ColumnaDatos As Integer, ColumnaUbicar As Integer)
  Dim i As Integer
  Dim valor As Double
  
  For i = 2 To grdConsulta.Rows - 1
    If Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos))) > Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos + 1))) Then
      grdConsulta.TextMatrix(i, ColumnaUbicar) = Format(Abs(Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos))) - Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos + 1)))), "###,###,##0.00")
      grdConsulta.TextMatrix(i, ColumnaUbicar + 1) = ""
    ElseIf Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos))) < Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos + 1))) Then
      grdConsulta.TextMatrix(i, ColumnaUbicar + 1) = Format(Abs(Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos))) - Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos + 1)))), "###,###,##0.00")
      grdConsulta.TextMatrix(i, ColumnaUbicar) = ""
    Else
      grdConsulta.TextMatrix(i, ColumnaUbicar) = ""
      grdConsulta.TextMatrix(i, ColumnaUbicar + 1) = ""
    End If
  Next i
End Sub

Sub ObtenerTotales(ColumnaDatos As Integer)
  Dim i As Integer
  Dim Total1 As Double
  Dim Total2 As Double
  
  Total1 = 0
  Total2 = 0
  For i = 2 To grdConsulta.Rows - 1
    If grdConsulta.TextMatrix(i, ColumnaDatos) <> "" Then
      Total1 = Total1 + Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos)))
    End If
    
    If grdConsulta.TextMatrix(i, ColumnaDatos + 1) <> "" Then
      Total2 = Total2 + Val(Var37(grdConsulta.TextMatrix(i, ColumnaDatos + 1)))
    End If
  Next i
    
  grdConsulta.TextMatrix(i - 1, ColumnaDatos) = Format(Total1, "###,###,##0.00")
  grdConsulta.TextMatrix(i - 1, ColumnaDatos + 1) = Format(Total2, "###,###,##0.00")
End Sub

Sub CalcularCTS()
  If Left(FechaFinConsulta, 6) < Left(FechaFinConsulta, 4) & "05" Then
    FechaFinConsulta = AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4)))
  End If
' /*/*/*/*/*/*/*/*/*
' Mayor
' /*/*/*/*/*/*/*/*/*
  
  TotalCTSCorte = 0
  TotalCTSAnt = 0
  TotalCTSAct = 0
  If (Left(FechaFinConsulta, 4) & "04" > Left(FechaInicioConsulta, 6)) And (Left(FechaFinConsulta, 4) & "04" < Left(FechaFinConsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE ABRIL
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "04" & "' AND '" & Left(FechaFinConsulta, 4) & "04" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSCorte.SQL = sSQL
  
    If mcurCTSCorte.Abrir = HAY_ERROR Then
      End
    End If
    
    ' OBTENIENDO LA GRATIFICACION DE DICIEMBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "' AND '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionCorte.SQL = sSQL
  
    If mcurGratificacionCorte.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSCorte = TotalCTSCorte + ((mcurCTSCorte.campo(0) + (1 / 6 * mcurGratificacionCorte.campo(0))) / 12) * 5
    
    mcurCTSCorte.Cerrar
    mcurGratificacionCorte.Cerrar
  End If
  
  If (Left(FechaFinConsulta, 4) & "10" > Left(FechaInicioConsulta, 6)) And (Left(FechaFinConsulta, 4) & "10" < Left(FechaFinConsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE OCTUBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "10" & "' AND '" & Left(FechaFinConsulta, 4) & "10" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSCorte.SQL = sSQL
  
    If mcurCTSCorte.Abrir = HAY_ERROR Then
      End
    End If
    
    ' OBTENIENDO LA GRATIFICACION DE JULIO
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "07" & "' AND '" & Left(FechaFinConsulta, 4) & "07" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionCorte.SQL = sSQL
  
    If mcurGratificacionCorte.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSCorte = TotalCTSCorte + ((mcurCTSCorte.campo(0) + (1 / 6 * mcurGratificacionCorte.campo(0))) / 12) * 5
    
    mcurCTSCorte.Cerrar
    mcurGratificacionCorte.Cerrar
  End If
  
' /*/*/*/*/*/*/*/*/*
' Mes Anterior
' /*/*/*/*/*/*/*/*/*

If (Left(FechaFinConsulta, 4) & "04" > Left(FechaInicioConsulta, 6)) And (Left(FechaFinConsulta, 4) & "04" < Left(FechaAnteriorFinconsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE ABRIL
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "04" & "' AND '" & Left(FechaFinConsulta, 4) & "04" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSAnt.SQL = sSQL
  
    If mcurCTSAnt.Abrir = HAY_ERROR Then
      End
    End If
    
    ' OBTENIENDO LA GRATIFICACION DE DICIEMBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "' AND '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionAnt.SQL = sSQL
  
    If mcurGratificacionAnt.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSAnt = TotalCTSAnt + ((mcurCTSAnt.campo(0) + (1 / 6 * mcurGratificacionAnt.campo(0))) / 12) * 5
    
    mcurCTSAnt.Cerrar
    mcurGratificacionAnt.Cerrar
  End If
  
  If (Left(FechaFinConsulta, 4) & "10" > Left(FechaInicioConsulta, 6)) And (Left(FechaFinConsulta, 4) & "10" < Left(FechaAnteriorFinconsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE OCTUBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "10" & "' AND '" & Left(FechaFinConsulta, 4) & "10" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSAnt.SQL = sSQL
  
    If mcurCTSAnt.Abrir = HAY_ERROR Then
      End
    End If
  
    ' OBTENIENDO LA GRATIFICACION DE JULIO
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "07" & "' AND '" & Left(FechaFinConsulta, 4) & "07" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionAnt.SQL = sSQL
  
    If mcurGratificacionAnt.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSAnt = TotalCTSAnt + ((mcurCTSAnt.campo(0) + (1 / 6 * mcurGratificacionAnt.campo(0))) / 12) * 5
    
    mcurCTSAnt.Cerrar
    mcurGratificacionAnt.Cerrar
  End If
  
' /*/*/*/*/*/*/*/*/*
' Mes Actual
' /*/*/*/*/*/*/*/*/*

If (Left(FechaFinConsulta, 4) & "04" = Left(FechaAnteriorFinconsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE ABRIL
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "04" & "' AND '" & Left(FechaFinConsulta, 4) & "04" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSAct.SQL = sSQL
  
    If mcurCTSAct.Abrir = HAY_ERROR Then
      End
    End If
  
    ' OBTENIENDO LA GRATIFICACION DE DICIEMBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "' AND '" & AnioMesAnterior(1, Val(Left(FechaFinConsulta, 4))) & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionAct.SQL = sSQL
  
    If mcurGratificacionAct.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSAct = TotalCTSAct + ((mcurCTSAct.campo(0) + (1 / 6 * mcurGratificacionAct.campo(0))) / 12) * 5
    
    mcurCTSAct.Cerrar
    mcurGratificacionAct.Cerrar
  End If
  
  If (Left(FechaFinConsulta, 4) & "10" = Left(FechaAnteriorFinconsulta, 6)) Then
    ' OBTENIENDO EL TOTAL DE INGRESOS DE OCTUBRE
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "10" & "' AND '" & Left(FechaFinConsulta, 4) & "10" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '14*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '38511') AND " _
        & "(PLN_VALOR_NOMINA.CODCONTABLE NOT LIKE '4*') AND (PLN_VALOR_NOMINA.CODCONTABLE <> '62711') "
        
    ' Ejecuta la consulta
    mcurCTSAct.SQL = sSQL
  
    If mcurCTSAct.Abrir = HAY_ERROR Then
      End
    End If
    
    ' OBTENIENDO LA GRATIFICACION DE JULIO
    sSQL = ""
    sSQL = "SELECT SUM(PLN_VALOR_NOMINA.VALOR) " _
        & "FROM PLN_VALOR_NOMINA " _
        & "WHERE (PLN_VALOR_NOMINA.CODPLANILLA BETWEEN '" & Left(FechaFinConsulta, 4) & "07" & "' AND '" & Left(FechaFinConsulta, 4) & "07" & "') AND " _
        & "(" & InstrucPersonal & ") AND (PLN_VALOR_NOMINA.CODCONTABLE = '62511') "
        
    ' Ejecuta la consulta
    mcurGratificacionAct.SQL = sSQL
  
    If mcurGratificacionAct.Abrir = HAY_ERROR Then
      End
    End If
    
    TotalCTSAct = TotalCTSAct + ((mcurCTSAct.campo(0) + (1 / 6 * mcurGratificacionAct.campo(0))) / 12) * 5
    
    mcurCTSAct.Cerrar
    mcurGratificacionAct.Cerrar
  End If
End Sub

Sub CargarCTS()
  Dim i As Integer
  
  For i = 0 To grdConsulta.Rows - 1
    If grdConsulta.TextMatrix(i, 0) = "47111" Then
      grdConsulta.TextMatrix(i, 2) = Format(TotalCTSCorte, "###,###,##0.00")
      grdConsulta.TextMatrix(i, 3) = Format(TotalCTSCorte, "###,###,##0.00")
      
      grdConsulta.TextMatrix(i, 4) = Format(TotalCTSAnt, "###,###,##0.00")
      grdConsulta.TextMatrix(i, 5) = Format(TotalCTSAnt, "###,###,##0.00")
      
      grdConsulta.TextMatrix(i, 6) = Format(TotalCTSAct, "###,###,##0.00")
      grdConsulta.TextMatrix(i, 7) = Format(TotalCTSAct, "###,###,##0.00")
    End If
    
    If grdConsulta.TextMatrix(i, 0) = "68611" Then
      grdConsulta.TextMatrix(i, 2) = Format(TotalCTSCorte, "###,###,##0.00")
      
      grdConsulta.TextMatrix(i, 4) = Format(TotalCTSAnt, "###,###,##0.00")
      
      grdConsulta.TextMatrix(i, 6) = Format(TotalCTSAct, "###,###,##0.00")
    End If
  Next i
End Sub
