VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCBConsulDiarioBancos 
   Caption         =   "SGCcaijo-Consulta Diario de Bancos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   71
   Icon            =   "SCCBConsulDiarioBancos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   360
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   400
      Left            =   8640
      TabIndex        =   4
      Top             =   8160
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10320
      TabIndex        =   5
      Top             =   8160
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedRows       =   0
      HighLight       =   0
      FillStyle       =   1
      MergeCells      =   4
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   960
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   11415
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   4935
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   330
            Left            =   1245
            TabIndex        =   0
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
            Left            =   3525
            TabIndex        =   1
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
            Caption         =   "Fecha &Fin:"
            Height          =   195
            Left            =   2640
            TabIndex        =   9
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha &Inicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   285
            Width           =   915
         End
      End
      Begin MSMask.MaskEdBox mskFecConsulta 
         Height          =   315
         Left            =   9960
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Fecha consulta:"
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComctlLib.ProgressBar prgInforme 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   8205
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   2160
      TabIndex        =   11
      Top             =   8160
      Width           =   1815
   End
End
Attribute VB_Name = "frmCBConsulDiarioBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Colecciones para la carga de la consulta
Private mcolSaldoIni As New Collection
Private mcolCtasBanc As New Collection
Private msCtaCaja As String

' Cursores para la carga de la consulta
Private mcurCtasBanc As New clsBD2

Private mcurIngresos As New clsBD2
Private mcurIngrTraslado As New clsBD2
Private mcurIngrDevPrest As New clsBD2
Private mcurIngrMovPers As New clsBD2
Private mcurIngrMovTerc As New clsBD2
Private mcurIngrAnulados As New clsBD2

Private mcurEgresos As New clsBD2
Private mcurEgreTraslado As New clsBD2
Private mcurEgreCAProd As New clsBD2
Private mcurEgreCAServ As New clsBD2
Private mcurEgreCAIngrImptRet As New clsBD2
Private mcurEgreAdelt As New clsBD2
Private mcurEgrePrest As New clsBD2
Private mcurEgrePlanll As New clsBD2
Private mcurEgreMovPers As New clsBD2
Private mcurEgreMovTerc As New clsBD2
Private mcurEgreAnulados As New clsBD2

'Variable de fondo de restriccion
Private mcurAsiDetraccion As New clsBD2

' Variable para la carga de la colección de saldos
Private mFecha As Variant
Dim mdblIngreso As Double
Dim mdblEgreso As Double
Public FechaAperturaCta As String
Public FechaAnulacionCta As String
' Variable para determinar el Saldo de FOndo de Restriccion
Dim SaldoDisponible As Double


Private Sub cmdInforme_Click()
Dim rptGastosProvDocDet As New clsBD4

' Deshabilita el botón aceptar
  cmdInforme.Enabled = False

' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón informe
       cmdInforme.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla consulta de boletas de pago
  LlenaTablaConsul

' Genera el reporte
' Formulario
  Set rptGastosProvDocDet.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptGastosProvDocDet.AsignarRpt

' Clausula WHERE de las relaciones del rpt.
  rptGastosProvDocDet.FiltroSelectionFormula = ""

  If CDate(mskFechaIni.Text) < CDate("01/01/2007") Or CDate(mskFechaFin.Text) < CDate("01/01/2007") Then
    ' Nombre del fichero
    rptGastosProvDocDet.NombreRPT = "rptCBDiarioBancosAntiguo.rpt"
  Else
    ' Nombre del fichero
    rptGastosProvDocDet.NombreRPT = "rptCBDiarioBancosNuevo.rpt"
  End If
  
' Presentación preliminar del Informe
  rptGastosProvDocDet.PresentancionPreliminar

' Elimina los datos de la tabla
  BorraDatosTablaConsul

' Elimina los datos de la BD
  Var43 gsFormulario

' Habilita el botón aceptar
  cmdInforme.Enabled = True
  
End Sub

Private Sub BorraDatosTablaConsul()
'------------------------------------------------------------
' Propósito: Borra los datos de las tablas RPTCBCONSULSALDOINICIALDIARIOCAJA y _
             RPTCBSALDOINICIALDIARIOBANCOS
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3

' Carga la sentencia
sSQL = "DELETE * FROM RPTCBSALDOINICIALDIARIOBANCOS"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Carga la sentencia
sSQL = "DELETE * FROM RPTCBDIARIOBANCOSDET"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

 'Carga la sentencia
sSQL = "DELETE * FROM RPTCBFONDORESTRICCION"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End


' Cierra la componente
modTablaConsul.Cerrar



End Sub

Private Sub LlenaTablaConsul()
'------------------------------------------------------------
' Propósito: LLena las tablas RPTCBSALDOINICIALDIARIOBANCOS y _
             RPTCBDIARIOBANCOSDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3
'Modulo para asiento de manuales(Fondo de retraccion)
Dim modTablaRetraccion As New clsBD3
Dim i As Long
' Recorre los datos del grid
For i = 0 To grdConsulta.Rows - 1
' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
  Ingreso , Egreso, Orden, CtaBanc, Var
            
    If grdConsulta.TextMatrix(i, 11) = "G" Then
     If Var37(grdConsulta.TextMatrix(i, 6)) = mcurAsiDetraccion.campo(10) Then
      ' Carga la sentencia sSQL para grabar los saldos iniciales de Bancos
      sSQL = "INSERT INTO RPTCBSALDOINICIALDIARIOBANCOS VALUES('" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & "','" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & grdConsulta.TextMatrix(i, 10) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 5)) & "','" _
        & grdConsulta.TextMatrix(i, 6) & "','" _
        & grdConsulta.TextMatrix(i, 4) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 8)) & ")"
       Else
       sSQL = "INSERT INTO RPTCBSALDOINICIALDIARIOBANCOS VALUES('" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & "','" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & grdConsulta.TextMatrix(i, 10) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 5)) & "','" _
        & grdConsulta.TextMatrix(i, 6) & "','" _
        & grdConsulta.TextMatrix(i, 4) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 8)) & ")"
       End If
      ' Ejecuta la sentencia
      modTablaConsul.SQL = sSQL
      If modTablaConsul.Ejecutar = HAY_ERROR Then End
      modTablaConsul.Cerrar
        
    ElseIf grdConsulta.TextMatrix(i, 11) = "D" Then
      ' Carga la sentencia sSQL para grabar los Movimientos de Bancos

      sSQL = "INSERT INTO RPTCBDIARIOBANCOSDET VALUES('" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & "','" _
        & FechaAMD(grdConsulta.TextMatrix(i, 0)) & grdConsulta.TextMatrix(i, 10) & "','" _
        & grdConsulta.TextMatrix(i, 9) & "','" _
        & grdConsulta.TextMatrix(i, 1) & "','" _
        & grdConsulta.TextMatrix(i, 2) & "','" _
        & grdConsulta.TextMatrix(i, 3) & "','" _
        & grdConsulta.TextMatrix(i, 4) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 5)) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 6)) & "'," _
        & Var31(Var37(grdConsulta.TextMatrix(i, 7))) & "," _
        & Var31(Var37(grdConsulta.TextMatrix(i, 8))) & ")"
      ' Ejecuta la sentencia
      modTablaConsul.SQL = sSQL
      If modTablaConsul.Ejecutar = HAY_ERROR Then End
      modTablaConsul.Cerrar
        
    End If
  
Next i

 sSQL = "INSERT INTO RPTCBFONDORESTRICCION VALUES('" _
        & mcurAsiDetraccion.campo(5) & "','" _
        & mcurAsiDetraccion.campo(0) & "','" _
        & mcurAsiDetraccion.campo(1) & "','" _
        & mcurAsiDetraccion.campo(2) & "','" _
        & mcurAsiDetraccion.campo(3) & "'," _
        & mcurAsiDetraccion.campo(4) & ",'" _
        & mcurAsiDetraccion.campo(6) & "','" _
        & mcurAsiDetraccion.campo(7) & "','" _
        & mcurAsiDetraccion.campo(8) & "','" _
        & mcurAsiDetraccion.campo(9) & "','" _
        & mcurAsiDetraccion.campo(10) & "','" _
        & mcurAsiDetraccion.campo(11) & "')"
      ' Ejecuta la sentencia
      modTablaRetraccion.SQL = sSQL
      If modTablaRetraccion.Ejecutar = HAY_ERROR Then End
      modTablaRetraccion.Cerrar
      


End Sub


Private Sub cmdSalir_Click()
  mcurAsiDetraccion.Cerrar
' Descarga el formulario
  Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Carga los tamaños de las 12 columnas
' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
  Ingreso , Egreso, Orden, CtaBanc, Var
aTamañosColumnas = Array(1000, 400, 1000, 1000, 800, 3000, 3000, 1200, 1200, 1100, 0, 0)
CargarGridTamanios grdConsulta, aTamañosColumnas
    
' Inicia alineamieto de la columna 3
grdConsulta.ColAlignment(6) = 1
    
' Carga la fecha de consulta
mskFecConsulta = gsFecTrabajo

' Establece los campos obligatorios
EstableceCamposObligatorios

' Deshabilita el botón generar informe
cmdInforme.Enabled = False

End Sub

Private Sub EstableceCamposObligatorios()
' ------------------------------------------------------------
' Propósito: Muestra de color amarillo los campos obligatorios
' Recibe: Nada
' Entrega:Nada
' ------------------------------------------------------------
mskFechaIni.BackColor = Obligatorio
mskFechaFin.BackColor = Obligatorio

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Cierra las colecciones
Set mcolSaldoIni = Nothing
Set mcolCtasBanc = Nothing

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
    grdConsulta.Rows = 0
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
  End If
' Inicia progreso
   prgInforme.Max = 24
   prgInforme.Min = 0
   prgInforme.Value = 0
  
' Cargar coleciones y variables de modulo para la consulta
   If CargaColCtasBanc = False Then
      ' No se tiene ctas bancarias que reportar
      MsgBox "No se tienen Ctas bancarias en soles para consultar", , "SGCcaijo-Consulta el diario de bancos"
      ' Sale de el proceso
      Exit Sub
   End If
  prgInforme.Value = prgInforme.Value + 1
   CargaCtaCaja
  prgInforme.Value = prgInforme.Value + 1
   CargaColSaldosIniciales
  prgInforme.Value = prgInforme.Value + 1
' Carga los cursores de ingreso
   CargaIngresos
  prgInforme.Value = prgInforme.Value + 1
   CargaIngrTraslados
  prgInforme.Value = prgInforme.Value + 1
   CargaIngrDevPrest
  prgInforme.Value = prgInforme.Value + 1
   CargaIngrMovPersonal
  prgInforme.Value = prgInforme.Value + 1
   CargaIngrMovTerceros
  prgInforme.Value = prgInforme.Value + 1
   CargaIngrAnulados
  prgInforme.Value = prgInforme.Value + 1
   
' Carga los cursores de egreso
   CargaEgresos
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreTraslados
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreCAProds
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreCAServs
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreIngrCAImpRet
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreAdelantos
  prgInforme.Value = prgInforme.Value + 1
   CargaEgrePrestamos
  prgInforme.Value = prgInforme.Value + 1
   CargaEgrePlanilla
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreMovPersonal
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreMovTerceros
  prgInforme.Value = prgInforme.Value + 1
   CargaEgreAnulados
  prgInforme.Value = prgInforme.Value + 1
  ' Fondo de restriccion
   CargaAsientoDetraccion
  prgInforme.Value = prgInforme.Value + 1
  
'  Carga el grid consulta
   CargarGridConsulta
   prgInforme.Value = prgInforme.Value + 1
   prgInforme.Value = 0
   
' Deshabilita el botón generar informe
  If grdConsulta.Rows > 0 Then
    cmdInforme.Enabled = True
  Else
    cmdInforme.Enabled = False
  End If
   
End Sub

Private Sub CargaAsientoDetraccion()
' ----------------------------------------------------
' Propósito: Consulta apra determinar monto de fondo de restriccion del banco (caso exclusivo)
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT AD.NUMASIENTO, AD.IDASIENTO, AD.CODCONTABLE, AD.DEBEHABER, AD.MONTO, A.FECHA, A.GLOSA, A.ANULADO, A.PROCORIGEN,CB.IDBANCO, CB.DESCCTA, TB.DESCBANCO " _
& "FROM TIPO_BANCOS TB INNER JOIN(TIPO_CUENTASBANC CB INNER JOIN (PLAN_CONTABLE PC INNER JOIN (CTB_ASIENTOS_DET AS AD INNER JOIN CTB_ASIENTOS AS A ON A.NUMASIENTO=AD.NUMASIENTO) " _
& "ON PC.CODCONTABLE=AD.CODCONTABLE) ON PC.CODCONTABLE=CB.CODCONT)ON TB.IDBANCO=CB.IDBANCO " _
& "WHERE A.PROCORIGEN='AM' AND A.GLOSA='POR FONDO SUJETO A RESTRICCION' AND A.ANULADO='NO'"
' Ejecuta la sentencia
mcurAsiDetraccion.SQL = sSQL
If mcurAsiDetraccion.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreAnulados()
' ----------------------------------------------------
' Propósito: Carga el cursor de los egresos anulados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov, E.IdCta,E.Orden " _
     & "FROM EGRESOS E " _
     & "WHERE (E.Orden like 'BA*') and E.Anulado='SI' and " _
          & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov, E.IdCta,E.Orden"
     
' Ejecuta la sentencia
mcurEgreAnulados.SQL = sSQL
If mcurEgreAnulados.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreMovTerceros()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de egresos a Bancos _
            que fueron generados por MovTerceros
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
' Carga la consulta
sSQL = "SELECT E.FecMov, E.IdCta,E.Orden, E.CodContable, T.DescTerc," _
            & "MCB.DescConCB, E.MontoCB " _
     & "FROM MOV_TERCEROS MT, EGRESOS E ,TIPO_TERCEROS T, TIPO_MOVCB MCB " _
     & "WHERE MT.Orden=E.Orden and MT.IdTercero=T.IdTerc and " _
           & "E.CodMov=MCB.IdConCB and (E.Orden like 'BA*') and E.Anulado='NO' and " _
          & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov, E.IdCta,E.Orden"
     
' Ejecuta la sentencia
mcurEgreMovTerc.SQL = sSQL
If mcurEgreMovTerc.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreMovPersonal()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos a Bancos _
            que fueron generados por MovPersonal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
'sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, E.CodContable,P.Apellidos + ' ' + P.Nombre," _
            & "MCB.DescConCB, E.MontoCB " _
     & "FROM MOV_PERSONAL MP,EGRESOS E ,PLN_PERSONAL P, TIPO_MOVCB MCB " _
     & "WHERE MP.Orden=E.Orden and MP.IdPersona=P.IdPersona and " _
           & "E.CodMov=MCB.IdConCB and (E.Orden like 'BA*') and E.Anulado='NO' and " _
          & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov,E.IdCta, E.Orden"
          
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, E.CodContable,TB.DescBanco," _
            & "MCB.DescConCB, E.MontoCB " _
     & "FROM MOV_PERSONAL MP, EGRESOS E, TIPO_MOVCB MCB, TIPO_CUENTASBANC TCB, TIPO_BANCOS TB " _
     & "WHERE MP.Orden=E.Orden and E.IdCta = TCB.IdCta and TCB.IdBanco = TB.IdBanco and " _
           & "E.CodMov=MCB.IdConCB and (E.Orden like 'BA*') and E.Anulado='NO' and " _
          & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov,E.IdCta, E.Orden"

' Ejecuta la sentencia
mcurEgreMovPers.SQL = sSQL
If mcurEgreMovPers.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgrePlanilla()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos que pagan la planilla al personal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, P.CodContable,  " _
            & "PP.DescPlanilla, MCB.DescConCB, P.Monto " _
    & "FROM PAGO_PLANILLAS P, EGRESOS E, " _
         & "PLN_PLANILLAS PP , TIPO_MOVCB MCB " _
    & "WHERE P.Orden=E.Orden and " _
          & "P.CodPlanilla=PP.CodPlanilla and P.Orden like 'BA*' and " _
          & "E.Anulado='NO' and E.CodMov=MCB.IdConCB and " _
         & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
    & "ORDER BY E.FecMov,E.IdCta, E.Orden"
' Ejecuta la sentencia
mcurEgrePlanll.SQL = sSQL
If mcurEgrePlanll.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgrePrestamos()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos que pagan prestamos al personal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov, E.IdCta, E.Orden, CO.CodContable, P.Apellidos+ ' ' + P.Nombre," _
            & "MCB.DescConCB, E.MontoCB " _
    & "FROM PAGO_PRESTAMOS PP, EGRESOS E, PLNCONCEPTOS_OTROS CO, " _
         & "PLN_PERSONAL P , TIPO_MOVCB MCB " _
    & "WHERE PP.Orden=E.Orden and PP.IdPersona=P.IdPersona and " _
          & "PP.IdConPl=CO.IdConPl and PP.Orden like 'BA*' and " _
          & "E.Anulado='NO' and E.CodMov=MCB.IdConCB and " _
         & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
    & "ORDER BY E.FecMov, E.IdCta, E.Orden"
' Ejecuta la sentencia
mcurEgrePrest.SQL = sSQL
If mcurEgrePrest.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreAdelantos()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos que pagan adelantos al personal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, CO.CodContable, P.Apellidos+ ' ' + P.Nombre, " _
            & "MCB.DescConCB, E.MontoCB " _
    & "FROM ADELANTOS A, EGRESOS E, PLNCONCEPTOS_OTROS CO, " _
         & "PLN_PERSONAL P , TIPO_MOVCB MCB " _
    & "WHERE A.Orden=E.Orden and A.IdPersona=P.IdPersona and " _
          & "A.IdConPl=CO.IdConPl and A.Orden like 'BA*' and " _
          & "E.Anulado='NO' and E.CodMov=MCB.IdConCB and " _
         & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
    & "ORDER BY E.FecMov,E.IdCta, E.Orden"
' Ejecuta la sentencia
mcurEgreAdelt.SQL = sSQL
If mcurEgreAdelt.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreCAProds()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos con afectación que pagan prods
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, P.CodCont, PRV.DescProveedor," _
            & "P.DescProd, G.Monto " _
   & "FROM GASTOS G, EGRESOS E, PRODUCTOS P, PROVEEDORES PRV " _
   & "WHERE G.Concepto='P' and  G.Orden=E.Orden and " _
         & "G.CodConcepto=P.IdProd and (G.Orden like 'BA*') and " _
         & "E.Anulado='NO' and E.IdProveedor=PRV.IdProveedor and " _
        & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
   & "ORDER BY E.FecMov,E.IdCta, E.Orden"
' Ejecuta la sentencia
mcurEgreCAProd.SQL = sSQL
If mcurEgreCAProd.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreCAServs()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos con afectación que pagan Serv
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, S.CodCont, PRV.DescProveedor," _
            & "S.DescServ, G.Monto " _
   & "FROM GASTOS G, EGRESOS E, SERVICIOS S, PROVEEDORES PRV " _
   & "WHERE G.Concepto='S' and  G.Orden=E.Orden and " _
         & "G.CodConcepto=S.IdServ and G.Orden like 'BA*'and " _
         & "E.Anulado='NO' and E.IdProveedor=PRV.IdProveedor and " _
        & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
   & "ORDER BY E.FecMov,E.IdCta, E.Orden"

' Ejecuta la sentencia
mcurEgreCAServ.SQL = sSQL
If mcurEgreCAServ.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreIngrCAImpRet()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos de Bancos _
             con movimientos con afectación que Retienen Impts
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
'sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, TI.CodContable, PRV.DescProveedor," _
            & "TI.DescImp, I.Monto " _
   & "FROM MOV_IMPUESTOS I, EGRESOS E, TIPO_IMPUESTOS TI, PROVEEDORES PRV " _
   & "WHERE (I.RelacTributo='Retiene' or I.RelacTributo='Paga') and " _
          & "I.Orden=E.Orden and I.IdImp=TI.IdImp and I.Orden like 'BA*' and " _
          & "E.Anulado='NO' and E.IdProveedor=PRV.IdProveedor and " _
         & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
   & "ORDER BY E.FecMov,E.IdCta, E.Orden"
 
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, I.CodContable, PRV.DescProveedor," _
            & "I.DescImp, I.Monto, E.IdTipoDoc " _
   & "FROM MOV_IMPUESTOS I, EGRESOS E, PROVEEDORES PRV " _
   & "WHERE (I.RelacTributo='Retiene' or I.RelacTributo='Paga') and " _
          & "I.Orden=E.Orden and I.Orden like 'BA*' and " _
          & "E.Anulado='NO' and E.IdProveedor=PRV.IdProveedor and " _
         & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
   & "ORDER BY E.FecMov,E.IdCta, E.Orden, I.DescImp, I.Monto"
 
' Ejecuta la sentencia
mcurEgreCAIngrImptRet.SQL = sSQL
If mcurEgreCAIngrImptRet.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgreTraslados()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de Egresos _
            que fueron generados por traslados de bancos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la sentencia
'sSQL = "SELECT E.FecMov,E.IdCta,E.Orden,I.Orden,I.IdCta,P.Apellidos + ' ' + P.Nombre," _
            & "MCB.DescConCB, E.MontoCB " _
     & "FROM CTB_TRASLADOCAJABANCOS T,INGRESOS I,EGRESOS E, PLN_PERSONAL P, TIPO_MOVCB MCB " _
     & "WHERE T.OrdenEgreso=E.Orden and T.OrdenIngreso=I.Orden and " _
     & "T.OrdenEgreso like 'BA*' and T.IdPersona=P.IdPersona and " _
     & "E.Anulado='NO' and E.CodMov=MCB.IdConCB and " _
     & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov,E.IdCta,E.Orden"
         
sSQL = "SELECT distinct E.FecMov,E.IdCta,E.Orden,I.Orden,I.IdCta,TB.DescBanco," _
            & "MCB.DescConCB, E.MontoCB " _
     & "FROM CTB_TRASLADOCAJABANCOS T,INGRESOS I,EGRESOS E, TIPO_MOVCB MCB, TIPO_CUENTASBANC TCB, TIPO_BANCOS TB " _
     & "WHERE T.OrdenEgreso=E.Orden and T.OrdenIngreso=I.Orden and " _
     & "T.OrdenEgreso like 'BA*' and E.IdCta = TCB.IdCta and TCB.IdBanco = TB.IdBanco and " _
     & "E.Anulado='NO' and E.CodMov=MCB.IdConCB and " _
     & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov,E.IdCta,E.Orden"
' Ejecuta la sentencia
mcurEgreTraslado.SQL = sSQL
If mcurEgreTraslado.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgresos()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de egresos a bancos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la sentencia
sSQL = "SELECT E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque " _
     & "FROM EGRESOS E, TIPO_DOCUM TD " _
     & "WHERE Orden like 'BA*' and " _
           & "E.IdTipoDoc=TD.IdTipoDoc and " _
          & "(E.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY E.FecMov,E.IdCta, E.Orden"
' Ejecuta la sentencia
mcurEgresos.SQL = sSQL
If mcurEgresos.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngrMovPersonal()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de ingresos a bancos _
            que fueron generados por MovPersonal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la consulta
'sSQL = "SELECT I.FecMov, I.IdCta, I.Orden, I.CodContable,P.Apellidos+ ' ' + P.Nombre, " _
           & " MCB.DescConCB, I.Monto " _
     & "FROM MOV_PERSONAL MP,INGRESOS I ,PLN_PERSONAL P, TIPO_MOVCB MCB " _
     & "WHERE MP.Orden=I.Orden and MP.IdPersona=P.IdPersona and " _
           & "I.CodMov=MCB.IdConCB  and (I.Orden like 'BA*') and I.Anulado='NO' and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov, I.IdCta, I.Orden"
     
sSQL = "SELECT I.FecMov, I.IdCta, I.Orden, I.CodContable,TB.DescBanco, " _
           & " MCB.DescConCB, I.Monto " _
     & "FROM MOV_PERSONAL MP,INGRESOS I , TIPO_MOVCB MCB, TIPO_CUENTASBANC TCB, TIPO_BANCOS TB " _
     & "WHERE MP.Orden=I.Orden and I.IdCta = TCB.IdCta and TCB.IdBanco = TB.IdBanco and " _
           & "I.CodMov=MCB.IdConCB  and (I.Orden like 'BA*') and I.Anulado='NO' and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov, I.IdCta, I.Orden"
' Ejecuta la sentencia
mcurIngrMovPers.SQL = sSQL
If mcurIngrMovPers.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngrAnulados()
' ----------------------------------------------------
' Propósito: Carga el cursor de los ingresos anulados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT I.FecMov,I.IdCta, I.Orden " _
     & "FROM INGRESOS I " _
     & "WHERE (I.Orden like 'BA*') and I.Anulado='SI' and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov,I.IdCta, I.Orden"
' Ejecuta la sentencia
mcurIngrAnulados.SQL = sSQL
If mcurIngrAnulados.Abrir = HAY_ERROR Then End
     
End Sub

Private Sub CargaIngrMovTerceros()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de ingresos a Bancos _
            que fueron generados por MovTerceros
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT I.FecMov,I.IdCta, I.Orden, I.CodContable, T.DescTerc, " _
            & "MCB.DescConCB, I.Monto " _
     & "FROM MOV_TERCEROS MT, INGRESOS I ,TIPO_TERCEROS T, TIPO_MOVCB MCB " _
     & "WHERE MT.Orden=I.Orden and MT.IdTercero=T.IdTerc and " _
           & "I.CodMov=MCB.IdConCB and (I.Orden like 'BA*') and I.Anulado='NO' and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov,I.IdCta, I.Orden"
' Ejecuta la sentencia
mcurIngrMovTerc.SQL = sSQL
If mcurIngrMovTerc.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngrDevPrest()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de ingresos a bancos _
            que fueron generados por la devolución de prestamos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la consulta
sSQL = "SELECT I.FecMov,I.IdCta,I.Orden,CO.CodContable,P.Apellidos + ' ' + P.Nombre," _
            & "MCB.DescConCB, I.Monto " _
     & "FROM DEVOLUCION_PRESTAMOSCB DP, INGRESOS I, PLN_PERSONAL P," _
     & "TIPO_MOVCB MCB , PLNCONCEPTOS_OTROS CO " _
     & "WHERE DP.Orden=I.Orden and DP.IdConPl=CO.IdConPl and DP.IdPersona=P.IdPersona and " _
           & "I.CodMov=MCB.IdConCB and (I.Orden like 'BA*') and I.Anulado='NO' and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov,I.IdCta,I.Orden"
' Ejecuta la sentencia
mcurIngrDevPrest.SQL = sSQL
If mcurIngrDevPrest.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngrTraslados()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de ingresos _
            que fueron generados por traslados a Bancos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la sentencia
'sSQL = "SELECT I.FecMov,I.IdCta,I.Orden,E.Orden,E.IdCta,P.Apellidos + ' ' + P.Nombre," _
            & "MCB.DescConCB, I.Monto " _
     & "FROM CTB_TRASLADOCAJABANCOS T,INGRESOS I,EGRESOS E, PLN_PERSONAL P, TIPO_MOVCB MCB  " _
     & "WHERE T.OrdenIngreso=I.Orden and T.OrdenEgreso=E.Orden and " _
          & "(T.OrdenIngreso like 'BA*') and T.IdPersona=P.IdPersona and " _
           & "I.Anulado='NO' and I.CodMov=MCB.IdConCB and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov,I.IdCta,I.Orden"
     
sSQL = "SELECT distinct I.FecMov,I.IdCta,I.Orden,E.Orden,E.IdCta,TB.DescBanco," _
            & "MCB.DescConCB, I.Monto " _
     & "FROM CTB_TRASLADOCAJABANCOS T,INGRESOS I,EGRESOS E, TIPO_MOVCB MCB, TIPO_CUENTASBANC TCB, TIPO_BANCOS TB  " _
     & "WHERE T.OrdenIngreso=I.Orden and T.OrdenEgreso=E.Orden and " _
          & "(T.OrdenIngreso like 'BA*') and I.IdCta = TCB.IdCta and TCB.IdBanco = TB.IdBanco and " _
           & "I.Anulado='NO' and I.CodMov=MCB.IdConCB and " _
          & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov,I.IdCta,I.Orden"
' Ejecuta la sentencia
mcurIngrTraslado.SQL = sSQL
If mcurIngrTraslado.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngresos()
' ----------------------------------------------------
' Propósito: Carga el cursor de los movimientos de ingresos a Bancos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
' Carga la sentencia
sSQL = "SELECT I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc  " _
     & "FROM INGRESOS I, TIPO_DOCUM TD " _
     & "WHERE (Orden like 'BA*') and " _
           & "I.IdTipoDoc=TD.IdTipoDoc and " _
           & "(I.FecMov between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
     & "ORDER BY I.FecMov, I.IdCta, I.Orden"
' Ejecuta la sentencia
mcurIngresos.SQL = sSQL
If mcurIngresos.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaColSaldosIniciales()
' ----------------------------------------------------
' Propósito: Carga la colección que almacena los datos de los _
             saldos iniciales diarios de el intervalo selecionado
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL, sFecha As String
Dim sFecFinal As String
Dim brecorre As Boolean
Dim dblIngreso As Double
Dim dblEgreso As Double
Dim curfechasTrabIng As New clsBD2
Dim curfechasTrabEgr As New clsBD2

' Averigua la fecha inicial,Final
sFecha = FechaAMD(mskFechaIni)
sFecFinal = FechaAMD(mskFechaFin)

' Averigua los días que se usó Bancos en ingresos
sSQL = "SELECT DISTINCT FecMov FROM INGRESOS " _
     & "WHERE Left(Orden,2)='BA'  And " _
     & "FecMov BETWEEN '" & sFecha & "' And '" & sFecFinal & "' " _
     & "ORDER BY FecMov"
' Ejecuta la sentencia
curfechasTrabIng.SQL = sSQL
If curfechasTrabIng.Abrir = HAY_ERROR Then End

' Averigua los días que se usó Bancos en egresos
sSQL = "SELECT DISTINCT FecMov FROM EGRESOS " _
     & "WHERE Left(Orden,2)='BA' And " _
     & "FecMov BETWEEN '" & sFecha & "' And '" & sFecFinal & "' " _
     & "ORDER BY FecMov"
' Ejecuta la sentencia
curfechasTrabEgr.SQL = sSQL
If curfechasTrabEgr.Abrir = HAY_ERROR Then End

' Recorre la fecha hasta la fecha final
brecorre = True
Do While brecorre ' Recorre el intervalo de fechas
    
    ' Verifica si se terminó de recorrer los cursores de fechas ingr y egr
    If curfechasTrabIng.EOF And curfechasTrabEgr.EOF Then
         ' Termina de recorrer los cursores
          brecorre = False
    
    Else ' Algún cursor esta lleno todavía
        If curfechasTrabIng.EOF Then ' Cursor egr tiene fechas
                ' Asigna la fecha del registro actual del cursor Egr
                sFecha = curfechasTrabEgr.campo(0)
                ' Mueve al siguiente registro del cursor
                curfechasTrabEgr.MoverSiguiente
                
        ElseIf curfechasTrabEgr.EOF Then ' Cursor Ingr tiene fechas
                ' Asigna la fecha del registro actual del cursor Ing
                sFecha = curfechasTrabIng.campo(0)
                ' Mueve al siguiente registro del cursor
                curfechasTrabIng.MoverSiguiente
        
        Else ' Ambos cursores tienen fechas
            ' Verifica si los cursores son iguales
            If curfechasTrabIng.campo(0) = curfechasTrabEgr.campo(0) Then
                ' Asigna la fecha del registro actual del cursor
                sFecha = curfechasTrabIng.campo(0)
                ' Mueve al siguiente registro del cursor
                curfechasTrabIng.MoverSiguiente
                curfechasTrabEgr.MoverSiguiente
            ' Verifica si la fecha del cursor ingreso es Menor
            ElseIf curfechasTrabIng.campo(0) < curfechasTrabEgr.campo(0) Then
                ' Asigna la fecha del registro actual del cursor Ing
                sFecha = curfechasTrabIng.campo(0)
                ' Mueve al siguiente registro del cursor
                curfechasTrabIng.MoverSiguiente
            Else ' la fecha del cursor Egreso es Menor
                ' Asigna la fecha del registro actual del cursor Egr
                sFecha = curfechasTrabEgr.campo(0)
                ' Mueve al siguiente registro del cursor
                curfechasTrabEgr.MoverSiguiente
            End If ' Fin de verificar si los cursores son iguales
        
        End If ' Fin de recorrer cursor ingr o egr que tiene registros
        
        Do While Not (mcurCtasBanc.EOF) ' coloca los saldos iniciales para cada cta
            ' Averigua los ingresos hasta el día anteior a caja
            dblIngreso = AveriguaIngresos(AnioMesDiaAnterior(sFecha), mcurCtasBanc.campo(0))
            ' Averigua los egresos hasta el día anteriro de caja
            dblEgreso = AveriguaEgresos(AnioMesDiaAnterior(sFecha), mcurCtasBanc.campo(0))
            ' Añade un elemento al la colección
            mcolSaldoIni.Add Item:=sFecha & "¯" & mcurCtasBanc.campo(0) & "¯" _
                                 & Format(dblIngreso - dblEgreso, "##0.00"), _
                             Key:=sFecha & "¯" & mcurCtasBanc.campo(0)
            ' Mueve al siguiente cuenta
            mcurCtasBanc.MoverSiguiente
        Loop ' Fin de recorrer las ctas bancarias
        mcurCtasBanc.MoverPrimero ' Mueve a la primera cta bancaria
        
    End If ' Fin de verificar si los cursores son vacíos

Loop ' bucle recorre intervalo de fechas

End Sub

Private Function AveriguaIngresos(sFecha As String, sCodCtb As String) As Double
' ----------------------------------------------------
' Propósito: Averigua los ingresos del periodo hasta la fecha pasada
' Recibe: sFecha Fecha que se quiere averiguar
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curIngreso As New clsBD2
' Carga la sentencia
sSQL = "SELECT SUM(Monto) " _
     & " FROM INGRESOS " _
     & " WHERE FecMov<='" & sFecha _
     & "' and Anulado='NO' " _
     & " and IdCta='" & sCodCtb _
     & "' and (Orden like 'BA*')"

' Ejecuta la sentencia
curIngreso.SQL = sSQL
If curIngreso.Abrir = HAY_ERROR Then End

' Verifica si es vacío
If curIngreso.EOF Then
   ' Envía 0.00 como resultado
   AveriguaIngresos = 0
Else
  If IsNull(curIngreso.campo(0)) Then
     ' Envía 0.00 como resultado
     AveriguaIngresos = 0
  Else
    ' Envía la suma de los ingresos
    AveriguaIngresos = curIngreso.campo(0)
  End If
End If

' Cierra el cursor
curIngreso.Cerrar

End Function

Private Function AveriguaEgresos(sFecha As String, sCodCtb As String) As Double
' ----------------------------------------------------
' Propósito: Averigua los egresos del periodo hasta la fecha pasada
' Recibe: sFecha que se quiere averiguar
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curEgreso As New clsBD2
Dim curEmpresas As New clsBD2
Dim EmpresasExistentes As String
Dim InstrucEmpresas As String
Dim TotalEgresoProyectos As Double
Dim TotalEgresoEmpresasSinRH As Double
Dim TotalEgresoEmpresasSoloRHCB As Double

'*-*-*-*-*
'*-*-*-*-*  EMPRESAS EXISTENTES
'*-*-*-*-*
sSQL = "SELECT IdProy " _
     & " FROM PROYECTOS " _
     & " WHERE (PROYECTOS.Tipo = 'EMPR') ORDER BY IdProy "

' Ejecuta la sentencia
curEmpresas.SQL = sSQL
If curEmpresas.Abrir = HAY_ERROR Then End

EmpresasExistentes = ""
If Not curEmpresas.EOF Then
  Do While Not curEmpresas.EOF
    EmpresasExistentes = EmpresasExistentes & curEmpresas.campo(0) & "@"
    
    curEmpresas.MoverSiguiente
  Loop
End If

InstrucEmpresas = ""
Do While InStr(1, EmpresasExistentes, "@")
  InstrucEmpresas = InstrucEmpresas & "IDPROY <> '" & Left(EmpresasExistentes, 2) & "' AND "
  EmpresasExistentes = Mid(EmpresasExistentes, 4, Len(EmpresasExistentes))
Loop

InstrucEmpresas = Left(InstrucEmpresas, Len(InstrucEmpresas) - 4)

curEmpresas.Cerrar

' Carga la sentencia
'*-*-*-*-*
'*-*-*-*-*  TOTAL DE EGRESOS PARA PROYECTOS CON AFECTACION Y SIN AFECTACION
'*-*-*-*-*
sSQL = "SELECT SUM(MontoCB) " _
     & " FROM EGRESOS " _
     & " WHERE FecMov<='" & sFecha _
     & "' and Anulado='NO' " _
     & " and IdCta='" & sCodCtb _
     & "' and (Orden like 'BA*') AND " & InstrucEmpresas

' Ejecuta la sentencia
curEgreso.SQL = sSQL
If curEgreso.Abrir = HAY_ERROR Then End

' Verifica si es vacío
If curEgreso.EOF Then
   ' Envía 0.00 como resultado
   TotalEgresoProyectos = 0
Else
  If IsNull(curEgreso.campo(0)) Then
     ' Envía 0.00 como resultado
     TotalEgresoProyectos = 0
  Else
    ' Envía la suma de los ingresos
    TotalEgresoProyectos = curEgreso.campo(0)
  End If
End If

' Cierra el cursor
curEgreso.Cerrar

' Carga la sentencia
'*-*-*-*-*
'*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SIN RH
'*-*-*-*-*
sSQL = "SELECT SUM(MontoAfectado) " _
     & " FROM EGRESOS, PROYECTOS " _
     & " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And FecMov<='" & sFecha _
     & "' and Anulado='NO' " _
     & " and EGRESOS.IdCta='" & sCodCtb _
     & "' and (Orden like 'BA*') And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc<>'02') "

' Ejecuta la sentencia
curEgreso.SQL = sSQL
If curEgreso.Abrir = HAY_ERROR Then End

' Verifica si es vacío
If curEgreso.EOF Then
   ' Envía 0.00 como resultado
   TotalEgresoEmpresasSinRH = 0
Else
  If IsNull(curEgreso.campo(0)) Then
     ' Envía 0.00 como resultado
     TotalEgresoEmpresasSinRH = 0
  Else
    ' Envía la suma de los ingresos
    TotalEgresoEmpresasSinRH = curEgreso.campo(0)
  End If
End If

' Cierra el cursor
curEgreso.Cerrar

' Carga la sentencia
'*-*-*-*-*
'*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SOLO RH
'*-*-*-*-*
sSQL = "SELECT SUM(MontoCB) " _
     & " FROM EGRESOS, PROYECTOS " _
     & " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And FecMov<='" & sFecha _
     & "' and Anulado='NO' " _
     & " and EGRESOS.IdCta='" & sCodCtb _
     & "' and (Orden like 'BA*') And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc='02') "

' Ejecuta la sentencia
curEgreso.SQL = sSQL
If curEgreso.Abrir = HAY_ERROR Then End

' Verifica si es vacío
If curEgreso.EOF Then
   ' Envía 0.00 como resultado
   TotalEgresoEmpresasSoloRHCB = 0
Else
  If IsNull(curEgreso.campo(0)) Then
     ' Envía 0.00 como resultado
     TotalEgresoEmpresasSoloRHCB = 0
  Else
    ' Envía la suma de los ingresos
    TotalEgresoEmpresasSoloRHCB = curEgreso.campo(0)
  End If
End If

' Cierra el cursor
curEgreso.Cerrar

AveriguaEgresos = TotalEgresoProyectos + TotalEgresoEmpresasSinRH + TotalEgresoEmpresasSoloRHCB

End Function

Private Function CargaColCtasBanc() As Boolean
' ----------------------------------------------------
' Propósito: Carga la colección que carga los datos de las Ctas _
             Bancarias en soles
' Recibe: Nada
' Entrega: Booleano que indica si se tiene ctas bancarias que reportar
' ----------------------------------------------------
Dim sSQL As String

' VADICK MODIFICANDO LA CONSULTA PARA PODER OBTENER EL SALDO DE TODAS LAS CUENTAS BANCARIAS
' TENGAN O NO MOVIMIENTOS EN EL AÑO
' NO MUESTRA LAS CUENTAS ANULADAS LA   67  Y  LA  88
' LA CONSULTA ORIGINAL SOLO PERMITIA OBTENER SALDOS DE CUENTAS CON MOVIMIENTOS ENTRE LAS FECHAS
' DE LA CONSULTA

' Carga la sentencia
'sSQL = "SELECT CT.IdCta,B.DescBanco,CT.DescCta, CT.CodCont " _
    & "FROM TIPO_CUENTASBANC CT, TIPO_BANCOS B " _
    & "WHERE CT.IdMoneda='SOl' and CT.IdBanco=B.IdBanco AND CT.FECHAANULADO >= '" & FechaAMD(mskFechaIni.Text) & "'" _
    & "ORDER BY CT.idCta "




' VADICK MODIFICANDO LA CONSULTA PARA PODER OBTENER EL SALDO DE TODAS LAS CUENTAS BANCARIAS
' TENGAN O NO MOVIMIENTOS EN EL AÑO
' LA CONSULTA ORIGINAL SOLO PERMITIA OBTENER SALDOS DE CUENTAS CON MOVIMIENTOS ENTRE LAS FECHAS
' DE LA CONSULTA

' Carga la sentencia
sSQL = "SELECT CT.IdCta,B.DescBanco,CT.DescCta, CT.CodCont " _
    & "FROM TIPO_CUENTASBANC CT, TIPO_BANCOS B " _
    & "WHERE CT.IdMoneda='SOl' and CT.IdBanco=B.IdBanco " _
    & "ORDER BY CT.idCta "
    
    
    
    
    
    
' /*/*/*/
' CONSULTA PARA OBTENER LAS CUENTAS QUE TUVIERON MOVIMIENTOS EN EL AÑO DE LA CONSULTA
' /*/*/*/
'sSQL = "SELECT CT.IdCta,B.DescBanco,CT.DescCta, CT.CodCont " _
    & "FROM TIPO_CUENTASBANC CT, TIPO_BANCOS B " _
    & "WHERE CT.IdMoneda='SOl' and CT.IdBanco=B.IdBanco and " _
           & "(CT.idCta IN (SELECT DISTINCT I.IdCta " _
                        & "FROM INGRESOS I " _
                        & "WHERE I.IdCta<>'' and " _
                        & "I.FecMov like '" & Right(mskFechaFin, 4) & "*' ) or " _
           & "CT.idCta IN (SELECT DISTINCT E.IdCta " _
                        & "FROM EGRESOS E " _
                        & "WHERE E.IdCta<>'' and " _
                        & "E.FecMov like '" & Right(mskFechaFin, 4) & "*' ) ) " _
    & "ORDER BY CT.idCta "
    
    
    
    
    
    
' /*/*/*/
' CONSULTA PARA OBTENER LAS CUENTAS QUE TUVIERON MOVIMIENTOS ENTRE LAS FECHAS DE LA CONSULTA
' /*/*/*/
'sSQL = "SELECT CT.IdCta,B.DescBanco,CT.DescCta, CT.CodCont " _
    & "FROM TIPO_CUENTASBANC CT, TIPO_BANCOS B " _
    & "WHERE CT.IdMoneda='SOl' and CT.IdBanco=B.IdBanco and " _
           & "(CT.idCta IN (SELECT DISTINCT I.IdCta " _
                        & "FROM INGRESOS I " _
                        & "WHERE I.IdCta<>'' and " _
                        & "I.FecMov between '" & FechaAMD(mskFechaIni.Text) & "' and '" & FechaAMD(mskFechaFin.Text) & "') or " _
           & "CT.idCta IN (SELECT DISTINCT E.IdCta " _
                        & "FROM EGRESOS E " _
                        & "WHERE E.IdCta<>'' and " _
                        & "E.FecMov between '" & FechaAMD(mskFechaIni.Text) & "' and '" & FechaAMD(mskFechaFin.Text) & "') ) " _
    & "ORDER BY CT.idCta "



' Ejecuta la sentencia
mcurCtasBanc.SQL = sSQL
If mcurCtasBanc.Abrir = HAY_ERROR Then End

' Verifica si el cursor es vacío
If mcurCtasBanc.EOF Then
    mcurCtasBanc.Cerrar
    CargaColCtasBanc = False
    Exit Function
Else
  Do While Not mcurCtasBanc.EOF
     ' Añade un elemento a la coleción
     mcolCtasBanc.Add Item:=mcurCtasBanc.campo(0) & "¯" & _
                            mcurCtasBanc.campo(1) & "¯" & _
                            mcurCtasBanc.campo(2) & "¯" & _
                            mcurCtasBanc.campo(3), _
                      Key:=mcurCtasBanc.campo(0)
                           
     ' Mueve al siguiente registro
     mcurCtasBanc.MoverSiguiente
  Loop
    ' Mueve al primer elemento
    mcurCtasBanc.MoverPrimero
End If

' Devuelve función
CargaColCtasBanc = True

End Function

Private Sub CargaCtaCaja()
' ----------------------------------------------------
' Propósito: Carga el código contable de caja el la variable de modulo
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curCaja As New clsBD2
' Carga la sentencia
sSQL = "SELECT DescCaja,CodContable " _
       & "FROM CTB_CAJA " _
       & "WHERE IdCTBCaja='01'"
' Ejecuta la sentencia
curCaja.SQL = sSQL
If curCaja.Abrir = HAY_ERROR Then End

' Carga la variable
msCtaCaja = curCaja.campo(0) & "¯" & curCaja.campo(1)

End Sub

Sub RecuperarFechaAperturaCta(CtaAConsultar As String)
  ' ----------------------------------------------------
  ' Propósito: Recuperar la apertura de la cta
  ' Recibe: Numero de cta
  ' Entrega: Fecha apertura
  ' ----------------------------------------------------
  Dim sSQL As String
  Dim curFechaAperturaCta As New clsBD2
  
  FechaAperturaCta = ""
  ' Carga la sentencia
  sSQL = "SELECT MIN(FECMOV) " _
         & "FROM INGRESOS " _
         & "WHERE IDCTA ='" & CtaAConsultar & "'"
  ' Ejecuta la sentencia
  curFechaAperturaCta.SQL = sSQL
  
  If curFechaAperturaCta.Abrir = HAY_ERROR Then
    End
  End If
  
  ' Carga la variable
  If IsNull(curFechaAperturaCta.campo(0)) Then
    FechaAperturaCta = "20110101"
  Else
    FechaAperturaCta = curFechaAperturaCta.campo(0)
  End If
  
  'FechaAperturaCta = curFechaAperturaCta.campo(0)
  
  curFechaAperturaCta.Cerrar
End Sub

Private Sub CargarGridConsulta()
' ----------------------------------------------------
' Propósito: Arma la consulta en el grid
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim bRecorreOrden As Boolean

' Inicializa el grid
grdConsulta.Rows = 0
grdConsulta.ScrollBars = flexScrollBarNone
grdConsulta.Visible = True

' Recorre la colección de saldos iniciales
For Each mFecha In mcolSaldoIni
    ' Inicializa las variables acumuladas
    mdblIngreso = 0: mdblEgreso = 0
    ' grd: Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
    Ingreso , Egreso, Orden, CtaBanc, Var
    ' colcta: CT.idCta,B.DescBanco,CT.DescCta, CT.CodCont
  RecuperarFechaAperturaCta Var30(mFecha, 2)
  RecuperarFechaAnulacionCta Var30(mFecha, 2)
  'If (FechaAperturaCta <= Var30(mFecha, 1)) Then
' VADICK MODIFICACION DE LA CONDICIONAL PARA VALIDAR LA FECHA DE ANULACION DE LA CTA Y COLOCAR EN EL GRID
  If (FechaAperturaCta <= Var30(mFecha, 1)) And (FechaAnulacionCta >= Var30(mFecha, 1)) Then
    grdConsulta.AddItem FechaDMA(Var30(mFecha, 1)) & vbTab & vbTab & vbTab _
                & "CTA.BANC." & vbTab _
                & Var30(mcolCtasBanc(Var30(mFecha, 2)), 4) & vbTab _
                & Var30(mcolCtasBanc(Var30(mFecha, 2)), 2) & vbTab _
                & Var30(mcolCtasBanc(Var30(mFecha, 2)), 3) & vbTab _
                & "SALD.ANT:" & vbTab & Format(Val(Var30(mFecha, 3)), "###,###,##0.00") & vbTab _
                & vbTab & Var30(mFecha, 2) & vbTab & "G"
   ' Coloca color al grid
    grdConsulta.Row = grdConsulta.Rows - 1
    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
    grdConsulta.AddItem "Fecha" & vbTab & "Doc" & vbTab & "Número" & vbTab & "Cheque" _
                & vbTab & "Cod.Cta." & vbTab & "Proveedor.Ejecutor" & vbTab & "Detalle.Movimiento" _
                & vbTab & "Ingreso" & vbTab & "Egreso" & vbTab & "Orden"
   ' Coloca color al grid
    grdConsulta.Row = grdConsulta.Rows - 1
    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
                
    'Recorre los movimientos de Caja Bancos
    bRecorreOrden = True
    Do While bRecorreOrden
        ' Establece el orden de Ingreso y Egreso
        If mcurIngresos.EOF And mcurEgresos.EOF Then ' Cursores vacios
            bRecorreOrden = False
        Else ' Algún cursor no es vacío
            If mcurIngresos.EOF Then ' Ingresos es vacío
                If mcurEgresos.campo(0) = Var30(mFecha, 1) And _
                   mcurEgresos.campo(1) = Var30(mFecha, 2) Then
                   ' Misma fecha y misma cta bancaria
                    CargaRegEgreso
                    mcurEgresos.MoverSiguiente
                Else ' No es la misma fecha
                    bRecorreOrden = False
                End If ' Fin de verificar si la fecha es la misma
            ElseIf mcurEgresos.EOF Then ' Egresos es vacío
                If mcurIngresos.campo(0) = Var30(mFecha, 1) And _
                   mcurIngresos.campo(1) = Var30(mFecha, 2) Then
                   ' Misma fecha y misma cta bancaria
                    CargaRegIngreso
                    mcurIngresos.MoverSiguiente
                Else ' No es la misma fecha
                    bRecorreOrden = False
                End If ' Fin de verificar si la fecha es la misma
            Else ' Ninguno es vacío
                ' Verifica si la fecha y la cta banc de los cursores es la misma que la colección de saldos
                If mcurEgresos.campo(0) = Var30(mFecha, 1) And _
                   mcurEgresos.campo(1) = Var30(mFecha, 2) And _
                   mcurIngresos.campo(0) = Var30(mFecha, 1) And _
                   mcurIngresos.campo(1) = Var30(mFecha, 2) Then
                   ' Averigua cual cursor tiene el orden Menor
                   If mcurIngresos.campo(2) < mcurEgresos.campo(2) Then
                     'Ingresos tiene el orden Menor
                      CargaRegIngreso
                      mcurIngresos.MoverSiguiente
                   Else ' Egresos tiene el orden Menor
                      CargaRegEgreso
                      mcurEgresos.MoverSiguiente
                   End If ' fin de verificar el orden Menor
                Else 'Algún cursor no es Igual a la fecha y la cta de saldos iniciales
                    If mcurEgresos.campo(0) = Var30(mFecha, 1) And _
                       mcurEgresos.campo(1) = Var30(mFecha, 2) Then
                      ' Igual la fecha y ctabanc ,Carga egresos
                      CargaRegEgreso
                      mcurEgresos.MoverSiguiente
                    ElseIf mcurIngresos.campo(0) = Var30(mFecha, 1) And _
                           mcurIngresos.campo(1) = Var30(mFecha, 2) Then
                      ' Igual la fecha y ctabanc ,Carga ingresos
                      CargaRegIngreso
                      mcurIngresos.MoverSiguiente
                    Else ' Ninguna fecha es Igual a la fecha de los saldos iniciales
                       bRecorreOrden = False
                    End If ' Fin de verificar cual cursor tiene la fecha Igual a los saldos
                End If ' Fin de verificar las fechas iguales a saldos iniciales
            End If ' Fin de verificar si algún cursor es vacío
        End If ' Fin de verificar si los cursores son vacios
    Loop ' Fin de recorrer mov caja bancos
    ' Cargar los totales en el grid
    grdConsulta.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL:" & vbTab & _
                        Format(mdblIngreso, "###,###,##0.00") & vbTab & _
                        Format(mdblEgreso, "###,###,##0.00")
   
   

   ' Coloca color al grid con Saldo disponible del Fondo de restriccion
    grdConsulta.Row = grdConsulta.Rows - 1
    MarcarFilaGRID grdConsulta, vbBlack, vbGray
    
    ' Si EL banco, cuenta bancaria y fecha estan definidos segun el asinto automatico del fondo de restriccion
    If Var30(mFecha, 1) >= mcurAsiDetraccion.campo(5) And Var30(mcolCtasBanc(Var30(mFecha, 2)), 2) = mcurAsiDetraccion.campo(11) And Var30(mcolCtasBanc(Var30(mFecha, 2)), 3) = mcurAsiDetraccion.campo(10) Then
      SaldoDisponible = Format(mdblIngreso + Val(Var30(mFecha, 3)) - mdblEgreso - mcurAsiDetraccion.campo(4), "###,###,##0.00")
      grdConsulta.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SALDO DISPONIBLE:" & vbTab & Format(SaldoDisponible, "###,###,##0.00")
               
    End If
   
   ' Coloca color al grid
    grdConsulta.Row = grdConsulta.Rows - 1
    MarcarFilaGRID grdConsulta, vbBlack, vbGray
    
    grdConsulta.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SALDO ACTUAL:" & vbTab & _
                        Format(mdblIngreso + Val(Var30(mFecha, 3)) - mdblEgreso _
                        , "###,###,##0.00")
   ' Coloca color al grid
    grdConsulta.Row = grdConsulta.Rows - 1
    MarcarFilaGRID grdConsulta, vbBlack, vbGray
  End If
Next mFecha ' Fin de recorrer saldos iniciales

' Coloca las barras de desplazamiento
grdConsulta.ScrollBars = flexScrollBarBoth
grdConsulta.Visible = True

' Cierra los cursores y colecciones
' Colecciones para la carga de la consulta
Set mcolSaldoIni = Nothing
Set mcolCtasBanc = Nothing

mcurCtasBanc.Cerrar

mcurIngresos.Cerrar
mcurIngrTraslado.Cerrar
mcurIngrDevPrest.Cerrar
mcurIngrMovPers.Cerrar
mcurIngrMovTerc.Cerrar
mcurIngrAnulados.Cerrar

mcurEgresos.Cerrar
mcurEgreTraslado.Cerrar
mcurEgreCAProd.Cerrar
mcurEgreCAServ.Cerrar
mcurEgreCAIngrImptRet.Cerrar
mcurEgreAdelt.Cerrar
mcurEgrePrest.Cerrar
mcurEgrePlanll.Cerrar
mcurEgreMovPers.Cerrar
mcurEgreMovTerc.Cerrar
mcurEgreAnulados.Cerrar


End Sub

Private Sub CargaRegEgreso()
' ----------------------------------------------------
' Propósito: Carga los datos del orden que generó egresos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' Verifica en cada proceso de egreso si pertenece al orden
' Muestra Proceso de traslados
  MuestraTrasladosEgresos
' Muestra Proceso con Afectación
  MuestraEgresoCAProds
  MuestraEgresoCAServs
  MuestraEgresoCARetImpt
' Muestra Proceso de Adelantos, Prestamos
  MuestraEgrAdelantos
  MuestraEgrPrestamos
' Muestra Proceso de planillas
  MuestraEgrPlanillas
' Muestra procesos de personal y terceros
  MuestraEgrMovPersonal
  MuestraEgrMovTerceros
  MuestraEgrAnulados
End Sub

Private Sub MuestraEgrAnulados()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreAnulados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov, E.IdCta,E.Orden
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreAnulados.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreAnulados.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreAnulados.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreAnulados.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreAnulados.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab & vbTab _
                      & "ANULADO" & vbTab _
                      & "ANULADO" & vbTab & vbTab _
                      & "0.00" & vbTab _
                      & mcurEgreAnulados.campo(2) & vbTab _
                      & mcurEgreAnulados.campo(1) & vbTab & "D"
     ' Mueve al siguiente elemento del cursor
     mcurEgreAnulados.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraEgrMovTerceros()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreMovTerceros
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov, E.IdCta,E.Orden, E.CodContable, T.DescTerc " _
     & ", MCB.DescConCB, E.MontoCB
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreMovTerc.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreMovTerc.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreMovTerc.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreMovTerc.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreMovTerc.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgreMovTerc.campo(3) & vbTab _
                      & mcurEgreMovTerc.campo(4) & vbTab _
                      & mcurEgreMovTerc.campo(5) & vbTab & vbTab _
                      & Format(mcurEgreMovTerc.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgreMovTerc.campo(2) & vbTab _
                      & mcurEgreMovTerc.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreMovTerc.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgreMovTerc.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraEgrMovPersonal()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreMovPersonal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, E.CodContable,P.Apellidos " _
     & "+ ' ' + P.Nombre, MCB.DescConCB, E.MontoCB
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreMovPers.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreMovPers.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreMovPers.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreMovPers.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreMovPers.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgreMovPers.campo(3) & vbTab _
                      & mcurEgreMovPers.campo(4) & vbTab _
                      & mcurEgreMovPers.campo(5) & vbTab & vbTab _
                      & Format(mcurEgreMovPers.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgreMovPers.campo(2) & vbTab _
                      & mcurEgreMovPers.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreMovPers.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgreMovPers.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraEgrPlanillas()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgrePlanillas
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, PT.CodContable,  " _
    & "PP.DescPlanilla, MCB.DescConCB, PT.Monto
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgrePlanll.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgrePlanll.campo(0) And _
      mcurEgresos.campo(1) = mcurEgrePlanll.campo(1) And _
      mcurEgresos.campo(2) = mcurEgrePlanll.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgrePlanll.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgrePlanll.campo(3) & vbTab _
                      & mcurEgrePlanll.campo(4) & vbTab _
                      & mcurEgrePlanll.campo(5) & vbTab & vbTab _
                      & Format(mcurEgrePlanll.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgrePlanll.campo(2) & vbTab _
                      & mcurEgrePlanll.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgrePlanll.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgrePlanll.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraEgrPrestamos()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgrePrestamos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov, E.IdCta, E.Orden, CO.CodContable, P.Apellidos " _
    & "+ ' ' + P.Nombre, MCB.DescConCB, E.MontoCB
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgrePrest.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgrePrest.campo(0) And _
      mcurEgresos.campo(1) = mcurEgrePrest.campo(1) And _
      mcurEgresos.campo(2) = mcurEgrePrest.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgrePrest.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgrePrest.campo(3) & vbTab _
                      & mcurEgrePrest.campo(4) & vbTab _
                      & mcurEgrePrest.campo(5) & vbTab & vbTab _
                      & Format(mcurEgrePrest.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgrePrest.campo(2) & vbTab _
                      & mcurEgrePrest.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgrePrest.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgrePrest.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub


Private Sub MuestraEgrAdelantos()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreAdelantos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, CO.CodContable, P.Apellidos " _
    & "+ ' ' + P.Nombre, MCB.DescConCB, E.MontoCB
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreAdelt.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreAdelt.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreAdelt.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreAdelt.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreAdelt.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgreAdelt.campo(3) & vbTab _
                      & mcurEgreAdelt.campo(4) & vbTab _
                      & mcurEgreAdelt.campo(5) & vbTab & vbTab _
                      & Format(mcurEgreAdelt.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgreAdelt.campo(2) & vbTab _
                      & mcurEgreAdelt.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreAdelt.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgreAdelt.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub


Private Sub MuestraEgresoCARetImpt()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreCARetImpt
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, TI.CodContable, PRV.DescProveedor," _
   & "TI.DescImp, I.Monto
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreCAIngrImptRet.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreCAIngrImptRet.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreCAIngrImptRet.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreCAIngrImptRet.campo(2) Then
      If (LTrim(RTrim(mcurEgreCAIngrImptRet.campo(5))) = LTrim(RTrim("IMPUESTO GENERAL DE LAS VENTAS (IGV) 18%"))) And ((mcurEgreCAIngrImptRet.campo(7) = "01") Or (mcurEgreCAIngrImptRet.campo(7) = "12") Or (mcurEgreCAIngrImptRet.campo(7) = "13")) Then
                ' muestra registro en le grid
        grdConsulta.AddItem FechaDMA(mcurEgreCAIngrImptRet.campo(0)) & vbTab _
                         & mcurEgresos.campo(3) & vbTab _
                         & mcurEgresos.campo(4) & vbTab _
                         & mcurEgresos.campo(5) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(3) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(4) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(5) & vbTab & vbTab _
                         & Format(mcurEgreCAIngrImptRet.campo(6), "###,###,##0.00") & vbTab _
                         & mcurEgreCAIngrImptRet.campo(2) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(1) & vbTab & "D"
        ' Acumula en la variable de egresos
        'mdblIngreso = mdblIngreso + mcurEgreCAIngrImptRet.campo(6)
        mdblEgreso = mdblEgreso + mcurEgreCAIngrImptRet.campo(6)
      Else
        ' muestra registro en le grid
        grdConsulta.AddItem FechaDMA(mcurEgreCAIngrImptRet.campo(0)) & vbTab _
                         & mcurEgresos.campo(3) & vbTab _
                         & mcurEgresos.campo(4) & vbTab _
                         & mcurEgresos.campo(5) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(3) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(4) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(5) & vbTab _
                         & Format(mcurEgreCAIngrImptRet.campo(6), "###,###,##0.00") & vbTab & vbTab _
                         & mcurEgreCAIngrImptRet.campo(2) & vbTab _
                         & mcurEgreCAIngrImptRet.campo(1) & vbTab & "D"
        ' Acumula en la variable de egresos
        mdblIngreso = mdblIngreso + mcurEgreCAIngrImptRet.campo(6)
      End If
     ' Mueve al siguiente elemento del cursor
     mcurEgreCAIngrImptRet.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub


Private Sub MuestraEgresoCAServs()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreServs
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, S.CodCont, PRV.DescProveedor," _
   & "S.DescServ, G.Monto
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreCAServ.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreCAServ.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreCAServ.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreCAServ.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreCAServ.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgreCAServ.campo(3) & vbTab _
                      & mcurEgreCAServ.campo(4) & vbTab _
                      & mcurEgreCAServ.campo(5) & vbTab & vbTab _
                      & Format(mcurEgreCAServ.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgreCAServ.campo(2) & vbTab _
                      & mcurEgreCAServ.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreCAServ.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgreCAServ.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraEgresoCAProds()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreProds
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta, E.Orden, P.CodCont, PRV.DescProveedor," _
   & "P.DescProd, G.Monto
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreCAProd.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreCAProd.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreCAProd.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreCAProd.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreCAProd.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & mcurEgreCAProd.campo(3) & vbTab _
                      & mcurEgreCAProd.campo(4) & vbTab _
                      & mcurEgreCAProd.campo(5) & vbTab & vbTab _
                      & Format(mcurEgreCAProd.campo(6), "###,###,##0.00") & vbTab _
                      & mcurEgreCAProd.campo(2) & vbTab _
                      & mcurEgreCAProd.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreCAProd.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurEgreCAProd.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraTrasladosEgresos()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurEgreTraslados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: E.FecMov,E.IdCta,E.Orden,I.Orden,I.IdCta,P.Apellidos + ' ' + P.Apellidos," _
     & "MCB.DescConCB, E.MontoCB
' Egr: E.FecMov,E.IdCta, E.Orden, TD.Abreviatura, E.NumDoc, E.NumCheque
Dim sCodCta As String
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurEgreTraslado.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurEgresos.campo(0) = mcurEgreTraslado.campo(0) And _
      mcurEgresos.campo(1) = mcurEgreTraslado.campo(1) And _
      mcurEgresos.campo(2) = mcurEgreTraslado.campo(2) Then
     ' Carga el codigo contable
     If Left(mcurEgreTraslado.campo(3), 2) = "CA" Then
        sCodCta = Var30(msCtaCaja, 2)
     Else
        sCodCta = Var30(mcolCtasBanc.Item(mcurEgreTraslado.campo(4)), 4)
     End If
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurEgreTraslado.campo(0)) & vbTab _
                      & mcurEgresos.campo(3) & vbTab _
                      & mcurEgresos.campo(4) & vbTab _
                      & mcurEgresos.campo(5) & vbTab _
                      & sCodCta & vbTab _
                      & mcurEgreTraslado.campo(5) & vbTab _
                      & mcurEgreTraslado.campo(6) & vbTab & vbTab _
                      & Format(mcurEgreTraslado.campo(7), "###,###,##0.00") & vbTab _
                      & mcurEgreTraslado.campo(2) & vbTab _
                      & mcurEgreTraslado.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblEgreso = mdblEgreso + mcurEgreTraslado.campo(7)
     ' Mueve al siguiente elemento del cursor
     mcurEgreTraslado.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub CargaRegIngreso()
' ----------------------------------------------------
' Propósito: Carga los datos del orden que generó ingresos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' Verifica en cada proceso de ingreso si pertenece al orden
' Muestra Proceso de traslados
  MuestraTrasladosIngresos
' Muestra Proceso de Prestamos
  MuestraIngrDevPrestamos
' Muestra procesos de personal y terceros
  MuestraIngrMovPersonal
  MuestraIngrMovTerceros
  MuestraIngrAnulados
End Sub

Private Sub MuestraIngrAnulados()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurIngrAnulados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: I.FecMov,I.IdCta, I.Orden
' Ingr: I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurIngrAnulados.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurIngresos.campo(0) = mcurIngrAnulados.campo(0) And _
      mcurIngresos.campo(1) = mcurIngrAnulados.campo(1) And _
      mcurIngresos.campo(2) = mcurIngrAnulados.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurIngrAnulados.campo(0)) & vbTab _
                      & mcurIngresos.campo(3) & vbTab _
                      & mcurIngresos.campo(4) & vbTab & vbTab & vbTab _
                      & "ANULADO" & vbTab _
                      & "ANULADO" & vbTab _
                      & "0.00" & vbTab & vbTab _
                      & mcurIngrAnulados.campo(2) & vbTab _
                      & mcurIngrAnulados.campo(1) & vbTab & "D"

     ' Mueve al siguiente elemento del cursor
     mcurIngrAnulados.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraIngrMovTerceros()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurIngrMovTerceros
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: I.FecMov,I.IdCta, I.Orden, I.CodContable, T.DescTerc " _
     & ", MCB.DescConCB, I.Monto
' Ingr: I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurIngrMovTerc.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurIngresos.campo(0) = mcurIngrMovTerc.campo(0) And _
      mcurIngresos.campo(1) = mcurIngrMovTerc.campo(1) And _
      mcurIngresos.campo(2) = mcurIngrMovTerc.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurIngrMovTerc.campo(0)) & vbTab _
                      & mcurIngresos.campo(3) & vbTab _
                      & mcurIngresos.campo(4) & vbTab & vbTab _
                      & mcurIngrMovTerc.campo(3) & vbTab _
                      & mcurIngrMovTerc.campo(4) & vbTab _
                      & mcurIngrMovTerc.campo(5) & vbTab _
                      & Format(mcurIngrMovTerc.campo(6), "###,###,##0.00") & vbTab & vbTab _
                      & mcurIngrMovTerc.campo(2) & vbTab _
                      & mcurIngrMovTerc.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblIngreso = mdblIngreso + mcurIngrMovTerc.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurIngrMovTerc.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraIngrMovPersonal()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurIngrMovPersonal
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: I.FecMov, I.IdCta, I.Orden, I.CodContable,P.Apellidos " _
     & "+ ' ' + P.Nombre, MCB.DescConCB, I.Monto
' Ingr: I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurIngrMovPers.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurIngresos.campo(0) = mcurIngrMovPers.campo(0) And _
      mcurIngresos.campo(1) = mcurIngrMovPers.campo(1) And _
      mcurIngresos.campo(2) = mcurIngrMovPers.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurIngrMovPers.campo(0)) & vbTab _
                      & mcurIngresos.campo(3) & vbTab _
                      & mcurIngresos.campo(4) & vbTab & vbTab _
                      & mcurIngrMovPers.campo(3) & vbTab _
                      & mcurIngrMovPers.campo(4) & vbTab _
                      & mcurIngrMovPers.campo(5) & vbTab _
                      & Format(mcurIngrMovPers.campo(6), "###,###,##0.00") & vbTab & vbTab _
                      & mcurIngrMovPers.campo(2) & vbTab _
                      & mcurIngrMovPers.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblIngreso = mdblIngreso + mcurIngrMovPers.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurIngrMovPers.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraIngrDevPrestamos()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurIngrDevPrestamos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: I.FecMov,I.IdCta,I.Orden,CO.CodContable,P.Apellidos + ' '" _
     & " + P.Nombre, MCB.DescConCB, I.Monto
' Ingr: I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurIngrDevPrest.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurIngresos.campo(0) = mcurIngrDevPrest.campo(0) And _
      mcurIngresos.campo(1) = mcurIngrDevPrest.campo(1) And _
      mcurIngresos.campo(2) = mcurIngrDevPrest.campo(2) Then
     ' muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurIngrDevPrest.campo(0)) & vbTab _
                      & mcurIngresos.campo(3) & vbTab _
                      & mcurIngresos.campo(4) & vbTab & vbTab _
                      & mcurIngrDevPrest.campo(3) & vbTab _
                      & mcurIngrDevPrest.campo(4) & vbTab _
                      & mcurIngrDevPrest.campo(5) & vbTab _
                      & Format(mcurIngrDevPrest.campo(6), "###,###,##0.00") & vbTab & vbTab _
                      & mcurIngrDevPrest.campo(2) & vbTab _
                      & mcurIngrDevPrest.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblIngreso = mdblIngreso + mcurIngrDevPrest.campo(6)
     ' Mueve al siguiente elemento del cursor
     mcurIngrDevPrest.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Sub MuestraTrasladosIngresos()
' ----------------------------------------------------
' Propósito: Carga los datos de mcurIngreTraslados
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
' grd: ' Fecha,TipDoc,NroDoc,Cheque,Cta.Ctb.,Prov.Ejecutor,Detalle.Movimiento, _
         Ingreso , Egreso, Orden
' cur: I.FecMov,I.IdCta,I.Orden,E.Orden,E.IdCta,P.Apellidos + ' ' + P.Apellidos," _
      "MCB.DescConCB, I.Monto
' Ingr: I.FecMov, I.IdCta, I.Orden, TD.Abreviatura, I.NumDoc
Dim sCodCta As String
Dim brecorre As Boolean
' Inicializa la variable recorre
 brecorre = True
 Do While brecorre = True
  ' Verifica si se ha recorrido todo
  If mcurIngrTraslado.EOF Then ' final del cursor
     brecorre = False
  Else ' El cursor tiene datos
   ' Verifica si es Igual al orden de egreso del cursor egresos
   If mcurIngresos.campo(0) = mcurIngrTraslado.campo(0) And _
      mcurIngresos.campo(1) = mcurIngrTraslado.campo(1) And _
      mcurIngresos.campo(2) = mcurIngrTraslado.campo(2) Then
     ' Carga el código contable
     If Left(mcurIngrTraslado.campo(3), 2) = "CA" Then
        sCodCta = Var30(msCtaCaja, 2)
     Else
        sCodCta = Var30(mcolCtasBanc.Item(mcurIngrTraslado.campo(4)), 4)
     End If
     ' Muestra registro en le grid
     grdConsulta.AddItem FechaDMA(mcurIngrTraslado.campo(0)) & vbTab _
                      & mcurIngresos.campo(3) & vbTab _
                      & mcurIngresos.campo(4) & vbTab & vbTab _
                      & sCodCta & vbTab _
                      & mcurIngrTraslado.campo(5) & vbTab _
                      & mcurIngrTraslado.campo(6) & vbTab _
                      & Format(mcurIngrTraslado.campo(7), "###,###,##0.00") & vbTab & vbTab _
                      & mcurIngrTraslado.campo(2) & vbTab _
                      & mcurIngrTraslado.campo(1) & vbTab & "D"
     ' Acumula en la variable de egresos
     mdblIngreso = mdblIngreso + mcurIngrTraslado.campo(7)
     ' Mueve al siguiente elemento del cursor
     mcurIngrTraslado.MoverSiguiente
   Else ' no es Igual al orden
     brecorre = False
   End If ' Fin de verificar si si es el mismo orden
  End If ' Verifica si es el final del cursor
 Loop ' Recorre tralados
End Sub

Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito: Verifica si esta bien los datos para ejecutar _
            la consulta
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------

If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
' Verifica que la fecha de inicio sea Menor a la fecha final
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = True Then
        fbOkDatosIntroducidos = False
        Exit Function
    End If
Else ' Alguna fecha es obligatorio
        fbOkDatosIntroducidos = False
        Exit Function
End If

' Verifica que el año de la fecha  de inicio sea Igual al año de la fecha fin
If Right(mskFechaIni, 4) <> Right(mskFechaFin, 4) Then
   ' Msg Mismo año
   MsgBox "La consulta debe pertenecer al mismo periodo ", , "SGCcaijo-Consulta Diario de Caja"
   mskFechaIni.SetFocus
   fbOkDatosIntroducidos = False
   Exit Function
End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function


Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  ' Carga consulta
  CargaConsulta
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 0
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
End If

End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub mskFechaIni_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
    mskFechaIni.BackColor = vbWhite
    ' Carga consulta
    CargaConsulta
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 0
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
  
End If

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Sub RecuperarFechaAnulacionCta(CtaAConsultar As String)
  ' ----------------------------------------------------
  ' Propósito: Recuperar la Fecha de Anulacion de la cta
  ' Recibe: Numero de cta
  ' Entrega: Fecha Anulacion
  ' ----------------------------------------------------
  
' VADICK CONSULTA PARA OBTENER LA FECHA DE ANULACION DE LA CTA

  Dim sSQL As String
  Dim curFechaAnulacionCta As New clsBD2
      
  FechaAnulacionCta = ""
  ' Carga la sentencia
  sSQL = "SELECT FECHAANULADO " _
         & "FROM TIPO_CUENTASBANC " _
         & "WHERE IDCTA ='" & CtaAConsultar & "'"
  ' Ejecuta la sentencia
  curFechaAnulacionCta.SQL = sSQL
  
  If curFechaAnulacionCta.Abrir = HAY_ERROR Then
    End
  End If
  
  ' Carga la variable
  If IsNull(curFechaAnulacionCta.campo(0)) Then
    FechaAnulacionCta = "21001231"
  Else
    FechaAnulacionCta = curFechaAnulacionCta.campo(0)
  End If
  
  curFechaAnulacionCta.Cerrar
End Sub


