VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALConsulIngresosCompras 
   Caption         =   "SGCcaijo-Consulta Ingresos a almacén por compras"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   97
   Icon            =   "SCALIngresosCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   360
      Top             =   7680
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
      TabIndex        =   6
      Top             =   8160
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10320
      TabIndex        =   7
      Top             =   8160
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   960
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   11535
      Begin VB.OptionButton optNumero 
         Caption         =   "Por Número"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Por Fechas"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNumDoc 
         Height          =   315
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   6240
         TabIndex        =   9
         Top             =   120
         Width           =   4935
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   330
            Left            =   1245
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
            Left            =   3525
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
            Caption         =   "Fecha &Fin:"
            Height          =   195
            Left            =   2640
            TabIndex        =   11
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha &Inicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   285
            Width           =   915
         End
      End
      Begin VB.Frame N 
         Height          =   735
         Left            =   2160
         TabIndex        =   14
         Top             =   120
         Width           =   3735
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número de Salida"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   300
            Width           =   1260
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      HighLight       =   0
      FillStyle       =   1
      MergeCells      =   4
   End
   Begin ComctlLib.ProgressBar prgInforme 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   8160
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
      TabIndex        =   12
      Top             =   8160
      Width           =   1815
   End
End
Attribute VB_Name = "frmALConsulIngresosCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Colección para la carga de la consulta
Private mcolDocumentosVerif As New Collection

' Cursores para la carga de la consulta
Private mcurIngresoActivos As New clsBD2
Private mcurIngresoMercaderias As New clsBD2

' Variable para la carga de los Totales
Private mdblPrecioTotal As Double

Private Sub cmdInforme_Click()
Dim rptMayor As New clsBD4

' Deshabilita el botón informe
  cmdInforme.Enabled = False
  
' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón informe
       cmdInforme.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla consulta
  LlenaTablaConsul
  
' Genera el reporte
' Formulario
  Set rptMayor.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptMayor.AsignarRpt

' Clausula WHERE de las relaciones del rpt.
  rptMayor.FiltroSelectionFormula = ""

' Nombre del fichero
  rptMayor.NombreRPT = "rptALIngresoVerif.rpt"

' Presentación preliminar del Informe
  rptMayor.PresentancionPreliminar

' Elimina los datos de la tabla
  BorraDatosTablaConsul

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
  cmdInforme.Enabled = True
 
End Sub

Private Sub BorraDatosTablaConsul()
'------------------------------------------------------------
' Propósito: Borra los datos de las tablas RPTALINGRESOVERIFDOC y _
             RPTALINGRESOVERIFDOCDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3

' Carga la sentencia
sSQL = "DELETE * FROM RPTALINGRESOVERIFDET"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Carga la sentencia
sSQL = "DELETE * FROM RPTALINGRESOVERIFDOC"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modTablaConsul.Cerrar

End Sub

Private Sub LlenaTablaConsul()
'------------------------------------------------------------
' Propósito: LLena las tablas RPTALINGRESOVERIFDOC y _
             RPTALINGRESOVERIFDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3
Dim i As Long
Dim sReg As Variant

' A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura,D.DescTipoDoc,E.NumDoc,E.FecMov,NúmeroVerif
' Recorre los documentos verificados
For Each sReg In mcolDocumentosVerif
      
      ' Carga la sentencia sSQL guarda el detalle en la tabla detalle
      sSQL = "INSERT INTO RPTALINGRESOVERIFDOC VALUES('" _
        & Var30(sReg, 1) & Var30(sReg, 2) & "','" _
        & Var30(sReg, 2) & "','" _
        & FechaDMA(Var30(sReg, 1)) & "','" _
        & Var30(sReg, 5) & "','" _
        & Var30(sReg, 6) & "','" _
        & Var9(Var30(sReg, 3)) & "','" _
        & FechaDMA(Var30(sReg, 7)) & "','" _
        & Var30(sReg, 8) & "')"
      
      ' Ejecuta la sentencia
      modTablaConsul.SQL = sSQL
      If modTablaConsul.Ejecutar = HAY_ERROR Then End
      modTablaConsul.Cerrar

Next sReg

' Recorre los datos del grid llena el detalle
For i = 1 To grdConsulta.Rows - 1
' "Fecha", "Orden", "Cantidad", "Unidad" , "Cod.", "Descripción", "Prec.Unit.", "Total", "Tipo"
    If grdConsulta.TextMatrix(i, 0) <> Empty Then
      
      ' Carga la sentencia sSQL guarda el detalle en la tabla detalle
      sSQL = "INSERT INTO RPTALINGRESOVERIFDET VALUES('" _
        & grdConsulta.TextMatrix(i, 0) & grdConsulta.TextMatrix(i, 1) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 2)) & ",'" _
        & grdConsulta.TextMatrix(i, 3) & "','" _
        & grdConsulta.TextMatrix(i, 4) & "','" _
        & grdConsulta.TextMatrix(i, 5) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 6)) & "," _
        & Var37(grdConsulta.TextMatrix(i, 7)) & ",'" _
        & grdConsulta.TextMatrix(i, 8) & "')"
      
      ' Ejecuta la sentencia
      modTablaConsul.SQL = sSQL
      If modTablaConsul.Ejecutar = HAY_ERROR Then End
      modTablaConsul.Cerrar
        
    End If
  
Next i

End Sub

Private Sub cmdSalir_Click()

' Descarga el formulario
  Unload Me

End Sub



Private Sub Form_Load()
Dim sSQL As String

' Carga los tamaños de las 9 columnas
' "Fecha", "Orden", "Cantidad", "Unidad" , "Cod.", "Descripción", "Prec.Unit.", "Total", "Tipo"
aTamañosColumnas = Array(0, 0, 1200, 2000, 950, 3500, 1200, 1200, 1050)
aTitulosColGrid = Array("Fecha", "Orden", "Cantidad", "Unidad", "Cod.", "Descripción", "Prec.Unit.", "Total", "Tipo")
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas
    
' Inicia alineamieto de la columna 5
grdConsulta.ColAlignment(3) = 1
grdConsulta.ColAlignment(5) = 1

' Habilita el optnumero
optNumero.Value = True

' Deshabilita el botón generar informe
cmdInforme.Enabled = False

End Sub

Private Sub optFechas_Click()
    ' Habilita controles por número
    HabilitaParametros
    'Limpia los controles
    LimpiarParametros
    ' Establece los campos obligatorios
    EstableceCamposObligatorios
    
End Sub

Private Sub optNumero_Click()
    ' Habilita controles por número
    HabilitaParametros
    'Limpia los controles
    LimpiarParametros
    ' Establece los campos obligatorios
    EstableceCamposObligatorios
End Sub

Private Sub LimpiarParametros()
'Limpia los controles
If optNumero.Value Then
    mskFechaIni.Text = "__/__/____"
    mskFechaFin.Text = "__/__/____"
Else
    txtNumDoc.Text = Empty
End If
End Sub

Private Sub HabilitaParametros()
' habilita controles dependiendo de los opt
If optNumero.Value Then
    txtNumDoc.Enabled = True
    mskFechaIni.Enabled = False
    mskFechaFin.Enabled = False
Else
    txtNumDoc.Enabled = False
    mskFechaIni.Enabled = True
    mskFechaFin.Enabled = True
End If
End Sub

Private Sub txtNumDoc_Change()
' Se valida el tamaño del número ingresado
If Len(txtNumDoc) = txtNumDoc.MaxLength Then
    txtNumDoc.BackColor = vbWhite
    ' Carga consulta
    CargaConsulta
Else
  txtNumDoc.BackColor = Obligatorio
  grdConsulta.Rows = 1
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
End If

End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)

'Si se presiona enter va al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    'valida entradas
    Var35 txtNumDoc, KeyAscii
End If
End Sub

Private Sub EstableceCamposObligatorios()
' ------------------------------------------------------------
' Propósito: Muestra de color amarillo los campos obligatorios
' Recibe: Nada
' Entrega:Nada
' ------------------------------------------------------------
If optNumero.Value Then
    txtNumDoc.BackColor = Obligatorio
    mskFechaIni.BackColor = vbWhite
    mskFechaFin.BackColor = vbWhite
Else
    txtNumDoc.BackColor = vbWhite
    mskFechaIni.BackColor = Obligatorio
    mskFechaFin.BackColor = Obligatorio
End If
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
    grdConsulta.Rows = 1
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
  End If

' Inicia progreso
   prgInforme.Max = 4
   prgInforme.Min = 0
   prgInforme.Value = 0

' Cargar los documetos verificados
  CargaDocumentosVerificados
  If mcolDocumentosVerif.Count = 0 Then
      MsgBox "No se tienen ingresos por compras", , "SGCcaijo-Amacén, Consulta ingresos por compras"
      ' Sale de el proceso y limpia el grid
      grdConsulta.Rows = 1
      ' Deshabilita el botón generar informe
      cmdInforme.Enabled = False
     ' Sale de el proceso
      Exit Sub
  End If
   prgInforme.Value = prgInforme.Value + 1
   
' Carga cursores de Almacén
   CargaIngrMercaderias
   prgInforme.Value = prgInforme.Value + 1
   CargaIngrActivos
   prgInforme.Value = prgInforme.Value + 1
   
'  Carga el grid consulta, inicia progreso
   CargarGridConsulta
   prgInforme.Value = prgInforme.Value + 1
   prgInforme.Value = 0
   
' Deshabilita el botón generar informe
  If grdConsulta.Rows > 1 Then
    cmdInforme.Enabled = True
  Else
    cmdInforme.Enabled = False
  End If
   
End Sub

Private Sub CargaIngrMercaderias()
' ----------------------------------------------------
' Propósito: Carga el cursor con los productos en almacén
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim sFechaIni As String
Dim sFechaFin As String
    
' Intevalo de fechas
sFechaIni = FechaAMD(mskFechaIni.Text)
sFechaFin = FechaAMD(mskFechaFin.Text)

If optNumero.Value = False Then
    ' Carga la consulta
    sSQL = "SELECT A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad," _
         & "A.PrecioUnit,G.Monto,P.Medida " _
         & "FROM ALMACEN_INGRESOS A,PRODUCTOS P, GASTOS G, " _
         & "EGRESOS E " _
         & "WHERE A.Orden=G.Orden and A.IdProd=G.CodConcepto and " _
         & "A.IdProd=P.IdProd and " _
         & "G.Orden=E.Orden and E.Anulado='NO' and " _
         & "(A.Fecha between '" & sFechaIni & "' and '" & sFechaFin & "') " _
         & "ORDER BY A.Fecha, A.Orden"
Else
     ' Carga la consulta
    sSQL = "SELECT A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad," _
         & "A.PrecioUnit,G.Monto,P.Medida " _
         & "FROM ALMACEN_INGRESOS A,PRODUCTOS P, GASTOS G, " _
         & "EGRESOS E " _
         & "WHERE A.Orden=G.Orden and A.IdProd=G.CodConcepto and " _
         & "A.IdProd=P.IdProd and " _
         & "G.Orden=E.Orden and E.Anulado='NO' and " _
         & "A.Fecha='" & Var30(mcolDocumentosVerif.Item(1), 1) & "' And " _
         & "A.Orden='" & Var30(mcolDocumentosVerif.Item(1), 2) & "' " _
         & "ORDER BY A.Fecha, A.Orden"
End If
' Ejecuta la sentencia
mcurIngresoMercaderias.SQL = sSQL
If mcurIngresoMercaderias.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaIngrActivos()
' ----------------------------------------------------
' Propósito: Carga el cursor de los activos fijos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
If optNumero.Value = False Then

    ' Carga la consulta
    sSQL = "SELECT A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad," _
         & "G.Monto/G.Cantidad,G.Monto,P.Medida " _
         & "FROM ACTIVOFIJO_INGRESOS A,PRODUCTOS P, GASTOS G, " _
         & "EGRESOS E " _
         & "WHERE A.Orden=G.Orden and A.IdProd=G.CodConcepto and " _
         & "A.IdProd=P.IdProd and " _
         & "G.Orden=E.Orden and E.Anulado='NO' and " _
         & "(A.Fecha between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
         & "ORDER BY A.Fecha, A.Orden"
Else
       ' Carga la consulta
    sSQL = "SELECT A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad," _
         & "G.Monto/G.Cantidad,G.Monto,P.Medida " _
         & "FROM ACTIVOFIJO_INGRESOS A,PRODUCTOS P, GASTOS G, " _
         & "EGRESOS E " _
         & "WHERE A.Orden=G.Orden and A.IdProd=G.CodConcepto and " _
         & "A.IdProd=P.IdProd and " _
         & "G.Orden=E.Orden and E.Anulado='NO' and " _
         & "A.Fecha='" & Var30(mcolDocumentosVerif.Item(1), 1) & "' And " _
         & "A.Orden='" & Var30(mcolDocumentosVerif.Item(1), 2) & "' " _
         & "ORDER BY A.Fecha, A.Orden"

End If
' Ejecuta la sentencia
mcurIngresoActivos.SQL = sSQL
If mcurIngresoActivos.Abrir = HAY_ERROR Then End

End Sub

Private Function AveriguaNroVerifInicial() As Long
Dim brecorre As Boolean
Dim NroVerifInicial As Long
Dim sDiaInicial As String
Dim sSQL As String
Dim curOrdenActivo As New clsBD2
Dim curOrdenMercad As New clsBD2

' Inicializa el nro
NroVerifInicial = 0
sDiaInicial = "01/01/" & Right(mskFechaIni, 4)

' Averigua el Nro, si el día de inicio es Plan28 al 01/01/AAAA
If mskFechaIni <> sDiaInicial Then
    ' Carga la sentencia que carga activos
    sSQL = "SELECT DISTINCT A.Fecha,A.Orden " _
        & "FROM ACTIVOFIJO_INGRESOS A, EGRESOS E " _
        & "WHERE A.Orden=E.Orden and E.Anulado='NO' and " _
        & "(A.Fecha Between '" & FechaAMD(sDiaInicial) & "' and '" & AnioMesDiaAnterior(FechaAMD(mskFechaIni)) & "') " _
        & "ORDER BY A.Fecha,A.Orden"
    
    ' Ejecuta la sentencia
    curOrdenActivo.SQL = sSQL
    If curOrdenActivo.Abrir = HAY_ERROR Then End
    
    ' Carga la sentencia que carga Mercaderias
    sSQL = "SELECT DISTINCT A.Fecha,A.Orden " _
        & "FROM ALMACEN_INGRESOS A, EGRESOS E " _
        & "WHERE A.Orden=E.Orden and E.Anulado='NO' and " _
        & "(A.Fecha Between '" & FechaAMD(sDiaInicial) & "' and '" & AnioMesDiaAnterior(FechaAMD(mskFechaIni)) & "') " _
        & "ORDER BY A.Fecha,A.Orden"
    
    ' Ejecuta la sentencia
    curOrdenMercad.SQL = sSQL
    If curOrdenMercad.Abrir = HAY_ERROR Then End
    
    ' Inicializa recorrer
    brecorre = True
    ' Recorre los cursores y lo añade a la colección de modulo
    Do While brecorre = True
        ' Verifica si ambos cursores son vacios
        If curOrdenActivo.EOF And curOrdenMercad.EOF Then
            ' Sale de recorrer
            brecorre = False
        Else
            If curOrdenActivo.EOF Then ' Recorre Mercaderías
                ' Siguiente número
                NroVerifInicial = NroVerifInicial + 1
                ' Mueve al siguiente elemento de mercaderías
                curOrdenMercad.MoverSiguiente
            ElseIf curOrdenMercad.EOF Then ' Recorre Activos
                ' Siguiente número
                NroVerifInicial = NroVerifInicial + 1
                ' Mueve al siguiente elemento de Activos
                curOrdenActivo.MoverSiguiente
            Else ' Ninguno de los cursores es vacío
                ' Añade el Menor
                If curOrdenMercad.campo(0) & curOrdenMercad.campo(1) _
                  = curOrdenActivo.campo(0) & curOrdenActivo.campo(1) Then
                        ' Siguiente número
                        NroVerifInicial = NroVerifInicial + 1
                       ' Mueve al siguiente elemento de mercaderías
                        curOrdenMercad.MoverSiguiente
                        ' Mueve al siguiente elemento de Activos
                        curOrdenActivo.MoverSiguiente
                ElseIf curOrdenMercad.campo(0) & curOrdenMercad.campo(1) _
                  < curOrdenActivo.campo(0) & curOrdenActivo.campo(1) Then
                    ' Siguiente número
                    NroVerifInicial = NroVerifInicial + 1
                    ' Mueve al siguiente elemento de mercaderías
                    curOrdenMercad.MoverSiguiente
                Else
                    ' Siguiente número
                    NroVerifInicial = NroVerifInicial + 1
                    ' Mueve al siguiente elemento de Activos
                    curOrdenActivo.MoverSiguiente
                End If
            End If
        End If
    Loop
    
    ' Cierra los cursore
    curOrdenActivo.Cerrar
    curOrdenMercad.Cerrar
    
End If

' Devuelve el número inicial de verificaciones
AveriguaNroVerifInicial = NroVerifInicial

End Function

Private Function CargaDocumentosVerificados() As Boolean
' ----------------------------------------------------
' Propósito: Carga la colección que almacena los documentos verificados por fecha
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curOrdenActivo As New clsBD2
Dim curOrdenMercad As New clsBD2
Dim brecorre As Boolean
Dim iNroInicial As Long
Dim sAnio As String
Dim sNumDoc As String
Dim sFechaIni As String
Dim sFechaFin As String
On Error GoTo mnjError

' Limpia la colección
Set mcolDocumentosVerif = Nothing

'Verifica el año
If optNumero.Value Then
    ' Averigua el número de salida inicial de la consulta
    iNroInicial = 0
    sAnio = Left(txtNumDoc, 4)
    sFechaIni = sAnio & "0101"
    sFechaFin = sAnio & "1231"
Else
    ' Averigua el número de salida inicial de la consulta
    iNroInicial = AveriguaNroVerifInicial
    sAnio = Right(mskFechaIni, 4)
    sFechaIni = FechaAMD(mskFechaIni.Text)
    sFechaFin = FechaAMD(mskFechaFin.Text)
End If

    ' Carga la sentencia que carga activos
    sSQL = "SELECT DISTINCT A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura," _
        & "D.DescTipoDoc,E.NumDoc,E.FecMov " _
        & "FROM ACTIVOFIJO_INGRESOS A, EGRESOS E,PROVEEDORES P,TIPO_DOCUM D " _
        & "WHERE A.Orden=E.Orden and E.Anulado='NO' and  E.IdProveedor=P.IdProveedor and " _
        & "E.IdTipoDoc=D.IdTipoDoc and " _
        & "(A.Fecha Between '" & sFechaIni & "' and '" & sFechaFin & "') " _
        & "ORDER BY A.Fecha,A.Orden"
    
    ' Ejecuta la sentencia
    curOrdenActivo.SQL = sSQL
    If curOrdenActivo.Abrir = HAY_ERROR Then End
    
    ' Carga la sentencia que carga Mercaderias
    sSQL = "SELECT DISTINCT A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura," _
        & "D.DescTipoDoc,E.NumDoc,E.FecMov " _
        & "FROM ALMACEN_INGRESOS A, EGRESOS E,PROVEEDORES P,TIPO_DOCUM D " _
        & "WHERE A.Orden=E.Orden and E.Anulado='NO' and  E.IdProveedor=P.IdProveedor and " _
        & "E.IdTipoDoc=D.IdTipoDoc and " _
        & "(A.Fecha Between '" & sFechaIni & "' and '" & sFechaFin & "') " _
        & "ORDER BY A.Fecha,A.Orden"

' Ejecuta la sentencia
curOrdenMercad.SQL = sSQL
If curOrdenMercad.Abrir = HAY_ERROR Then End

' Inicializa recorrer
brecorre = True
' Recorre los cursores y lo añade a la colección de modulo
Do While brecorre = True
    ' Verifica si ambos cursores son vacios
    If curOrdenActivo.EOF And curOrdenMercad.EOF Then
        ' Sale de recorrer
        brecorre = False
    Else
        If curOrdenActivo.EOF Then ' Recorre Mercaderías
            ' Incrementa el número inicial
            iNroInicial = iNroInicial + 1
            mcolDocumentosVerif.Add Item:=curOrdenMercad.campo(0) _
                    & "¯" & curOrdenMercad.campo(1) _
                    & "¯" & curOrdenMercad.campo(2) _
                    & "¯" & curOrdenMercad.campo(3) _
                    & "¯" & curOrdenMercad.campo(4) _
                    & "¯" & curOrdenMercad.campo(5) _
                    & "¯" & curOrdenMercad.campo(6) _
                    & "¯" & sAnio & Format(iNroInicial, "0000000#"), _
                    Key:=sAnio & Format(iNroInicial, "0000000#")
            ' Mueve al siguiente elemento de mercaderías
            curOrdenMercad.MoverSiguiente
        ElseIf curOrdenMercad.EOF Then ' Recorre Activos
            ' Incrementa el número inicial
            iNroInicial = iNroInicial + 1
            mcolDocumentosVerif.Add Item:=curOrdenActivo.campo(0) _
                    & "¯" & curOrdenActivo.campo(1) _
                    & "¯" & curOrdenActivo.campo(2) _
                    & "¯" & curOrdenActivo.campo(3) _
                    & "¯" & curOrdenActivo.campo(4) _
                    & "¯" & curOrdenActivo.campo(5) _
                    & "¯" & curOrdenActivo.campo(6) _
                    & "¯" & sAnio & Format(iNroInicial, "0000000#"), _
                    Key:=sAnio & Format(iNroInicial, "0000000#")
                    
            ' Mueve al siguiente elemento de Activos
            curOrdenActivo.MoverSiguiente
        Else ' Ninguno de los cursores es vacío
            ' Añade el Menor
            If curOrdenMercad.campo(0) & curOrdenMercad.campo(1) _
              = curOrdenActivo.campo(0) & curOrdenActivo.campo(1) Then
                    ' Incrementa el número inicial
                    iNroInicial = iNroInicial + 1
                    ' Añade mercaderias y avanza ambos
                    mcolDocumentosVerif.Add Item:=curOrdenMercad.campo(0) _
                            & "¯" & curOrdenMercad.campo(1) _
                            & "¯" & curOrdenMercad.campo(2) _
                            & "¯" & curOrdenMercad.campo(3) _
                            & "¯" & curOrdenMercad.campo(4) _
                            & "¯" & curOrdenMercad.campo(5) _
                            & "¯" & curOrdenMercad.campo(6) _
                            & "¯" & sAnio & Format(iNroInicial, "0000000#"), _
                            Key:=sAnio & Format(iNroInicial, "0000000#")
                   ' Mueve al siguiente elemento de mercaderías
                    curOrdenMercad.MoverSiguiente
                    ' Mueve al siguiente elemento de Activos
                    curOrdenActivo.MoverSiguiente
            ElseIf curOrdenMercad.campo(0) & curOrdenMercad.campo(1) _
              < curOrdenActivo.campo(0) & curOrdenActivo.campo(1) Then
                ' Incrementa el número inicial
                iNroInicial = iNroInicial + 1
                mcolDocumentosVerif.Add Item:=curOrdenMercad.campo(0) _
                        & "¯" & curOrdenMercad.campo(1) _
                        & "¯" & curOrdenMercad.campo(2) _
                        & "¯" & curOrdenMercad.campo(3) _
                        & "¯" & curOrdenMercad.campo(4) _
                        & "¯" & curOrdenMercad.campo(5) _
                        & "¯" & curOrdenMercad.campo(6) _
                        & "¯" & sAnio & Format(iNroInicial, "0000000#"), _
                        Key:=sAnio & Format(iNroInicial, "0000000#")
                ' Mueve al siguiente elemento de mercaderías
                curOrdenMercad.MoverSiguiente
            Else
                ' Incrementa el número inicial
                iNroInicial = iNroInicial + 1
                mcolDocumentosVerif.Add Item:=curOrdenActivo.campo(0) _
                        & "¯" & curOrdenActivo.campo(1) _
                        & "¯" & curOrdenActivo.campo(2) _
                        & "¯" & curOrdenActivo.campo(3) _
                        & "¯" & curOrdenActivo.campo(4) _
                        & "¯" & curOrdenActivo.campo(5) _
                        & "¯" & curOrdenActivo.campo(6) _
                        & "¯" & sAnio & Format(iNroInicial, "0000000#"), _
                        Key:=sAnio & Format(iNroInicial, "0000000#")
                ' Mueve al siguiente elemento de Activos
                curOrdenActivo.MoverSiguiente
            End If
        End If
    End If
Loop

' Cierra los cursore
curOrdenActivo.Cerrar
curOrdenMercad.Cerrar

If optNumero.Value = True Then
    sNumDoc = mcolDocumentosVerif.Item(txtNumDoc.Text)
    Set mcolDocumentosVerif = Nothing
    mcolDocumentosVerif.Add sNumDoc
End If
Exit Function
'----------
mnjError:
Set mcolDocumentosVerif = Nothing
End Function

Private Sub CargarGridConsulta()
' ----------------------------------------------------
' Propósito: Arma la consulta en el grid
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim brecorre As Boolean
Dim sReg As Variant
' grd:"Fecha", "Orden", "Cantidad", "Cod.", "Descripción", "Unidad", "Prec.Unit.", "Total", "Tipo"
' cur:A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad,A.PrecioUnit,G.Monto, Unidad
' Col:A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura,D.DescTipoDoc,E.NumDoc,E.FecMov,Número
' Inicializa el grid
grdConsulta.Rows = 1
grdConsulta.ScrollBars = flexScrollBarNone
grdConsulta.Visible = False

' Recorre los documentos verificados por día
For Each sReg In mcolDocumentosVerif
    
   mdblPrecioTotal = 0
   ' Añade el documento al grid
   grdConsulta.AddItem vbTab & vbTab & "Nro:" & vbTab & _
        Var30(sReg, 8) & vbTab & "Documento:" & vbTab & _
        Var30(sReg, 4) & "/" & Var30(sReg, 6) & vbTab & _
        vbTab & "Fecha:" & vbTab & FechaDMA(Var30(sReg, 1))
  ' Colorea el grid
   grdConsulta.Row = grdConsulta.Rows - 1
   MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
   ' Añade los Activos
   MuestraActivos (sReg)
   ' Añade las Mercaderías
   MuestraMercaderias (sReg)
   ' Muestra el total del documento
   ' Añade el documento al grid
   grdConsulta.AddItem vbTab & vbTab & vbTab & vbTab & _
        "Total" & vbTab & vbTab & vbTab & _
         Format(mdblPrecioTotal, "###,###,##0.00")
   'Colorea el grid
   grdConsulta.Row = grdConsulta.Rows - 1
   MarcarFilaGRID grdConsulta, vbBlack, vbGray

Next sReg ' Siguiente registro

' Coloca las barras de desplazamiento
grdConsulta.ScrollBars = flexScrollBarBoth
grdConsulta.Visible = True

' Cierra los cursores generales
mcurIngresoActivos.Cerrar
mcurIngresoMercaderias.Cerrar

End Sub

Private Sub MuestraActivos(sOrden As String)
'---------------------------------------------------------------
' Propósito: Muestra los productos que entraron como activos
' Recibe: sOrden, Documento con el que se compró
' Entrga:Nada
'---------------------------------------------------------------
' grd:"Fecha", "Orden", "Cantidad", "Unidad" , "Cod.", "Descripción", "Prec.Unit.", "Total", "Tipo"
' cur:A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad,A.PrecioUnit,G.Monto,Unidad
' sOrden:A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura,D.DescTipoDoc,E.NumDoc,E.FecMov
Dim brecorre As Boolean
' Carga los datos de ingreso de activo
brecorre = True
Do While brecorre
    ' Verifica si el cursor es vacío
    If mcurIngresoActivos.EOF Then
        brecorre = False
    Else
        ' Verifica que los datos son de la Cta y fecha
        If mcurIngresoActivos.campo(0) = Var30(sOrden, 1) And _
           mcurIngresoActivos.campo(1) = Var30(sOrden, 2) Then
              ' Añade el Activo
              grdConsulta.AddItem mcurIngresoActivos.campo(0) & vbTab & _
                               mcurIngresoActivos.campo(1) & vbTab & _
                               Format(mcurIngresoActivos.campo(4), "##0.00") & vbTab & _
                               mcurIngresoActivos.campo(7) & vbTab & _
                               mcurIngresoActivos.campo(2) & vbTab & _
                               mcurIngresoActivos.campo(3) & vbTab & _
                               Format(mcurIngresoActivos.campo(5), "###,###,##0.00") & vbTab & _
                               Format(mcurIngresoActivos.campo(6), "###,###,##0.00") & vbTab & _
                               "Activo"
              ' Acumula en el debe
              mdblPrecioTotal = mdblPrecioTotal + Val(mcurIngresoActivos.campo(6))
              ' Pasa al siguiente control
              mcurIngresoActivos.MoverSiguiente
        Else
            ' Sale de recorrer el cursor
            brecorre = False
        End If
     End If ' fin de ver si es vacío
Loop ' fin de recorrer cursor

End Sub

Private Sub MuestraMercaderias(sOrden As String)
'---------------------------------------------------------------
' Propósito: Muestra los productos que entraron como Mercaderías
' Recibe: sOrden, Documento con el que se compró
' Entrga:Nada
'---------------------------------------------------------------
' grd:' "Fecha", "Orden", "Cantidad", "Cod.", "Descripción", "Unidad", "Prec.Unit.", "Total", "Tipo"
' cur:A.Fecha,A.Orden,A.IdProd,P.DescProd,G.Cantidad,A.PrecioUnit,G.Monto,Unidad
' sOrden:A.Fecha,A.Orden,P.DescProveedor,D.Abreviatura,D.DescTipoDoc,E.NumDoc,E.FecMov
Dim brecorre As Boolean
' Carga los datos de ingreso de activo
brecorre = True
Do While brecorre
    ' Verifica si el cursor es vacío
    If mcurIngresoMercaderias.EOF Then
        brecorre = False
    Else
        ' Verifica que los datos son de la Cta y fecha
        If mcurIngresoMercaderias.campo(0) = Var30(sOrden, 1) And _
           mcurIngresoMercaderias.campo(1) = Var30(sOrden, 2) Then
              ' Añade el Activo
              ' Añade el Activo
              grdConsulta.AddItem mcurIngresoMercaderias.campo(0) & vbTab & _
                               mcurIngresoMercaderias.campo(1) & vbTab & _
                               Format(mcurIngresoMercaderias.campo(4), "##0.00") & vbTab & _
                               mcurIngresoMercaderias.campo(7) & vbTab & _
                               mcurIngresoMercaderias.campo(2) & vbTab & _
                               mcurIngresoMercaderias.campo(3) & vbTab & _
                               Format(mcurIngresoMercaderias.campo(5), "###,###,##0.00") & vbTab & _
                               Format(mcurIngresoMercaderias.campo(6), "###,###,##0.00") & vbTab & _
                               "Mercadería"
              ' Acumula en el debe
              mdblPrecioTotal = mdblPrecioTotal + Val(mcurIngresoMercaderias.campo(6))
              ' Pasa al siguiente control
              mcurIngresoMercaderias.MoverSiguiente
        Else
            ' Sale de recorrer el cursor
            brecorre = False
        End If
    End If ' fin de verificar si esta vacío
Loop ' fin de recorrer cursor

End Sub

Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito: Verifica si esta bien los datos para ejecutar _
            la consulta
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curContb As New clsBD2
If optNumero.Value = False Then
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
End If
' Datos correctos
fbOkDatosIntroducidos = True
 
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Libera la colección
Set mcolDocumentosVerif = Nothing

End Sub

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  ' Carga consulta
  CargaConsulta
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
End If

End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

'Si se presiona enter va al siguiente control
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
  grdConsulta.Rows = 1
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
  
End If

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

'Si se presiona enter va al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
