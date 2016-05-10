VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCBConsulRegistroCompras 
   Caption         =   "Consulta de registro de compras"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   72
   Icon            =   "SCCBConsulRegistroCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   960
      Top             =   8160
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
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   11415
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   7440
         TabIndex        =   9
         Top             =   180
         Width           =   3015
         Begin MSMask.MaskEdBox mskFechaConsulta 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Fecha de consulta:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   285
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consulta"
         Height          =   735
         Left            =   480
         TabIndex        =   7
         Top             =   180
         Width           =   5055
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   315
            Left            =   3240
            TabIndex        =   1
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   2880
            TabIndex        =   11
            Top             =   255
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consulta del "
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   285
            Width           =   915
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6735
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FillStyle       =   1
   End
End
Attribute VB_Name = "frmCBConsulRegistroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de Colecciones
Dim mcolCodNroCta As New Collection
Dim mcolCodDesNroCta As New Collection
'Variable del CodCuenta
Dim msNroCta As String

Private Sub cmdInforme_Click()
Dim sSQL As String
Dim rptRegCompras As New clsBD4

' Deshabilita el botón informe
  cmdInforme.Enabled = False
  
' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón informe
       cmdInforme.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla con datos
  LlenarTablaRPTCBREGCOMPRAS

' Formulario
  Set rptRegCompras.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptRegCompras.AsignarRpt

' Formula/s de Crystal.
  rptRegCompras.Formulas.Add "Fecha='DEL " & mskFechaIni.Text & " AL " & mskFechaFin.Text & "'"
  'rptRegCompras.Formulas.Add "TotalInformes='" & txtTotalIngresos.Text & "'"
  
' Clausula WHERE de las relaciones del rpt.
  rptRegCompras.FiltroSelectionFormula = ""

' Nombre del fichero
  rptRegCompras.NombreRPT = "RPTCBREGISTROCOMPRAS.rpt"

' Presentación preliminar del Informe
  rptRegCompras.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTCBREGCOMPRAS"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True

End Sub

Private Sub LlenarTablaRPTCBREGCOMPRAS()
'-----------------------------------------------------
'Propósito  : Llena la tabla con los datos del grdConsulta
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modRegCompras As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
    'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
     sSQL = "INSERT INTO RPTCBREGCOMPRAS VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 3) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 4) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 6)) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 7)) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 8))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 9))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 10))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 11) & "', " _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modRegCompras.SQL = sSQL
    
    'Verifica si hay error
    If modRegCompras.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modRegCompras.Cerrar

Next i

End Sub


Private Sub cmdSalir_Click()
'Descarga el formulario
Unload Me
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
End If
' Verifica si los datos obligatorios se ha llenado
If mskFechaIni.BackColor <> vbWhite Or _
   mskFechaFin.BackColor <> vbWhite Then
   fbOkDatosIntroducidos = False
   Exit Function
End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function



Private Sub Form_Load()
Dim sSQL As String

'Establece la ubicación del formulario
Me.Top = 0

' Carga los títulos del grid
'"FECHA", "COMPROBANTE NRO.", "TIPO NUMERO", "NUMERO RUC", "IDPROV", "PROVEEDOR", "DESCRIPCION", "VALOR VENTA", "IMPUESTO", "TOTAL", "NRO. ORDEN"
aTitulosColGrid = Array("FECHA MOV", "FECHA DOC", "N°COMPROBANTE", "TIPO DOC", "N°RUC/DNI", "IDPROV", "PROVEEDOR", "DESCRIPCION", "VALOR VENTA", "IMPUESTO", "TOTAL", "NRO. ORDEN")
aTamañosColumnas = Array(1000, 1000, 1600, 900, 1200, 0, 3500, 2500, 1300, 1300, 1300, 1300)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Carga la fecha del sistema
mskFechaConsulta.Text = gsFecTrabajo

'Deshabilita el control cmdInforme
cmdInforme.Enabled = False

'Establece campos obligatorios
EstableceCamposObligatorios

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

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  ' Carga consulta
  CargaRegistroCompras
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
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
    
  'Carga los registros de compras
  CargaRegistroCompras
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If
End Sub

Private Sub CargaRegistroCompras()
' ----------------------------------------------------
' Propósito : Arma la consulta de salidas de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curRegCompras As New clsBD2
Dim sProdServ As String
Dim dblImpuestos, dblMontoTotal As Double
Dim sFecha As String
Dim sFecDoc As String
Dim sComprobante, sTipoNro, sNroRuc, sIdProv, sProv, sDescripcion, sOrden, sOrdenSig As String

' Verifica los datos introducidos para la consulta
If fbOkDatosIntroducidos = False Then
  'Limpia el grdConsulta
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
  Exit Sub
End If

'Fecha, Comprobante, TipoDoc ,NumRuc ,IdProveedor, Proveedor, Descripción, Impuesto, MontoTotal, Orden--Se aumenta FECDOC
sSQL = "SELECT Distinct E.FecMov, E.IdTipoDoc +'  ' +E.NumDoc, P.RUC_DNI,P.Numero,E.IdProveedor, P.DescProveedor, " & _
               "G.Concepto, I.Monto, E.MontoAfectado, E.Orden, E.FecDoc " & _
       "FROM (((EGRESOS E INNER JOIN PROVEEDORES P ON E.IdProveedor= P.IdProveedor) " & _
             "INNER JOIN GASTOS G ON E.Orden= G.Orden ) " & _
             "LEFT OUTER JOIN MOV_IMPUESTOS I ON E.Orden=I.Orden) " & _
       "WHERE E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " & _
       " E.Anulado='NO' And E.IdProy <> Null " & _
       "ORDER BY E.FecMov ,E.Orden "
        
' Ejecuta la sentencia
curRegCompras.SQL = sSQL
If curRegCompras.Abrir = HAY_ERROR Then End
'Inicializa la  variable
dblImpuestos = 0
dblMontoTotal = 0
'Verifica que no hay registros en la consulta
If curRegCompras.EOF Then
    'Mensaje no hay exsitencias en almacén
    MsgBox "No hay registro de compras entre estas fechas", vbInformation + vbOKOnly, "Caja Bancos- Consulta de Registro de compras"
    
    ' cierra la consulta
    curRegCompras.Cerrar
    
    'Limpiar Grid
    grdConsulta.Rows = 1
    
    'Termina la ejecución del procedimiento
    Exit Sub
Else
     'Asigna el Orden del primer registro
     'Fecha, Comprobante, TipoDoc ,NumRuc, IdProveedor, _
     'Proveedor, Descripción, Impuesto, MontoTotal, Orden
     sFecha = curRegCompras.campo(0)
     sComprobante = curRegCompras.campo(1)
     sTipoNro = curRegCompras.campo(2)
     sNroRuc = curRegCompras.campo(3)
     sIdProv = curRegCompras.campo(4)
     sProv = curRegCompras.campo(5)
     sDescripcion = curRegCompras.campo(6)
     dblMontoTotal = curRegCompras.campo(8)
     sOrden = curRegCompras.campo(9)
     sFecDoc = curRegCompras.campo(10)
     
     'Verifica si es nulo el curRegCOmpras.campo(7)
     If IsNull(curRegCompras.campo(7)) Then
        'Es nulo el curRegCompras
        dblImpuestos = 0
     Else
        'No es nulo
        dblImpuestos = curRegCompras.campo(7)
     End If
     
    ' Recorre el cursor ingresos a cta
    Do While Not curRegCompras.EOF
        
        'Mueve al siguiente registro
        curRegCompras.MoverSiguiente
                
        ' Verifica que no sea el final del cursor
            If Not curRegCompras.EOF Then
                'Asigna el cursor al sIdPersona
                sOrdenSig = curRegCompras.campo(9)
            Else
                sOrdenSig = ""
            End If
            
            '****************************************************
            ' SI cambia de Orden
            If sOrden <> sOrdenSig Then
            
                'Verifica si es producto
                If sDescripcion = "P" Then
                    sProdServ = "COMPRA DE PRODUCTOS"
                Else
                    sProdServ = "PAGO DE SERVICIOS"
                End If
                
                'Fecha, Comprobante, TipoDoc ,NumRuc, IdProveedor, Proveedor, Descripción, Impuesto, MontoTotal, Orden
                '"FECHA", "COMPROBANTE NRO.", "TIPO NUMERO", "NUMERO RUC", "IDPROV", "PROVEEDOR", "DESCRIPCION", "VALOR VENTA", "IMPUESTO", "TOTAL", "NRO. ORDEN"
                grdConsulta.AddItem FechaDMA(sFecha) & vbTab & _
                                    FechaDMA(sFecDoc) & vbTab & _
                                    sComprobante & vbTab & _
                                    sTipoNro & vbTab & _
                                    sNroRuc & vbTab & _
                                    sIdProv & vbTab & _
                                    sProv & vbTab & _
                                    sProdServ & vbTab & _
                                    Format((dblMontoTotal - dblImpuestos), "###,###,##0.00") & vbTab & _
                                    Format(dblImpuestos, "###,###,##0.00") & vbTab & _
                                    Format(dblMontoTotal, "###,###,##0.00") & vbTab & _
                                    sOrden & vbTab & _
                                    sFecDoc
        
                                                      
                'Inicializa los totales, el nombre e identificador de la nueva persona
                If Not curRegCompras.EOF Then
                    sOrden = sOrdenSig
                    'Asigna el Orden del primer registro
                    'Fecha, Comprobante, TipoDoc ,NumRuc, IdProveedor, _
                    'Proveedor, Descripción, Impuesto, MontoTotal, Orden
                    sFecha = curRegCompras.campo(0)
                    sComprobante = curRegCompras.campo(1)
                    sTipoNro = curRegCompras.campo(2)
                    sNroRuc = curRegCompras.campo(3)
                    sIdProv = curRegCompras.campo(4)
                    sProv = curRegCompras.campo(5)
                    sDescripcion = curRegCompras.campo(6)
                    dblMontoTotal = curRegCompras.campo(8)
                    sFecDoc = curRegCompras.campo(10)
                    'Verifica si es nulo el cursor
                    If IsNull(curRegCompras.campo(7)) Then
                        dblImpuestos = 0
                    Else
                        'Copia el valor del cursor
                        dblImpuestos = curRegCompras.campo(7)
                    End If
                End If
            Else
                'Suma los impuestos
                If IsNull(curRegCompras.campo(7)) Then
                    dblImpuestos = 0
                Else
                    'Acumula los impuestos
                    dblImpuestos = dblImpuestos + curRegCompras.campo(7)
                End If
            End If
    Loop
    
End If

'Cierra el cursor
curRegCompras.Cerrar

'Habilita el cmdInforme
cmdInforme.Enabled = True

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
