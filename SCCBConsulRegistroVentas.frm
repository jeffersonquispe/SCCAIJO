VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCBConsulRegistroVentas 
   Caption         =   "Consulta de registro de ventas"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   72
   Icon            =   "SCCBConsulRegistroVentas.frx":0000
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
      Left            =   10560
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
      Cols            =   17
      FillStyle       =   1
   End
End
Attribute VB_Name = "frmCBConsulRegistroVentas"
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
Dim rptRegVentas As New clsBD4

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
  LlenarTablaRPTCBREGVentas

' Formulario
  Set rptRegVentas.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptRegVentas.AsignarRpt

' Formula/s de Crystal.
  rptRegVentas.Formulas.Add "Fecha='DEL " & mskFechaIni.Text & " AL " & mskFechaFin.Text & "'"
  'rptRegVentas.Formulas.Add "TotalInformes='" & txtTotalIngresos.Text & "'"
  
' Clausula WHERE de las relaciones del rpt.
  rptRegVentas.FiltroSelectionFormula = ""

' Nombre del fichero
  rptRegVentas.NombreRPT = "RPTCBREGISTROVENTAS.rpt"

' Presentación preliminar del Informe
  rptRegVentas.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
sSQL = "DELETE * FROM RPTCBREGVENTAS"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True

End Sub

Private Sub LlenarTablaRPTCBREGVentas()
'-----------------------------------------------------
'Propósito  : Llena la tabla con los datos del grdConsulta
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modRegVentas As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
    'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
     sSQL = "INSERT INTO RPTCBREGVENTAS VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 4) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 5) & "', " _
      & "'" & grdConsulta.TextMatrix(i, 6) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 7)) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 8)) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 9) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 10) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 11) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 12) & "'," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 13))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 14))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 15))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 16) & "'," _
     & "'" & grdConsulta.TextMatrix(i, 3) & "'," _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modRegVentas.SQL = sSQL
    
    'Verifica si hay error
    If modRegVentas.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modRegVentas.Cerrar

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
'Periodo, Fecha de Comprobante, TipoDoc ,Serie,NumeroDoc,TipoDoc,NumRucDni ,DescCliente, Monto SubTotal,
'MontoIgv , MontoTotal, Orden, FecMov
aTitulosColGrid = Array("ORDEN", "PERIODO", "FECHA MOV.", "FECHA DOC.", "TIPO", "DOCUMENTO", "SERIE", "NUM. DOC", "TIPO", "DNI/RUC", "NUMERO", "CLIENTE", "DESC. VENTA", "VALOR VENTA", "IGV", "TOTAL", "FECHA")
aTamañosColumnas = Array(1200, 1000, 1100, 1100, 600, 1200, 600, 900, 700, 700, 1200, 4000, 2750, 1300, 1300, 1300, 1200)
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
  CargaRegistroVentas
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
    
  'Carga los registros de Ventas
  CargaRegistroVentas
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If
End Sub

Private Sub CargaRegistroVentas()
' ----------------------------------------------------
' Propósito : Arma la consulta de salidas de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curRegVentas As New clsBD2
Dim sProdServ As String
Dim dblIGV, dblSubTotal, dblMontoTotal As Double
Dim sFecha, sFecDoc, sPeriodo As String
Dim sFechaComprobante, sTipoDoc, sSerie, sNumDoc, sTipoRUC_DNI, sNumero, sDescripcion, sOrden As String
Dim sOrdenSig, sFecMov, sDescVenta, sRUCDNI, sDocumento As String

' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
  'Limpia el grdConsulta
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
  Exit Sub
End If

'Periodo, Fecha de Comprobante, TipoDoc ,Serie,NumeroDoc,TipoDocRD,NumRucDni ,DescCliente, Monto SubTotal,
'MontoIgv , MontoTotal, Orden, FecMov
 
'  sSQL = "SELECT DISTINCT Mid( V.FecMov,1,6)+'00' AS Periodo,CDate(Mid(V.FecMov,7,2) & " / " & Mid(V.FecMov,5,2) & " / " & Mid(V.FecMov,3,2)) AS [Fecha E Comprobante]," _
'  & " CInt(V.IdTipoDoc) AS TipoDocumento, iif(InStr(V.NumDoc," - ")=0,'000',left(V.NumDoc,InStr(V.NumDoc," - ")-1)) AS Serie" _
'  & ",Mid(V.NumDoc, InStr(V.NumDoc, " - ") + 1)As NumeroDoc," _
'  & " iif(VC.RUC_DNI='RUC',6,1) AS [Tipo de Documento], VC.Numero AS NumeroRD, VC.DescCliente, V.MontoSubtotal," _
'  & " V.MontoIgv , V.MontoTotal, V.Orden, V.FecMov" _
'  & " FROM Ventas AS V INNER JOIN Ventas_CLientes AS VC ON V.IdCLiente=VC.IdCliente" _
'  & " WHERE V.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And V.Anulado='NO'" _
'  & " ORDER BY V.FecMov, V.Orden"

sSQL = "SELECT DISTINCT V.Orden, Mid(V.FecMov,1,6)+'00' AS Periodo,CDate(Mid(V.FecMov,7,2) + '/' + Mid(V.FecMov,5,2) + '/' + Mid(V.FecMov,3,2))," _
& " CInt(V.IdTipoDoc),iif(InStr(V.NumDoc,'-')=0,'000',left(V.NumDoc,InStr(V.NumDoc,'-')-1))," _
& " Mid(V.NumDoc, InStr(V.NumDoc, '-') + 1)," _
& " iif(VC.RUC_DNI='RUC',6,1), VC.Numero, VC.DescCliente, V.MontoSubtotal," _
& " V.MontoIgv , V.MontoTotal, V.FecMov,VS.DescServ,CDate(Mid(V.FecDoc,7,2) + '/' + Mid(V.FecDoc,5,2) + '/' + Mid(V.FecDoc,3,2))" _
& " FROM Ventas_Servicios as VS INNER JOIN (Ventas_Det as  VD INNER JOIN(Ventas AS V INNER JOIN Ventas_CLientes AS VC ON V.IdCLiente=VC.IdCliente)" _
& " ON V.Orden=VD.Orden) ON VS.IdServ=VD.CodConcepto" _
& " WHERE V.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And V.Anulado='NO'" _
& " ORDER BY V.FecMov, V.Orden"
      
      
' Ejecuta la sentencia
curRegVentas.SQL = sSQL
If curRegVentas.Abrir = HAY_ERROR Then End
'Inicializa la  variable
dblIGV = 0
dblSubTotal = 0
dblMontoTotal = 0
'Verifica que no hay registros en la consulta
If curRegVentas.EOF Then
    'Mensaje no hay exsitencias en almacén
    MsgBox "No hay registro de Ventas entre estas fechas", vbInformation + vbOKOnly, "Ventas - Consulta de Registro de Ventas"
    
    ' cierra la consulta
    curRegVentas.Cerrar
    
    'Limpiar Grid
    grdConsulta.Rows = 1
    
    'Termina la ejecución del procedimiento
    Exit Sub
Else
     
     'Asigna el Orden del primer registro
     'Orden,Periodo, Fecha de Comprobante, TipoDoc ,Serie,NumeroDoc,TipoDoc,NumRucDni ,DescCliente, Monto SubTotal,
     'MontoIgv , MontoTotal,  FecMov
     sPeriodo = curRegVentas.campo(1)
     sFechaComprobante = curRegVentas.campo(2)
     sTipoDoc = curRegVentas.campo(3)
     sSerie = curRegVentas.campo(4)
     sNumDoc = curRegVentas.campo(5)
     sTipoRUC_DNI = curRegVentas.campo(6)
     sNumero = curRegVentas.campo(7)
     sDescripcion = curRegVentas.campo(8)
     dblSubTotal = curRegVentas.campo(9)
     dblMontoTotal = curRegVentas.campo(11)
     sOrden = curRegVentas.campo(0)
     sFecMov = curRegVentas.campo(12)
     sDescVenta = curRegVentas.campo(13)
     sFecDoc = curRegVentas.campo(14)
     'Determina si campo es factura o boleta
      If (curRegVentas.campo(3) = "1") Then
          sDocumento = "FACTURA"
      Else
          sDocumento = "BOLETA VENTA"
      End If
      'Determina si campo es RUC o DNI
      If (curRegVentas.campo(6) = "6") Then
          sRUCDNI = "RUC"
      Else
          sRUCDNI = "DNI"
      End If
     'Verifica si es nulo el curRegVentas.campo(7)
     If IsNull(curRegVentas.campo(10)) Then
        'Es nulo el curRegVentas
        dblIGV = 0
     Else
        'No es nulo
        dblIGV = curRegVentas.campo(10)
     End If
     
    ' Recorre el cursor ingresos a cta
    Do While Not curRegVentas.EOF
        
        'Mueve al siguiente registro
        curRegVentas.MoverSiguiente
                
        ' Verifica que no sea el final del cursor
            If Not curRegVentas.EOF Then
                'Asigna el cursor al sIdPersona
                sOrdenSig = curRegVentas.campo(0)
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
                
                'Asigna el Orden del primer registro
                'Periodo, Fecha de Comprobante, TipoDoc ,Serie,NumeroDoc,TipoDoc,NumRucDni ,DescCliente, Monto SubTotal,
                'MontoIgv , MontoTotal, Orden, FecMov
     
     
                            grdConsulta.AddItem sOrden & vbTab & _
                                    sPeriodo & vbTab & _
                                    sFechaComprobante & vbTab & _
                                    sFecDoc & vbTab & _
                                    sTipoDoc & vbTab & _
                                    sDocumento & vbTab & _
                                    sSerie & vbTab & _
                                    sNumDoc & vbTab & _
                                    sTipoRUC_DNI & vbTab & _
                                    sRUCDNI & vbTab & _
                                    sNumero & vbTab & _
                                    sDescripcion & vbTab & _
                                    sDescVenta & vbTab & _
                                    Format(dblSubTotal, "###,###,##0.00") & vbTab & _
                                    Format(dblIGV, "###,###,##0.00") & vbTab & _
                                    Format(dblMontoTotal, "###,###,##0.00") & vbTab & _
                                    sFecMov


        
                                                      
                'Inicializa los totales, el nombre e identificador de la nueva persona
                If Not curRegVentas.EOF Then
                    sOrden = sOrdenSig
                    'Asigna el Orden del primer registro
                    'Fecha, Comprobante, TipoDoc ,NumRuc, IdProveedor, _
                    'Proveedor, Descripción, Impuesto, MontoTotal, Orden
                    sPeriodo = curRegVentas.campo(1)
                    sFechaComprobante = curRegVentas.campo(2)
                    sTipoDoc = curRegVentas.campo(3)
                    sSerie = curRegVentas.campo(4)
                    sNumDoc = curRegVentas.campo(5)
                    sTipoRUC_DNI = curRegVentas.campo(6)
                    sNumero = curRegVentas.campo(7)
                    sDescripcion = curRegVentas.campo(8)
                    dblSubTotal = curRegVentas.campo(9)
                    dblMontoTotal = curRegVentas.campo(11)
                    sOrden = curRegVentas.campo(0)
                    sFecMov = curRegVentas.campo(12)
                    sDescVenta = curRegVentas.campo(13)
                    sFecDoc = curRegVentas.campo(14)
                    'Determina si campo es factura o boleta
                    If (curRegVentas.campo(3) = "1") Then
                        sDocumento = "FACTURA"
                    Else
                        sDocumento = "BOLETA VENTA"
                    End If
                    'Determina si campo es RUC o DNI
                    If (curRegVentas.campo(6) = "6") Then
                        sRUCDNI = "RUC"
                    Else
                        sRUCDNI = "DNI"
                    End If
                    
                    'Verifica si es nulo el cursor
                    If IsNull(curRegVentas.campo(10)) Then
                        dblIGV = 0
                    Else
                        'Copia el valor del cursor
                        dblIGV = curRegVentas.campo(10)
                    End If
                End If
            Else
                'Suma los impuestos
                If IsNull(curRegVentas.campo(10)) Then
                    dblIGV = 0
                Else
                    'Acumula los impuestos
                    dblIGV = dblIGV + curRegVentas.campo(10)
                End If
            End If
    Loop
    
End If

'Cierra el cursor
curRegVentas.Cerrar

'Habilita el cmdInforme
cmdInforme.Enabled = True

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
