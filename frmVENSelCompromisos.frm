VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVENSelCompromisos 
   Caption         =   "Ventas - Selección de Compromisos de Venta"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   HelpContextID   =   83
   Icon            =   "frmVENSelCompromisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDocVerificar 
      Caption         =   "Seleccione los documentos"
      Height          =   3600
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11910
      Begin MSFlexGridLib.MSFlexGrid grdIngreso 
         Height          =   3255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         HighLight       =   0
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10920
      TabIndex        =   2
      Top             =   3765
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   9720
      TabIndex        =   1
      Top             =   3765
      Width           =   1000
   End
End
Attribute VB_Name = "frmVENSelCompromisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ipos As Long

Private Sub cmdAceptar_Click()
  'Se comprueba que se haya marcado algún salida de Almacen
  If grdIngreso.Row < 1 Then
    MsgBox "Debe seleccionar algún documento", vbInformation + vbOKOnly, "SGCcaijo - Compromisos"
    Exit Sub
  End If
  
  'Guarda al formulario en la variable
  gsFormulario = "044"
  
  'Carga en la BD el formulario en uso y usuario
  Var42 ("044")
  
  'Verifica si es exclusivo y esta en uso
  If Var47("044") Then
      'Elimna la sesionn de la BD
      Var43 "044"
      'Termina la ejecucion del procedimiento
      Exit Sub
  End If
  
  'Muestra el formulario de ingresos a Caja o a Bancos
  gsTipoOperacionIngreso = "Nuevo"
  CancelarVenta = True
  frmCBIngresos.OrdenVenta = grdIngreso.TextMatrix(ipos, 0)
  frmCBIngresos.CodigoVenta = "I026"
  frmCBIngresos.CodigoTerceroVenta = grdIngreso.TextMatrix(ipos, 1)
  frmCBIngresos.VentaTotal = grdIngreso.TextMatrix(ipos, 6)
  frmCBIngresos.VentaPagada = grdIngreso.TextMatrix(ipos, 7)
  frmCBIngresos.VentaSaldo = grdIngreso.TextMatrix(ipos, 8)
  frmCBIngresos.Show vbModal, Me
  
  'Deshabilita el botón aceptar
  cmdAceptar.Enabled = False
  
  'Elimina la sesion de la BD
  Var43 "044"
  
  'Carga los Documentos no verificados en almacén, para verificarlos
  CargarCompromisos
End Sub

Private Sub cmdSalir_Click()
  'Termina la ejecucion del formulario

  Unload Me
End Sub

Private Sub Form_Load()
  'Carga los Documentos no verificados en almacén, para verificarlos
  CargarCompromisos
  ' Inicializa el grid
  ipos = 0
  gbCambioCelda = False
  grdIngreso.ColAlignment(2) = 1
  
  ' Deshabilita el aceptar
  cmdAceptar.Enabled = False
End Sub

Private Sub grdIngreso_Click()
  If grdIngreso.Row > 0 And grdIngreso.Row < grdIngreso.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdIngreso, vbWhite, vbDarkBlue
    ' Habilita aceptar
    cmdAceptar.Enabled = True
  End If
End Sub

Private Sub grdIngreso_DblClick()
  'Hace llamado al evento click del aceptar
  cmdAceptar_Click
End Sub

Private Sub grdIngreso_EnterCell()
  If ipos <> grdIngreso.Row Then
    '  Verifica si es la última fila
    If grdIngreso.Row > 0 And grdIngreso.Row < grdIngreso.Rows Then
      If gbCambioCelda = False Then
        gbCambioCelda = True
        ' Marca la fila
        MarcarSoloUnaFilaGrid grdIngreso, ipos
        gbCambioCelda = False
        cmdAceptar.Enabled = True
      End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdIngreso.Row
  End If
End Sub

Private Sub grdIngreso_KeyPress(KeyAscii As Integer)
  ' Verifica si se apretó el enter
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
End Sub

Private Sub CargarCompromisos()
  Dim sSQL As String
  Dim sIntervalo As String
  Dim mcurVentasPendientes As New clsBD2
  Dim mcurTotalPagos As New clsBD2
  
  'Limpia el grid grdSalida, inicializa la variable intervalo
  grdIngreso.Rows = 1
  sIntervalo = Empty
  'Se seleccionan los Documentos de Almacen que no hayan sido verificados
  'sSQL = "SELECT DISTINCT V.Orden, VC.IdTercero, VC.DescProveedor, VD.DescTipoDoc, V.NumDoc, V.FecMov, " & _
         "V.MontoTotal, SUM(VP.MONTOTOTAL), (V.MontoTotal - SUM(VP.MONTOTOTAL)) " & _
         "FROM VENTAS V, VENTAS_TIPO_DOCUM VD, VENTAS_CLIENTES VC, VENTAS_PAGOS VP " & _
         "WHERE V.Cancelado ='NO' and VP.ORDEN=V.ORDEN and " & _
         "V.IdTipoDoc=VD.IdTipoDoc and V.IdCliente=VC.IdProveedor " & _
         "GROUP BY V.Orden, VC.IdTercero, VC.DescProveedor, VD.DescTipoDoc, V.NumDoc, V.FecMov, V.MontoTotal " & _
         "ORDER BY V.FecMov, V.Orden "
  sSQL = ""
  sSQL = "SELECT DISTINCT V.Orden, VC.IdTerc, VC.DescCliente, VD.DescTipoDoc, V.NumDoc, V.FecMov, " & _
         "V.MontoTotal " & _
         "FROM VENTAS V, VENTAS_TIPO_DOCUM VD, VENTAS_CLIENTES VC " & _
         "WHERE V.Cancelado ='NO'and V.Anulado='NO' and " & _
         "V.IdTipoDoc=VD.IdTipoDoc and V.IdCliente=VC.IdCliente " & _
         "ORDER BY V.FecMov, V.Orden "
  
  ' Ejecuta la sentencia
  mcurVentasPendientes.SQL = sSQL
  If mcurVentasPendientes.Abrir = HAY_ERROR Then End
  
  ' Se carga un array con los títulos de las columnas y otro con los tamaños para
  'pasárselos a la función que carga el grid
  aTitulosColGrid = Array("Orden", "IdTercero", "Cliente", "Documento", "N° Documento", "Fecha Mov", "Monto Total", "Monto Pagado", "Saldo")
  aTamañosColumnas = Array(1050, 0, 4000, 1100, 1300, 900, 900, 1100, 900)
  CargarGridTitulos grdIngreso, aTitulosColGrid, aTamañosColumnas
  
  Do While Not mcurVentasPendientes.EOF
    sSQL = ""
    sSQL = "SELECT VP.Orden, SUM(VP.MONTOTOTAL) " & _
         "FROM VENTAS_PAGOS VP " & _
         "WHERE VP.ORDEN = '" & mcurVentasPendientes.campo(0) & "' " & _
         "GROUP BY VP.Orden "
             
    ' Ejecuta la sentencia
    mcurTotalPagos.SQL = sSQL
    If mcurTotalPagos.Abrir = HAY_ERROR Then End
    '"Orden", "IdTercero", "Cliente", "Documento", "N° Documento", "Fecha Mov", "Monto Total", "Monto Pagado", "Saldo")
    
    If Not mcurTotalPagos.EOF Then
      grdIngreso.AddItem mcurVentasPendientes.campo(0) & vbTab & mcurVentasPendientes.campo(1) _
                      & vbTab & mcurVentasPendientes.campo(2) & vbTab & mcurVentasPendientes.campo(3) _
                      & vbTab & mcurVentasPendientes.campo(4) & vbTab & FechaDMA(mcurVentasPendientes.campo(5)) _
                      & vbTab & Format(mcurVentasPendientes.campo(6), "###,###,##0.00") _
                      & vbTab & Format(mcurTotalPagos.campo(1), "###,###,##0.00") _
                      & vbTab & Format((Val(mcurVentasPendientes.campo(6)) - Val(mcurTotalPagos.campo(1))), "###,###,##0.00")
    Else
      grdIngreso.AddItem mcurVentasPendientes.campo(0) & vbTab & mcurVentasPendientes.campo(1) _
                      & vbTab & mcurVentasPendientes.campo(2) & vbTab & mcurVentasPendientes.campo(3) _
                      & vbTab & mcurVentasPendientes.campo(4) & vbTab & FechaDMA(mcurVentasPendientes.campo(5)) _
                      & vbTab & Format(mcurVentasPendientes.campo(6), "###,###,##0.00") _
                      & vbTab & "0.00" _
                      & vbTab & Format(mcurVentasPendientes.campo(6), "###,###,##0.00")
    End If
    
    mcurVentasPendientes.MoverSiguiente
    mcurTotalPagos.Cerrar
  Loop
  
  mcurVentasPendientes.Cerrar
  'aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Fecha", "fmt_Importe", "fmt_Importe", "fmt_Importe")
  
  'CargarGridConFormatos grdIngreso, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
  
  If grdIngreso.Rows = 1 Then ' no hay registros en la consulta
    'Mensaje de No existen registros que mostrar
    MsgBox "No existen Compromisos", _
            vbInformation + vbOKOnly, "SGCcaijo - Compromisos"
  End If
End Sub
