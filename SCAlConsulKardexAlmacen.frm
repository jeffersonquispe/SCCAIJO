VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAlConsulKardexAlmacen 
   Caption         =   "Consulta de kardex de almacén                            -----  PROYECTOS  -----"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   95
   Icon            =   "SCAlConsulKardexAlmacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   1080
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPProducto 
      Height          =   255
      Left            =   6810
      Picture         =   "SCAlConsulKardexAlmacen.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   220
   End
   Begin VB.ComboBox cboProducto 
      Height          =   315
      Left            =   2130
      Style           =   1  'Simple Combo
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      Height          =   735
      Left            =   7440
      TabIndex        =   9
      Top             =   240
      Width           =   4020
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   2475
         TabIndex        =   11
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consulta del:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Producto"
      Height          =   1095
      Left            =   280
      TabIndex        =   7
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtUnidad 
         Height          =   315
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtProducto 
         Height          =   315
         Left            =   1005
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   11535
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   2
      FillStyle       =   1
      MergeCells      =   3
   End
End
Attribute VB_Name = "frmAlConsulKardexAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcolidprod As New Collection
Dim mcolCodDesProd As New Collection
Dim mcolDesMedidaProd As New Collection
Dim mdblMontoIngresos As Double
Dim mdblCantidadIngresos As Double
Dim mdblMontoSalidas As Double
Dim mdblCantidadSalidas As Double
Dim mcurIngresos As New clsBD2
Dim mcurSalidas As New clsBD2
Dim mblnIngresos, mblnSalidas As Boolean

Private Sub cboProducto_Change()

' verifica si lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProducto) = True Then SendKeys "{down}"

End Sub

Private Sub cboProducto_Click()

' Verifica si el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProducto.ListIndex) = False Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboProducto_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica si es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboProducto_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProducto, vbWhite) = True Then
   ' Se actualiza código (TextBox) correspondiente a descripción introducida
   CD_ActCod cboProducto.Text, txtProducto, mcolidprod, mcolCodDesProd
Else
   'Coloca a obligatorio el txt
   txtProducto.Text = Empty
End If

'Cambia el alto del combo
cboProducto.Height = CBONORMAL

End Sub


Private Sub cmdInforme_Click()
Dim sSQL As String
Dim rptKardexAlmacen As New clsBD4

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
  LlenarTablaRPTALKARDEX
 
' Formulario
  Set rptKardexAlmacen.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptKardexAlmacen.AsignarRpt

' Formula/s de Crystal.
  rptKardexAlmacen.Formulas.Add "Fecha='" & mskFechaIni & " AL " & mskFechaFin & "'"
  rptKardexAlmacen.Formulas.Add "IdProd='" & txtProducto & "'"
  rptKardexAlmacen.Formulas.Add "DescProd='" & cboProducto.Text & "'"
  rptKardexAlmacen.Formulas.Add "Unidad='" & txtUnidad & "'"
 
' Clausula WHERE de las relaciones del rpt.
  rptKardexAlmacen.FiltroSelectionFormula = ""

' Nombre del fichero
  rptKardexAlmacen.NombreRPT = "RPTALKARDEX.rpt"

' Presentación preliminar del Informe
  rptKardexAlmacen.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTALKARDEX"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True


End Sub

Private Sub LlenarTablaRPTALKARDEX()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modAlKardex As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1

    'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
     sSQL = "INSERT INTO RPTALKARDEX VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 2))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 3))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 4) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 5)) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 6))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 7))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 8))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 9))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 10) & "', " _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modAlKardex.SQL = sSQL
    
    'Verifica si hay error
    If modAlKardex.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modAlKardex.Cerrar

Next i

End Sub


Private Sub cmdPProducto_Click()
'Cambia el alto del cboProducto
If cboProducto.Enabled Then
    ' alto
     cboProducto.Height = CBOALTO
    ' focus a cbo
    cboProducto.SetFocus
End If
End Sub


Private Sub cmdSalir_Click()

'Descarga el formulario
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

'Establece la ubicación del formulario
Me.Top = 0

'Carga la colección de producto
CargarColProducto

' Limpia el combo del producto
cboProducto.Clear

'Carga el cboProducto de acuerdo a la relación
CargarCboCols cboProducto, mcolCodDesProd

' Carga los títulos del grid
'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
aTitulosColGrid = Array("FECHA", "N° COMPROBANTE", "COMPRA CANTIDAD", "PRECIO COMPRA", "N° SALIDA", "ENTREGADO A", "SALIDA CANTIDAD", "PRECIO SALIDA", "SALDO CANTIDAD", "PRECIO SALDO", "ORDEN")
aTamañosColumnas = Array(1000, 1550, 1550, 1550, 1100, 3500, 1550, 1500, 1500, 1500, 1200)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Deshabilita el control cmdInforme
cmdInforme.Enabled = False

'Los campos coloca a color amarillo
EstableceCamposObligatorios

End Sub

Private Sub CargarColProducto()
'---------------------------------------------------------------
'Propósito : Carga la colección de Productos con su medida
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curMedidaProd As New clsBD2

'Sentencia para cargar del combo producto
sSQL = "SELECT P.Idprod,P.DescProd,P.Medida " _
        & " FROM PRODUCTOS P " _
        & " WHERE Tipo='PROY' " _
        & " ORDER BY DescProd"

'Carga la colección de descripcion y medida de los productos
curMedidaProd.SQL = sSQL
If curMedidaProd.Abrir = HAY_ERROR Then
  End
End If
Do While Not curMedidaProd.EOF
    ' Se carga la colección de descripciones + unidades de los productos con la 1º y 2º
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolDesMedidaProd.Add Key:=curMedidaProd.campo(0), _
                              Item:=curMedidaProd.campo(2)
    
    'colección de producto y su descripción
    mcolidprod.Add curMedidaProd.campo(0)
    mcolCodDesProd.Add curMedidaProd.campo(1), curMedidaProd.campo(0)

    ' Se avanza a la siguiente fila del cursor
    curMedidaProd.MoverSiguiente
Loop
'Cierra el cursor de medida de productos
curMedidaProd.Cerrar

End Sub


Private Sub EstableceCamposObligatorios()
' ------------------------------------------------------------
' Propósito : Muestra de color amarillo los campos obligatorios
' Recibe    : Nada
' Entrega   :Nada
' ------------------------------------------------------------
mskFechaIni.BackColor = Obligatorio
mskFechaFin.BackColor = Obligatorio
txtProducto.BackColor = Obligatorio
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Vacia la colección
Set mcolidprod = Nothing
Set mcolCodDesProd = Nothing
Set mcolDesMedidaProd = Nothing

End Sub

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  
  ' Carga los Kardex de almacén
  CargaKardexAlmacen
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito : Verifica si esta bien los datos para ejecutar _
              la consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio And txtProducto.BackColor <> Obligatorio Then
' Verifica que la fecha de inicio sea Menor a la fecha final
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = True Then
        fbOkDatosIntroducidos = False
        Exit Function
    End If
End If
' Verifica si los datos obligatorios se ha llenado
If mskFechaIni.BackColor <> vbWhite Or _
   mskFechaFin.BackColor <> vbWhite Or _
   txtProducto.BackColor <> vbWhite Then
   fbOkDatosIntroducidos = False
   Exit Function
End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub mskFechaIni_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
  mskFechaIni.BackColor = vbWhite
  
  ' Carga los Kardex de almacén
  CargaKardexAlmacen
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub CargaKardexAlmacen()
' ----------------------------------------------------
' Propósito : Carga los kardex de Almacén entre las
'             fecha de consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------

' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
    grdConsulta.Rows = 1
    'Deshabilita el cmdInforme
    cmdInforme.Enabled = False
    Exit Sub
  End If

'Limpia el grdConsulta
grdConsulta.Rows = 1

'Carga los ingresos de almacén
IngresosAlmacenAntesFechaIni

'Carga las salidas de almacén
SalidasAlmacenAntesFechaIni

'Carga los saldos al grid
CargaSaldos

'Carga los ingresos y egresos entre estas fechas
CargaIngresosSalidas

End Sub

Private Sub CargaIngresosSalidas()
' ----------------------------------------------------
' Propósito : Carga los ingresos, egresos y saldos al grid
'             entre las fechas de consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim blnCargaIng, blnCargaSalida As Boolean
Dim blnRecorreCursores As Boolean
Dim dblCantSaldo, dblMontoSaldo As Double

'Carga el cursor con los ingresos
DeterminarIngresos

'Carga el cursor con las salidas
DeterminarSalidas

'Verifica si hay ingresos y salidas
If mblnIngresos = False And mblnSalidas = False Then
    'No hay salidas e ingreso de almacen
    cmdInforme.Enabled = False
    
    'Cierra los cursores
    mcurIngresos.Cerrar
    mcurSalidas.Cerrar
    
    'Termina la ejecución
    Exit Sub
Else
    'Hay ingreso o salida de almacen
    cmdInforme.Enabled = True
End If
'Agrega los datos al grid
'Verifica si es que no sea fin de registro
blnRecorreCursores = True
Do While blnRecorreCursores = True

   ' Verifica si se ha terminado de recorrer todos los cursores
   If mcurIngresos.EOF And mcurSalidas.EOF Then
       ' Sale de recorrer cursor
       blnRecorreCursores = False
       
   Else
        'Verifica que ninguno de los cursores sea el final del registro
        If Not mcurIngresos.EOF And Not mcurSalidas.EOF Then
        
            'Verifica si el ingreso es antes de la salida
           If mcurIngresos.campo(0) < mcurSalidas.campo(0) Then
                blnCargaIng = True
                blnCargaSalida = False
                
            'Verifica si la salida es antes del ingreso
            ElseIf mcurIngresos.campo(0) > mcurSalidas.campo(0) Then
                blnCargaIng = False
                blnCargaSalida = True
                
            'Ingresos y egresos tienen la misma fecha
            ElseIf mcurIngresos.campo(0) = mcurSalidas.campo(0) Then
                blnCargaIng = True
                blnCargaSalida = False
            End If
        
        'El mcurIngresos no es fin del registro
        ElseIf Not mcurIngresos.EOF Then
            blnCargaIng = True
            blnCargaSalida = False
            
        'El mcurSalidas no es fin del registro
        ElseIf Not mcurSalidas.EOF Then
            blnCargaSalida = True
            blnCargaIng = False
        End If
             
        ' añade una fila al grid
        If blnCargaIng Then
            
            'Determina el saldo y PrecioSaldo
            dblCantSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 8))) + Val(mcurIngresos.campo(3))
            dblMontoSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 9))) + Val(mcurIngresos.campo(4))
            
            ' Coloca al grid el concepto de ingreso
            grdConsulta.AddItem FechaDMA(mcurIngresos.campo(0)) & vbTab & _
                                fsAsignarDoc & vbTab & _
                    Format(mcurIngresos.campo(3), "###,###,###,##0.00") & vbTab & _
                    Format(mcurIngresos.campo(4), "###,###,###,##0.00") & vbTab & _
                                vbTab & vbTab & vbTab & vbTab & _
                    Format(dblCantSaldo, "###,###,###,##0.00") & vbTab & _
                    Format(dblMontoSaldo, "###,###,###,##0.00") & vbTab & mcurIngresos.campo(5)
            
            'Coloca el solor a los ingresos
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, &H80000003

            ' Mueve al siguiente concepto de ingreso
            mcurIngresos.MoverSiguiente
        End If
        
        'Agrega fila al grid
        If blnCargaSalida Then
        
            'Determina el saldo y PrecioSaldo
            dblCantSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 8))) - Val(mcurSalidas.campo(3))
            dblMontoSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 9))) - Val(mcurSalidas.campo(4))
            
            ' Coloca al grid el concepto de ingreso
            grdConsulta.AddItem FechaDMA(mcurSalidas.campo(0)) & vbTab _
                                & vbTab & vbTab & vbTab & _
                                mcurSalidas.campo(1) & vbTab & _
                                mcurSalidas.campo(2) & vbTab & _
                    Format(mcurSalidas.campo(3), "###,###,###,##0.00") & vbTab & _
                    Format(mcurSalidas.campo(4), "###,###,###,##0.00") & vbTab & _
                    Format(dblCantSaldo, "###,###,###,##0.00") & vbTab & _
                    Format(dblMontoSaldo, "###,###,###,##0.00") & vbTab & _
                                mcurSalidas.campo(5)
                                
            'Coloca el color del egreso
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbBlack, &H80000005
            
            ' Mueve al siguiente concpto de ingreso
            mcurSalidas.MoverSiguiente
        End If
    End If
Loop 'Fin de hacer mientras sea fin de cursor

'Cierra los cursores
mcurIngresos.Cerrar
mcurSalidas.Cerrar

End Sub

Private Function fsAsignarDoc() As String
' ----------------------------------------------------
' Propósito : Concatena los documento de ingreso y ingreso por Balance
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
'Inicializa la variable
fsAsignarDoc = Empty
'Verifica si los cursores son vacios
If Not IsNull(mcurIngresos.campo(1)) Then
    fsAsignarDoc = mcurIngresos.campo(1)
End If
If Not IsNull(mcurIngresos.campo(2)) Then
    fsAsignarDoc = mcurIngresos.campo(2)
End If

End Function

Private Sub DeterminarIngresos()
' ----------------------------------------------------
' Propósito : Determina los ingresos y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = "SELECT A.Fecha,B.NumDoc,E.NumDoc, G.Cantidad, G.Monto, A.Orden " & _
     "FROM (((ALMACEN_INGRESOS A left outer join EGRESOS E on A.Orden=E.Orden ) left outer join ALMACEN_BALANCE B on A.Orden=B.IdBalance) left outer join GASTOS G on A.Orden=G.Orden ) " & _
     "WHERE  A.IdProd='" & txtProducto.Text & "' And A.IdProd=G.CodConcepto And " & _
     "A.Fecha BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "'  " & _
     "ORDER BY A.Fecha, A.NroIngreso "
      

'Ejecuta la sentencia SQL
mcurIngresos.SQL = sSQL

'Verifica si hay error al ejecutar la sentencia
If mcurIngresos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si el cursor es nulo
If mcurIngresos.EOF Then
    'No hay ingresos
    mblnIngresos = False
Else
    'Hay ingresos
    mblnIngresos = True
End If

End Sub

Private Sub DeterminarSalidas()
' ----------------------------------------------------
' Propósito : Determina las salidas y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT SA.Fecha, AD.IdSalida, ( PP.Apellidos & ', ' & PP.Nombre), SUM(AD.Cantidad), " & _
       "SUM(AD.Precio),AD.Orden,I.NroIngreso " & _
       "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD, PLN_PERSONAL PP, ALMACEN_INGRESOS I " & _
       "WHERE AD.IdProd='" & txtProducto.Text & "' And  SA.Fecha BETWEEN " & _
       "'" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " & _
       "SA.IdSalida=AD.IdSalida And SA.IdPersona=PP.IdPersona And SA.Anulado='NO' " & _
       "And AD.Orden=I.Orden And AD.IdProd=I.IdProd " & _
       "GROUP BY SA.Fecha, AD.IdSalida, PP.IdPersona, ( PP.Apellidos & ', ' & PP.Nombre),AD.Orden, I.NroIngreso " & _
       "ORDER BY SA.Fecha, AD.IdSalida, I.NroIngreso "
    
'Ejecuta la sentencia SQL
mcurSalidas.SQL = sSQL

'Verifica si hay error al ejecutar la sentencia
If mcurSalidas.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si el cursor es nulo
If mcurSalidas.EOF Then
    'No hay Salidas
    mblnSalidas = False
Else
    'Hay Salidas
    mblnSalidas = True
End If

End Sub

Private Sub CargaSaldos()
' ----------------------------------------------------
' Propósito : Determina los saldos y carga al grid antes
'             de la fecha de Inicio
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim dblSaldoCantidad As Double
Dim dblSaldoMonto As Double

' Calcula el saldo
 dblSaldoCantidad = mdblCantidadIngresos - mdblCantidadSalidas
 dblSaldoMonto = mdblMontoIngresos - mdblMontoSalidas
 
' Carga al grid el saldo
' Fecha, Comprobante, CompraCantidad,PrecioCompra, NumSalida, SalidadCantida
' Entregado a, PrecioSalida, SaldoUnidades, PrecioSaldo
grdConsulta.AddItem "" & vbTab & "Saldo Anterior" & vbTab & vbTab & _
                    vbTab & vbTab & vbTab & vbTab & vbTab & _
                    Format(dblSaldoCantidad, "###,###,###,##0.00") & vbTab & _
                    Format(dblSaldoMonto, "###,###,###,##0.00") & vbTab

'Coloca el color a la fila
grdConsulta.Row = grdConsulta.Rows - 1
MarcarFilaGRID grdConsulta, &H80000012, &H80000000

End Sub

Private Sub IngresosAlmacenAntesFechaIni()
' ----------------------------------------------------
' Propósito : Determina los ingresos a Almacén ante de
'             la fecha de consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curIngresoAlmacen As New clsBD2

'Fecha, IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total
sSQL = "SELECT Sum(G.Cantidad),Sum(G.Monto) " & _
       "FROM ALMACEN_INGRESOS A, GASTOS G " & _
       "WHERE G.CodConcepto='" & txtProducto.Text & "' and A.Fecha < '" & FechaAMD(mskFechaIni) & "' And " & _
       "A.Orden=G.Orden And A.IdProd=G.CodConcepto "
              
' Ejecuta la sentencia
curIngresoAlmacen.SQL = sSQL
If curIngresoAlmacen.Abrir = HAY_ERROR Then End

'Inicializa la variable
mdblMontoIngresos = 0
mdblCantidadIngresos = 0

'Verifica que no haya ingresos
If curIngresoAlmacen.EOF Then
    ' Recorre el cursor ingresos a cta
    mdblMontoIngresos = 0
    mdblCantidadIngresos = 0
    
Else
    If IsNull(curIngresoAlmacen.campo(0)) And IsNull(curIngresoAlmacen.campo(1)) Then
        'No hay ningun dato
        mdblMontoIngresos = 0
        mdblCantidadIngresos = 0
    Else
        'Asigna los montos a las variables
        mdblMontoIngresos = Val(curIngresoAlmacen.campo(1))
        mdblCantidadIngresos = Val(curIngresoAlmacen.campo(0))
    End If
End If

End Sub

Private Sub SalidasAlmacenAntesFechaIni()
' ----------------------------------------------------
' Propósito : Determina las salidas de Almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curSalidasAlmacen As New clsBD2

' Carga la sentencia
'Fecha, IdSalida, IdProd,IdPersona ,Nombre, DescProd, Cantidad,PrecioUni,Total, CodCont, Orden
sSQL = "SELECT  SUM(AD.Cantidad) ,SUM(AD.Precio) " & _
       "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD " & _
       "WHERE AD.IdProd='" & txtProducto.Text & "' And SA.Fecha < '" & FechaAMD(mskFechaIni) & "' And " & _
       "SA.IdSalida=AD.IdSalida And SA.Anulado='NO' "
       
' Ejecuta la sentencia
curSalidasAlmacen.SQL = sSQL
If curSalidasAlmacen.Abrir = HAY_ERROR Then End

'Inicializa la variable
mdblCantidadSalidas = 0
mdblMontoSalidas = 0

'Verifica si no hay salidas
If curSalidasAlmacen.EOF Then
    mdblCantidadSalidas = 0
    mdblMontoSalidas = 0
Else
    
    If IsNull(curSalidasAlmacen.campo(0)) And IsNull(curSalidasAlmacen.campo(1)) Then
        'No hay ningun dato
        'Los Campos son nulos
        mdblCantidadSalidas = 0
        mdblMontoSalidas = 0
    Else
        'Copia el total de salidas
        mdblCantidadSalidas = Val(curSalidasAlmacen.campo(0))
        mdblMontoSalidas = Val(curSalidasAlmacen.campo(1))
      
    End If
End If

'Cierra el cursor
curSalidasAlmacen.Cerrar

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtProducto_Change()

'Verifica si es munuscula el texto ingresado
If UCase(txtProducto.Text) = txtProducto.Text Then

    ' Si procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboProducto, txtProducto, mcolCodDesProd
     
     ' Verifica si el campo esta vacio
    If txtProducto.Text <> "" And cboProducto.Text <> "" Then
       ' Los campos coloca a color blanco
       txtProducto.BackColor = vbWhite
       ' Muestra la unidad
       txtUnidad = Var30(mcolDesMedidaProd.Item(Trim(txtProducto)), 1)
       ' Carga los Kardex de almacén
       CargaKardexAlmacen
    Else
       'Los campos coloca a color amarillo
       txtUnidad = Empty
       txtProducto.BackColor = Obligatorio
       cmdInforme.Enabled = False
       grdConsulta.Rows = 1
    End If

Else
    If Len(txtProducto.Text) = txtProducto.MaxLength Then
        'comvertimos a mayuscula
        txtProducto.Text = UCase(txtProducto.Text)
    End If
End If

End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtUnidad_GotFocus()

' Cuando ingresa al control unidad sale inmediatamente
SendKeys vbTab

End Sub
