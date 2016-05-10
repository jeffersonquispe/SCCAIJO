VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAlConsulSalidaProd 
   Caption         =   "Consulta de salidas de almacén por producto"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   HelpContextID   =   96
   Icon            =   "SCAlConsulSalidaProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   4080
      Top             =   7440
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
      Left            =   6690
      Picture         =   "SCAlConsulSalidaProd.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   220
   End
   Begin VB.ComboBox cboProducto 
      Height          =   315
      Left            =   1860
      Style           =   1  'Simple Combo
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   5085
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Top             =   7650
      Width           =   1095
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   7650
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      Height          =   735
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   4140
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
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtUnidad 
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   15
         Top             =   760
         Width           =   1815
      End
      Begin VB.TextBox txtProducto 
         Height          =   315
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
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
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   11415
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6015
      Left            =   405
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10610
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FillStyle       =   1
      MergeCells      =   1
   End
   Begin VB.Frame Frame4 
      Height          =   6375
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   11295
   End
End
Attribute VB_Name = "frmAlConsulSalidaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcolidprod As New Collection
Dim mcolCodDesProd As New Collection
Dim mcolDesMedidaProd As New Collection

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
Dim rptALSalidasProd As New clsBD4

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
  LlenarTablaRPTALSALIDASPROD
  
' Formulario
  Set rptALSalidasProd.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptALSalidasProd.AsignarRpt

' Formula/s de Crystal.
  rptALSalidasProd.Formulas.Add "Fecha='DEL " & mskFechaIni & " AL " & mskFechaFin.Text & "'"
  rptALSalidasProd.Formulas.Add "IdProd='" & txtProducto.Text & "'"
  rptALSalidasProd.Formulas.Add "DescProd='" & cboProducto.Text & "'"
  rptALSalidasProd.Formulas.Add "Unidad='" & txtUnidad & "'"
  
' Clausula WHERE de las relaciones del rpt.
  rptALSalidasProd.FiltroSelectionFormula = ""

' Nombre del fichero
  rptALSalidasProd.NombreRPT = "RPTALSALIDASPROD.rpt"

' Presentación preliminar del Informe
  rptALSalidasProd.PresentancionPreliminar

'Sentencia SQL
 sSQL = "DELETE * FROM RPTALSALIDASPROD"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True

End Sub

Private Sub LlenarTablaRPTALSALIDASPROD()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modAlSalidasProd As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1

'N° SALIDA,FEC.SALIDA,NOMBRES Y APELLIDOS,SALIDA CANT,PRECIO SALIDA,ORDEN INGRESO
     sSQL = "INSERT INTO RPTALSALIDASPROD VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 2)) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 3))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 4))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 5) & "', " & i & ")"
    
    'Copia la sentencia sSQL
    modAlSalidasProd.SQL = sSQL
    
    'Verifica si hay error
    If modAlSalidasProd.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modAlSalidasProd.Cerrar

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
'N° SALIDA,FEC.SALIDA,NOMBRES Y APELLIDOS,SALIDA CANT,PRECIO SALIDA,ORDEN INGRESO
aTitulosColGrid = Array("N° SALIDA", "FEC.SALIDA", "NOMBRES Y APELLIDOS", _
                        "SALIDA CANT.", "PRECIO SALIDA", "ORDEN INGRESO")
aTamañosColumnas = Array(1300, 1000, 3500, 1400, 1300, 1500)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Deshabilita el control cmdInforme
cmdInforme.Enabled = False

'Los campos coloca a color amarillo
EstableceCamposObligatorios

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
  CargaSalidaAlmacenProd
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
  CargaSalidaAlmacenProd
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub CargaSalidaAlmacenProd()
' ----------------------------------------------------
' Propósito : Carga las salidas de Almacén por producto entre las
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
  
'Vacia el grdConsulta
grdConsulta.Rows = 1
'Carga el cursor con las salidas
DeterminarSalidas

'Se agrupan por egreso de las ctas corrientes en dólares
grdConsulta.MergeCells = flexMergeRestrictAll
grdConsulta.MergeCol(0) = True
grdConsulta.MergeCol(1) = True

End Sub

Private Sub DeterminarSalidas()
' ----------------------------------------------------
' Propósito : Determina las salidas y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curSalidasAlmacen As New clsBD2

'Sentencia SQL
sSQL = "SELECT  ( PP.Apellidos & ', ' & PP.Nombre),AD.IdSalida, AD.Cantidad, " & _
       "AD.Precio, AD.Orden, SA.Fecha " & _
       "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD, PLN_PERSONAL PP " & _
       "WHERE AD.IdProd='" & txtProducto.Text & "' And  SA.Fecha BETWEEN " & _
       "'" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " & _
       "SA.IdSalida=AD.IdSalida And SA.IdPersona=PP.IdPersona And SA.Anulado='NO' " & _
       "ORDER BY AD.IdSalida, AD.Orden"
    
'Ejecuta la sentencia SQL
curSalidasAlmacen.SQL = sSQL

'Verifica si hay error al ejecutar la sentencia
If curSalidasAlmacen.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si es fin de registro
If curSalidasAlmacen.EOF Then
    'Mensaje
    MsgBox "No hay salidas de almacen de este producto", vbInformation + vbOKOnly, "Consulta Salida de Almacén"
    curSalidasAlmacen.Cerrar
    
    'Habilita el cmdInforme
    cmdInforme.Enabled = False

    Exit Sub
Else
    'Habilita el cmdInforme
    cmdInforme.Enabled = True
    
    'Agrega al grid los datos
    'Hacer mientras no sea fin del registro
    Do While Not curSalidasAlmacen.EOF
        'Agrega al grid los datos
        grdConsulta.AddItem curSalidasAlmacen.campo(1) & vbTab & _
                            FechaDMA(curSalidasAlmacen.campo(5)) & vbTab & _
                            curSalidasAlmacen.campo(0) & vbTab & _
                            Format(curSalidasAlmacen.campo(2), "###,###,###,##0.00") & vbTab & _
                            Format(curSalidasAlmacen.campo(3), "###,###,###,##0.00") & vbTab & _
                            curSalidasAlmacen.campo(4)
                           

        ' Mueve al siguiente concpto de ingreso
        curSalidasAlmacen.MoverSiguiente
    Loop
    
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
       grdConsulta.Rows = 1
       ' Carga los salidas de almacén
       CargaSalidaAlmacenProd
       
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
