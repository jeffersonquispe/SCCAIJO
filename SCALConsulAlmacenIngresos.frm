VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALConsulAlmacenIngresos 
   Caption         =   "Consulta de ingresos a almacén"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   92
   Icon            =   "SCALConsulAlmacenIngresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Opciones de informe"
      Height          =   615
      Left            =   5520
      TabIndex        =   16
      Top             =   7800
      Width           =   3255
      Begin VB.OptionButton optSinOrden 
         Caption         =   "SIN ORDEN"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "CON ORDEN"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport rptInformes 
      Left            =   480
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
   Begin VB.TextBox txtTotalIngresos 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2820
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7965
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   8
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11630
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   7440
         TabIndex        =   13
         Top             =   180
         Width           =   3015
         Begin MSMask.MaskEdBox mskFechaConsulta 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            TabStop         =   0   'False
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
            TabIndex        =   14
            Top             =   285
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         Height          =   735
         Left            =   480
         TabIndex        =   11
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
            TabIndex        =   15
            Top             =   255
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consulta del "
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   285
            Width           =   915
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL INGRESOS:"
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   8040
      Width           =   1455
   End
End
Attribute VB_Name = "frmALConsulAlmacenIngresos"
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
Dim rptIngresoAlmacen As New clsBD4

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
  LlenarTablaRPTALINGRESO
  
' Formulario
  Set rptIngresoAlmacen.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptIngresoAlmacen.AsignarRpt

' Formula/s de Crystal.
  rptIngresoAlmacen.Formulas.Add "Fecha='DEL " & mskFechaIni.Text & " AL " & mskFechaFin.Text & "'"
  
' Clausula WHERE de las relaciones del rpt.
  rptIngresoAlmacen.FiltroSelectionFormula = ""
  
' Verifica si el informe es con orden
  If optOrden.Value Then
     ' Nombre del fichero
     rptIngresoAlmacen.NombreRPT = "RPTALMACENINGRESOCONORDEN.rpt"
  Else
     ' Nombre del fichero
     rptIngresoAlmacen.NombreRPT = "RPTALMACENINGRESOSINORDEN.rpt"
  End If

' Presentación preliminar del Informe
  rptIngresoAlmacen.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTALINGRESO"

'Borra la tabla
 Var21 sSQL
 
' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True
 
End Sub

Private Sub LlenarTablaRPTALINGRESO()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modAlIngreso As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
           
     sSQL = "INSERT INTO RPTALINGRESO VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 3)) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 4) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 5))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 6))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 7))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 8) & "', " _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modAlIngreso.SQL = sSQL
    
    'Verifica si hay error
    If modAlIngreso.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modAlIngreso.Cerrar

Next i

End Sub

Private Sub cmdSalir_Click()
'Descarga el formulario
Unload Me
End Sub

Private Sub CargaAlmacenIngresos()
' ----------------------------------------------------
' Propósito : Determina el ingreso de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curIngresoAlmacen As New clsBD2
Dim curIngresoALBalance As New clsBD2
Dim dblTotalIngresos As Double
Dim blnCargado As Boolean

' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
    txtTotalIngresos.Text = "0.00"
    grdConsulta.Rows = 1
    'Deshabilita el cmdInforme
    cmdInforme.Enabled = False
    Exit Sub
  End If

'Limpia el grdConsulta
grdConsulta.Rows = 1
'Inicializa la variable
dblTotalIngresos = 0

' Carga la sentencia
'Fecha, IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total
sSQL = "SELECT A.Fecha, T.Abreviatura,E.NumDoc,P.IdProd, P.DescProd, P.Medida, G.Cantidad, " & _
       "A.PrecioUnit, G.Monto,E.Orden, A.NroIngreso " & _
       "FROM ALMACEN_INGRESOS A, PRODUCTOS P, EGRESOS E, GASTOS G,TIPO_DOCUM T " & _
       "WHERE A.Fecha BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " & _
       "A.Orden=G.Orden And A.IdProd=G.CodConcepto And G.Orden=E.Orden And E.IdTipoDoc=T.IdTipoDoc And " & _
       "A.IdProd=P.IdProd " & _
       "ORDER BY A.Fecha, A.NroIngreso"
       
' Ejecuta la sentencia
curIngresoAlmacen.SQL = sSQL
If curIngresoAlmacen.Abrir = HAY_ERROR Then End

'Determina los ingresos a almacén por balance
' Carga la sentencia
'Fecha, IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total
sSQL = "SELECT A.Fecha, B.NumDoc ,P.IdProd, P.DescProd, P.Medida, G.Cantidad, " & _
       "A.PrecioUnit, G.Monto,A.Orden, A.NroIngreso " & _
       "FROM ALMACEN_INGRESOS A, PRODUCTOS P, GASTOS G, ALMACEN_BALANCE B " & _
       "WHERE A.Fecha BETWEEN '" & FechaAMD(mskFechaIni) & "' And  " & _
       "'" & FechaAMD(mskFechaFin) & "' And A.Orden=G.Orden And   " & _
       "A.IdProd=G.CodConcepto And G.Orden=B.IdBalance And B.Anulado='NO' " & _
       " And A.IdProd=P.IdProd " & _
       "ORDER BY A.Fecha, A.NroIngreso"
       
' Ejecuta la sentencia
curIngresoALBalance.SQL = sSQL
If curIngresoALBalance.Abrir = HAY_ERROR Then End

'Inicializa la variable
blnCargado = False

'Verifica que no hay registros en la consulta
If curIngresoAlmacen.EOF And curIngresoALBalance.EOF Then
    'Mensaje no hay exsitencias en almacén
    MsgBox "No hay ingresos en almacén entre estas fechas", , "Almacén - Consulta de Ingresos"
    ' cierra la consulta
    curIngresoAlmacen.Cerrar
    curIngresoALBalance.Cerrar
    'Limpiar Grid
    grdConsulta.Rows = 1
    'Termina la ejecución del procedimiento
    Exit Sub
Else
    ' Hacer mientras blnCargado=false
    Do While blnCargado = False
        If Not curIngresoAlmacen.EOF And Not curIngresoALBalance.EOF Then
            'Compara los numeros de ingreso
            If curIngresoAlmacen.campo(10) < curIngresoALBalance.campo(9) Then
            
                 'Añade el elemento al grid
                 'Fecha, NroDoc, IdProd, DescProd, Medida, Cantidad,PrecioUnit, Total
                 grdConsulta.AddItem FechaDMA(curIngresoAlmacen.campo(0)) & vbTab & _
                                     curIngresoAlmacen.campo(1) & "/" & _
                                     curIngresoAlmacen.campo(2) & vbTab & _
                                     curIngresoAlmacen.campo(3) & vbTab & _
                                     curIngresoAlmacen.campo(4) & vbTab & _
                                     curIngresoAlmacen.campo(5) & vbTab & _
                                     Format(curIngresoAlmacen.campo(6), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoAlmacen.campo(7), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoAlmacen.campo(8), "###,###,##0.00") & vbTab & _
                                     curIngresoAlmacen.campo(9)
                                     
                'Acumula los ingresos a la cuenta
                dblTotalIngresos = dblTotalIngresos + Val(curIngresoAlmacen.campo(8))
                 
                ' Mueve al siguiente programa
                curIngresoAlmacen.MoverSiguiente
            Else
                'Añade el elemento al grid
                 'Fecha, NroDoc, IdProd, DescProd, Medida, Cantidad,PrecioUnit, Total
                 grdConsulta.AddItem FechaDMA(curIngresoALBalance.campo(0)) & vbTab & _
                                     curIngresoALBalance.campo(1) & vbTab & _
                                     curIngresoALBalance.campo(2) & vbTab & _
                                     curIngresoALBalance.campo(3) & vbTab & _
                                     curIngresoALBalance.campo(4) & vbTab & _
                                     Format(curIngresoALBalance.campo(5), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoALBalance.campo(6), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoALBalance.campo(7), "###,###,##0.00") & vbTab & _
                                     curIngresoALBalance.campo(8)
                                     
                'Acumula los ingresos a la cuenta
                dblTotalIngresos = dblTotalIngresos + Val(curIngresoALBalance.campo(7))
                ' Mueve al siguiente programa
                curIngresoALBalance.MoverSiguiente
            End If
        ElseIf Not curIngresoAlmacen.EOF Then
            'Añade el elemento al grid
                 'Fecha, Abreviatura, NroDoc, IdProd, DescProd, Medida, Cantidad,PrecioUnit, Total
                 grdConsulta.AddItem FechaDMA(curIngresoAlmacen.campo(0)) & vbTab & _
                                     curIngresoAlmacen.campo(1) & "/" & _
                                     curIngresoAlmacen.campo(2) & vbTab & _
                                     curIngresoAlmacen.campo(3) & vbTab & _
                                     curIngresoAlmacen.campo(4) & vbTab & _
                                     curIngresoAlmacen.campo(5) & vbTab & _
                                     Format(curIngresoAlmacen.campo(6), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoAlmacen.campo(7), "###,###,##0.00") & vbTab & _
                                     Format(curIngresoAlmacen.campo(8), "###,###,##0.00") & vbTab & _
                                     curIngresoAlmacen.campo(9)
                                     
                'Acumula los ingresos a la cuenta
                dblTotalIngresos = dblTotalIngresos + Val(curIngresoAlmacen.campo(8))
                ' Mueve al siguiente programa
                curIngresoAlmacen.MoverSiguiente
                
        ElseIf Not curIngresoALBalance.EOF Then
            'Añade el elemento al grid
             'Fecha, NroDoc, IdProd, DescProd, Medida, Cantidad,PrecioUnit, Total
             grdConsulta.AddItem FechaDMA(curIngresoALBalance.campo(0)) & vbTab & _
                                 curIngresoALBalance.campo(1) & vbTab & _
                                 curIngresoALBalance.campo(2) & vbTab & _
                                 curIngresoALBalance.campo(3) & vbTab & _
                                 curIngresoALBalance.campo(4) & vbTab & _
                                 Format(curIngresoALBalance.campo(5), "###,###,##0.00") & vbTab & _
                                 Format(curIngresoALBalance.campo(6), "###,###,##0.00") & vbTab & _
                                 Format(curIngresoALBalance.campo(7), "###,###,##0.00") & vbTab & _
                                 curIngresoALBalance.campo(8)
                                 
            'Acumula los ingresos a la cuenta
            dblTotalIngresos = dblTotalIngresos + Val(curIngresoALBalance.campo(7))
            ' Mueve al siguiente programa
            curIngresoALBalance.MoverSiguiente
        End If
        
        'Verifica si los cursores estan vacios
        If curIngresoAlmacen.EOF And curIngresoALBalance.EOF Then
            'Actualiza la variable
            blnCargado = True
        End If
    Loop
    
    ' Muestra el total de los proyectos
    txtTotalIngresos.Text = Format(dblTotalIngresos, "###,###,##0.00")
    
End If

'Cierra el cursor
curIngresoAlmacen.Cerrar
curIngresoALBalance.Cerrar
'Habilita el cmdInforme
cmdInforme.Enabled = True

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
'IdIngreso, IdCta, IdProy, DescProy, Fec_ing, MontoDol, Concepto
aTitulosColGrid = Array("FECHA", "DOC.", "CODIGO", "DESCRIPCION", "UNIDAD", "CANTIDAD", "PRECIO UNI", "TOTAL", "ORDEN")
aTamañosColumnas = Array(950, 1600, 800, 3400, 800, 900, 1000, 1050, 1050)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Carga la fecha del sistema
mskFechaConsulta.Text = gsFecTrabajo

'Muestra el dato al lado derecho
grdConsulta.ColAlignment(1) = 1

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
  CargaAlmacenIngresos
    
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  txtTotalIngresos.Text = "0.00"
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

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
  ' Carga las existencias de almacén
  CargaAlmacenIngresos
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  txtTotalIngresos.Text = "0.00"
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If
End Sub


Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
