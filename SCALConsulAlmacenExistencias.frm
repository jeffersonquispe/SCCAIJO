VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALConsulAlmacenExistencias 
   Caption         =   "Consulta de existencias de almacén por productos"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   94
   Icon            =   "SCALConsulAlmacenExistencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Elegir el reporte de existencias:"
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   7080
      Width           =   4935
      Begin VB.OptionButton optGeneral 
         Caption         =   "Reporte General"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optPorCuenta 
         Caption         =   "Reporte Por Cuenta Contable"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin Crystal.CrystalReport rptInformes 
      Left            =   11400
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10920
      Top             =   240
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fecha del sistema"
      Height          =   735
      Left            =   4200
      TabIndex        =   10
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtHora 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mskFecReporte 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora de consulta:"
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha de consulta:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   285
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6255
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FillStyle       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   735
      Left            =   200
      TabIndex        =   8
      Top             =   0
      Width           =   3855
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1560
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   285
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmALConsulAlmacenExistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración de Colecciones
Dim mcolCodNroCta As New Collection
Dim mcolCodDesNroCta As New Collection

'Declaración de los cursores de modulo
Dim mcurIngresoAlmacen As New clsBD2
Dim mcurEgresoAlmacen As New clsBD2

'Variable del CodCuenta
Dim msNroCta As String

Private Sub cmdInforme_Click()
Dim sSQL As String
Dim rptExistenciaAlmacen As New clsBD4

' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
  
' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón informe
       cmdInforme.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla con datos
  LlenarTablaRPTALEXISTENCIA
  
' Formulario
  Set rptExistenciaAlmacen.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptExistenciaAlmacen.AsignarRpt

' Formula/s de Crystal.
  rptExistenciaAlmacen.Formulas.Add "Fecha='AL " & mskFecha.Text & "'"

' Clausula WHERE de las relaciones del rpt.
  rptExistenciaAlmacen.FiltroSelectionFormula = ""

' Nombre del fichero
  If optPorCuenta.Value Then
    ' El reporte es por cuenta contable
    rptExistenciaAlmacen.NombreRPT = "RPTALEXISTENCIASCUENTA.rpt"
  Else
    ' El reporte es general
    rptExistenciaAlmacen.NombreRPT = "RPTALMACENEXISTENCIAS.rpt"
  End If
  
' Presentación preliminar del Informe
  rptExistenciaAlmacen.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTALEXISTENCIA"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Deshabilita el botón generar informe
 cmdInforme.Enabled = True

End Sub

Private Sub LlenarTablaRPTALEXISTENCIA()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modAlExistencia As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
    
     'GRID IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total, CodCont
     'TABLA IdProd,DescProd, Medida, Cantidad, MontoTotal, CodCont
     sSQL = "INSERT INTO RPTALEXISTENCIA VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 1)) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 3))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 5))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 6) & "')"
    
    'Copia la sentencia sSQL
    modAlExistencia.SQL = sSQL
    
    'Verifica si hay error
    If modAlExistencia.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modAlExistencia.Cerrar

Next i

End Sub


Private Sub cmdSalir_Click()
'Descarga el formulario
Unload Me
End Sub

Private Sub CargaAlmacenExistencias()
' ----------------------------------------------------
' Propósito : Determina las existencias de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim dblTotalExistencia As Double

' Verifica los datos introducidos para la consulta
If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
    grdConsulta.Rows = 1
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
 End If

'Limpiar Grid
grdConsulta.Rows = 1

'Carga los ingresos a Almacén a las fechas de consulta
CargaIngresoAlmacen

'Carga los egresos de Almacen a las fechas de consulta
CargaEgresoAlmacen

'Carga los ingresos y egresos al grdConsulta
CargarExistenciaGrid

'Verifica si el grd tiene datos
If grdConsulta.Rows > 1 Then
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = True
End If
End Sub

Private Sub CargarExistenciaGrid()
' ----------------------------------------------------
' Propósito : Carga las existencias de almacén
'             al grdConsulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim dblTotalExistencia As Double

'Carga los datos al grd
Do While Not mcurIngresoAlmacen.EOF
    'Verifica si es fin del cursor
    If Not mcurEgresoAlmacen.EOF Then
    
        'Compara los codigo de los productos de ingreso y egreso de almacén
        If mcurIngresoAlmacen.campo(0) = mcurEgresoAlmacen.campo(0) Then
            'Compara los ingresos y egresos
            If Val(Format(mcurIngresoAlmacen.campo(4), "####0.00")) <> Val(Format(mcurEgresoAlmacen.campo(3), "####0.00")) Then
            
                ' Añade el elemento al grid
                 'IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total, CodCont
                 grdConsulta.AddItem mcurIngresoAlmacen.campo(0) & vbTab & _
                                     mcurIngresoAlmacen.campo(1) & vbTab & _
                                     mcurIngresoAlmacen.campo(2) & vbTab & _
                                     Format(Val(mcurIngresoAlmacen.campo(3)) - Val(mcurEgresoAlmacen.campo(2)), "###,###,##0.00") & vbTab & _
                                     Format((Val(mcurIngresoAlmacen.campo(4)) - Val(mcurEgresoAlmacen.campo(3))) / (Val(mcurIngresoAlmacen.campo(3)) - Val(mcurEgresoAlmacen.campo(2))), "###,###,##0.00") & vbTab & _
                                     Format(Val(mcurIngresoAlmacen.campo(4)) - Val(mcurEgresoAlmacen.campo(3)), "###,###,##0.00") & vbTab & _
                                     mcurIngresoAlmacen.campo(5)
                             
                'Acumula los ingresos a la cuenta
                dblTotalExistencia = Val(dblTotalExistencia) + (Val(mcurIngresoAlmacen.campo(4)) - Val(mcurEgresoAlmacen.campo(3)))
                
            End If
            
            ' Mueve al siguiente programa
            mcurIngresoAlmacen.MoverSiguiente
            mcurEgresoAlmacen.MoverSiguiente
            
        ElseIf mcurIngresoAlmacen.campo(0) < mcurEgresoAlmacen.campo(0) Then
            ' Añade el elemento al grid
             'IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total, CodCont
             grdConsulta.AddItem mcurIngresoAlmacen.campo(0) & vbTab & _
                                 mcurIngresoAlmacen.campo(1) & vbTab & _
                                 mcurIngresoAlmacen.campo(2) & vbTab & _
                                 Format(Val(mcurIngresoAlmacen.campo(3)), "###,###,##0.00") & vbTab & _
                                 Format(Val(mcurIngresoAlmacen.campo(4)) / Val(mcurIngresoAlmacen.campo(3)), "###,###,##0.00") & vbTab & _
                                 Format(Val(mcurIngresoAlmacen.campo(4)), "###,###,##0.00") & vbTab & _
                                 mcurIngresoAlmacen.campo(5)
                         
            'Acumula los ingresos a la cuenta
            dblTotalExistencia = Val(dblTotalExistencia) + Val(mcurIngresoAlmacen.campo(4))
             
            ' Mueve al siguiente programa
            mcurIngresoAlmacen.MoverSiguiente
        End If
    Else
        ' Añade el elemento al grid
        'IdProd,P.DescProd, P.Medida, Cantidad, PrecioU, Total, CodCont
        grdConsulta.AddItem mcurIngresoAlmacen.campo(0) & vbTab & _
                            mcurIngresoAlmacen.campo(1) & vbTab & _
                            mcurIngresoAlmacen.campo(2) & vbTab & _
                            Format(Val(mcurIngresoAlmacen.campo(3)), "###,###,##0.00") & vbTab & _
                            Format(Val(mcurIngresoAlmacen.campo(4)) / Val(mcurIngresoAlmacen.campo(3)), "###,###,##0.00") & vbTab & _
                            Format(Val(mcurIngresoAlmacen.campo(4)), "###,###,##0.00") & vbTab & _
                            mcurIngresoAlmacen.campo(5)
                     
        'Acumula los ingresos a la cuenta
        dblTotalExistencia = dblTotalExistencia + Val(mcurIngresoAlmacen.campo(4))
         
        ' Mueve al siguiente programa
        mcurIngresoAlmacen.MoverSiguiente
    End If
            
Loop
 
 ' Muestra el total de los proyectos
 grdConsulta.AddItem vbTab & "TOTAL EXISTENCIA VALORIZADO : " _
                   & vbTab & vbTab & vbTab & vbTab & _
                   Format(dblTotalExistencia, "###,###,##0.00")
                   
'Colorea el grid
grdConsulta.Row = grdConsulta.Rows - 1
MarcarFilaGRID grdConsulta, vbBlack, vbGray
   
End Sub

Private Sub CargaIngresoAlmacen()
' ----------------------------------------------------
' Propósito : Carga los ingresos a almacén a la fecha de
'             ingreso a almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia de Ingreso a Almacén
'IdProd,P.DescProd, P.Medida, Suma de Cantidad,Suma del Total
sSQL = "SELECT P.IdProd, P.DescProd, P.Medida, Sum(G.Cantidad), " & _
       "Sum(G.Monto), P.CodCont " & _
       "FROM ALMACEN_INGRESOS A, PRODUCTOS P, GASTOS G " & _
       "WHERE A.Fecha <= '" & FechaAMD(mskFecha) & "' And " & _
       "A.Orden=G.Orden And A.IdProd=G.CodConcepto And " & _
       "A.IdProd=P.IdProd " & _
       "GROUP BY P.IdProd, P.DescProd, P.Medida, P.CodCont " & _
       "ORDER BY P.IdProd "
       
' Ejecuta la sentencia
mcurIngresoAlmacen.SQL = sSQL

'Verifica si hay error el ingreso a almacén
If mcurIngresoAlmacen.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaEgresoAlmacen()
' ----------------------------------------------------
' Propósito : Carga los Egresos de almacén a la fecha de
'             Egreso de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
'IdProd, DescProd, Suma de Cantidad,Suma de Total, CodCont
sSQL = "SELECT AD.IdProd, P.DescProd, SUM(AD.Cantidad), " & _
       "SUM(AD.Precio),P.CodCont " & _
       "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD,PRODUCTOS P " & _
       "WHERE SA.Fecha <= '" & FechaAMD(mskFecha) & "' And " & _
       "SA.IdSalida=AD.IdSalida And AD.IdProd=P.IdProd And SA.Anulado='NO' " & _
       "GROUP BY AD.IdProd, P.DescProd, P.CodCont " & _
       "ORDER BY AD.IdProd "
       
' Ejecuta la sentencia
mcurEgresoAlmacen.SQL = sSQL

'Verifica si hay error el ingreso a almacén
If mcurEgresoAlmacen.Abrir = HAY_ERROR Then End

End Sub

Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito: Verifica si esta bien los datos para ejecutar _
            la consulta
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
If mskFecha.BackColor <> vbWhite Then
    ' La fecha no es correcta
    fbOkDatosIntroducidos = False
    Exit Function
End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function

Private Sub Form_Load()
Dim sSQL As String

'Establece la posicion del formulario
Me.Top = 0

' Carga los títulos del grid
'IdIngreso, IdCta, IdProy, DescProy, Fec_ing, MontoDol, Concepto
aTitulosColGrid = Array("CODIGO", "NOMBRE DEL PRODUCTO", "UNIDAD", "CANTIDAD", "PRECIO UNIT.", "TOTAL", "COD_CONT")
aTamañosColumnas = Array(1000, 3500, 1300, 1000, 1300, 1300, 1300)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Carga la fecha del sistema
mskFecReporte.Text = gsFecTrabajo

'Coloca a Obligatorio
mskFecha.BackColor = Obligatorio

'Deshabilita el cmdInforme
cmdInforme.Enabled = False

End Sub

Private Sub mskFecha_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFecha) Then
  ' Coloca el color a la fecha
  mskFecha.BackColor = vbWhite
  
  ' Carga las existencias de almacén
  CargaAlmacenExistencias

Else
  ' limpia el formulario
  mskFecha.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub Timer1_Timer()

'Muesta la hora del sistema
txtHora.Text = Format(Time, "hh:mm:ss AMPM")

End Sub
