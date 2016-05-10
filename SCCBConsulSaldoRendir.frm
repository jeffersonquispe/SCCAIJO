VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCBConsulRendirSaldo 
   Caption         =   "Consulta de Saldos de Cuentas a Rendir del Personal"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   94
   Icon            =   "SCCBConsulSaldoRendir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6255
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      FillStyle       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   855
      Left            =   200
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      Begin VB.Frame Frame1 
         Caption         =   "Opciones"
         Height          =   615
         Left            =   3120
         TabIndex        =   9
         Top             =   120
         Width           =   4095
         Begin VB.OptionButton optAmbos 
            Caption         =   "Ambos"
            Height          =   255
            Left            =   3120
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optRendido 
            Caption         =   "Rendidos"
            Height          =   255
            Left            =   1680
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optDeudor 
            Caption         =   "Deudores"
            Height          =   255
            Left            =   360
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
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
         Caption         =   "A la Fecha:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   405
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmCBConsulRendirSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaración de los cursores de modulo
Dim mcurSaldosRendir As New clsBD2

Private Sub cmdInforme_Click()
Dim sSQL As String
Dim rptRendir As New clsBD4

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
  LlenarTabla
  
' Formulario
  Set rptRendir.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptRendir.AsignarRpt

' Formula/s de Crystal.
  rptRendir.Formulas.Add "Fecha='AL " & mskFecha.Text & "'"

' Clausula WHERE de las relaciones del rpt.
  rptRendir.FiltroSelectionFormula = ""

' Nombre del fichero
  rptRendir.NombreRPT = "RPTCBSALDOSRENDIR.rpt"
  
' Presentación preliminar del Informe
  rptRendir.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTSALDOSRENDIR"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Deshabilita el botón generar informe
 cmdInforme.Enabled = True

End Sub

Private Sub LlenarTabla()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modRendir As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
    
     'GRID codigo,apellidos,nombres,monto
     'TABLA codigo, apellidos, nombre, monto
     sSQL = "INSERT INTO RPTSALDOSRENDIR VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "','" _
     & grdConsulta.TextMatrix(i, 1) & "','" _
     & grdConsulta.TextMatrix(i, 2) & "'," _
     & Var37(grdConsulta.TextMatrix(i, 3)) & ")"
    
    'Copia la sentencia sSQL
    modRendir.SQL = sSQL
    
    'Verifica si hay error
    If modRendir.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modRendir.Cerrar

Next i

End Sub


Private Sub cmdSalir_Click()
'Descarga el formulario
Unload Me
End Sub

Private Sub CargaConsulta()
' ----------------------------------------------------
' Propósito : Determina los saldos de las cuentas a rendir
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim dblTotalSaldo As Double

' Verifica los datos introducidos para la consulta
If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
    grdConsulta.Rows = 1
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
 End If

'Carga los ingresos a Almacén a las fechas de consulta
If CargaSaldosRendir = False Then
    ' Sale de el proceso y limpia el grid
    grdConsulta.Rows = 1
    ' Cierra el cursor
    mcurSaldosRendir.Cerrar
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
End If

'Incicializa el grid
grdConsulta.Rows = 1


' Muestra los datos
CargarGridConsulta

'Verifica si el grd tiene datos
If grdConsulta.Rows > 1 Then
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = True
End If
End Sub

Private Sub CargarGridConsulta()
' ----------------------------------------------------
' Propósito : Muestra la consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim dblTotalRendir As Double

' Inicializa la variable
dblTotalRendir = 0

' Carga los datos al grd
Do While Not mcurSaldosRendir.EOF
    
   ' Codigo,Apellidos,Nombres,SaldoRendir
    grdConsulta.AddItem mcurSaldosRendir.campo(0) & vbTab & _
                        mcurSaldosRendir.campo(1) & vbTab & _
                        mcurSaldosRendir.campo(2) & vbTab & _
                        Format(Val(mcurSaldosRendir.campo(3)), "###,###,##0.00")
                                                     
    ' Acumula los ingresos a la cuenta
    dblTotalRendir = dblTotalRendir + Val(mcurSaldosRendir.campo(3))
                    
    'Mueve al siguiente elemento del cursor
    mcurSaldosRendir.MoverSiguiente
            
Loop
 
 ' Muestra el total la cuenta a rendir a la fecha
 grdConsulta.AddItem vbTab & "TOTAL A RENDIR : " _
                   & vbTab & vbTab & _
                   Format(dblTotalRendir, "###,###,##0.00")
' Colorea el grid
grdConsulta.Row = grdConsulta.Rows - 1
MarcarFilaGRID grdConsulta, vbBlack, vbGray
   
End Sub

Private Function CargaSaldosRendir() As Boolean
' ----------------------------------------------------
' Propósito : Carga los saldos de las cuentas a rendir
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
sSQL = "SELECT M.IdPersona, P.Apellidos, P.Nombre, (Sum(M.Ingreso)-Sum(M.Egreso)) " & _
       "FROM MOV_ENTREG_RENDIR M, PLN_PERSONAL P " & _
       "WHERE M.IdPersona=P.IdPersona and M.Anulado='NO' and " & _
       "M.Fecha <= '" & FechaAMD(mskFecha) & "' " & _
       "GROUP BY M.IdPersona, P.Apellidos, P.Nombre "
' Verifica las opciones
If optDeudor.Value = True Then
    ' Añade seleccione los deudores
    sSQL = sSQL & "HAVING (Sum(M.Ingreso)-Sum(M.Egreso)) > 0.001 ORDER BY P.Apellidos"
ElseIf optRendido.Value = True Then
    ' Añade seleccione los cancelados
    sSQL = sSQL & "HAVING (Sum(M.Ingreso)-Sum(M.Egreso)) <= 0.001 ORDER BY P.Apellidos"
ElseIf optAmbos.Value = True Then ' No hace nada
    sSQL = sSQL & "ORDER BY P.Apellidos"
End If

' Ejecuta la sentencia
mcurSaldosRendir.SQL = sSQL

' Verifica si hay error
If mcurSaldosRendir.Abrir = HAY_ERROR Then End

' Mensaje Fecha incorrecta
If mcurSaldosRendir.EOF Then
    ' Mendaje y sale de la función
    MsgBox "No existen cuentas a rendir a la fecha", vbInformation + vbOKOnly, "Consulta Saldos a Rendir"
    ' Devuelve la función
    CargaSaldosRendir = False
    ' Sale
    Exit Function
End If

' Devuelve la función
CargaSaldosRendir = True

End Function

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
aTitulosColGrid = Array("CODIGO", "APELLIDOS", "NOMBRES", "SALDO A RENDIR")
aTamañosColumnas = Array(1000, 5000, 3000, 1600)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Coloca a Obligatorio
mskFecha.BackColor = Obligatorio

'Deshabilita el cmdInforme
cmdInforme.Enabled = False

End Sub

Private Sub grdConsulta_DblClick()

'Recupera el codigo del personal y la fecha
frmCBConsulEntregaRendir.txtPersonal = grdConsulta.TextMatrix(grdConsulta.Row, 0)
frmCBConsulEntregaRendir.mskFechaFin.Text = mskFecha.Text

'Muestra el formulario del seguimiento a rendir
frmCBConsulEntregaRendir.Show vbModal, Me

End Sub

Private Sub mskFecha_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFecha) Then
    mskFecha.BackColor = vbWhite
    ' Carga consulta
    CargaConsulta
Else
  mskFecha.BackColor = Obligatorio
  grdConsulta.Rows = 1
  ' Deshabilita el botón generar informe
  cmdInforme.Enabled = False
End If

End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optAmbos_Click()

' Carga la consulta
CargaConsulta

End Sub

Private Sub optDeudor_Click()

' Carga la consulta
CargaConsulta

End Sub

Private Sub optRendido_Click()

' Carga la consulta
CargaConsulta

End Sub
