VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALConsulAlmacenSalidas 
   Caption         =   "Consulta de salidas de almacén"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   93
   Icon            =   "SCALConsulAlmacenSalidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   360
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
   Begin VB.TextBox txtTotalSalidas 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5700
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8085
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   7440
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   285
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         Height          =   735
         Left            =   480
         TabIndex        =   9
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
            TabIndex        =   13
            Top             =   255
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consulta del "
            Height          =   195
            Left            =   360
            TabIndex        =   10
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
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL SALIDAS:"
      Height          =   195
      Left            =   4200
      TabIndex        =   8
      Top             =   8160
      Width           =   1290
   End
End
Attribute VB_Name = "frmALConsulAlmacenSalidas"
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
Dim rptSalidasAlmacen As New clsBD4

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
  LlenarTablaRPTALSALIDA
  
' Formulario
  Set rptSalidasAlmacen.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptSalidasAlmacen.AsignarRpt

' Formula/s de Crystal.
  rptSalidasAlmacen.Formulas.Add "Fecha='DEL " & mskFechaIni.Text & " AL " & mskFechaFin.Text & "'"
 
' Clausula WHERE de las relaciones del rpt.
  rptSalidasAlmacen.FiltroSelectionFormula = ""

' Nombre del fichero
  rptSalidasAlmacen.NombreRPT = "RPTALMACENSALIDA.rpt"

' Presentación preliminar del Informe
  rptSalidasAlmacen.PresentancionPreliminar

'Sentencia SQL
 sSQL = ""
 sSQL = "DELETE * FROM RPTALSALIDA"

'Borra la tabla
 Var21 sSQL
 
' Elimina los datos de la BD
  Var43 gsFormulario
  
 ' Habilita el botón informe
 cmdInforme.Enabled = True


End Sub

Private Sub LlenarTablaRPTALSALIDA()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modAlSalida As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1
    'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
     sSQL = "INSERT INTO RPTALSALIDA VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 4)) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & "'" & Var9(grdConsulta.TextMatrix(i, 5)) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 6))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 7))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 8))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 9) & "', " _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modAlSalida.SQL = sSQL
    
    'Verifica si hay error
    If modAlSalida.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modAlSalida.Cerrar

Next i

End Sub


Private Sub cmdSalir_Click()
'Descarga el formulario
Unload Me
End Sub


Private Sub CargaAlmacenSalidas()
' ----------------------------------------------------
' Propósito : Arma la consulta de salidas de almacén
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curSalidasAlmacen As New clsBD2
Dim dblTotalSalidas As Double

' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
    txtTotalSalidas.Text = "0.00"
    grdConsulta.Rows = 1
    'Deshabilita el cmdInforme
    cmdInforme.Enabled = False
    Exit Sub
  End If
  
'Vacia el grdConsulta
grdConsulta.Rows = 1

' Carga la sentencia
'Fecha, IdSalida, IdProd,IdPersona ,Nombre, DescProd, Cantidad,PrecioUni,Total, CodCont, Orden
sSQL = "SELECT SA.Fecha, AD.IdSalida, AD.IdProd, PP.IdPersona, ( PP.Apellidos & ', ' & PP.Nombre), P.DescProd, SUM(AD.Cantidad), " & _
       "SUM(AD.Precio),P.CodCont " & _
       "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD,PRODUCTOS P, PLN_PERSONAL PP " & _
       "WHERE SA.Fecha BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " & _
       "SA.IdSalida=AD.IdSalida And AD.IdProd=P.IdProd And SA.IdPersona=PP.IdPersona And SA.Anulado='NO' " & _
       "GROUP BY SA.Fecha, AD.IdSalida, AD.IdProd, PP.IdPersona, ( PP.Apellidos & ', ' & PP.Nombre), P.DescProd, P.CodCont " & _
       "ORDER BY SA.Fecha "
       
' Ejecuta la sentencia
curSalidasAlmacen.SQL = sSQL
If curSalidasAlmacen.Abrir = HAY_ERROR Then End

'Inicializa la variable
dblTotalSalidas = 0

'Verifica que no hay registros en la consulta
If curSalidasAlmacen.EOF Then
    'Mensaje no hay exsitencias en almacén
    MsgBox "No hay salidas de almacén entre estas fechas", , "Almacén - Consulta de Salidas"
    ' cierra la consulta
    curSalidasAlmacen.Cerrar
    'Limpiar Grid
    grdConsulta.Rows = 1
    'Termina la ejecución del procedimiento
    Exit Sub
Else

    ' Recorre el cursor ingresos a cta
    Do While Not curSalidasAlmacen.EOF
    
        ' Añade el elemento al grid
        'Fecha, IdSalida, IdProd, IdPersona ,Nombre, DescProd, Cantidad,PrecioUni,Total, CtaContable
        grdConsulta.AddItem FechaDMA(curSalidasAlmacen.campo(0)) & vbTab & _
                            curSalidasAlmacen.campo(1) & vbTab & _
                            curSalidasAlmacen.campo(2) & vbTab & _
                            curSalidasAlmacen.campo(3) & vbTab & _
                            curSalidasAlmacen.campo(4) & vbTab & _
                            curSalidasAlmacen.campo(5) & vbTab & _
                            Format(curSalidasAlmacen.campo(6), "###,###,##0.00") & vbTab & _
                            Format(curSalidasAlmacen.campo(7) / curSalidasAlmacen.campo(6), "###,###,##0.00") & vbTab & _
                            Format(curSalidasAlmacen.campo(7), "###,###,##0.00") & vbTab & _
                            curSalidasAlmacen.campo(8)
                             
       'Acumula los ingresos a la cuenta
       dblTotalSalidas = dblTotalSalidas + Val(curSalidasAlmacen.campo(7))
        
       ' Mueve al siguiente programa
       curSalidasAlmacen.MoverSiguiente
    Loop
    
    ' Muestra el total de los proyectos
    txtTotalSalidas.Text = Format(dblTotalSalidas, "###,###,##0.00")

End If

'Cierra el cursor
curSalidasAlmacen.Cerrar

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

'Establece la ubicación del formulario
Me.Top = 0

' Carga los títulos del grid
'Fecha, Documento, Insumo, IdPersona,Persona, DescProd, Cantidad, PrecioUni, Total, CodCont
aTitulosColGrid = Array("FECHA", "Nº SALIDA", "INSUMO", "IDPERSONA", "USUARIO", "DESCRIPCION", "CANTIDAD", "PRECIO UNIT.", "TOTAL", "COD.CONT")
aTamañosColumnas = Array(1000, 1200, 800, 0, 3500, 3500, 900, 1000, 1200, 1000)
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
  CargaAlmacenSalidas
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  txtTotalSalidas.Text = "0.00"
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
  CargaAlmacenSalidas
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  txtTotalSalidas.Text = "0.00"
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
