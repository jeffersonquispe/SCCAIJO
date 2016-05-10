VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALConsulSalidasDoc 
   Caption         =   "SGCcaijo-Consulta Salidas de Almacén por Documento"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   98
   Icon            =   "SCALSalidasDoc.frx":0000
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
      Begin VB.TextBox txtNumDoc 
         Height          =   315
         Left            =   4200
         MaxLength       =   11
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Por Fechas"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optNumero 
         Caption         =   "Por Número"
         Height          =   255
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   6360
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
         Left            =   2520
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
      Cols            =   7
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
Attribute VB_Name = "frmALConsulSalidasDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Colección para la carga de la consulta
Private mcolDocumentos As New Collection

' Cursores para la carga de la consulta
Private mcurDetSalida As New clsBD2

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

' Genera el reporte formulario
  Set rptMayor.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptMayor.AsignarRpt

' Clausula WHERE de las relaciones del rpt.
  rptMayor.FiltroSelectionFormula = ""

' Nombre del fichero
  rptMayor.NombreRPT = "rptALSalidaDoc.rpt"

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
' Propósito: Borra los datos de las tablas RPTALSALIDADOC y _
             RPTALSALIDADOCDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3

' Carga la sentencia
sSQL = "DELETE * FROM RPTALSALIDADOCDET"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Carga la sentencia
sSQL = "DELETE * FROM RPTALSALIDADOC"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modTablaConsul.Cerrar

End Sub

Private Sub LlenaTablaConsul()
'------------------------------------------------------------
' Propósito: LLena las tablas RPTALSALIDADOC y _
             RPTALSALIDADOCDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3
Dim i As Long
Dim sReg As Variant

' grd: "IdSalida","Cantidad","Unidad", "IdProd", "Descripción", "Precio Unit.", "Total"
' cur:AD.IdSalida,AD.Cantidad.,P.Medida,P.IdProd," _
     & "P.DescProd,AD.Precio/AD.Cantidad,A.Precio
' Col:A.IdSalida,A.Fecha,PL.Apellidos + ' ' + PL.Nombre,P.DescProy' Recorre los documentos verificados
For Each sReg In mcolDocumentos
      
      ' Carga la sentencia sSQL guarda el detalle en la tabla detalle
      sSQL = "INSERT INTO RPTALSALIDADOC VALUES('" _
        & Var30(sReg, 1) & "','" _
        & FechaDMA(Var30(sReg, 2)) & "','" _
        & Var9(Var30(sReg, 3)) & "','" _
        & Var9(Var30(sReg, 4)) & "')"
      
      ' Ejecuta la sentencia
      modTablaConsul.SQL = sSQL
      If modTablaConsul.Ejecutar = HAY_ERROR Then End
      modTablaConsul.Cerrar

Next sReg

' Recorre los datos del grid llena el detalle
For i = 1 To grdConsulta.Rows - 1
' grd: "IdSalida","Cantidad","Unidad", "IdProd", "Descripción", "Precio Unit.", "Total"
    If grdConsulta.TextMatrix(i, 0) <> Empty Then
      
      ' Carga la sentencia sSQL guarda el detalle en la tabla detalle
      sSQL = "INSERT INTO RPTALSALIDADOCDET VALUES('" _
        & grdConsulta.TextMatrix(i, 0) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 1)) & ",'" _
        & Var9(grdConsulta.TextMatrix(i, 2)) & "','" _
        & grdConsulta.TextMatrix(i, 3) & "','" _
        & Var9(grdConsulta.TextMatrix(i, 4)) & "'," _
        & Var37(grdConsulta.TextMatrix(i, 5)) & "," _
        & Var37(grdConsulta.TextMatrix(i, 6)) & ")"
      
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

' Carga los tamaños de las 7 columnas
' "IdSalida","Cantidad","Unidad", "IdProd", "Descripción", "Precio Unit.", "Total"
aTamañosColumnas = Array(0, 1300, 2100, 1000, 4200, 1300, 1300)
aTitulosColGrid = Array("IdSalida", "Cantidad", "Unidad", "IdProd", "Descripción", "Precio Unit.", "Total")
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas
    
' Inicia alineamieto de la columna 3
grdConsulta.ColAlignment(2) = 1
grdConsulta.ColAlignment(4) = 1
    
' Habilita el optnumero
optNumero.Value = True

' Deshabilita el botón generar informe
cmdInforme.Enabled = False

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
   prgInforme.Max = 3
   prgInforme.Min = 0
   prgInforme.Value = 0

' Cargar los documetos de salida
  CargaDocumentosSalida
  If mcolDocumentos.Count = 0 Then
      MsgBox "No se tienen Salidas de almacén", , "SGCcaijo-Amacén, Consulta salidas de Almacén"
     ' Sale de el proceso
      Exit Sub
  End If
   prgInforme.Value = prgInforme.Value + 1
   
' Carga cursores de Almacén
   CargaSalidaMercaderias
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

Private Sub CargaSalidaMercaderias()
' ----------------------------------------------------
' Propósito: Carga el cursor con el detalle de salidas de almacén
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
If optNumero.Value = False Then
    ' Carga la consulta
    sSQL = "SELECT AD.IdSalida,SUM(AD.Cantidad),P.Medida,P.IdProd," _
         & "P.DescProd,SUM(AD.Precio/AD.Cantidad),SUM(AD.Precio) " _
         & "FROM ALMACEN_SALIDAS A,ALMACEN_SAL_DET AD,PRODUCTOS P " _
         & "WHERE A.IdSalida=AD.IdSalida and A.Anulado='NO' and " _
         & "(A.Fecha between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') and " _
         & "AD.IdProd=P.IdProd " _
         & "GROUP BY AD.IdSalida,P.Medida,P.IdProd,P.DescProd " _
         & "ORDER BY AD.IdSalida"
Else
    ' Carga la consulta
    sSQL = "SELECT AD.IdSalida,SUM(AD.Cantidad),P.Medida,P.IdProd," _
         & "P.DescProd,SUM(AD.Precio/AD.Cantidad),SUM(AD.Precio) " _
         & "FROM ALMACEN_SALIDAS A,ALMACEN_SAL_DET AD,PRODUCTOS P " _
         & "WHERE A.IdSalida=AD.IdSalida and A.Anulado='NO' and " _
         & "A.IdSalida='" & txtNumDoc & "' and " _
         & "AD.IdProd=P.IdProd " _
         & "GROUP BY AD.IdSalida,P.Medida,P.IdProd,P.DescProd " _
         & "ORDER BY AD.IdSalida"
End If
' Ejecuta la sentencia
mcurDetSalida.SQL = sSQL
If mcurDetSalida.Abrir = HAY_ERROR Then End

End Sub


Private Function CargaDocumentosSalida() As Boolean
' ----------------------------------------------------
' Propósito: Carga la colección que almacena los documentos de salida
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curOrdenSalida As New clsBD2

' Limpia la colección
Set mcolDocumentos = Nothing
If optNumero.Value = False Then
    ' Carga la sentencia que carga los Documentos
    sSQL = "SELECT A.IdSalida,A.Fecha,PL.Apellidos + ' ' + PL.Nombre,P.DescProy " _
        & "FROM ALMACEN_SALIDAS A, PLN_PERSONAL PL,PROYECTOS P " _
        & "WHERE A.IdPersona=PL.IdPersona and A.IdProy=P.IdProy and " _
        & "A.Anulado='NO' and " _
        & "(A.Fecha between '" & FechaAMD(mskFechaIni) & "' and '" & FechaAMD(mskFechaFin) & "') " _
        & "ORDER BY A.IdSalida"
Else
    ' Carga la sentencia que carga los Documentos
    sSQL = "SELECT A.IdSalida,A.Fecha,PL.Apellidos + ' ' + PL.Nombre,P.DescProy " _
        & "FROM ALMACEN_SALIDAS A, PLN_PERSONAL PL,PROYECTOS P " _
        & "WHERE A.IdPersona=PL.IdPersona and A.IdProy=P.IdProy and " _
        & "A.Anulado='NO' and " _
        & "A.IdSalida='" & txtNumDoc & "' " _
        & "ORDER BY A.IdSalida"

End If
' Ejecuta la sentencia
curOrdenSalida.SQL = sSQL
If curOrdenSalida.Abrir = HAY_ERROR Then End

' Recorre el cursor de salidas
Do While Not curOrdenSalida.EOF
    ' Añade un documento de salida
    mcolDocumentos.Add curOrdenSalida.campo(0) _
            & "¯" & curOrdenSalida.campo(1) _
            & "¯" & curOrdenSalida.campo(2) _
            & "¯" & curOrdenSalida.campo(3)
    
    ' Mueve al siguiente elemento de mercaderías
    curOrdenSalida.MoverSiguiente
Loop

' Cierra los cursore
curOrdenSalida.Cerrar

End Function

Private Sub CargarGridConsulta()
' ----------------------------------------------------
' Propósito: Arma la consulta en el grid
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim brecorre As Boolean
Dim sReg As Variant
' grd: "IdSalida","Cantidad","Unidad", "IdProd", "Descripción", "Precio Unit.", "Total"
' cur:AD.IdSalida,AD.Cantidad.,P.Medida,P.IdProd," _
     & "P.DescProd,AD.Precio/AD.Cantidad,A.Precio
' Col:A.IdSalida,A.Fecha,PL.Apellidos + ' ' + PL.Nombre,P.DescProy
' Inicializa el grid
grdConsulta.Rows = 1
grdConsulta.ScrollBars = flexScrollBarNone
grdConsulta.Visible = False
' Recorre los documentos verificados por día
For Each sReg In mcolDocumentos
   mdblPrecioTotal = 0
   ' Añade el documento al grid
   grdConsulta.AddItem vbTab & "Nro.Salida:" & vbTab & _
        Var30(sReg, 1) & vbTab & "Proyecto:" & vbTab & _
        Var30(sReg, 4) & vbTab & _
        "Fecha:" & vbTab & FechaDMA(Var30(sReg, 2))
  ' Colorea el grid
   grdConsulta.Row = grdConsulta.Rows - 1
   MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
   ' Añade las Mercaderías
   MuestraMercaderias (sReg)
   ' Muestra el total del documento
   ' Añade el documento al grid
   grdConsulta.AddItem vbTab & vbTab & vbTab & vbTab & _
        "Total" & vbTab & vbTab & Format(mdblPrecioTotal, "###,###,##0.00")
   'Colorea el grid
   grdConsulta.Row = grdConsulta.Rows - 1
   MarcarFilaGRID grdConsulta, vbBlack, vbGray

Next sReg ' Siguiente registro

' Coloca las barras de desplazamiento
grdConsulta.ScrollBars = flexScrollBarBoth
grdConsulta.Visible = True

' Cierra los cursores generales
mcurDetSalida.Cerrar

End Sub

Private Sub MuestraMercaderias(sOrdenSal As String)
'---------------------------------------------------------------
' Propósito: Muestra los productos que salieron en almacén
' Recibe: sOrden, Documento con el que se compró
' Entrga:Nada
'---------------------------------------------------------------
' grd: "IdSalida","Cantidad", "Unidad", "IdProd", "Descripción", "Precio Unit.", "Total"
' cur:AD.IdSalida,AD.Cantidad.,P.Medida,P.IdProd," _
     & "P.DescProd,AD.Precio/AD.Cantidad,A.Precio
' col:A.IdSalida,A.Fecha,PL.Apellidos + ' ' + PL.Nombre,P.DescProy
Dim brecorre As Boolean
' Carga los datos de ingreso de activo
brecorre = True
Do While brecorre
    ' Verifica si el cursor es vacío
    If mcurDetSalida.EOF Then
        brecorre = False
    Else
        ' Verifica que los datos son de IdSalida
        If mcurDetSalida.campo(0) = Var30(sOrdenSal, 1) Then
              ' Añade Mercadería
              grdConsulta.AddItem mcurDetSalida.campo(0) & vbTab & _
                               Format(mcurDetSalida.campo(1), "##0.00") & vbTab & _
                               mcurDetSalida.campo(2) & vbTab & _
                               mcurDetSalida.campo(3) & vbTab & _
                               mcurDetSalida.campo(4) & vbTab & _
                               Format(mcurDetSalida.campo(5), "###,###,##0.00") & vbTab & _
                               Format(mcurDetSalida.campo(6), "###,###,##0.00")
              ' Acumula en el debe
              mdblPrecioTotal = mdblPrecioTotal + Val(mcurDetSalida.campo(6))
              ' Pasa al siguiente control
              mcurDetSalida.MoverSiguiente
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
Set mcolDocumentos = Nothing

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
