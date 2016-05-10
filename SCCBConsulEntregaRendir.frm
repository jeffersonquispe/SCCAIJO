VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCBConsulEntregaRendir 
   Caption         =   "Consulta de kardex de almacén"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   95
   Icon            =   "SCCBConsulEntregaRendir.frx":0000
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
   Begin VB.CommandButton cmdPPersonal 
      Height          =   255
      Left            =   6810
      Picture         =   "SCCBConsulEntregaRendir.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   585
      Width           =   220
   End
   Begin VB.ComboBox cboPersonal 
      Height          =   315
      Left            =   2130
      Style           =   1  'Simple Combo
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   555
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
      Caption         =   "Datos del Personal"
      Height          =   1095
      Left            =   280
      TabIndex        =   7
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtActivo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         TabIndex        =   15
         Top             =   705
         Width           =   750
      End
      Begin VB.TextBox txtPersonal 
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   0
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Activo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Personal:"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   360
         Width           =   660
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
      Cols            =   8
      FixedCols       =   2
      FillStyle       =   1
      MergeCells      =   3
   End
End
Attribute VB_Name = "frmCBConsulEntregaRendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcolIdPers As New Collection
Dim mcolCodDesPers As New Collection
Dim mcolDesEstadoPers As New Collection
Dim mdblMontoIngresos As Double
Dim mdblMontoSalidas As Double
Dim mcurIngresos As New clsBD2
Dim mcurSalidasEgre As New clsBD2
Dim mcurSalidasIngre As New clsBD2
Dim mblnIngresos, mblnSalidasIngre, mblnSalidasEgre As Boolean
Dim mcolDesTipoDocum As New Collection

Private Sub cboPersonal_Change()

' verifica si lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboPersonal) = True Then SendKeys "{down}"

End Sub

Private Sub cboPersonal_Click()

' Verifica si el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboPersonal.ListIndex) = False Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboPersonal_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica si es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboPersonal_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboPersonal, vbWhite) = True Then
   ' Se actualiza código (TextBox) correspondiente a descripción introducida
   CD_ActCod cboPersonal.Text, txtPersonal, mcolIdPers, mcolCodDesPers
Else
   'Coloca a obligatorio el txt
   txtPersonal.Text = Empty
End If

'Cambia el alto del combo
cboPersonal.Height = CBONORMAL

End Sub

Private Sub cmdInforme_Click()
Dim sSQL As String
Dim rptKardexRendir As New clsBD4

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
  LlenarTablaRPTCBRENDIRKARDEX
 
' Formulario
  Set rptKardexRendir.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptKardexRendir.AsignarRpt

' Formula/s de Crystal.
  rptKardexRendir.Formulas.Add "Fecha='" & mskFechaIni & " AL " & mskFechaFin & "'"
  rptKardexRendir.Formulas.Add "CodPersonal='" & txtPersonal & "'"
  rptKardexRendir.Formulas.Add "Personal='" & cboPersonal.Text & "'"
  rptKardexRendir.Formulas.Add "Estado='" & txtActivo.Text & "'"
 
' Clausula WHERE de las relaciones del rpt.
  rptKardexRendir.FiltroSelectionFormula = ""

' Nombre del fichero
  rptKardexRendir.NombreRPT = "RPTCBRENDIRKARDEX.rpt"

' Presentación preliminar del Informe
  rptKardexRendir.PresentancionPreliminar

'Sentencia SQL
 sSQL = "DELETE * FROM RPTCBRENDIRKARDEX"

'Borra la tabla
 Var21 sSQL

' Elimina los datos de la BD
  Var43 gsFormulario
  
' Habilita el botón informe
 cmdInforme.Enabled = True


End Sub

Private Sub LlenarTablaRPTCBRENDIRKARDEX()
'-----------------------------------------------------
'Propósito  :Llena la tabla con los datos del grdConsulta
'Recibe     :Nada
'Devuelve   :Nada
'-----------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim modRendirKardex As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdConsulta.Rows - 1

    'Fecha, Tipo Doc, Nro Doc., Ingreso, Egreso,Movimiento,Saldo Rendir, Orden
     sSQL = "INSERT INTO RPTCBRENDIRKARDEX VALUES " _
     & "('" & grdConsulta.TextMatrix(i, 0) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 1) & "', " _
     & "'" & grdConsulta.TextMatrix(i, 2) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 3))) & "," _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 4))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 7) & "', " _
     & " " & Val(Var37(grdConsulta.TextMatrix(i, 5))) & "," _
     & "'" & grdConsulta.TextMatrix(i, 6) & "'," _
     & " " & i & ")"
    
    'Copia la sentencia sSQL
    modRendirKardex.SQL = sSQL
    
    'Verifica si hay error
    If modRendirKardex.Ejecutar = HAY_ERROR Then
      End
    End If
    
    'Se cierra la query
    modRendirKardex.Cerrar

Next i

End Sub


Private Sub cmdPPersonal_Click()
'Cambia el alto del cboPersonal
If cboPersonal.Enabled Then
    ' alto
     cboPersonal.Height = CBOALTO
    ' focus a cbo
    cboPersonal.SetFocus
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

'Carga la colección de personal
CargarColPersonal

'Carga la colección del tipo de documento
CargarColTipoDocum

' Limpia el combo del personal
cboPersonal.Clear

'Carga el cboPersonal de acuerdo a la relación
CargarCboCols cboPersonal, mcolCodDesPers

' Carga los títulos del grid
'Fecha, Tipo Doc,Nro Doc., Ingreso, Egreso,Movimiento,Saldo Rendir, Orden
aTitulosColGrid = Array("FECHA", "TIPO DE DOCUMENTO", "NRO. DOCUMENTO", "INGRESO", "EGRESO", "SALDO RENDIR", "ORDEN", "MOVIMIENTO")
aTamañosColumnas = Array(1000, 2300, 1600, 1200, 1200, 1400, 1200, 1200)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

'Deshabilita el control cmdInforme
cmdInforme.Enabled = False

'Los campos coloca a color amarillo
EstableceCamposObligatorios

End Sub

Private Sub CargarColPersonal()
'---------------------------------------------------------------
'Propósito  : Carga la colección de personal con su medida
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curPersonal As New clsBD2
        
sSQL = "SELECT DISTINCT P.IdPersona, ( P.Apellidos & ', ' & P.Nombre), PR.Activo " _
               & "FROM PLN_PERSONAL P, PLN_PROFESIONAL PR ,MOV_ENTREG_RENDIR ME " _
               & "WHERE P.IdPersona=PR.IdPersona  and P.IdPersona=ME.IdPersona And ME.Anulado='NO' "

'Carga la colección de descripcion y medida de los productos
curPersonal.SQL = sSQL
If curPersonal.Abrir = HAY_ERROR Then
  End
End If
Do While Not curPersonal.EOF
    'Colección de Estado del personal
    mcolDesEstadoPers.Add Key:=curPersonal.campo(0), _
                         Item:=curPersonal.campo(2)
                              
    'Colección de producto y su descripción
    mcolIdPers.Add curPersonal.campo(0)
    mcolCodDesPers.Add curPersonal.campo(1), curPersonal.campo(0)

    ' Se avanza a la siguiente fila del cursor
    curPersonal.MoverSiguiente
Loop

'Cierra el cursor de medida de productos
curPersonal.Cerrar

End Sub

Private Sub CargarColTipoDocum()
'---------------------------------------------------------------
'Propósito  : Carga la colección de tipos de documentos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curTipoDocum As New clsBD2
        
sSQL = "SELECT DISTINCT IdTipoDoc,DescTipoDoc " _
               & "FROM TIPO_DOCUM  "

'Carga la colección de descripcion y medida de los productos
curTipoDocum.SQL = sSQL
If curTipoDocum.Abrir = HAY_ERROR Then
  End
End If
Do While Not curTipoDocum.EOF
    'Colección de Estado del personal
    mcolDesTipoDocum.Add Key:=curTipoDocum.campo(0), _
                         Item:=curTipoDocum.campo(1)
                              
    ' Se avanza a la siguiente fila del cursor
    curTipoDocum.MoverSiguiente
Loop

'Cierra el cursor de medida de productos
curTipoDocum.Cerrar

End Sub

Private Sub EstableceCamposObligatorios()
' ------------------------------------------------------------
' Propósito : Muestra de color amarillo los campos obligatorios
' Recibe    : Nada
' Entrega   :Nada
' ------------------------------------------------------------
mskFechaIni.BackColor = Obligatorio
mskFechaFin.BackColor = Obligatorio
txtPersonal.BackColor = Obligatorio
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Vacia la colección
Set mcolIdPers = Nothing
Set mcolCodDesPers = Nothing
Set mcolDesEstadoPers = Nothing
Set mcolDesTipoDocum = Nothing

End Sub

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  
  ' Carga los Kardex de almacén
  CargaKardexEntrega
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
If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio And txtPersonal.BackColor <> Obligatorio Then
' Verifica que la fecha de inicio sea Menor a la fecha final
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = True Then
        fbOkDatosIntroducidos = False
        Exit Function
    End If
End If
' Verifica si los datos obligatorios se ha llenado
If mskFechaIni.BackColor <> vbWhite Or _
   mskFechaFin.BackColor <> vbWhite Or _
   txtPersonal.BackColor <> vbWhite Then
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
  CargaKardexEntrega
  
Else
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub CargaKardexEntrega()
' ----------------------------------------------------
' Propósito : Carga los kardex de entrega a rendir entre
'             las fecha de consulta
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

'Carga los ingresos de entregas a rendir
IngresosEntregasAntesFechaIni

'Carga las salidas de entregas a rendir
SalidasEntregasAntesFechaIni

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
Dim blnCargaIng, blnCargaSalidaEgre, blnCargaSalidaIngre As Boolean
Dim blnRecorreCursores As Boolean
Dim strMenor As String
Dim dblMontoSaldo As Double

'Carga el cursor con los ingresos
DeterminarIngresos

'Carga el cursor con las salidas
DeterminarSalidasIngre

'Carga el cursor con las salidas
DeterminarSalidasEgre

'Verifica si hay ingresos y salidas
If mblnIngresos = False And mblnSalidasEgre = False And mblnSalidasIngre = False Then
    'No hay salidas e ingreso de almacen
    cmdInforme.Enabled = False
    
    'Cierra los cursores
    mcurIngresos.Cerrar
    mcurSalidasEgre.Cerrar
    mcurSalidasIngre.Cerrar
    'Termina la ejecución
    Exit Sub
Else
    'Hay ingreso o salida de almacen
    cmdInforme.Enabled = True
End If
'Agrega los datos al grid
'Verifica que no sea fin de registro
blnRecorreCursores = True
Do While blnRecorreCursores = True

   ' Verifica si se ha terminado de recorrer todos los cursores
   If mcurIngresos.EOF And mcurSalidasEgre.EOF And mcurSalidasIngre.EOF Then
       ' Sale de recorrer cursor
       blnRecorreCursores = False
       
   Else
        'Verifica que ninguno de los cursores sea el final del registro
        If Not mcurIngresos.EOF And Not mcurSalidasEgre.EOF And Not mcurSalidasIngre.EOF Then
            'Asigna el valor del Menor al mcurIngreso
            strMenor = mcurIngresos.campo(5)
            blnCargaIng = True
            blnCargaSalidaIngre = False
            blnCargaSalidaEgre = False
            'Verifica si el strMenor es Menor que mcurSalidasEgre
            If strMenor > mcurSalidasIngre.campo(5) Then
                strMenor = mcurSalidasIngre.campo(5)
                blnCargaIng = False
                blnCargaSalidaIngre = True
                blnCargaSalidaEgre = False
            End If
            
            If strMenor > mcurSalidasEgre.campo(5) Then
                strMenor = mcurSalidasEgre.campo(5)
                blnCargaIng = False
                blnCargaSalidaIngre = False
                blnCargaSalidaEgre = True
            End If
        
        'El mcurIngresos no es fin del registro
        ElseIf Not mcurIngresos.EOF And Not mcurSalidasIngre.EOF Then
            strMenor = mcurIngresos.campo(5)
            blnCargaIng = True
            blnCargaSalidaIngre = False
            blnCargaSalidaEgre = False
            'Verifica si el strMenor es Menor que mcurSalidasEgre
            If strMenor > mcurSalidasIngre.campo(5) Then
                strMenor = mcurSalidasIngre.campo(5)
                blnCargaIng = False
                blnCargaSalidaIngre = True
                blnCargaSalidaEgre = False
            End If
            
        'El mcurSalidasEgre no es fin del registro
        ElseIf Not mcurIngresos.EOF And Not mcurSalidasEgre.EOF Then
            strMenor = mcurIngresos.campo(5)
            blnCargaIng = True
            blnCargaSalidaIngre = False
            blnCargaSalidaEgre = False
            'Verifica si el strMenor es Menor que mcurSalidasEgre
            If strMenor > mcurSalidasEgre.campo(5) Then
                strMenor = mcurSalidasEgre.campo(5)
                blnCargaIng = False
                blnCargaSalidaIngre = False
                blnCargaSalidaEgre = True
            End If
        'El mcurSalidasEgre no es fin del registro
        ElseIf Not mcurSalidasIngre.EOF And Not mcurSalidasEgre.EOF Then
            strMenor = mcurSalidasIngre.campo(5)
            blnCargaIng = False
            blnCargaSalidaIngre = True
            blnCargaSalidaEgre = False
            'Verifica si el strMenor es Menor que mcurSalidasEgre
            If strMenor > mcurSalidasEgre.campo(5) Then
                strMenor = mcurSalidasEgre.campo(5)
                blnCargaIng = False
                blnCargaSalidaIngre = False
                blnCargaSalidaEgre = True
            End If
        ElseIf Not mcurIngresos.EOF Then
            blnCargaIng = True
            blnCargaSalidaIngre = False
            blnCargaSalidaEgre = False
        ElseIf Not mcurSalidasIngre.EOF Then
            blnCargaIng = False
            blnCargaSalidaIngre = True
            blnCargaSalidaEgre = False
        ElseIf Not mcurSalidasEgre.EOF Then
            blnCargaIng = False
            blnCargaSalidaIngre = False
            blnCargaSalidaEgre = True
        End If
             
        ' añade una fila al grid
        If blnCargaIng Then
            
            'Determina el saldo y PrecioSaldo
            dblMontoSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5))) + Val(mcurIngresos.campo(4))
            
            ' Coloca al grid el concepto de ingreso
            grdConsulta.AddItem FechaDMA(mcurIngresos.campo(0)) & vbTab & _
                                fsAsignarDoc(mcurIngresos.campo(2)) & vbTab & mcurIngresos.campo(3) & vbTab & _
                    Format(mcurIngresos.campo(4), "###,###,###,##0.00") & vbTab & vbTab & _
                    Format(dblMontoSaldo, "###,###,###,##0.00") & vbTab & _
                    mcurIngresos.campo(5) & vbTab & "ENTREGA"
            
            'Coloca el solor a los ingresos
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, &H80000003

            ' Mueve al siguiente concepto de ingreso
            mcurIngresos.MoverSiguiente
        End If
        
        'Agrega fila al grid
        If blnCargaSalidaIngre Then
        
            'Determina el saldo y PrecioSaldo
            dblMontoSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5))) - Val(mcurSalidasIngre.campo(4))
                                                       
            ' Coloca al grid el concepto de ingreso
            grdConsulta.AddItem FechaDMA(mcurSalidasIngre.campo(0)) & vbTab & _
                                fsAsignarDoc(mcurSalidasIngre.campo(2)) & vbTab & mcurSalidasIngre.campo(3) & vbTab _
                    & vbTab & Format(mcurSalidasIngre.campo(4), "###,###,###,##0.00") & vbTab & _
                    Format(dblMontoSaldo, "###,###,###,##0.00") & vbTab & _
                    mcurSalidasIngre.campo(5) & vbTab & "DEVOLUCION"
                    
            'Coloca el color del egreso
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbBlack, &H80000005
            
            ' Mueve al siguiente concpto de ingreso
            mcurSalidasIngre.MoverSiguiente
        End If
        
        'Agrega fila al grid
        If blnCargaSalidaEgre Then
        
            'Determina el saldo y PrecioSaldo
            dblMontoSaldo = Val(Var37(grdConsulta.TextMatrix(grdConsulta.Rows - 1, 5))) - Val(mcurSalidasEgre.campo(4))
                                                       
            ' Coloca al grid el concepto de ingreso
            grdConsulta.AddItem FechaDMA(mcurSalidasEgre.campo(0)) & vbTab & _
                                fsAsignarDoc(mcurSalidasEgre.campo(2)) & vbTab & mcurSalidasEgre.campo(3) & vbTab _
                    & vbTab & Format(mcurSalidasEgre.campo(4), "###,###,###,##0.00") & vbTab & _
                    Format(dblMontoSaldo, "###,###,###,##0.00") & vbTab & _
                    mcurSalidasEgre.campo(5) & vbTab & "RENDICION"
                                
            'Coloca el color del egreso
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbBlack, &H80000005
            
            ' Mueve al siguiente concpto de ingreso
            mcurSalidasEgre.MoverSiguiente
        End If
    End If
Loop 'Fin de hacer mientras sea fin de cursor

'Cierra los cursores
mcurIngresos.Cerrar
mcurSalidasEgre.Cerrar
mcurSalidasIngre.Cerrar

End Sub

Private Function fsAsignarDoc(strTipoDoc As String) As String
' ----------------------------------------------------
' Propósito : Determina el Tipo de documento
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------

'Verifica si los cursores son vacios
fsAsignarDoc = Var30(mcolDesTipoDocum.Item(strTipoDoc), 2)

End Function

Private Sub DeterminarIngresos()
' ----------------------------------------------------
' Propósito : Determina los ingresos y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = "SELECT E.FecMov,M.IdPersona ,E.IdTipoDoc, E.NumDoc, E.MontoCB, E.Orden " & _
     "FROM MOV_ENTREG_RENDIR M left outer join EGRESOS E on M.Orden=E.Orden " & _
     "WHERE  M.IdPersona='" & txtPersonal.Text & "' And E.Anulado='NO' And " & _
     "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' " & _
     "And M.Operacion='I' " & _
     "ORDER BY E.Orden,E.FecMov "
      
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

Private Sub DeterminarSalidasEgre()
' ----------------------------------------------------
' Propósito : Determina las salidas y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = "SELECT E.FecMov,ER.IdPersona ,E.IdTipoDoc, E.NumDoc, E.MontoCB, E.Orden " & _
       "FROM MOV_ENTREG_RENDIR ER LEFT OUTER JOIN EGRESOS E ON " & _
       "ER.Orden=E.Orden " & _
       "WHERE ER.IdPersona='" & txtPersonal.Text & "' " & _
       "And ER.Operacion='E' And E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' " & _
       "And '" & FechaAMD(mskFechaFin) & "' And ER.Anulado='NO' " & _
       "ORDER BY E.Orden, E.FecMov "
       
'Ejecuta la sentencia SQL
mcurSalidasEgre.SQL = sSQL

'Verifica si hay error al ejecutar la sentencia
If mcurSalidasEgre.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si el cursor es nulo
If mcurSalidasEgre.EOF Then
    'No hay Salidas
    mblnSalidasEgre = False
Else
    'Hay Salidas
    mblnSalidasEgre = True
End If

End Sub

Private Sub DeterminarSalidasIngre()
' ----------------------------------------------------
' Propósito : Determina las salidas y carga en un cursor
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = "SELECT I.FecMov,ER.IdPersona ,I.IdTipoDoc, I.NumDoc, I.Monto, I.Orden " & _
       "FROM MOV_ENTREG_RENDIR ER LEFT OUTER JOIN INGRESOS I ON " & _
       "ER.Orden=I.Orden " & _
       "WHERE ER.IdPersona='" & txtPersonal.Text & "' " & _
       "And ER.Operacion='E' And I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' " & _
       "And '" & FechaAMD(mskFechaFin) & "' And ER.Anulado='NO' " & _
       "ORDER BY I.Orden, I.FecMov "
       
'Ejecuta la sentencia SQL
mcurSalidasIngre.SQL = sSQL

'Verifica si hay error al ejecutar la sentencia
If mcurSalidasIngre.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si el cursor es nulo
If mcurSalidasIngre.EOF Then
    'No hay Salidas
    mblnSalidasIngre = False
Else
    'Hay Salidas
    mblnSalidasIngre = True
End If

End Sub
Private Sub CargaSaldos()
' ----------------------------------------------------
' Propósito : Determina los saldos y carga al grid antes
'             de la fecha de Inicio
' Recibe    : Nada
' Entrega   : Nada
' ---------------------------------------------------
 
' Carga al grid el saldo
' Fecha, Comprobante, CompraCantidad,PrecioCompra, NumSalida, SalidadCantida
' Entregado a, PrecioSalida, SaldoUnidades, PrecioSaldo
grdConsulta.AddItem vbTab & "Saldo Anterior" & vbTab & vbTab & _
                    vbTab & vbTab & Format(mdblMontoIngresos - mdblMontoSalidas, "###,###,###,##0.00") & vbTab & vbTab

'Coloca el color a la fila
grdConsulta.Row = grdConsulta.Rows - 1
MarcarFilaGRID grdConsulta, &H80000012, &H80000000

End Sub

Private Sub IngresosEntregasAntesFechaIni()
' ----------------------------------------------------
' Propósito : Determina los ingresos de entrega a rendir
'             antes de la fecha de consulta
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curIngresoRendir As New clsBD2

'Monto ingresos a rendir
sSQL = "SELECT Sum(ER.Ingreso) " & _
       "FROM MOV_ENTREG_RENDIR ER, EGRESOS E " & _
       "WHERE ER.IdPersona='" & txtPersonal.Text & "' " & _
       "And ER.Operacion='I' And ER.Orden=E.Orden And " & _
       "E.FecMov < '" & FechaAMD(mskFechaIni) & "' And " & _
       "E.Anulado='NO' "
              
' Ejecuta la sentencia
curIngresoRendir.SQL = sSQL
If curIngresoRendir.Abrir = HAY_ERROR Then End

'Inicializa la variable
mdblMontoIngresos = 0

'Verifica que no haya ingresos
If curIngresoRendir.EOF Then
    ' Recorre el cursor ingresos a cta
    mdblMontoIngresos = 0
    
Else
    
    If IsNull(curIngresoRendir.campo(0)) Then
        'No hay ningun dato
        mdblMontoIngresos = 0
    Else
        'Asigna los montos a las variables
    mdblMontoIngresos = Val(curIngresoRendir.campo(0))

    End If
End If

End Sub

Private Sub SalidasEntregasAntesFechaIni()
' ----------------------------------------------------
' Propósito : Determina las salidas de entregas a rendir
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------------
Dim sSQL As String
Dim curSalidasEntrega As New clsBD2

sSQL = "SELECT Sum(ER.Egreso) " & _
       "FROM ((MOV_ENTREG_RENDIR ER LEFT OUTER JOIN EGRESOS E ON " & _
       "ER.Orden=E.Orden) LEFT OUTER JOIN INGRESOS I ON " & _
       "ER.Orden=I.Orden) " & _
       "WHERE ER.IdPersona='" & txtPersonal.Text & "' " & _
       "And ER.Operacion='E' And (E.FecMov < '" & FechaAMD(mskFechaIni) & "' " & _
       "Or I.FecMov < '" & FechaAMD(mskFechaIni) & "') And " & _
       "ER.Anulado='NO'"
       
' Ejecuta la sentencia
curSalidasEntrega.SQL = sSQL
If curSalidasEntrega.Abrir = HAY_ERROR Then End

'Inicializa la variable
mdblMontoSalidas = 0

'Verifica si no hay salidas
If curSalidasEntrega.EOF Then
    mdblMontoSalidas = 0
Else
      
   If IsNull(curSalidasEntrega.campo(0)) Then
        'No hay ningun dato
        mdblMontoSalidas = 0
    Else
        'Copia el total de salidas
        mdblMontoSalidas = Val(curSalidasEntrega.campo(0))
    End If
    
End If

'Cierra el cursor
curSalidasEntrega.Cerrar

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtPersonal_Change()

'Verifica si es munuscula el texto ingresado
If UCase(txtPersonal.Text) = txtPersonal.Text Then

    ' Si procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboPersonal, txtPersonal, mcolCodDesPers
     
     ' Verifica si el campo esta vacio
    If txtPersonal.Text <> Empty And cboPersonal.Text <> Empty Then
       ' Los campos coloca a color blanco
       txtPersonal.BackColor = vbWhite
       ' Muestra la unidad
       txtActivo = Var30(mcolDesEstadoPers.Item(Trim(txtPersonal)), 1)
       ' Carga los Kardex de entregas a rendir
       CargaKardexEntrega
    Else
       'Los campos coloca a color amarillo
       txtActivo = Empty
       txtPersonal.BackColor = Obligatorio
       cmdInforme.Enabled = False
       grdConsulta.Rows = 1
    End If

Else
    If Len(txtPersonal.Text) = txtPersonal.MaxLength Then
        'comvertimos a mayuscula
        txtPersonal.Text = UCase(txtPersonal.Text)
    End If
End If

End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter sale del control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

