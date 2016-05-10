VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPRConsulGastosProveedor 
   Caption         =   "SGCcaijo-Consulta de gastos de proyectos por RUC de Proveedor"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   13
   Icon            =   "frmPRConsulGastosProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGastoProy 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   14
      Top             =   8100
      Width           =   1575
   End
   Begin VB.CommandButton cmdPProyecto 
      Height          =   255
      Left            =   9480
      Picture         =   "frmPRConsulGastosProveedor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   220
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   400
      Left            =   8640
      TabIndex        =   5
      Top             =   8100
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10320
      TabIndex        =   6
      Top             =   8100
      Width           =   1000
   End
   Begin VB.ComboBox cboProyecto 
      Height          =   315
      Left            =   2450
      Style           =   1  'Simple Combo
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   7275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor"
      Height          =   1275
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11655
      Begin VB.TextBox TxTelefono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txtProyecto 
         Height          =   315
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   7080
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   960
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   11655
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   4935
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   330
            Left            =   1320
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
            Left            =   3600
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
            TabIndex        =   12
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha &Inicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   285
            Width           =   915
         End
      End
      Begin MSMask.MaskEdBox mskFecConsulta 
         Height          =   315
         Left            =   9960
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Fecha de Trabajo:"
         Height          =   255
         Left            =   8520
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   5655
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2280
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   4
      HighLight       =   0
      FillStyle       =   1
      MergeCells      =   4
   End
   Begin Crystal.CrystalReport rptInformes 
      Left            =   120
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total de gastos :"
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   2160
      TabIndex        =   15
      Top             =   8160
      Width           =   1815
   End
End
Attribute VB_Name = "frmPRConsulGastosProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Colecciones para la carga del combo de Proyectos
Private mcolCodProy As New Collection
Private mcolCodDesProy As New Collection

'Cursores para la carga de la consulta
Private mcurGastosProy As New clsBD2
Private mcurGastosProg As New clsBD2
Private mcurGastosLineas As New clsBD2
Private mcurGastosActiv As New clsBD2
Private mcurRegGastos As New clsBD2
Public IdProveedor As String

Private Sub cboProyecto_Change()
' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProyecto) = True Then SendKeys "{down}"

End Sub

Private Sub cboProyecto_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProyecto.ListIndex) = False And cboProyecto.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboProyecto_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboProyecto_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProyecto, vbWhite) = True Then
    
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboProyecto.Text, txtProyecto, mcolCodProy, mcolCodDesProy
    
Else '  Vaciar Controles enlazados al combo
    txtProyecto.Text = Empty
End If

'Cambia el alto del combo
cboProyecto.Height = CBONORMAL

End Sub

Private Sub cmdInforme_Click()
Dim rptGastosProvDocDet As New clsBD4

' Deshabilita el botón informe
  cmdInforme.Enabled = False
  
' Verifica la transaccion
  If Var46 Then
     ' Deshabilita el botón imprimir
       cmdInforme.Enabled = True
     ' Termina la ejecución del procedimiento
       Exit Sub
  End If

' Llena la tabla consulta de seguimiento presupuestal
  LlenaTablaConsul
  
' Genera el reporte
' Formulario
  Set rptGastosProvDocDet.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptGastosProvDocDet.AsignarRpt

' Formula/s de Crystal.
  rptGastosProvDocDet.Formulas.Add "Proveedor='" & txtProyecto & "   " & cboProyecto & " ' "
  'rptGastosProvDocDet.Formulas.Add "CodProy='" & txtProyecto & "'"
  rptGastosProvDocDet.Formulas.Add "Periodo='" & mskFechaIni & " AL " & mskFechaFin & "'"

' Clausula WHERE de las relaciones del rpt.
  rptGastosProvDocDet.FiltroSelectionFormula = ""

' Nombre del fichero
  rptGastosProvDocDet.NombreRPT = "rptPRSegGastosProvedor.rpt"

' Presentación preliminar del Informe
  rptGastosProvDocDet.PresentancionPreliminar

' Elimina los datos de la tabla
  BorraDatosTablaConsul

' Elimina los datos de la BD
  Var43 gsFormulario

' Habilita el botón informe
  cmdInforme.Enabled = True

End Sub

Private Sub BorraDatosTablaConsul()
'------------------------------------------------------------
' Propósito: Borra los datos de la tabla RPTPRGASTOSACTIVDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3

' Carga la sentencia
sSQL = "DELETE * FROM RPTPRGASTOSPROVEEDOR"

' Ejecuta la sentencia
modTablaConsul.SQL = sSQL
If modTablaConsul.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modTablaConsul.Cerrar

End Sub

Private Sub LlenaTablaConsul()
'------------------------------------------------------------
' Propósito: LLena la tabla RPTPRGASTOSACTIVDET
' Recibe : Nada
' Entrega :Nada
'------------------------------------------------------------
Dim sSQL As String
Dim modTablaConsul As New clsBD3
Dim i As Long
' Recorre los datos del grid
For i = 1 To grdConsulta.Rows - 1
  '("IdProy", "IdProg", "IdLinea", "IdActividad", "TipDoc", "Nro.Doc", "Fecha", "Descripción Gasto", "Monto S/.", "Orden")
  ' Carga la sentencia sSQL
   sSQL = "INSERT INTO RPTPRGASTOSPROVEEDOR VALUES('" _
      & grdConsulta.TextMatrix(i, 0) & "','" _
      & grdConsulta.TextMatrix(i, 1) & "','" _
      & grdConsulta.TextMatrix(i, 2) & "','" _
      & grdConsulta.TextMatrix(i, 3) & "','" _
      & grdConsulta.TextMatrix(i, 4) & "/" _
      & grdConsulta.TextMatrix(i, 5) & "','" _
      & FechaAMD(grdConsulta.TextMatrix(i, 6)) & "','" _
      & grdConsulta.TextMatrix(i, 7) & "'," _
      & Var37(grdConsulta.TextMatrix(i, 8)) & ",'" _
      & grdConsulta.TextMatrix(i, 9) & "')"
  ' Ejecuta la sentencia
  modTablaConsul.SQL = sSQL
  If modTablaConsul.Ejecutar = HAY_ERROR Then End
  modTablaConsul.Cerrar
Next i

End Sub

Private Sub cmdPProyecto_Click()

If cboProyecto.Enabled Then
    ' alto
     cboProyecto.Height = CBOALTO
    ' focus a cbo
    cboProyecto.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()

' Descarga el formulario
  Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Carga los títulos del grid
aTitulosColGrid = Array("IdProy", "IdProg", "IdLinea", "IdActividad", "TipDoc", "Nro.Doc", "Fecha", "Descripción Gasto", "Monto S/.", "Orden")
'aTitulosColGrid = Array("IdProy", "IdProg", "IdLinea", "IdActividad", "Descripción Afecta", "TipDoc", "Nro.", "Fecha", "Descripción Gasto", "Monto S/.", "Orden")
'aTamañosColumnas = Array(0, 0, 0, 0, 3400, 400, 1000, 1000, 3100, 1200, 1080)
aTamañosColumnas = Array(600, 600, 650, 900, 650, 1000, 1100, 3400, 1200, 1080)
'aTamañosColumnas = Array(0, 0, 0, 0, 400, 1000, 1000, 3100, 1200, 1080)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

' Inicia alineamieto de la columna 3
grdConsulta.ColAlignment(4) = 1
    
'Se carga el combo de Proyectos
sSQL = "SELECT Numero, DescProveedor FROM PROVEEDORES WHERE RUC_DNI = 'RUC' ORDER BY DescProveedor "
CD_CargarColsCbo cboProyecto, sSQL, mcolCodProy, mcolCodDesProy

' Carga la fecha de consulta
mskFecConsulta = gsFecTrabajo

' Establece los campos obligatorios
EstableceCamposObligatorios

' Deshabilita el botón generar informe
cmdInforme.Enabled = False

End Sub

Private Sub EstableceCamposObligatorios()
' ------------------------------------------------------------
' Propósito: Muestra de color amarillo los campos obligatorios
' Recibe: Nada
' Entrega:Nada
' ------------------------------------------------------------
txtProyecto.BackColor = Obligatorio
mskFechaIni.BackColor = Obligatorio
mskFechaFin.BackColor = Obligatorio

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Cierra las colecciones
Set mcolCodProy = Nothing
Set mcolCodDesProy = Nothing

End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

' Si presiona enter entonces manda al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

' Si presiona enter entonces manda al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtProyecto_Change()
' Verifica si el código de proyecto esta en mayusculas
If UCase(txtProyecto) = txtProyecto Then
    ' Si procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboProyecto, txtProyecto, mcolCodDesProy

    ' Verifica si el campo esta vacio
    If txtProyecto.Text <> "" And cboProyecto.Text <> "" Then
        ' ok
          txtProyecto.BackColor = vbWhite
        ' Carga datos de proyecto
          ActualizarDatosProveedor
    Else
        ' Obligatorio
          txtProyecto.BackColor = Obligatorio
        ' Actualiza los datos de el proyecto ELIMINADOS
    End If
Else
   ' Cambia a mayúsculas el código del proyecto
   txtProyecto = UCase(txtProyecto)
End If

' Carga consulta
CargaConsulta

End Sub

Private Sub CargaConsulta()
' -------------------------------------------------------
' Propósito : Verifica los datos y carga la consulta
' Recibe : Nada
' Entrega : Nada
' -------------------------------------------------------
txtGastoProy.Text = "0.00"
grdConsulta.Rows = 1
' Verifica los datos introducidos para la consulta
  If fbOkDatosIntroducidos = False Then
    ' Sale de el proceso y limpia el grid
    txtGastoProy.Text = "0.00"
    grdConsulta.Rows = 1
    ' Deshabilita el botón generar informe
    cmdInforme.Enabled = False
    Exit Sub
  End If
  grdConsulta.Visible = False
' Cargar gastos totales de proyecto
   'CargaGastosProy
' Cargar gastos totales de programa
   'CargaGastosProg
' Cargar gastos totales de Lineas
   'CargaGastosLinea
' Cargar gastos totles de Actividades
   'CargaGastosActividades
' Cargar registro de gastos por Documento Pagado
   CargaRegGastos
' Carga el grid consulta
   CargarGridConsulta
  grdConsulta.Visible = True
' Deshabilita el botón generar informe
  If grdConsulta.Rows > 1 Then
    cmdInforme.Enabled = True
  Else
    cmdInforme.Enabled = False
  End If
   
End Sub

Private Sub CargarGridConsulta()
' ----------------------------------------------------
' Propósito: Arma la consulta en el grid
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim bRecorreProg As Boolean
Dim bRecorreLineas As Boolean
Dim bRecorreActiv As Boolean
Dim bRecorreRegGastos As Boolean
Dim dblGastoProy As Double
Dim sDescGasto As String
Dim ValorIGVRecu As Double
Dim ValorDetalleConIGV As Double
Dim sSQL As String
Dim mcurProyectos As New clsBD2

' Inicializa la variable
dblGastoProy = 0
' Recorre el cursor gastos programas

'("IdProy", "IdProg", "IdLinea", "IdActividad", "TipDoc", "Nro.", "Fecha", "Descripción Gasto", "Monto S/.", "Orden")
'E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,D.Abreviatura,E.NumDoc,E.FecMov,PR.DescProd,SV.DescServ,G.Monto,E.Orden
Do While Not mcurRegGastos.EOF
  ' Añade el elemento al grid
  sDescGasto = Empty
  If Not IsNull(mcurRegGastos.campo(7)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(7)
  If Not IsNull(mcurRegGastos.campo(8)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(8)
    
    sSQL = ""
    sSQL = "SELECT Tipo " & _
           "FROM Proyectos WHERE idproy = '" & mcurRegGastos.campo(0) & "' "
    
    ' ejecuta la sentencia
    mcurProyectos.SQL = sSQL
    If mcurProyectos.Abrir = HAY_ERROR Then End
    
    If Not mcurProyectos.EOF Then
      If mcurProyectos.campo(0) = "PROY" Then
        TipoEgreso = "PROY"
      ElseIf mcurProyectos.campo(0) = "EMPR" Then
        TipoEgreso = "EMPR"
      End If
    End If
    
    mcurProyectos.Cerrar
     
'    If Val(mcurRegGastos.campo(9)) <> Val(mcurRegGastos.campo(12)) Then  ' EMPRESA
    If TipoEgreso = "EMPR" Then ' EMPRESA
      If mcurRegGastos.campo(4) = "F" Or mcurRegGastos.campo(4) = "LC" Or mcurRegGastos.campo(4) = "BT" Or mcurRegGastos.campo(4) = "CART" Or mcurRegGastos.campo(4) = "TIC" Or mcurRegGastos.campo(4) = "SP" Then ' EMPRESA
        ValorIGVRecu = Format((Val(mcurRegGastos.campo(9)) - Val(mcurRegGastos.campo(12))) / Val(mcurRegGastos.campo(12)), "###,###,##0.00")
        ValorDetalleConIGV = Format(Val(mcurRegGastos.campo(11)) * (1 + ValorIGVRecu), "###,###,##0.00")
        ' Añade el registro de gastos
        grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & mcurRegGastos.campo(3) & _
            vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
            vbTab & sDescGasto & vbTab & Format(ValorDetalleConIGV, "###,###,##0.00") & vbTab & mcurRegGastos.campo(10)
          
        'grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
                  vbTab & mcurRegGastos.campo(7) & vbTab & sDescGasto & vbTab & Format(ValorDetalleConIGV, "###,###,##0.00") & vbTab & mcurRegGastos.campo(3)
      
        ' Acumula en la variable
        dblGastoProy = dblGastoProy + ValorDetalleConIGV
      Else
        ValorDetalleConIGV = Format(mcurRegGastos.campo(11), "###,###,##0.00")
        ' Añade el registro de gastos
        grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & mcurRegGastos.campo(3) & _
            vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
            vbTab & sDescGasto & vbTab & Format(ValorDetalleConIGV, "###,###,##0.00") & vbTab & mcurRegGastos.campo(10)
          
        'grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
                  vbTab & mcurRegGastos.campo(7) & vbTab & sDescGasto & vbTab & Format(ValorDetalleConIGV, "###,###,##0.00") & vbTab & mcurRegGastos.campo(3)
      
        ' Acumula en la variable
        dblGastoProy = dblGastoProy + ValorDetalleConIGV
      End If
    Else
      ' Añade el registro de gastos
      grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & mcurRegGastos.campo(3) & _
          vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
          vbTab & sDescGasto & vbTab & Format(mcurRegGastos.campo(11), "###,###,##0.00") & vbTab & mcurRegGastos.campo(10)
      ' Acumula en la variable
      dblGastoProy = dblGastoProy + Val(mcurRegGastos.campo(11))
    End If
  ' siguiente registro de gastos
  mcurRegGastos.MoverSiguiente
Loop

'Do While Not mcurGastosProy.EOF
'  ' IdProg", "IdLinea", "IdActividad", "Descripción Afecta", "TipDoc", "Nro.", "Fecha", "Proveedor", "Monto S/.", "Orden"
'  ' Añade el elemento al grid gastos de programas
'  grdConsulta.AddItem mcurGastosProy.campo(0) & mcurGastosProy.campo(1) & vbTab & vbTab & vbTab & vbTab & _
'  mcurGastosProy.campo(0) & " " & mcurGastosProy.campo(2) & _
'  vbTab & vbTab & vbTab & vbTab & vbTab & Format(mcurGastosProy.campo(3), "###,###,##0.00")
'  ' Coloca color al grid
'  grdConsulta.Row = grdConsulta.Rows - 1
'  MarcarFilaGRID grdConsulta, vbWhite, vbVerdePetroleo
'
'  '///////////////////
'  ' Inicializa la variable
'  bRecorreProg = True
'  ' Recorre el cursor gastos programas
'  Do While bRecorreProg = True
'  ' Verifica si es el final del cursor
'    If mcurGastosProg.EOF Then ' Final de programas
'      bRecorreProg = False
'    Else  ' No es el final de programas
'      ' Verifica si se cambio de Proyecto
'      If mcurGastosProy.campo(0) <> mcurGastosProg.campo(0) Then
'        bRecorreProg = False
'      Else ' mismo proyecto
'        ' Añade el elemento al grid gastos programa
'        grdConsulta.AddItem mcurGastosProg.campo(0) & mcurGastosProy.campo(1) & vbTab & mcurGastosProg.campo(1) & _
'        vbTab & vbTab & vbTab & mcurGastosProy.campo(0) & mcurGastosProg.campo(1) & _
'        " " & mcurGastosProg.campo(2) & vbTab & vbTab & vbTab & vbTab & vbTab & _
'        Format(mcurGastosProg.campo(3), "###,###,##0.00")
'        ' Coloca color al grid
'        grdConsulta.Row = grdConsulta.Rows - 1
'        MarcarFilaGRID grdConsulta, vbWhite, &H80000003
'        'MarcarFilaGRID grdConsulta, &H80000012, &HC0C0C0
'
'
'
'
'        ' Inicializa la variable
'        bRecorreLineas = True
'        ' Recorre el cursor gastos lineas
'        Do While bRecorreLineas = True
'          ' Verifica si es el final del cursor
'          If mcurGastosLineas.EOF Then ' Final de Lineas
'              bRecorreLineas = False
'          Else  ' No es el final de lineas
'            ' Verifica si se cambio de Programa
'            If mcurGastosProy.campo(0) <> mcurGastosLineas.campo(0) Or _
'                mcurGastosProg.campo(1) <> mcurGastosLineas.campo(1) Then
'                bRecorreLineas = False
'            Else ' mismo programa
'              ' Añade el elemento al grid gastos lineas
'              grdConsulta.AddItem mcurGastosLineas.campo(0) & mcurGastosProy.campo(1) & vbTab & mcurGastosLineas.campo(1) & vbTab & mcurGastosLineas.campo(2) & _
'              vbTab & vbTab & mcurGastosProy.campo(0) & mcurGastosLineas.campo(1) & mcurGastosLineas.campo(2) & _
'              " " & mcurGastosLineas.campo(3) & vbTab & vbTab & vbTab & vbTab & vbTab & _
'              Format(mcurGastosLineas.campo(4), "###,###,##0.00")
'              ' Coloca color al grid
'              grdConsulta.Row = grdConsulta.Rows - 1
'              MarcarFilaGRID grdConsulta, &H80000012, &HC0C0C0
'
'
'
'              ' Inicializa la variable
'              bRecorreActiv = True
'              ' Recorre el cursor gastos actividades
'              Do While bRecorreActiv = True
'                ' Verifica el final de cursor
'                If mcurGastosActiv.EOF Then ' final Actividades
'                    bRecorreActiv = False
'                Else ' no es el final de actividades
'                  ' Verifica si se cambio de actividad
'                  If mcurGastosProy.campo(0) <> mcurGastosActiv.campo(0) Or _
'                      mcurGastosProg.campo(1) <> mcurGastosActiv.campo(1) Or _
'                      mcurGastosLineas.campo(2) <> mcurGastosActiv.campo(2) Then
'                       bRecorreActiv = False
'                  Else ' misma linea
'                    ' Añade un elemento al grid gastos actividades
'                    grdConsulta.AddItem mcurGastosActiv.campo(0) & mcurGastosProy.campo(1) & vbTab & mcurGastosActiv.campo(1) & vbTab _
'                    & mcurGastosActiv.campo(2) & vbTab & mcurGastosActiv.campo(3) & vbTab & mcurGastosProy.campo(0) & mcurGastosActiv.campo(1) _
'                    & mcurGastosActiv.campo(2) & mcurGastosActiv.campo(3) & " " & mcurGastosActiv.campo(4) & vbTab & vbTab & vbTab & vbTab & _
'                    vbTab & Format(mcurGastosActiv.campo(5), "###,###,##0.00")
'                    ' Coloca color al grid
'                    grdConsulta.Row = grdConsulta.Rows - 1
'                    MarcarFilaGRID grdConsulta, &H80000012, &HE0E0E0
'
'
'
'
'                    ' Incializa la variable
'                    bRecorreRegGastos = True
'                    ' Recorre el cursor registro de gastos
'                    Do While bRecorreRegGastos = True
'                      ' Verifica si se es el final del registro de gastos
'                      If mcurRegGastos.EOF Then
'                        bRecorreRegGastos = False
'                      Else ' No es final de registro de gastos
'                        ' Verifica si se cambio de actividad
'                        If mcurGastosProy.campo(0) <> mcurRegGastos.campo(0) Or _
'                            mcurGastosProg.campo(1) <> mcurRegGastos.campo(1) Or _
'                            mcurGastosLineas.campo(2) <> mcurRegGastos.campo(2) Or _
'                            mcurGastosActiv.campo(3) <> mcurRegGastos.campo(3) Then
'                           bRecorreRegGastos = False
'                        Else ' misma actividad
'                          ' grid IdProg", "IdLinea", "IdActividad", "Descripción Afecta", "TipDoc", "Nro.", "Fecha", "Proveedor","Desc Gasto", "Monto S/.", "Orden"
'                          ' cursor E.IdProg,E.IdLinea,E.IdActiv,E.Orden,D.DescTipoDoc,E.NumDoc,E.FecMov,P.DescProveedor,DescProd,DescServ,G.Monto"
'                          sDescGasto = Empty
'                          If Not IsNull(mcurRegGastos.campo(8)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(8)
'                          If Not IsNull(mcurRegGastos.campo(9)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(9)
'                          ' Añade el registro de gastos
'                          grdConsulta.AddItem mcurRegGastos.campo(0) & mcurGastosProy.campo(1) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & mcurRegGastos.campo(3) & vbTab & vbTab & mcurRegGastos.campo(5) & vbTab & mcurRegGastos.campo(6) & vbTab & FechaDMA(mcurRegGastos.campo(7)) & _
'                                      vbTab & sDescGasto & vbTab & Format(mcurRegGastos.campo(10), "###,###,##0.00") & vbTab & mcurRegGastos.campo(4)
'                          'grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
'                                      vbTab & mcurRegGastos.campo(7) & vbTab & sDescGasto & vbTab & Format(mcurRegGastos.campo(10), "###,###,##0.00") & vbTab & mcurRegGastos.campo(3)
'
'                          ' Acumula en la variable
'                          dblGastoProy = dblGastoProy + Val(mcurRegGastos.campo(10))
'                          ' siguiente registro de gastos
'                          mcurRegGastos.MoverSiguiente
'                        End If
'                      End If
'                    Loop ' Siguiente reg gastos
'                    ' Siguiente actividad
'                    mcurGastosActiv.MoverSiguiente
'                  End If
'                End If
'              Loop
'              ' Siguiente linea
'              mcurGastosLineas.MoverSiguiente
'            End If
'          End If
'        Loop ' Repetir lineas
'        ' Mueve al siguiente programa
'        mcurGastosProg.MoverSiguiente
'      End If
'    End If
'  Loop
'  ' Mueve al siguiente proyecto
'  mcurGastosProy.MoverSiguiente
'Loop






'' Recorre el cursor gastos programas
'Do While Not mcurGastosProg.EOF
'    ' IdProg", "IdLinea", "IdActividad", "Descripción Afecta", "TipDoc", "Nro.", "Fecha", "Proveedor", "Monto S/.", "Orden"
'    ' Añade el elemento al grid gastos de programas
'    grdConsulta.AddItem mcurGastosProg.campo(0) & vbTab & vbTab & vbTab & _
'    txtProyecto & mcurGastosProg.campo(0) & " " & mcurGastosProg.campo(1) & _
'    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & Format(mcurGastosProg.campo(2), "###,###,##0.00")
'    ' Coloca color al grid
'    grdConsulta.Row = grdConsulta.Rows - 1
'    MarcarFilaGRID grdConsulta, vbWhite, &H80000003
'    ' Inicializa la variable
'    bRecorreLineas = True
'    ' Recorre el cursor gastos lineas
'    Do While bRecorreLineas = True
'        ' Verifica si es el final del cursor
'        If mcurGastosLineas.EOF Then ' Final de Lineas
'            bRecorreLineas = False
'        Else  ' No es el final de lineas
'            ' Verifica si se cambio de Programa
'            If mcurGastosProg.campo(0) <> mcurGastosLineas.campo(0) Then
'                bRecorreLineas = False
'            Else ' mismo programa
'                ' Añade el elemento al grid gastos lineas
'                grdConsulta.AddItem mcurGastosLineas.campo(0) & vbTab & mcurGastosLineas.campo(1) & _
'                vbTab & vbTab & txtProyecto & mcurGastosLineas.campo(0) & mcurGastosLineas.campo(1) & _
'                " " & mcurGastosLineas.campo(2) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & _
'                Format(mcurGastosLineas.campo(3), "###,###,##0.00")
'                ' Coloca color al grid
'                grdConsulta.Row = grdConsulta.Rows - 1
'                MarcarFilaGRID grdConsulta, &H80000012, &HC0C0C0
'                ' Inicializa la variable
'                bRecorreActiv = True
'                ' Recorre el cursor gastos actividades
'                Do While bRecorreActiv = True
'                    ' Verifica el final de cursor
'                    If mcurGastosActiv.EOF Then ' final Actividades
'                        bRecorreActiv = False
'                    Else ' no es el final de actividades
'                        ' Verifica si se cambio de actividad
'                        If mcurGastosProg.campo(0) <> mcurGastosActiv.campo(0) Or _
'                           mcurGastosLineas.campo(1) <> mcurGastosActiv.campo(1) Then
'                             bRecorreActiv = False
'                        Else ' misma linea
'                            ' Añade un elemento al grid gastos actividades
'                            grdConsulta.AddItem mcurGastosActiv.campo(0) & vbTab & mcurGastosActiv.campo(1) & vbTab _
'                            & mcurGastosActiv.campo(2) & vbTab & txtProyecto & mcurGastosActiv.campo(0) & mcurGastosActiv.campo(1) _
'                            & mcurGastosActiv.campo(2) & " " & mcurGastosActiv.campo(3) & vbTab & vbTab & vbTab & vbTab & vbTab & _
'                            vbTab & Format(mcurGastosActiv.campo(4), "###,###,##0.00")
'                            ' Coloca color al grid
'                            grdConsulta.Row = grdConsulta.Rows - 1
'                            MarcarFilaGRID grdConsulta, &H80000012, &HE0E0E0
'                            ' Incializa la variable
'                            bRecorreRegGastos = True
'                            ' Recorre el cursor registro de gastos
'                            Do While bRecorreRegGastos = True
'                               ' Verifica si se es el final del registro de gastos
'                               If mcurRegGastos.EOF Then
'                                  bRecorreRegGastos = False
'                               Else ' No es final de registro de gastos
'                                  ' Verifica si se cambio de actividad
'                                  If mcurGastosProg.campo(0) <> mcurRegGastos.campo(0) Or _
'                                     mcurGastosLineas.campo(1) <> mcurRegGastos.campo(1) Or _
'                                     mcurGastosActiv.campo(2) <> mcurRegGastos.campo(2) Then
'                                     bRecorreRegGastos = False
'                                  Else ' misma actividad
'                                    ' grid IdProg", "IdLinea", "IdActividad", "Descripción Afecta", "TipDoc", "Nro.", "Fecha", "Proveedor","Desc Gasto", "Monto S/.", "Orden"
'                                    ' cursor E.IdProg,E.IdLinea,E.IdActiv,E.Orden,D.DescTipoDoc,E.NumDoc,E.FecMov,P.DescProveedor,DescProd,DescServ,G.Monto"
'                                    sDescGasto = Empty
'                                    If Not IsNull(mcurRegGastos.campo(8)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(8)
'                                    If Not IsNull(mcurRegGastos.campo(9)) Then sDescGasto = sDescGasto + mcurRegGastos.campo(9)
'                                    ' Añade el registro de gastos
'                                    grdConsulta.AddItem mcurRegGastos.campo(0) & vbTab & mcurRegGastos.campo(1) & vbTab & mcurRegGastos.campo(2) & vbTab & vbTab & mcurRegGastos.campo(4) & vbTab & mcurRegGastos.campo(5) & vbTab & FechaDMA(mcurRegGastos.campo(6)) & _
'                                                vbTab & mcurRegGastos.campo(7) & vbTab & sDescGasto & vbTab & Format(mcurRegGastos.campo(10), "###,###,##0.00") & vbTab & mcurRegGastos.campo(3)
'                                    ' Acumula en la variable
'                                    'dblGastoProy = dblGastoProy + Val(mcurRegGastos.campo(10))
'                                    ' siguiente registro de gastos
'                                    mcurRegGastos.MoverSiguiente
'                                  End If
'                               End If
'                            Loop ' Siguiente reg gastos
'                            ' Siguiente actividad
'                            mcurGastosActiv.MoverSiguiente
'                        End If
'                    End If
'                Loop
'                ' Siguiente linea
'                mcurGastosLineas.MoverSiguiente
'            End If
'        End If
'
'    Loop ' Repetir lineas
'   ' Mueve al siguiente programa
'   mcurGastosProg.MoverSiguiente
'Loop

' Muestra el total de los proyectos
txtGastoProy = Format(dblGastoProy, "###,###,##0.00")

' Cierra el cursor
'mcurGastosProg.Cerrar
'mcurGastosLineas.Cerrar
'mcurGastosActiv.Cerrar
mcurRegGastos.Cerrar

End Sub

Private Sub CargaRegGastos()
' ----------------------------------------------------
' Propósito: Carga el registro de los gastos por documentos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String
'("IdProy", "IdProg", "IdLinea", "IdActividad", "TipDoc", "Nro.", "Fecha", "Descripción Gasto", "Monto S/.", "Orden")
' Carga la sentencia
sSQL = "SELECT E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,D.Abreviatura," _
     & "E.NumDoc,E.FecDoc,PR.DescProd,SV.DescServ, E.MontoAfectado, E.Orden, G.Monto, E.MONTOCB " _
     & "FROM ((EGRESOS E INNER  JOIN TIPO_DOCUM D ON E.IdTipoDoc=D.IdTipoDoc) " _
     & "INNER JOIN PROVEEDORES P ON E.IdProveedor=P.IdProveedor) " _
     & "INNER JOIN ((GASTOS G LEFT JOIN PRODUCTOS PR ON G.CodConcepto=PR.IdProd) " _
     & "LEFT JOIN SERVICIOS SV ON G.CodConcepto=SV.IdServ) ON E.Orden=G.Orden " _
     & "WHERE E.idProveedor='" & IdProveedor & "' and " _
     & "E.Anulado = 'NO' and E.FecDoc BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" _
     & FechaAMD(mskFechaFin) _
     & "' ORDER BY E.FecDoc,E.Orden,D.Abreviatura,E.NumDoc,E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,PR.DescProd "


'sSQL = "SELECT E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,E.Orden,D.Abreviatura," _
     & "E.NumDoc,E.FecMov,P.DescProveedor,PR.DescProd,SV.DescServ,G.Monto " _
     & "FROM ((EGRESOS E INNER  JOIN TIPO_DOCUM D ON E.IdTipoDoc=D.IdTipoDoc) " _
     & "INNER JOIN PROVEEDORES P ON E.IdProveedor=P.IdProveedor) " _
     & "INNER JOIN ((GASTOS G LEFT JOIN PRODUCTOS PR ON G.CodConcepto=PR.IdProd) " _
     & "LEFT JOIN SERVICIOS SV ON G.CodConcepto=SV.IdServ) ON E.Orden=G.Orden " _
     & "WHERE E.idProveedor='" & IdProveedor & "' and " _
     & "E.Anulado = 'NO' and E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" _
     & FechaAMD(mskFechaFin) _
     & "' ORDER BY E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,E.FecMov,E.Orden,PR.DescProd"
' Ejecuta la sentencia
mcurRegGastos.SQL = sSQL
If mcurRegGastos.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaGastosProy()
' ----------------------------------------------------
' Propósito: Carga los gastos acumulados para los proyectos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
'sSQL = "SELECT E.IdProg,P.DescProg, SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, PROGRAMAS P " & _
       "WHERE E.idProy='" & txtProyecto.Text & "' and " & _
       "E.IdProg=P.IdProg and E.Anulado = 'NO' and " & _
       "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProg,P.DescProg ORDER BY E.IdProg"

sSQL = "SELECT E.IdProy, P.IdFinan, F.DESCFINAN, SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, PROYECTOS P, TIPO_FINANCIERAS F " & _
       "WHERE E.idProveedor='" & IdProveedor & "' and " & _
       "E.IdProy=P.IdProy  and P.IDFINAN=F.IDFINAN and E.Anulado = 'NO' and " & _
       "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProy,P.IdFinan, F.DESCFINAN ORDER BY E.IdProy"
' Ejecuta la sentencia
mcurGastosProy.SQL = sSQL
If mcurGastosProy.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaGastosProg()
' ----------------------------------------------------
' Propósito: Carga los gastos acumulados para los programas
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
'sSQL = "SELECT E.IdProg,P.DescProg, SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, PROGRAMAS P " & _
       "WHERE E.idProy='" & txtProyecto.Text & "' and " & _
       "E.IdProg=P.IdProg and E.Anulado = 'NO' and " & _
       "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProg,P.DescProg ORDER BY E.IdProg"

sSQL = "SELECT E.IdProy,E.IdProg,P.DescProg, SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, PROGRAMAS P " & _
       "WHERE E.idProveedor='" & IdProveedor & "' and " & _
       "E.IdProg=P.IdProg and E.Anulado = 'NO' and " & _
       "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProy,E.IdProg,P.DescProg ORDER BY E.IdProy,E.IdProg"
' Ejecuta la sentencia
mcurGastosProg.SQL = sSQL
If mcurGastosProg.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaGastosLinea()
' ----------------------------------------------------
' Propósito: Carga los gastos acumulados para las lineas
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
sSQL = "SELECT E.IdProy,E.IdProg,E.IdLinea,L.DescLinea,SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, LINEAS L " & _
       "WHERE E.idProveedor ='" & IdProveedor & "' and " & _
       "E.IdLinea=L.IdLinea and E.Anulado = 'NO' and " & _
       "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProy,E.IdProg,E.IdLinea,L.DescLinea ORDER BY E.IdProy,E.IdProg,E.IdLinea"
' Ejecuta la sentencia
mcurGastosLineas.SQL = sSQL
If mcurGastosLineas.Abrir = HAY_ERROR Then End

End Sub

Private Sub CargaGastosActividades()
' ----------------------------------------------------
' Propósito: Carga los gastos acumulados para las actividades
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
Dim sSQL As String

' Carga la sentencia
sSQL = "SELECT E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,A.DescActiv, SUM(E.MontoAfectado) " _
     & "FROM EGRESOS E, ACTIVIDADES A " & _
       "WHERE idProveedor ='" & IdProveedor & "' and " & _
       "E.IdActiv=A.IdActiv and Anulado = 'NO' And " & _
       "FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' and '" & _
       FechaAMD(mskFechaFin) & _
       "' GROUP BY E.IdProy,E.IdProg,E.IdLinea,E.IdActiv,A.DescActiv ORDER BY E.IdProy,E.IdProg,E.IdLinea,E.IdActiv"
' Ejecuta la sentencia
mcurGastosActiv.SQL = sSQL
If mcurGastosActiv.Abrir = HAY_ERROR Then End

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
If txtProyecto.BackColor <> vbWhite Or _
   mskFechaIni.BackColor <> vbWhite Or _
   mskFechaFin.BackColor <> vbWhite Then
   fbOkDatosIntroducidos = False
   Exit Function
End If

' Verifica que el intervalo elegido este entre la fecha de inicio y fin del proyecto VADICK
'If fbOkintervalo(FechaAMD(mskFechaIni), FechaAMD(mskFechaFin), _
'        FechaAMD(mskFecInicioProy), _
'        Mid(CalcularPeriodo(Val(txtNumPeriodo), FechaAMD(mskFecInicioProy)), 17, 8)) = False Then
'        MsgBox "Las fechas elegidas deben estar dentro del intervalo del proyecto"
'        mskFechaIni.SetFocus
'        fbOkDatosIntroducidos = False
'        Exit Function
'End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function

Private Sub mskFechaFin_Change()

' Inicia el informe
grdConsulta.Rows = 1
txtGastoProy.Text = "0.00"
' Deshabilita el botón generar informe
cmdInforme.Enabled = False

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  ' Carga consulta
  CargaConsulta
Else ' Obligatorio
  mskFechaFin.BackColor = Obligatorio
End If

End Sub

Private Sub mskFechaIni_Change()
' Inicia el informe
grdConsulta.Rows = 1
txtGastoProy.Text = "0.00"
' Deshabilita el botón generar informe
cmdInforme.Enabled = False

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
    ' Fecha completa
    mskFechaIni.BackColor = vbWhite
    ' Carga consulta
    CargaConsulta
Else ' Obligatorio
  mskFechaIni.BackColor = Obligatorio
End If

End Sub

Public Sub ActualizarDatosProveedor()
'----------------------------------------------------------------------------
'PROPÓSITO: Actualizar los controles referentes a financiera, NroPeriodos, _
            fecha inicio de un proyecto
'Recibe:    nada
'Devuelve:  nada
'----------------------------------------------------------------------------
Dim sSQL As String
Dim curFinanPerioProy As New clsBD2

'Recupera financiera del proyecto seleccionado
sSQL = ""
    sSQL = "SELECT DESCPROVEEDOR, DIREC_PROVEEDOR, TEL_PROVEEDOR, IDPROVEEDOR " & _
           "FROM PROVEEDORES " & _
           "WHERE NUMERO =" & "'" & txtProyecto.Text & "'"
       
curFinanPerioProy.SQL = sSQL

' ejecuta la consulta y asignamos al txt de proyecto
If curFinanPerioProy.Abrir = HAY_ERROR Then
  Unload Me
  End
End If

'Carga las variables del modulo
IdProveedor = curFinanPerioProy.campo(3)
If IsNull(curFinanPerioProy.campo(1)) Then

Else
  TxDireccion.Text = curFinanPerioProy.campo(1)
End If

If IsNull(curFinanPerioProy.campo(2)) Then

Else
  TxTelefono.Text = curFinanPerioProy.campo(2)
End If
'mskFecInicioProy = FechaDMA(curFinanPerioProy.campo(2))

curFinanPerioProy.Cerrar 'Cierra el cursor

End Sub


Private Sub txtProyecto_KeyPress(KeyAscii As Integer)

' Si presiona enter entonces manda al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
