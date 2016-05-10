VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmALSelEntrada 
   Caption         =   "Almacén- Selección de Entradas a verificar"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   HelpContextID   =   83
   Icon            =   "SCALSelEntrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frDocVerificar 
      Caption         =   "Seleccione los documentos para verificar en almacén"
      Height          =   4440
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   11910
      Begin MSFlexGridLib.MSFlexGrid grdIngreso 
         Height          =   4095
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   1
         Cols            =   6
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
      Left            =   11040
      TabIndex        =   4
      Top             =   5205
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   9840
      TabIndex        =   3
      Top             =   5205
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   3495
      TabIndex        =   6
      Top             =   -10
      Width           =   5175
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   1245
         TabIndex        =   0
         Top             =   180
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
         Left            =   3720
         TabIndex        =   1
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecInicio 
         Caption         =   "Fecha &Inicio:"
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha &Fin:"
         Height          =   330
         Left            =   2850
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmALSelEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ipos As Long

Private Sub cmdAceptar_Click()

'Se comprueba que se haya marcado algún salida de Almacen
If grdIngreso.Row < 1 Then
  MsgBox "Debe seleccionar algún documento", vbInformation + vbOKOnly, "SGCcaijo-Selección de Registros"
  Exit Sub
End If

'Muestra el formulario frmAlSalida
frmALEntrVerif.Show vbModal, Me

'Deshabilita el botón aceptar
cmdAceptar.Enabled = False

'Actualiza los datos del grid
If gsTipoOperacionAlmacen = "Nuevo" Then
 CargarVerificados "NO"
ElseIf gsTipoOperacionAlmacen = "Modificar" Then
 CargarVerificados "SI"
End If

End Sub

Private Sub cmdSalir_Click()

'Termina la ejecucion del formulario
Unload Me

End Sub

Private Sub Form_Load()

'Coloca a obligatorio el mskFecIni
mskFechaIni.BackColor = Obligatorio

'Carga la fecha del sistema
mskFechaFin.Text = gsFecTrabajo

If gsTipoOperacionAlmacen = "Nuevo" Then
    'Oculta los controles fecha inicio y Fin
    lblFecInicio.Visible = False: lblFecFin.Visible = False
    mskFechaIni.Visible = False: mskFechaFin.Visible = False
    Frame1.Visible = False
    'Carga los Documentos no verificados en almacén, para verificarlos
    CargarVerificados "NO"
ElseIf gsTipoOperacionAlmacen = "Modificar" Then
    'Cambia el título al formulario
    frmALSelEntrada.Caption = "Almacén- Selección de Entradas a modificar"
End If

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

Private Sub CargarVerificados(sSINO As String)
Dim sSQL As String
Dim sIntervalo As String
 'Limpia el grid grdSalida, inicializa la variable intervalo
 grdIngreso.Rows = 1
 sIntervalo = Empty
 ' Verifica el tipo de operación para seleccionar los registros
 If gsTipoOperacionAlmacen = "Modificar" Then
  sIntervalo = "And E.FecMov BETWEEN '" & FechaAMD(mskFechaIni.Text) & _
           "' And '" & FechaAMD(mskFechaFin.Text) & "' "
 End If
 'Se seleccionan los Documentos de Almacen que no hayan sido verificados
 sSQL = "SELECT DISTINCT E.Orden,P.DescProy, E.NumDoc, TD.DescTipoDoc, PV.DescProveedor , E.FecMov " & _
        "FROM ALMACEN_VERIFICACION A, EGRESOS E ,PROYECTOS P, TIPO_DOCUM TD, PROVEEDORES PV " & _
        "WHERE A.Verificado='" & sSINO & "' and A.Orden=E.Orden and " _
        & "E.IdProy=P.IdProy and E.IdTipoDoc=TD.IdTipoDoc and E.IdProveedor=PV.IdProveedor " _
        & sIntervalo _
        & "ORDER BY  E.FecMov, E.Orden"

 ' Se carga un array con los títulos de las columnas y otro con los tamaños para
 'pasárselos a la función que carga el grid
 aTitulosColGrid = Array("Orden", "Proyecto", "Nº Documento", "Tipo Documento", "Proveedor", "Fecha Mov")
 aTamañosColumnas = Array(1050, 3600, 1200, 1600, 3100, 1000)
 aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Fecha")
 
 CargarGridConFormatos grdIngreso, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
 
 If grdIngreso.Rows = 1 Then ' no hay registros en la consulta
  If gsTipoOperacionAlmacen = "Nuevo" Then
   'Mensaje de No existen registros que mostrar
   MsgBox "No existen documentos a verificar en almacén", _
           vbInformation + vbOKOnly, "S.G.Ccaijo, Almacén verificación"
  Else
   'Mensaje de No existen registros que mostrar
   MsgBox "No existen documentos verificados en almacén en este lapso", _
           vbInformation + vbOKOnly, "S.G.Ccaijo, Almacén verificación"
  End If
 End If

End Sub

Private Sub mskFechaFin_Change()
    
' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  'Carga el grid
  CargarGrid
Else
  ' Maneja controles de la selección
  mskFechaFin.BackColor = Obligatorio
  grdIngreso.Rows = 1
  cmdAceptar.Enabled = False
End If


End Sub

Private Sub CargarGrid()
Dim sSQL As String

If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
    
    'Verifica si la fecha inicio es anterior a la fecha fin
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = False Then
        'Limpia el grid
        grdIngreso.Rows = 1
        ' Carga los Productos Verificados
        CargarVerificados "SI"
    Else
        'Limpia el grid
        grdIngreso.Rows = 1
        MsgBox "Fecha inicio es posterior a fecha fin", vbInformation + vbOKOnly, " Selección de Salidas de Almacen"
    End If
Else
    'Limpia el grid
    grdIngreso.Rows = 1
    'Deshabilita el botón aceptar
    cmdAceptar.Enabled = False
End If
End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub mskFechaIni_Change()
' Se svalida que la fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
    mskFechaIni.BackColor = vbWhite
    'Carga el grid
    CargarGrid
Else
  ' Maneja controles de la seleción
  mskFechaIni.BackColor = Obligatorio
  grdIngreso.Rows = 1
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)
  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
