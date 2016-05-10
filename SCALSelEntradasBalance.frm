VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmALSelEntradaBalance 
   Caption         =   "SGCcaijo - Entradas de almacén por balance"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   Icon            =   "SCALSelEntradasBalance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1200
      TabIndex        =   6
      Top             =   0
      Width           =   5175
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   1245
         TabIndex        =   0
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
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha &Fin:"
         Height          =   330
         Left            =   2850
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblFecInicio 
         Caption         =   "Fecha &Inicio:"
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   5160
      TabIndex        =   3
      Top             =   4230
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   6360
      TabIndex        =   4
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Frame frDocVerificar 
      Caption         =   "Seleccione la entrada a almacén por balance "
      Height          =   3435
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   7545
      Begin MSFlexGridLib.MSFlexGrid grdIngreso 
         Height          =   3015
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   1
         Cols            =   3
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
End
Attribute VB_Name = "frmALSelEntradaBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variable de modulo para el manejo del grid
Dim ipos As Long

Private Sub cmdAceptar_Click()

'Se comprueba que se haya marcado algúna fila
If grdIngreso.Row < 1 Then
  MsgBox "Debe marcar algún documento", vbInformation + vbOKOnly, "SGCcaijo-Selección de registros"
  Exit Sub
End If

' Muestra el formulario frmAlSalida
frmALEntrBalance.txtCodigo = grdIngreso.TextMatrix(grdIngreso.Row, 0)

' Sale del formulario
Unload Me

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

'Se carga un array con los títulos de las columnas y otro con los tamaños para
'pasárselos a la función que carga el grid
aTitulosColGrid = Array("Código", "Fecha de Balance", "Número Doc.")
aTamañosColumnas = Array(1050, 1200, 4600)
CargarGridTitulos grdIngreso, aTitulosColGrid, aTamañosColumnas

' Inicializa el grid
ipos = 0
gbCambioCelda = False
grdIngreso.ColAlignment(2) = 1

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
            ' Habilita aceptar
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

Private Sub CargarIngresosALBalance()
Dim sSQL As String
 
 'Limpia el grid grdIngreso, inicializa la variable intervalo
 grdIngreso.Rows = 1
  
 'Se seleccionan los Documentos de Almacen que no hayan sido verificados
 sSQL = "SELECT  IdBalance,Fecha,NumDoc " _
        & "FROM ALMACEN_BALANCE " _
        & "WHERE Fecha BETWEEN '" & FechaAMD(mskFechaIni) _
        & "' AND '" & FechaAMD(mskFechaFin) & "' AND Anulado='NO' " _
        & "ORDER BY IdBalance"

 ' Se carga un array con los títulos de las columnas y otro con los tamaños para
 'pasárselos a la función que carga el grid
 aTitulosColGrid = Array("IdBalance", "Fecha Balance", "Número Doc.")
 aTamañosColumnas = Array(1050, 1200, 4600)
 aFormatos = Array("fmt_Normal", "fmt_Fecha", "fmt_Normal")
 
 CargarGridConFormatos grdIngreso, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
 
 If grdIngreso.Rows = 1 Then ' no hay registros en la consulta
   MsgBox "No existen Ingresos a almacén por balance", _
           vbInformation + vbOKOnly, "S.G.Ccaijo, Almacén Ingresos por balance"
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
        ' Carga los Productos Verificados
        CargarIngresosALBalance
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
' Se valida que sla fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
    mskFechaIni.BackColor = vbWhite
    'Carga el grid
    CargarGrid
Else
  ' Maneja controles de la selección
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
