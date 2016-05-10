VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAlSelSalida 
   Caption         =   "Almacén- Selección de Salidas"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   HelpContextID   =   89
   Icon            =   "SCALSelSalida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   6015
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   1485
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
         Left            =   4440
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha &Fin:"
         Height          =   195
         Left            =   3570
         TabIndex        =   8
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha &Inicio:"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   105
      TabIndex        =   5
      Top             =   750
      Width           =   10080
      Begin MSFlexGridLib.MSFlexGrid grdSalida 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         HighLight       =   0
         FillStyle       =   1
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   3960
      Width           =   1000
   End
End
Attribute VB_Name = "frmAlSelSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' Colecciones para la carga productos de almacen que se desea anular
'Private mcolCodCantidadAlmacen As New Collection

'' Colecciones para la carga productos de almacen que se desea anular
'Private mcolAlmacenAnular As New Collection
'
'' Colecciones para la carga productos de almacen que se desea actualizar almacen
'Private mcolCantidadProducto As New Collection
'
'' Colección para la carga de productos de Almacén posteriores a la anulación
'Private mcolSalidasPosteriores As New Collection
'
'' Cursor para determinar las cantidades de los productos a repartir posteriores a la
'' Anulación de la salida de almacén
'Dim mcurRepartir As New clsBD2

' Variable de salida de almacén para la anulación
'Private msSalida As String

'' Variable que almacena el signo
'Private msSigno As String

' Variable de modulo para el manejo del grid
Dim ipos As Long

Private Sub cmdAceptar_Click()
'Se comprueba que se haya marcado algúna fila
If grdSalida.Row < 1 Then
  MsgBox "Debe marcar algún documento", vbInformation + vbOKOnly, "SGCcaijo-Selección de registros"
  Exit Sub
End If
    
'Carga el txtIdSalida con datos de la salida a modificar
frmALSalida.txtIdSalida = grdSalida.TextMatrix(grdSalida.Row, 0)
frmALSalida.mskFecha = grdSalida.TextMatrix(grdSalida.Row, 3)

'Carga con los datos de la salida
frmALSalida.CargarRegSalidaAlmacen
    
'Muestra el formulario frmAlSalida
frmALSalida.Show vbModal, Me

'Desabilita los controles
cmdAceptar.Enabled = False

'Actualiza los datos del grid
CargarGrid

End Sub



Private Sub cmdSalir_Click()

'Termina la ejecucion del formulario
Unload Me

End Sub

Private Sub Form_Load()

'Coloa a obligatorio
mskFechaIni.BackColor = Obligatorio

'Carga la fecha del sistema
mskFechaFin.Text = gsFecTrabajo
      
'Desabilita los controles
cmdAceptar.Enabled = False

' Inicializa el grid
ipos = 0
gbCambioCelda = False

End Sub

Private Sub grdSalida_Click()

If grdSalida.Row > 0 And grdSalida.Row < grdSalida.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdSalida, vbWhite, vbDarkBlue
    ' Habilita aceptar, Anular
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub grdSalida_DblClick()

'Hace llamado al evento click del aceptar
cmdAceptar_Click

End Sub

Private Sub grdSalida_EnterCell()

If ipos <> grdSalida.Row Then
    '  Verifica si es la última fila
    If grdSalida.Row > 0 And grdSalida.Row < grdSalida.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdSalida, ipos
            gbCambioCelda = False
            ' Habilita aceptar
            cmdAceptar.Enabled = True
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdSalida.Row
End If

End Sub

Private Sub grdSalida_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    SendKeys vbTab
 End If
 
End Sub

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
    mskFechaFin.BackColor = vbWhite
    
    'Carga el grid
    CargarGrid
  
Else
  ' Maneja controles de selección
  mskFechaFin.BackColor = Obligatorio
  grdSalida.Rows = 1
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub CargarGrid()
'------------------------------------------------------
'Propósito  : Cargar el grid entre las fechas seleccionadas
'             un registro
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------------------
' Nota el codigo de salida de Almacén es "99999"
Dim sSQL As String

If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
  
    'Verifica si la fecha inicio es anterior a la fecha fin
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = False Then
    
        'Limpia el grid grdSalida
        grdSalida.Rows = 1
        
        'Se seleccionan las salidas de almacén
        sSQL = "SELECT S.IdSalida, Y.DescProy, ( p.Apellidos & ', ' & P.Nombre), S.Fecha " & _
               "FROM PROYECTOS Y, PLN_Personal P,Almacen_Salidas S " & _
               "WHERE S.IdProy=Y.IdProy and S.IdPersona=P.IdPersona and S.Anulado='No' And S.Fecha BETWEEN " & _
               "'" & FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
               "ORDER BY  S.IdSalida"
        
        ' Se carga un array con los títulos de las columnas y otro con los tamaños para
        'pasárselos a la función que carga el grid
        aTitulosColGrid = Array("Nro Salida", "Proyecto", "Persona", "Fecha")
        aTamañosColumnas = Array(1100, 4000, 3300, 1000)
        aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Fecha")
         
        CargarGridConFormatos grdSalida, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
         
        If grdSalida.Rows = 1 Then
           MsgBox "No existen egreso de Almacén entre estas Fechas", _
                   vbInformation + vbOKOnly, "S.G.Ccaijo"
    
        End If
    Else
        'Limpia el grid
        grdSalida.Rows = 1
        MsgBox "Fecha inicio es posterior a fecha fin", vbInformation + vbOKOnly, " Selección de Salidas de Almacen"
    End If
Else
  ' Maneja controles de selección
  grdSalida.Rows = 1
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

' Valida  que la fecha Inicio de la consulta
If ValidarFecha(mskFechaIni) Then
  mskFechaIni.BackColor = vbWhite
  
  'Carga el grid
  CargarGrid
  
Else
  ' Maneja controles de selección
  mskFechaIni.BackColor = Obligatorio
  grdSalida.Rows = 1
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
