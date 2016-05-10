VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBEGPago_Prestamos 
   Caption         =   "Entrega de prestamos al personal"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "SCCBEGPagoPrestamos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   7440
      Picture         =   "SCCBEGPagoPrestamos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtDesc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   5640
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdPrestamos 
      Height          =   2175
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.TextBox txtPersonal 
      Height          =   315
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   7815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Personal:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmCBEGPago_Prestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim bSeleccionado As Boolean
Dim i As Integer

'Se comprueba que se haya marcado algún salida de Almacen
If ComprobarFilaMarcada(grdPrestamos) = False Then
  MsgBox "No se ha seleccionado ningún préstamo", _
          vbInformation + vbOKOnly, "S.G.Ccaijo-Modificación Préstamos"
  Exit Sub
End If

' Verifica si se puede dar adelantoss
If fbOkPrestamos = False Then Exit Sub

' Recorre el grid para determinar que fila esta seleccionada
 bSeleccionado = False
 i = 1
 Do While i <= grdPrestamos.Rows - 1 And bSeleccionado = False
  grdPrestamos.Row = i
  If grdPrestamos.CellBackColor = vbDarkBlue Then
    bSeleccionado = True
  End If
  i = i + 1
Loop

'carga la Colección "Prestamo", "Descripción", "Numero", "Monto", "CtaCtb")
 gcolDetMovCB.Add Item:=grdPrestamos.TextMatrix(grdPrestamos.Row, 0) & "¯" _
                       & grdPrestamos.TextMatrix(grdPrestamos.Row, 2) & "¯" _
                       & grdPrestamos.TextMatrix(grdPrestamos.Row, 4) & "¯" _
                       & Var37(grdPrestamos.TextMatrix(grdPrestamos.Row, 3)), _
                    Key:=grdPrestamos.TextMatrix(grdPrestamos.Row, 0) & "¯" _
                       & grdPrestamos.TextMatrix(grdPrestamos.Row, 2)

' Pasa la persona elegida y el monto a frmCBEgreSinAfecta
 frmCBEGSinAfecta.txtDesc.Text = txtDesc.Text
 frmCBEGSinAfecta.txtAfecta.MaxLength = Len(txtPersonal)
 frmCBEGSinAfecta.txtAfecta.Text = txtPersonal
 frmCBEGSinAfecta.txtMonto.Text = grdPrestamos.TextMatrix(grdPrestamos.Row, 3)
 
' Cierra el formulario
  Unload Me
  
End Sub

Private Function fbOkPrestamos() As Boolean
'---------------------------------------------------------------------
' Propósito : Verifica si se ha contabilizado alguna planilla _
              referente a las cuotas del prestamo definido
' Recibe  :  Nada
' Entrega : Nada
'-------------------------------------------------------------------
Dim curPlanilla As New clsBD2
Dim sSQL As String
'"Prestamo", "Descripción", "Numero", "Monto", "CtaCtb")
' Carga la sentencia
sSQL = "SELECT P.AnioMes FROM PRESTAMOS_CUOTAS P, PLN_CTB_TOTALES PC" _
     & " WHERE P.IdPersona='" & txtPersonal _
     & "' and P.IdConPL='" & grdPrestamos.TextMatrix(grdPrestamos.Row, 0) _
     & "' and P.NumPrestamo='" & grdPrestamos.TextMatrix(grdPrestamos.Row, 2) _
     & "' and P.AnioMes=PC.CodPlanilla"
     
' Ejecuta la sentencia
curPlanilla.SQL = sSQL
If curPlanilla.Abrir = HAY_ERROR Then End

' Verifica si es vacío
If curPlanilla.EOF Then
    ' Devuelve la función
    fbOkPrestamos = True
Else
    ' Mensaje
    MsgBox "No se puede dar el Prestamo, Alguna cuota del prestamo definido pertenece a planillas procesadas. " & Chr(13) _
    & "Se debe definir nuevamente el prestamo en Planillas", , "SGCcaijo-Pago de prestamos"
    ' Devuelve la función
    fbOkPrestamos = False
End If

End Function

Private Sub cmdBuscar_Click()
'Determina si tiene prestamos el personal
If gcolTabla.Count = 0 Then
    'Mensaje de nos hay prestamos al personal
    MsgBox "No hay prestamos definidos al personal", vbOKOnly + vbInformation, "SGCcaijo-Egresos sin Afectación"
    'Decarga el formulario
    Exit Sub
End If
' Carga los títulos del grid selección
  giNroColMNSel = 4
  aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Condición", "Activo")
  aTamañosColumnas = Array(1000, 4500, 1500, 600)
' Muestra el formulario de busqueda
  frmMNSeleccion.Show vbModal, Me
  
  ' Verifica si se eligió algun dato a modificar
    If gsCodigoMant <> Empty Then
      txtPersonal.Text = gsCodigoMant
      SendKeys "{tab}"
    Else ' No se eligió nada a modificar
      ' Verifica si txtcodigo es habilitado
      If txtPersonal.Enabled = True Then txtPersonal.SetFocus
    End If
    
End Sub

Private Sub cmdSalir_Click()

' Sale del formulario
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Carga las colecciones con el personal que tiene prestamos no entregados
  sSQL = "SELECT P.IdPersona, (P.Apellidos+', '+P.Nombre)as NombreCompl, " _
       & " PF.Condicion, PF.Activo " _
       & "FROM PLN_PERSONAL P, PLN_PROFESIONAL PF " _
       & "WHERE P.IdPersona=PF.IdPersona and PF.Activo='SI' and P.IdPersona IN " _
       & "(SELECT DISTINCT PR.IdPersona FROM PRESTAMOS PR WHERE PR.PagadoCB='NO') " _
       & "ORDER BY (P.Apellidos+', '+P.Nombre)"
 
' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla
 
' Pone el título al grid
aTitulosColGrid = Array("Prestamo", "Descripción", "Número", "Monto", "CtaCtb")
aTamañosColumnas = Array(800, 3800, 800, 1500, 0)

CargarGridTitulos grdPrestamos, aTitulosColGrid, aTamañosColumnas

' Limpia la colección EgresoSA detalle
Set gcolDetMovCB = Nothing

' Establece obligatorio
txtPersonal.BackColor = Obligatorio

' Inhabilita el botón aceptar
cmdAceptar.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Destruye la colección
Set gcolTabla = Nothing
If gcolDetMovCB.Count = 0 Then
  ' Pone vacia el concepto del formulario egreso
  frmCBEGSinAfecta.txtCodMov = Empty
End If

End Sub

Private Sub grdPrestamos_Click()

'SI se pincha el grid y está vacío NO hace nada
If grdPrestamos.Rows = 1 Then
  Exit Sub
End If

'Se marca o desmarca la fila que se ha pinchado
MarcarUnaFilaGrid grdPrestamos

End Sub

Private Sub grdPrestamos_DblClick()

'Hace llamado al evento click del aceptar
cmdAceptar_Click

End Sub


Private Sub txtPersonal_Change()
'Verifica si el tamaño del txt es Igual al tamaño definido
If Len(txtPersonal) = txtPersonal.MaxLength Then
    'Actualiza el txtDesc
    ActualizaDesc
Else
    'Limpia el txtDescAfecta
    txtDesc.Text = Empty
End If
    
 ' Verifica SI el campo esta vacio
If txtPersonal.Text <> Empty And txtDesc.Text <> Empty Then
   ' Los campos coloca a color blanco
     txtPersonal.BackColor = vbWhite
   ' Carga los datos del personal seleccionado
     CargaPrestamos
   ' Habilita cmdAceptar
    cmdAceptar.Enabled = True
Else
   'Los campos coloca a color amarillo
   txtPersonal.BackColor = Obligatorio

  'deshabilita boton aceptar
  cmdAceptar.Enabled = False

  'limpia grid
  grdPrestamos.Rows = 1
End If
    
End Sub

Private Sub ActualizaDesc()
'--------------------------------------------------------------
'PROPÓSITO  : Actualiza la descripcion de la persona
'Recive     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------
On Error GoTo mnjError
'Copia la descripción
txtDesc.Text = Var30(gcolTabla.Item(txtPersonal.Text), 2)

' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe ", , "SGCcaijo-Ingresos"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub CargaPrestamos()
'-----------------------------------------------------------
' Proposito: Carga los prestamos del personal  que no se han _
             entregado en Caja Bancos a el grid
' Recibe: Nada
' Entrega : Nada
'-----------------------------------------------------------
Dim sSQL As String
Dim curPrestamos As New clsBD2

'Limpia el grid de prestamos
grdPrestamos.Rows = 1

' Carga la consulta
sSQL = "SELECT P.IdConPL, PC.DescConPL, P.NumPrestamo, P.Monto, PCO.CodContable " _
     & "FROM PRESTAMOS P, PLN_CONCEPTOS PC, PLNCONCEPTOS_OTROS PCO  " _
     & "WHERE P.IdPersona='" & txtPersonal & "' and P.PagadoCB='NO' " _
     & "and P.IdConPL=PC.IdConPl and PC.IdConPl=PCO.IdConPl"
' Ejecuta la sentencia
curPrestamos.SQL = sSQL
If curPrestamos.Abrir = HAY_ERROR Then End

' Carga el grid
Do While Not curPrestamos.EOF
    grdPrestamos.AddItem curPrestamos.campo(0) _
                           & vbTab & curPrestamos.campo(1) _
                           & vbTab & curPrestamos.campo(2) _
                           & vbTab & Format(curPrestamos.campo(3), "###,###,##0.00") _
                           & vbTab & curPrestamos.campo(4)
    curPrestamos.MoverSiguiente
Loop
' cierra la componente
curPrestamos.Cerrar
End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub
