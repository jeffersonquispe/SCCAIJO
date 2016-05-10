VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBIDev_Prestamos 
   Caption         =   "Caja y Bancos - Ingreso, Selección de Devolución de Prestamos"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "SCCBIngreosDev_Prestamos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   7365
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6720
         Picture         =   "SCCBIngreosDev_Prestamos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   200
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   200
         Width           =   5140
      End
      Begin VB.TextBox txtPersonal 
         Height          =   315
         Left            =   960
         MaxLength       =   4
         TabIndex        =   0
         Top             =   200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   5280
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetPrestamos 
      Height          =   2055
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   5280
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdPrestamos 
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FillStyle       =   1
   End
   Begin VB.Label Label3 
      Caption         =   "PRESTAMOS:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "DETALLE DE PRESTAMOS:"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frmCBIDev_Prestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdblSaldoPrestamo As Double

Private Sub cmdAceptar_Click()

'Carga el código del personal
frmCBIngresos.txtDesc.Text = txtDesc.Text
frmCBIngresos.txtAfecta.Text = txtPersonal.Text
frmCBIngresos.lblEtiqueta.Caption = "Personal:"
frmCBIngresos.txtMonto.Text = Format(Val(mdblSaldoPrestamo), "###,###,###,##0.00")
   
'Vuelve al formulario anteriord
Unload Me

End Sub

Private Sub cmdBuscar_Click()

'Determina si tiene prestamos el personal
If gcolTabla.Count = 0 Then
    'Mensaje de nos hay prestamos al personal
    MsgBox "No hay prestamos a cancelar del personal", vbOKOnly + vbInformation, "SGCcaijo-Ingresos"
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
  End If
End Sub

Private Sub cmdCancelar_Click()

'Limpia los controles
txtPersonal.Text = Empty
grdPrestamos.Rows = 1
grdDetPrestamos.Rows = 1

'Ubica el control en txtPersonal
txtPersonal.SetFocus

'Vacia la colección
Set gcolAsientoDet = Nothing

End Sub

Private Sub cmdSalir_Click()

'Vacia las colecciones
Set gcolPrestamos = Nothing
Set gcolDetPrestamos = Nothing

'Vuelve al formulario anterior
Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim sSQL As String

'Coloca a obligatorio el txtCodPersonal
txtPersonal.BackColor = Obligatorio

'Deshabilita el boton cmdaceptar
cmdAceptar.Enabled = False

'Vacia la colección
Set gcolAsientoDet = Nothing

'Carga el combo con los nombres del personal, definido su remuneración
sSQL = "SELECT DISTINCT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
       & " PR.Condicion, PR.Activo " _
       & "FROM PLN_PERSONAL P, PLN_PROFESIONAL PR " _
       & "WHERE P.IdPersona=PR.IdPersona And " _
       & "P.IdPersona IN " _
            & "(SELECT DISTINCT PP.IdPersona " _
            & "FROM PRESTAMOS PP " _
            & "WHERE P.IdPersona=PP.IdPersona and PP.Cancelado='NO' and PP.PagadoCB='SI') " _
       & "ORDER BY  ( p.Apellidos & ', ' & P.Nombre)"

'   VADICK SE DEBE CONSIDERAR A PERSONAL QUE TIENE UNA DEUDA ESTA ACTIVO O INACTIVO
'sSQL = "SELECT DISTINCT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
'       & " PR.Condicion, PR.Activo " _
'       & "FROM PLN_PERSONAL P, PLN_PROFESIONAL PR " _
'       & "WHERE P.IdPersona=PR.IdPersona And PR.Activo='NO' And " _
'       & "P.IdPersona IN " _
'            & "(SELECT DISTINCT PP.IdPersona " _
'            & "FROM PRESTAMOS PP " _
'            & "WHERE P.IdPersona=PP.IdPersona and PP.Cancelado='NO' and PP.PagadoCB='SI') " _
'       & "ORDER BY  ( p.Apellidos & ', ' & P.Nombre)"

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

'Coloca el titulo al Grid
aTitulosColGrid = Array("IdConPL", "Tipo de Prestamo", "Nro Prestamo", "Monto Total", "Saldo", "Fecha", "CodContable")
aTamañosColumnas = Array(0, 3000, 1100, 1100, 1100, 1100, 0)
CargarGridTitulos grdPrestamos, aTitulosColGrid, aTamañosColumnas

'Coloca el titulo al Grid
aTitulosColGrid = Array("Tipo de Prestamo", "Nro Cuota", "Monto Cuota", "Monto Amortizado", "Fecha Cancelación", "Cancelado")
aTamañosColumnas = Array(0, 900, 1300, 1500, 1600, 1000)
CargarGridTitulos grdDetPrestamos, aTitulosColGrid, aTamañosColumnas

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Vacia las colecciones
Set gcolTabla = Nothing

If gcolPrestamos.Count = 0 Then
  ' Pone vacia el concepto del formulario egreso
  frmCBIngresos.txtCodMov = Empty
End If

End Sub

Private Sub MuestraDetallePrestamo()
'-------------------------------------------------------------------------
'Propósito  : Carga el grdDetallePrestamo con el detalle del prestamo
'Recibe     : Nada
'Devuelve   : Nada
'-------------------------------------------------------------------------
Dim sSQL As String
Dim curDetPrestamos As New clsBD2


'Sentencia sql
sSQL = ""
sSQL = "SELECT PC.IdConPl,PC.NumCuota, PC.Cuota, " _
      & "PC.Amortizado, PC.AnioMes, PC.Cancelado, PC.NumPrestamo " _
      & "FROM PRESTAMOS_CUOTAS PC, PLN_CONCEPTOS C " _
      & "WHERE PC.IdPersona= '" & txtPersonal.Text & "' And " _
      & "PC.NumPrestamo='" & grdPrestamos.TextMatrix(grdPrestamos.Row, 2) & "' " _
      & " And PC.IdConPL='" & grdPrestamos.TextMatrix(grdPrestamos.Row, 0) & "' " _
      & "And PC.IdConPL=C.IdConPL ORDER BY PC.NumCuota "
      
'Copia la sentencia sql
curDetPrestamos.SQL = sSQL

'Ejecuta la sentencia SQL
If curDetPrestamos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
Else
    'inicializa la variable dblSaldoPrestamo
    mdblSaldoPrestamo = 0
    
    'Copia los datos al grdPrestamos
    Do While Not curDetPrestamos.EOF
        
        'Agrega los datos al grdPrestamos
        grdDetPrestamos.AddItem curDetPrestamos.campo(0) & vbTab & _
                                curDetPrestamos.campo(1) & vbTab & _
                                Format(curDetPrestamos.campo(2), "###,###,###,##0.00") _
                                & vbTab & Format(curDetPrestamos.campo(3), "###,###,###,##0.00") _
                                & vbTab & FechaMMAAAA(curDetPrestamos.campo(4)) & vbTab & _
                                curDetPrestamos.campo(5)
        
        If curDetPrestamos.campo(5) = "NO" Then
        
            'Agrega a la coleccion los datos del grdDetPrestamos
            gcolDetPrestamos.Add Item:=curDetPrestamos.campo(0) & "¯" & _
                                    curDetPrestamos.campo(6) & "¯" & _
                                    curDetPrestamos.campo(4) & "¯" & _
                                    curDetPrestamos.campo(2), _
                                Key:=curDetPrestamos.campo(0) & "¯" & _
                                    curDetPrestamos.campo(6) & "¯" & _
                                    curDetPrestamos.campo(4)

            
            'Acumula los saldo que falta pagar
            mdblSaldoPrestamo = mdblSaldoPrestamo + Val(curDetPrestamos.campo(2))
            
        End If
        
        'Mueve al siguiete registro
        curDetPrestamos.MoverSiguiente
    Loop
    
End If

'Cierra el curPrestamosPersonal
curDetPrestamos.Cerrar

End Sub

Private Sub grdPrestamos_Click()
Dim i As Integer
Dim iFila As Integer
Dim bSeleccionado As Double

iFila = grdPrestamos.Row
i = 1

'SI se pincha el grid y está vacío NO hace nada
If grdPrestamos.Rows = 1 Then
   cmdAceptar.Enabled = False
  Exit Sub
End If

' Selecciona toda la iFila
If grdPrestamos.Rows > 1 Then

  'Se marca o desmarca la fila que se ha pinchado
   MarcarUnaFilaGrid grdPrestamos
    
    'Inicializa la variable bSeleccionado en falso
    bSeleccionado = False
    i = 1
    Do While i <= grdPrestamos.Rows - 1 And bSeleccionado = False
    
        'Vacia las colecciones
        Set gcolPrestamos = Nothing
        Set gcolDetPrestamos = Nothing
        
        'Ubica el registro seleccionado
        grdPrestamos.Row = i
        
        'Verifica si esta seleccionado
        If grdPrestamos.CellBackColor = vbDarkBlue Then
        
            'Despues de selccionar bSeleccionado coloca a True
            bSeleccionado = True
            
            'Habilita el boton aceptar
            cmdAceptar.Enabled = True
            
            'Borra el grdDetPrestamos
            grdDetPrestamos.Rows = 1
            
            'Carga coleccion con registro seleccionado en el grdPrestamo
            gcolPrestamos.Add Item:=grdPrestamos.TextMatrix(grdPrestamos.Row, 0) & "¯" & _
                              grdPrestamos.TextMatrix(grdPrestamos.Row, 2) & "¯" & _
                              grdPrestamos.TextMatrix(grdPrestamos.Row, 6) & "¯" & _
                              grdPrestamos.TextMatrix(grdPrestamos.Row, 3), _
                              Key:=grdPrestamos.TextMatrix(grdPrestamos.Row, 0)

            
        End If
        i = i + 1
    Loop
   ' Limpia el grid detalle
   grdDetPrestamos.Rows = 1
        
   'Muestra el detalle en el grdDetallePrestamo
   MuestraDetallePrestamo
   
End If

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
     'Limpia el grdPrestamos
        grdPrestamos.Rows = 1
        grdDetPrestamos.Rows = 1
        ' Los campos coloca a color blanco
        txtPersonal.BackColor = vbWhite
        'Carga el Grid
        CargaGrdPersonal
Else
    'Los campos coloca a color amarillo
   txtPersonal.BackColor = Obligatorio
   grdPrestamos.Rows = 1
   grdDetPrestamos.Rows = 1
End If

'Dehabilita el boton cmdAceptar
cmdAceptar.Enabled = False

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
        MsgBox "El código ingresado no existe en la BD", , "SGCcaijo-Prestamos"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub CargaGrdPersonal()
'-------------------------------------------------------------------------
'Propósito  : Carga el grdPrestamos con los prestamos no cancelados del per-
'             sonal seleccionado
'Recibe     : Nada
'Devuelve   : Nada
'-------------------------------------------------------------------------
Dim sSQL As String
Dim curPrestamosPersonal As New clsBD2

'Sentencia sql
sSQL = "SELECT P.IdConPl, C.DescConPl, P.NumPrestamo, P.Monto, Sum(PC.Cuota), P.Fecha, CO.CodContable " _
      & "FROM PRESTAMOS P, PRESTAMOS_CUOTAS PC, PLN_CONCEPTOS C, PLNCONCEPTOS_OTROS CO " _
      & "WHERE P.IdPersona= '" & txtPersonal.Text & "' And  P.Cancelado ='NO' And " _
      & "P.PagadoCB='SI' And P.IdPersona= PC.IdPersona And P.NumPrestamo=PC.NumPrestamo " _
      & "And P.IdConPL=PC.IdConPL And PC.Cancelado='NO' " _
      & " And P.IdPersona= PC.IdPersona And PC.IdConPL=C.IdConPl and C.IdConPl=CO.IdConpl " _
      & "GROUP BY P.IdConPl, C.DescConPl, P.NumPrestamo, P.Monto, P.Fecha,CO.CodContable " _
      & "ORDER BY P.NumPrestamo "
      
'Copia la sentencia sql
curPrestamosPersonal.SQL = sSQL

'Ejecuta la sentencia SQL
If curPrestamosPersonal.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
Else

    'Copia los datos al grdPrestamos
    Do While Not curPrestamosPersonal.EOF
        
        'Agrega los datos al grdPrestamos
        grdPrestamos.AddItem curPrestamosPersonal.campo(0) & vbTab & curPrestamosPersonal.campo(1) & vbTab & _
                            curPrestamosPersonal.campo(2) & vbTab & Format(curPrestamosPersonal.campo(3), "###,###,###,##0.00") & vbTab & _
                            Format(curPrestamosPersonal.campo(4), "###,###,###,##0.00") & vbTab & FechaDMA(curPrestamosPersonal.campo(5)) & vbTab & _
                            curPrestamosPersonal.campo(6)

        'Mueve al siguiete registro
        curPrestamosPersonal.MoverSiguiente
    Loop
    
End If

'Cierra el curPrestamosPersonal
curPrestamosPersonal.Cerrar

End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
