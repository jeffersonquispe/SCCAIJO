VERSION 5.00
Begin VB.Form frmCBEGPago_Adelantos 
   Caption         =   "Entrega de adelantos al personal"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   Icon            =   "SCCBEGPagoAdelantos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6720
         Picture         =   "SCCBEGPagoAdelantos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6960
         Picture         =   "SCCBEGPagoAdelantos.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   870
         Width           =   220
      End
      Begin VB.ComboBox cboAdelantos 
         Height          =   315
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   5035
      End
      Begin VB.TextBox txtPersonal 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   550
      End
      Begin VB.TextBox txtAdelantos 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   3
         Top             =   840
         Width           =   550
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Monto (S/):"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Adelanto:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Personal:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCBEGPago_Adelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcolCodPersonal As New Collection
Dim mcolCodDesPersonal As New Collection
Dim mcolCodAdelantos As New Collection
Dim mcolCodDesAdelantos As New Collection
Dim mcolDesCtbAdelt As New Collection

Private Sub cmdAceptar_Click()
Dim iResult As Integer

' Asigna el resultado de la verificación de las cuotas
iResult = Var10(0, txtPersonal, _
          FechaAM(frmCBEGSinAfecta.mskFecTrab), Val(Var37(txtMonto)))
Select Case iResult
Case 0 ' Seguir con el procesos
Case 1 ' Interrumpir el proceso
    ' Sale de la función
    Exit Sub
Case 2 ' No seguir validando las otras cuotas
End Select

' Verifica si se puede dar adelantos
If Var11(frmCBEGSinAfecta.mskFecTrab) = False Then Exit Sub

' Coloca el monto a el formulario egreso sin afectación
  gcolDetMovCB.Add Item:=txtAdelantos & "¯" _
                        & mcolDesCtbAdelt(txtAdelantos) & "¯" _
                        & Var37(txtMonto), _
                     Key:=txtAdelantos

' Pasa la persona elegida y el monto a frmCBEgreSinAfecta
 frmCBEGSinAfecta.txtDesc = txtDesc.Text
 frmCBEGSinAfecta.txtAfecta.MaxLength = Len(txtPersonal)
 frmCBEGSinAfecta.txtAfecta = txtPersonal
 frmCBEGSinAfecta.txtMonto = txtMonto
 
' Cierra el formulario
  Unload Me
  
End Sub



Private Sub cmdBuscar_Click()
'Determina si tiene prestamos el personal
If gcolTabla.Count = 0 Then
    'Mensaje de nos hay prestamos al personal
    MsgBox "No hay adelantos definidos al personal", vbOKOnly + vbInformation, "SGCcaijo-Egresos sin Afectación"
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

Private Sub Command1_Click()

If cboAdelantos.Enabled Then
    ' alto
     cboAdelantos.Height = CBOALTO
    ' focus a cbo
    cboAdelantos.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Carga las colecciones con el personal que tiene adelantos definidos en planillas
  sSQL = "SELECT P.IdPersona, (P.Apellidos+', '+P.Nombre)as NombreCompl, " _
       & " PF.Condicion, PF.Activo " _
       & "FROM PLN_PERSONAL P, PLN_PROFESIONAL PF " _
       & "WHERE P.IdPersona=PF.IdPersona and PF.Activo='SI' and P.IdPersona IN " _
       & "(SELECT DISTINCT PA.IdPersona FROM PLN_ADELANTOS PA ) " _
       & "ORDER BY (P.Apellidos+', '+P.Nombre)"
       
' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

' Limpia la colección EgresoSA detalle
Set gcolDetMovCB = Nothing

' Establece obligatorio
txtPersonal.BackColor = Obligatorio
txtAdelantos.BackColor = Obligatorio
txtMonto.BackColor = Obligatorio

' Inhabilita el botón aceptar
cmdAceptar.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Destruye las colecciones
Set mcolCodAdelantos = Nothing
Set mcolCodDesAdelantos = Nothing
Set mcolDesCtbAdelt = Nothing

If gcolDetMovCB.Count = 0 Then
  ' Pone vacia el concepto del formulario egreso
  frmCBEGSinAfecta.txtCodMov = Empty
End If

End Sub

Private Sub txtAdelantos_Change()
 
 ' Si procede, se actualiza descripción correspondiente a código introducido
 CD_ActDesc cboAdelantos, txtAdelantos, mcolCodDesAdelantos
  
  ' Verifica si el campo esta vacio
 If txtAdelantos.Text <> "" And cboAdelantos.Text <> "" Then
    ' Los campos coloca a color blanco
      txtAdelantos.BackColor = vbWhite
 Else
    'Los campos coloca a color amarillo
    txtAdelantos.BackColor = Obligatorio

 End If
    
' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub HabilitarBotonAceptar()

' Verifica los obligatorios
If txtPersonal.BackColor <> vbWhite Or _
   txtAdelantos.BackColor <> vbWhite Or _
   txtMonto.BackColor <> vbWhite Then
    cmdAceptar.Enabled = False
Else ' Todo esta Ok
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub txtAdelantos_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  
End Sub

Private Sub txtMonto_Change()

' Verifica SI el campo esta vacio
If txtMonto.Text <> "" Then
' El campos coloca a color blanco
   txtMonto.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtMonto.BackColor = Obligatorio
End If

'habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtMonto_GotFocus()

' Elimina las comas para dar un correcto formato a monto
txtMonto = Var37(txtMonto)
' Coloca el tamaño a 12
txtMonto.MaxLength = 12

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Monto ingresado en soles, luego se ubica en la componente inmediata
 Var33 txtMonto, KeyAscii
  
End Sub

Private Sub txtMonto_LostFocus()
   
txtMonto.MaxLength = 14
If txtMonto.Text <> "" Then
   'Da formato de moneda
   txtMonto.Text = Format(Val(Var37(txtMonto.Text)), "###,###,###,##0.00")
Else
   txtMonto.BackColor = Obligatorio
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
    ' Los campos coloca a color blanco
      txtPersonal.BackColor = vbWhite
      txtAdelantos.Text = Empty
    ' Carga los datos del personal seleccionado
      CargaAdelantos
 Else
    'Los campos coloca a color amarillo
    txtPersonal.BackColor = Obligatorio

    ' Limpia el cbo de Adelantos
    txtAdelantos.Text = Empty
    cboAdelantos.Clear
    Set mcolCodAdelantos = Nothing
    Set mcolCodDesAdelantos = Nothing
 End If
    
' Habilita el botón aceptar
HabilitarBotonAceptar
    
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

Private Sub CargaAdelantos()
'-----------------------------------------------------------
' Proposito: Carga los adelantos del personal
' Recibe: Nada
' Entrega : Nada
'-----------------------------------------------------------

Dim sSQL As String
Dim curAdelantos As New clsBD2

' Limpia el cbo de Adelantos
cboAdelantos.Clear
Set mcolCodAdelantos = Nothing
Set mcolCodDesAdelantos = Nothing
Set mcolDesCtbAdelt = Nothing

' Carga la consulta
sSQL = "SELECT PA.IdConPL, PC.DescConPL, PO.CodContable " & _
     "FROM PLN_ADELANTOS PA, PLN_CONCEPTOS PC, PLNCONCEPTOS_OTROS PO " & _
     "WHERE PA.IdPersona='" & txtPersonal & "' and PA.IdConPL=PC.IdConPl " & _
     "And PC.IdConPL=PO.IdConPL"

'Carga la colección de descripcion y medida de los productos
curAdelantos.SQL = sSQL
If curAdelantos.Abrir = HAY_ERROR Then
  End
End If
Do While Not curAdelantos.EOF
    ' Se carga la colección de descripciones + unidades de los productos con la 1º y 2º
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolDesCtbAdelt.Add Key:=curAdelantos.campo(0), _
                              Item:=curAdelantos.campo(2)
    
    'colección de producto y su descripción
    mcolCodAdelantos.Add curAdelantos.campo(0)
    mcolCodDesAdelantos.Add Item:=curAdelantos.campo(1), Key:=curAdelantos.campo(0)
    
    ' Carga el cbo adelantos
    cboAdelantos.AddItem curAdelantos.campo(1)
    
    ' Se avanza a la siguiente fila del cursor
    curAdelantos.MoverSiguiente
Loop
'Cierra el cursor de medida de productos
curAdelantos.Cerrar

End Sub

Private Sub cboAdelantos_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboAdelantos) = True Then SendKeys "{down}"

End Sub

Private Sub cboAdelantos_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboAdelantos.ListIndex) = False And cboAdelantos.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub


Private Sub cboAdelantos_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboAdelantos_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboAdelantos, vbWhite) = True Then
    
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboAdelantos.Text, txtAdelantos, mcolCodAdelantos, mcolCodDesAdelantos
    
Else '  Vaciar Controles enlazados al combo
    txtAdelantos.Text = Empty
End If

'Cambia el alto del combo
cboAdelantos.Height = CBONORMAL

End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
