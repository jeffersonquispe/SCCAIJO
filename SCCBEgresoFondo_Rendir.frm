VERSION 5.00
Begin VB.Form frmCBEGEntrega_Rendir 
   Caption         =   "Caja y Bancos -Egresos, Entregas a rendir a cuenta del trabajador"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "SCCBEgresoFondo_Rendir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   7365
      Begin VB.TextBox txtMonto 
         Height          =   330
         Left            =   1605
         TabIndex        =   3
         Top             =   675
         Width           =   1770
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6720
         Picture         =   "SCCBEgresoFondo_Rendir.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   200
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto a entregar:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   735
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta a Rendir:"
         Height          =   420
         Left            =   240
         TabIndex        =   7
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmCBEGEntrega_Rendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
' Verifica si es correcto el monto
If DatosOK = False Then
    ' Sale del procedimiento
    Exit Sub
End If

' Coloca el monto a el formulario esgreso sin afectación
' IdPersona,CtaCtb,Monto
  gcolDetMovCB.Add Item:=txtPersonal & "¯" _
                        & Conta25 & "¯" _
                        & Var37(txtMonto), _
                     Key:=txtPersonal

' Pasa la persona elegida y el monto a frmCBEgreSinAfecta
 frmCBEGSinAfecta.txtDesc = txtDesc.Text
 frmCBEGSinAfecta.txtAfecta.MaxLength = Len(txtPersonal)
 frmCBEGSinAfecta.txtAfecta = txtPersonal
 frmCBEGSinAfecta.txtMonto.MaxLength = 14
 frmCBEGSinAfecta.txtMonto = txtMonto
 
' Cierra el formulario
  Unload Me

End Sub

Private Function DatosOK() As Boolean
' ----------------------------------------------------------
' Propósito: Verifica si los datos ingresados son correctos
' ----------------------------------------------------------
If gsTipoOperacionEgreso = "Modificar" Then
    ' Verifica el saldo de la cuenta
    If VerificarMontoModificar = False Then
        ' No se puede realizar
        DatosOK = False
        Exit Function
    End If
End If
' Todo ok
DatosOK = True

End Function

Private Sub cmdBuscar_Click()

'Determina si existen cuentas a rendir del personal
If gcolTabla.Count = 0 Then
    'Mensaje de nos hay cuentas a rendir del personal
    MsgBox "No hay cuentas a rendir del personal", vbOKOnly + vbInformation, "SGCcaijo-Ingresos, fondos a rendir"
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

'Private Sub cmdCancelar_Click()
'
''Limpia los controles
'txtPersonal.Text = Empty
'
''Ubica el control en txtPersonal
'txtPersonal.SetFocus
'
''Vacia la colección
'Set gcolAsientoDet = Nothing
'
'End Sub

Private Sub cmdSalir_Click()

'Vuelve al formulario anterior
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Limpia la colección EgresoSA detalle
Set gcolDetMovCB = Nothing

'Coloca a obligatorio el txtCodPersonal
txtPersonal.BackColor = Obligatorio
txtMonto.BackColor = Obligatorio

'Deshabilita el boton cmdaceptar
cmdAceptar.Enabled = False

' Importa los datos del formulario de egresos
txtPersonal = frmCBEGSinAfecta.txtAfecta
txtMonto.MaxLength = 14
txtMonto = frmCBEGSinAfecta.txtMonto

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Limpia las colecciones
If gsTipoOperacionEgreso = "Nuevo" Then
    If gcolDetMovCB.Count = 0 Then
      ' Pone vacia el concepto del formulario egreso
      frmCBEGSinAfecta.txtAfecta = Empty
      frmCBEGSinAfecta.txtMonto = Empty
    End If
Else
' Coloca el monto a el formulario egreso sin afectación
' IdPersona,CtaCtb,Monto
    If gcolDetMovCB.Count = 0 Then
      gcolDetMovCB.Add Item:=gsCodAfectaAnterior & "¯" _
                            & Conta25 & "¯" _
                            & gdblMontoAnterior, _
                         Key:=gsCodAfectaAnterior
    ' Pasa la persona elegida y el monto a frmCBEgreSinAfecta
     frmCBEGSinAfecta.txtDesc = Var30(gcolTabla.Item(gsCodAfectaAnterior), 2)
     frmCBEGSinAfecta.txtAfecta.MaxLength = Len(gsCodAfectaAnterior)
     frmCBEGSinAfecta.txtAfecta = gsCodAfectaAnterior
     frmCBEGSinAfecta.txtMonto.MaxLength = 14
     frmCBEGSinAfecta.txtMonto = Format(gdblMontoAnterior, "###,###,##0.00")
 End If
End If

End Sub

Private Sub HabilitarBotonAceptar()

' Verifica los obligatorios
If txtPersonal.BackColor <> vbWhite Or _
   txtMonto.BackColor <> vbWhite Then
    'Inicializa la variable
    cmdAceptar.Enabled = False
    Exit Sub
End If

' Habilita el botón aceptar
cmdAceptar.Enabled = True

End Sub

Private Function VerificarMontoModificar() As Boolean
'---------------------------------------------------------------------------------------
'Propósito :Verificar si se puede aceptar el monto a rendir
'Recibe    : Nada
'Devuelve  :booleano que indica SI el monto esta conforme
'---------------------------------------------------------------------------------------
Dim dblMonto, dblMontoAnt, dblSaldo As Double
'Inicializamos la funcion asumiendo que el monto esta correcto
VerificarMontoModificar = True
dblMontoAnt = gdblMontoAnterior
dblMonto = Val(Var37(txtMonto.Text))
'Averigua el saldo de la cuenta a rendir
dblSaldo = Var6(gsCodAfectaAnterior)

'verifica SI es la misma cuenta a rendir del egreso Original
If txtPersonal = gsCodAfectaAnterior Then 'La misma
  'Verifica SI monto modificado excede el saldo de Caja o Bancos
  If (dblMonto - dblMontoAnt) + dblSaldo < -0.0001 Then
      MsgBox "No existe saldo suficiente en la cuenta a rendir", _
           vbInformation, "Caja-Banco- Modificación de entrega a rendir"
      txtMonto = gdblMontoAnterior
      txtMonto.SetFocus
      VerificarMontoModificar = False
      Exit Function
   End If
Else 'Se Cambio de CtaCte
  'Verifica SI monto excede el saldo de Caja o Bancos
  If Val(gdblMontoAnterior) > Val(dblSaldo) Then
      MsgBox "No se puede cambiar de cuenta a rendir." & Chr(13) & _
            "No existe saldo suficiente en la cuenta original del egreso", _
          vbInformation, "Caja-Banco- Modificación de entrega a rendir"
      txtPersonal = gsCodAfectaAnterior
      txtPersonal.SetFocus
      VerificarMontoModificar = False
      Exit Function
   End If
End If

End Function

Private Sub txtMonto_Change()

'Verifica SI el campo esta vacio
If txtMonto.Text <> Empty And Val(txtMonto.Text) <> 0 Then
  'El campos coloca a color blanco
   txtMonto.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtMonto.BackColor = Obligatorio
End If

'Habilita el botón aceptar en caso de estar lleno todos los campos
HabilitarBotonAceptar

End Sub

Private Sub txtMonto_GotFocus()

txtMonto.MaxLength = 12
'Elimina las comas
txtMonto.Text = Var37(txtMonto.Text)

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

'Maximo número de digitos para el monto
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
Else
    'Los campos coloca a color amarillo
   txtPersonal.BackColor = Obligatorio
End If

'Habilita el boton aceptar
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
        MsgBox "El código ingresado no existe en la BD", , "SGCcaijo-Egresos, Entrega a Rendir"
        'Limpia la descripción
        txtDesc.Text = Empty
        txtMonto.Text = Empty
    End If
End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
