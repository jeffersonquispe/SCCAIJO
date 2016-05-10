VERSION 5.00
Begin VB.Form frmCBIFondo_Rendir 
   Caption         =   "Caja y Bancos - Ingreso, Selección de Personal de Fondo a Rendir"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "SCCBIngresoFondo_Rendir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   7365
      Begin VB.TextBox txtMonto 
         Height          =   330
         Left            =   1590
         MaxLength       =   12
         TabIndex        =   2
         Top             =   645
         Width           =   1770
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6720
         Picture         =   "SCCBIngresoFondo_Rendir.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   200
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
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
         Caption         =   "Saldo a devolver:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   735
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5745
      TabIndex        =   4
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6705
      TabIndex        =   5
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4785
      TabIndex        =   3
      Top             =   1275
      Width           =   855
   End
End
Attribute VB_Name = "frmCBIFondo_Rendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()

'Carga el código del personal
frmCBIngresos.txtDesc.Text = txtDesc.Text
frmCBIngresos.txtAfecta.Text = txtPersonal.Text
frmCBIngresos.lblEtiqueta.Caption = "Personal:"
frmCBIngresos.txtMonto.MaxLength = 14
frmCBIngresos.txtMonto.Text = Format(Var37(txtMonto), "###,###,###,##0.00")
   
'Vuelve al formulario anteriord
Unload Me

End Sub

Private Sub cmdBuscar_Click()

'Determina si existen cuentas a rendir del personal
If gcolTabla.Count = 0 Then
    'Mensaje de nos hay cuentas a rendir del personal
    MsgBox "No hay cuentas a rendir del personal", vbOKOnly + vbInformation, "SGCcaijo-Ingresos, fondos a rendir"
    'Decarga el formulario
    Exit Sub
End If
' Asigna la colección
  Set gcolTRendir = gcolTabla
  txtPersonal.Text = Empty
  
' Carga los títulos del grid selección
  giNroColMNSel = 3
  aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Monto")
  aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Saldo a Rendir")
  aTamañosColumnas = Array(1000, 4500, 1200)
' Muestra el formulario de busqueda
  frmMNSelecERendir.Show vbModal, Me
  
' vacia la colección rendir
Set gcolTRendir = Nothing

' Verifica si se eligió algun dato a modificar
  If gsCodigoMant <> Empty Then
    txtPersonal.Text = gsCodigoMant
    SendKeys "{tab}"
  End If
End Sub

Private Sub cmdCancelar_Click()

'Limpia los controles
txtPersonal.Text = Empty

'Ubica el control en txtPersonal
txtPersonal.SetFocus

'Vacia la colección
Set gcolAsientoDet = Nothing

End Sub

Private Sub cmdSalir_Click()

'Vuelve al formulario anterior
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim di As Double
'Coloca a obligatorio el txtCodPersonal
txtPersonal.BackColor = Obligatorio
txtMonto.BackColor = Obligatorio

'Deshabilita el boton cmdaceptar
cmdAceptar.Enabled = False

If gsTipoOperacionIngreso = "Nuevo" Then
    'Carga el combo con los nombres del personal, que tienen entrega a rendir
    sSQL = "SELECT DISTINCT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
           & " SUM(ME.Ingreso)-SUM(ME.egreso) " _
           & "FROM PLN_PERSONAL P, MOV_ENTREG_RENDIR ME " _
           & "WHERE P.IdPersona=ME.IdPersona And ME.Anulado='NO' " _
           & "GROUP BY P.IdPersona, ( P.Apellidos & ', ' & P.Nombre) " _
           & "HAVING (SUM(ME.Ingreso)-SUM(ME.egreso))>0"
Else
    'Carga la colección de los trabajadores que tienen cuentas a rendir
    sSQL = "SELECT DISTINCT P.IdPersona, ( P.Apellidos & ', ' & P.Nombre), " _
               & "(SUM(ME.Ingreso)-SUM(ME.egreso)) " _
               & "FROM PLN_PERSONAL P, MOV_ENTREG_RENDIR ME " _
               & "WHERE P.IdPersona=ME.IdPersona and ME.Anulado='NO' " _
               & "GROUP BY P.IdPersona, ( P.Apellidos & ', ' & P.Nombre)"
               
End If

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros del personal que tiene entregas a rendir
Var17 sSQL, 3, gcolTabla

'Verifica el tipo de operación
If gsTipoOperacionIngreso = "Modificar" Then
    'Recupera los datos originales
    txtPersonal.Text = gsCodAfectaAnterior
    txtMonto.MaxLength = 14
    txtMonto.Text = Format(gdblMontoAnterior, "###,###,###,##0.00")
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If gcolTabla.Count = 0 Then
  ' Pone vacia el concepto del formulario egreso
  frmCBIngresos.txtCodMov = Empty
End If

'Vacia las colecciones
Set gcolTabla = Nothing

End Sub

Private Sub txtMonto_Change()
' Verifica SI el campo esta vacio
If txtMonto.Text <> Empty And Val(txtMonto.Text) <> 0 Then
    ' Los campos coloca a color blanco
    txtMonto.BackColor = vbWhite
Else
    'Los campos coloca a color amarillo
   txtMonto.BackColor = Obligatorio
End If

'Habilita el boton aceptar
HabilitarBotonAceptar

End Sub

Private Sub HabilitarBotonAceptar()
On Error GoTo mnjError

'Verificia el tipo de operación
If gsTipoOperacionIngreso = "Nuevo" Then
    ' Verifica los obligatorios
    If txtPersonal.BackColor <> Obligatorio And _
       txtMonto.BackColor <> Obligatorio Then
       'Verifica si el monto es Mayor al ingreso que dispone el personal
       If Val(Var37(txtMonto.Text)) <= Val(Var30(gcolTabla.Item(txtPersonal.Text), 3)) Then
            'Habilita el boton aceptar
            cmdAceptar.Enabled = True
            Exit Sub
       End If
    End If
Else
    ' Verifica los obligatorios
    If txtPersonal.BackColor <> Obligatorio And _
       txtMonto.BackColor <> Obligatorio Then
       If gsCodAfectaAnterior = txtPersonal.Text Then
            'Verifica si el monto es Mayor al ingreso que dispone el personal
            If Val(Var37(txtMonto.Text)) <= Val(gdblMontoAnterior) _
                 + Val(Var30(gcolTabla.Item(txtPersonal.Text), 3)) Then
                 'Habilita el boton aceptar
                 cmdAceptar.Enabled = True
                 Exit Sub
            End If
       Else
            'Verifica si el monto es Mayor al ingreso que dispone el personal
            If Val(Var37(txtMonto.Text)) <= Val(Var30(gcolTabla.Item(txtPersonal.Text), 3)) Then
                 'Habilita el boton aceptar
                 cmdAceptar.Enabled = True
                 Exit Sub
            End If
       End If
    End If
End If

mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe en la BD", , "SGCcaijo-Ingresos, Entrega a Rendir"
        'Limpia la descripción
        txtPersonal.Text = Empty
       
    End If
'Inicializa la variable
cmdAceptar.Enabled = False

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
    txtMonto.Text = Empty
End If

' Verifica SI el campo esta vacio
If txtPersonal.Text <> Empty And txtDesc.Text <> Empty Then
    ' Los campos coloca a color blanco
    txtPersonal.BackColor = vbWhite
    txtMonto.BackColor = vbWhite
    
Else
    'Los campos coloca a color amarillo
   txtPersonal.BackColor = Obligatorio
   txtMonto.BackColor = Obligatorio
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
txtMonto.MaxLength = 14
txtDesc.Text = Var30(gcolTabla.Item(txtPersonal.Text), 2)
txtMonto.Text = Format(Var30(gcolTabla.Item(txtPersonal.Text), 3), "###,###,###,##0.00")

' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe en la BD", , "SGCcaijo-Ingresos, Entrega a Rendir"
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

