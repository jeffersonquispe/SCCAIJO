VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBIngresos 
   Caption         =   "SGCcaijo - Ingresos a Caja Bancos "
   ClientHeight    =   5295
   ClientLeft      =   6555
   ClientTop       =   345
   ClientWidth     =   7815
   HelpContextID   =   59
   Icon            =   "SCCBIngresos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPCtaCte 
      Height          =   255
      Left            =   7155
      Picture         =   "SCCBIngresos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1230
      Width           =   220
   End
   Begin VB.ComboBox cboCtaCte 
      Height          =   315
      Left            =   5655
      Style           =   1  'Simple Combo
      TabIndex        =   21
      Top             =   1200
      Width           =   1740
   End
   Begin VB.CommandButton cmdPBanco 
      Height          =   255
      Left            =   4125
      Picture         =   "SCCBIngresos.frx":0BA2
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1260
      Width           =   220
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   1725
      Style           =   1  'Simple Combo
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1230
      Width           =   2655
   End
   Begin VB.CommandButton cmdPCodMov 
      Height          =   255
      Left            =   7125
      Picture         =   "SCCBIngresos.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1755
      Width           =   220
   End
   Begin VB.ComboBox cboCodMov 
      Height          =   315
      Left            =   1950
      Style           =   1  'Simple Combo
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Ingresos diversos"
      Top             =   1725
      Width           =   5430
   End
   Begin VB.CommandButton cmdPCodContable 
      Height          =   255
      Left            =   7140
      Picture         =   "SCCBIngresos.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2700
      Width           =   220
   End
   Begin VB.ComboBox cboCtaContable 
      Height          =   315
      Left            =   2055
      Style           =   1  'Simple Combo
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Banco de credito del peru"
      Top             =   2670
      Width           =   5325
   End
   Begin VB.CommandButton cmdPTipDoc 
      Height          =   255
      Left            =   4605
      Picture         =   "SCCBIngresos.frx":142A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3165
      Width           =   220
   End
   Begin VB.ComboBox cboTipDoc 
      Height          =   315
      Left            =   1740
      Style           =   1  'Simple Combo
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3135
      Width           =   3105
   End
   Begin VB.TextBox txtBanco 
      Height          =   315
      Left            =   1240
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1230
      Width           =   450
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4515
      TabIndex        =   24
      ToolTipText     =   "Vuelve al Menú Principal"
      Top             =   4815
      Width           =   1005
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "An&ular"
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      ToolTipText     =   "Graba los datos"
      Top             =   4800
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      ToolTipText     =   "Volver al Menú Principal"
      Top             =   4800
      Width           =   1005
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3435
      TabIndex        =   23
      ToolTipText     =   "Graba los datos"
      Top             =   4815
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   27
      Top             =   105
      Width           =   7560
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6765
         Picture         =   "SCCBIngresos.frx":1702
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2100
         Width           =   4920
      End
      Begin VB.Frame fraIngreso 
         Caption         =   "Ingreso a:"
         Height          =   615
         Left            =   3420
         TabIndex        =   28
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optCaja 
            Caption         =   "Ca&ja"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optBanco 
            Caption         =   "Ba&nco"
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtCodMov 
         Height          =   315
         Left            =   1140
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "i009"
         Top             =   1620
         Width           =   615
      End
      Begin VB.TextBox txtDocIngreso 
         Height          =   315
         Left            =   5820
         MaxLength       =   15
         TabIndex        =   14
         Top             =   3030
         Width           =   1455
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   15
         Top             =   3465
         Width           =   1455
      End
      Begin VB.TextBox txtObserv 
         Height          =   315
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   22
         Top             =   4035
         Width           =   6060
      End
      Begin VB.TextBox txtTipDoc 
         Height          =   315
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   11
         Top             =   3030
         Width           =   420
      End
      Begin VB.TextBox txtMontoPendiente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4155
         MaxLength       =   12
         TabIndex        =   16
         Top             =   3465
         Width           =   1455
      End
      Begin VB.TextBox txtCodContable 
         Height          =   315
         Left            =   1140
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "10712"
         Top             =   2565
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscarIngreso 
         Caption         =   "..."
         Height          =   255
         Left            =   2295
         TabIndex        =   1
         Top             =   390
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   6015
         TabIndex        =   4
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
      Begin VB.TextBox txtCodIngreso 
         Height          =   315
         Left            =   1160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox txtAfecta 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   2100
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   2760
         Picture         =   "SCCBIngresos.frx":1804
         Stretch         =   -1  'True
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mo&vimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha:"
         Height          =   195
         Left            =   5415
         TabIndex        =   39
         Top             =   375
         Width           =   495
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   165
         TabIndex        =   38
         Top             =   1155
         Width           =   510
      End
      Begin VB.Label lblCtaCte 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cuen&ta:"
         Height          =   195
         Left            =   4500
         TabIndex        =   37
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "&Doc. Ingreso:"
         Height          =   255
         Left            =   4830
         TabIndex        =   36
         Top             =   3030
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "&Monto (S/.) :"
         Height          =   255
         Left            =   150
         TabIndex        =   35
         Top             =   3540
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Observación:"
         Height          =   195
         Left            =   165
         TabIndex        =   34
         Top             =   4125
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "T&ipo Doc.:"
         Height          =   195
         Left            =   165
         TabIndex        =   33
         Top             =   3105
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso:"
         Height          =   195
         Left            =   165
         TabIndex        =   32
         Top             =   435
         Width           =   570
      End
      Begin VB.Label lblMontoPendiente 
         Caption         =   "Saldo &Pendiente :"
         Height          =   255
         Left            =   2835
         TabIndex        =   31
         Top             =   3495
         Width           =   1335
      End
      Begin VB.Label lblCodContable 
         AutoSize        =   -1  'True
         Caption         =   "CtaContab&le:"
         Height          =   240
         Left            =   165
         TabIndex        =   30
         Top             =   2610
         Width           =   915
      End
      Begin VB.Label lblEtiqueta 
         Height          =   375
         Left            =   210
         TabIndex        =   29
         Top             =   2085
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCBIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Colecciones para la carga del combo de Movimientos
Private mcolCodMov As New Collection
Private mcolCodDesCodMov As New Collection

'Colecciones para la carga del combo de Bancos
Private mcolCodBanco As New Collection
Private mcolCodDesBanco As New Collection

'Colecciones para la carga del combo de Cuentas
Private mcolCodCtaCte As New Collection
Private mcolCodDesCtaCte As New Collection

'Colecciones para la carga del Tipo de Documento
Private mcolCodTipDoc As New Collection
Private mcolCodDesTipDoc As New Collection

'Colección para la carga de Código contable y código del tipo de movimiento
Private mcolCodCont As New Collection
Private mcolDesCodCont As New Collection

'Colección para la carga de Afecta y código del tipo de movimiento
Private mcolCodAfecta As New Collection
Private mcolDesCodAfecta As New Collection

'Colección para la carga del código contable referente al tipo de movimiento
Private mcolCodPlanCont As New Collection
Private mcolDesCodPlanCont As New Collection

'Cursor que carga el registro de ingreso para su modificacion
Private curRegIngresoCajaBanco As New clsBD2

'Variable donde se carga el codigo equivalente al combobox recuperado
Private msCtaCte As String

'variable que identifica SI el ingreso es a caja o bancos
Private msCajaoBanco As String

'Variable que identifica el tipo de operación realizada con el registro "Ingreso Nuevo","Modificacion"
Private msOperacion As String

'Determina el maxlength del campo txtAfecta cuando es personal y terceros
Private miTamañoPer As Integer

'Variable que identifica a que Afecta el concepto(Tipo_Mov), Terceros o Personal
Private msAfecta As String '(T o P o Vacio)
Private msProceso As String

'Variable que determina el Mayor Tamaño de CodCont en Conceptos(TipoMov de la BD)
Private miTamañoCodCont As Integer

'Variable que determina el Mayor Tamaño de teceros en Conceptos(Terceros de la BD)
Private miTamañoTer As Integer
Private msIngresoDevPrestamos As String

'Variable que determina si el formulario esta cargado
Private mbIngresoCargado As Boolean

Public OrdenVenta As String
Public CodigoVenta As String
Public CodigoTerceroVenta As String
Public VentaTotal As String
Public VentaPagada As String
Public VentaSaldo As String

Private Sub cboBanco_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboBanco) = True Then SendKeys "{down}"

End Sub

Private Sub cboBanco_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboBanco.ListIndex) = False And cboBanco.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
   
End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboBanco_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboBanco, vbWhite) = True Then

  'Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboBanco.Text, txtBanco, mcolCodBanco, mcolCodDesBanco
Else
  txtBanco.Text = Empty
End If

'Cambia el alto del combo
 cboBanco.Height = CBONORMAL

End Sub

Public Sub ActualizarListcboCtaCte()

'----------------------------------------------------------------------------
'Propósito: Actualizar la lista del combo CtaCte de acuerdo al txtBanco
'Recibe:    Nada
'Devuelve:  Nada
'----------------------------------------------------------------------------
'Nota:      llamado desde el evento lostfocus de cboBanco y Change de txtBanco
Dim sSQL As String
Dim curCtaCte As New clsBD2

'Inicializa el cboCtaCte
cboCtaCte.Clear
cboCtaCte.BackColor = Obligatorio

 If txtBanco.BackColor <> Obligatorio And Len(txtBanco.Text) = txtBanco.MaxLength Then
  ' Carga la Sentencia para obtener las Ctas en dólares que pertenecen a el txtBanco
  If gsTipoOperacionIngreso = "Modificar" Then
    sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
       " WHERE c.idbanco = '" & txtBanco & "' " & _
       " AND  b.idBanco = c.IdBanco" & _
       " AND c.idmoneda= 'SOL' ORDER BY c.DescCta"
  Else
    sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
       " WHERE c.idbanco = '" & txtBanco & "' " & _
       " AND  b.idBanco = c.IdBanco" & _
       " AND c.idmoneda= 'SOL' AND C.ANULADO='NO' ORDER BY c.DescCta"
  End If
  curCtaCte.SQL = sSQL
  If curCtaCte.Abrir = HAY_ERROR Then
    End
  End If
    
  'Verifica SI existen cuentas asociadas a txtBanco
  If curCtaCte.EOF Then
    'NO existe cuentas asociadas
    MsgBox "NO existen cuentas en el banco seleccionado. Consulte al administrador", _
            vbInformation + vbOKOnly, "S.G.Ccaijo- Cuentas Bancarias"
    'end
  Else
    'Se carga el cboCtaCte con las cuentas asociadas
    Do While Not curCtaCte.EOF
  
      cboCtaCte.AddItem (curCtaCte.campo(0))
      curCtaCte.MoverSiguiente
    Loop
  
  End If
  
  'Se cierra el cursor
  curCtaCte.Cerrar
End If

End Sub

Private Sub cboCodMov_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCodMov) = True Then SendKeys "{down}"

End Sub

Private Sub cbocodMov_Click()

' Habilita el txtAfecta
 txtAfecta.Enabled = True

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCodMov.ListIndex) = False And cboCodMov.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCodMov_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCodMov_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCodMov, vbWhite) = True Then
  ' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboCodMov.Text, txtCodMov, mcolCodMov, mcolCodDesCodMov

Else '  Vaciar Controles enlazados al combo
    txtCodMov.Text = Empty
End If

'Cambia el alto del combo
cboCodMov.Height = CBONORMAL

End Sub

Private Sub cboCtaContable_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaContable) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaContable_Click()

' Verifica si el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaContable.ListIndex) = False And cboCtaContable.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCtaContable_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica si es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCtaContable_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCtaContable, vbWhite) = True Then
    
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
      CD_ActCod cboCtaContable.Text, txtCodContable, mcolCodPlanCont, mcolDesCodPlanCont
      
Else
  'NO se encuentra la CtaContable
  txtCodContable = Empty
End If

'Cambia el alto del combo
cboCtaContable.Height = CBONORMAL

End Sub

Private Sub cboCtaCte_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaCte) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaCte_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaCte.ListIndex) = False And cboCtaCte.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCtaCte_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCtaCte_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCtaCte, Obligatorio) = True Then
    
    'Se actualiza código (String) correspondiente a descripción introducida
    CD_ActCboVar cboCtaCte.Text, msCtaCte, mcolCodCtaCte, mcolCodDesCtaCte
      
    ' Verifica SI el campo esta vacio
    If cboCtaCte.Text <> "" Then
       ' Los campos coloca a color blanco
       cboCtaCte.BackColor = vbWhite
       txtBanco.BackColor = vbWhite
    Else
       'Marca los campos obligatorios
       cboCtaCte.BackColor = Obligatorio
    End If
Else
  'No se encuentra la CtaCte
  msCtaCte = ""
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

'Cambia el alto del combo
cboCtaCte.Height = CBONORMAL

End Sub

Private Sub cboTipDoc_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboTipDoc) = True Then SendKeys "{down}"

End Sub

Private Sub cboTipDoc_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboTipDoc.ListIndex) = False And cboTipDoc.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboTipDoc_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboTipDoc_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboTipDoc, vbWhite) = True Then
  ' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboTipDoc.Text, txtTipDoc, mcolCodTipDoc, mcolCodDesTipDoc

Else '  Vaciar Controles enlazados al combo
    txtTipDoc.Text = Empty
End If

'Cambia el alto del combo
cboTipDoc.Height = CBONORMAL

End Sub

Private Sub cmdAceptar_Click()

' Verifica si existe un documento duplicado
If VerificarDocExiste Then
  ' El documento ya ha sido ingresado mandamos mensaje
  If MsgBox("El Número de Documento está duplicado, ¿desea continuar con este mismo número de documento? ", _
        vbQuestion + vbYesNo, _
        "Caja-Bancos- Ingresos a Caja-Bancos") = vbNo Then
        ' Pone el focus a código del ingreso
        txtDocIngreso.SetFocus
        Exit Sub
  End If
End If

' Verifica si los datos son correctos
If fbVerificarDatosIntroducidos = False Then
    ' Algún dato es incorrecto
    Exit Sub
End If

' Verifica el tipo de operación a realizar
Select Case gsTipoOperacionIngreso
Case "Modificar"
    
        ' Mensaje de conformidad de los datos
         If MsgBox("¿Está conforme con las modificaciones realizadas en el Ingreso " & txtCodIngreso.Text & "?", _
                    vbQuestion + vbYesNo, "Caja-Bancos-Modificación de Ingresos") = vbYes Then
            'Actualiza la transaccion
             Var8 1, gsFormulario
           
            ' Modifica los datos ingresados en caja o bancos
            ModificarRegistro
         Else: Exit Sub
         End If
         
Case "Nuevo"
  
      ' Mensaje de conformidad
      If MsgBox("¿Está conforme con los datos?", vbQuestion + vbYesNo, _
                "Caja-Bancos- Ingresos") = vbYes Then
          'Actualiza la transaccion
          Var8 1, gsFormulario
                
          'Guarda el registro de ingreso en Caja o Bancos
          GuardarIngreso
      Else: Exit Sub
      End If

End Select

' Limpia la las cajas de texto del formulario
  LimpiarFormulario


Select Case gsTipoOperacionIngreso
Case "Modificar"
           
           'Oculta controles pendientes
           txtMontoPendiente.Visible = False
           lblMontoPendiente.Visible = False
           
            ' cierra el control egreso
            If mbIngresoCargado Then
             curRegIngresoCajaBanco.Cerrar
             mbIngresoCargado = False
             mskFecTrab = "__/__/____"
            End If
   
          'Prepara  el formulario para una nueva modificacion
           ModificarIngreso
           txtCodIngreso.SetFocus

Case "Nuevo"
                  
          'Prepara el Formulario para un nuevo ingreso
          NuevoIngreso

End Select

End Sub
    
Private Function fbVerificarDatosIntroducidos()
' -------------------------------------------------------
' Propósito: Verifica que los datos introducidos sean correctos
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------
Dim dblMontoDet As Double

'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    'Devuelve el resultado
    fbVerificarDatosIntroducidos = False
    Exit Function
End If

' Verifica que lo que sale de caja-bancos sea Menor que el saldo de Caja-bancos
 If gsTipoOperacionIngreso = "Modificar" Then
    
    ' Verifica si registro de ingreso esta relacionado con Egresos en dólares
    If txtMontoPendiente.Visible = True Then
        
        ' Verifica si monto modificado excede lo pendiente
        If Val(curRegIngresoCajaBanco.campo(4)) + Val(Var37(txtMontoPendiente.Text)) < Val(Var37(txtMonto.Text)) Then
            MsgBox "El Monto Ingresado excede el Pendiente de ingreso ", _
            vbExclamation + vbOKOnly, "Caja-Bancos- Modificación de Ingresos"
            txtMonto.SetFocus
            fbVerificarDatosIntroducidos = False
            Exit Function
        End If
        
    End If ' Verifica si es un ingreso generado por ctas extrangeras

    ' Verifica la conformidad con el saldo
    If msCajaoBanco = "CA" Then   'Caja
        If Val(CalcularTotalIngresosCB - Val(curRegIngresoCajaBanco.campo(4)) + Val(Var37(txtMonto.Text))) < Val(CalcularTotalEgresosCB) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "No se puede modificar, saldo insuficiente en Caja", , "SGCcaijo-Ingresos de Caja-Bancos"
            fbVerificarDatosIntroducidos = False
            If txtMonto.Enabled = True Then txtMonto.SetFocus
            Exit Function
        End If
    
    ElseIf (msCajaoBanco = "BA" And (curRegIngresoCajaBanco.campo(10) = msCtaCte)) Then 'La misma CtaCte
        If Val(CalcularTotalIngresosCB - Val(curRegIngresoCajaBanco.campo(4)) + Val(Var37(txtMonto.Text))) < Val(CalcularTotalEgresosCB) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "No se puede modificar, saldo insuficiente en la Cta-" & cboCtaCte.Text, , "SGCcaijo-Ingresos de CajaBancos"
            fbVerificarDatosIntroducidos = False
            If txtMonto.Enabled = True Then txtMonto.SetFocus
            Exit Function
        End If
    
    ElseIf (msCajaoBanco = "BA" And (curRegIngresoCajaBanco.campo(10) <> msCtaCte)) Then 'Se Cambio de CtaCte
        If Val(CalcularTotalIngresosCB - Val(curRegIngresoCajaBanco.campo(4))) < Val(CalcularTotalEgresosCB) Then
           ' Mensaje ,saldo insuficiente
           MsgBox "No se puede modificar, saldo insuficiente en la Cta-" & mcolCodDesCtaCte(curRegIngresoCajaBanco.campo(10)), , "SGCcaijo-Ingresos de CajaBancos"
           fbVerificarDatosIntroducidos = False
           If txtMonto.Enabled = True Then txtMonto.SetFocus
          Exit Function
        End If

    End If
 End If

' Verificados los datos
fbVerificarDatosIntroducidos = True

End Function
    
  Private Sub ModificarRegistro()
'----------------------------------------------------------------------------
'Propósito: Modifica los datos del registro Ingreso a Caja o a Bancos
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el Click Aceptar
Dim sSQL, sAfectaAnt As String
Dim modIngreCajaBanco As New clsBD3
Dim sCodEgresoCtasDol As String
Dim dblMontoAnt As Double
Dim sPendiente As String
' I.NumDoc,I.IdEgreso,I.IdTipoDoc,I.FecMov,I.Monto,I.CodMov, _
  I.CodContable, TM.Afecta, I.Observ, CTA.IdBanco,I.IdCta
  
' Asigna el código del egreso y el monto anterior
  sCodEgresoCtasDol = curRegIngresoCajaBanco.campo(1)
  dblMontoAnt = curRegIngresoCajaBanco.campo(4)
    
 'Carga la sentencia que modifica el registro de ingreso
If msCajaoBanco = "BA" Then
     ' Guardar los  datos
     sSQL = "UPDATE INGRESOS SET " & _
        "NumDoc='" & txtDocIngreso & "'," & _
        "IdTipoDoc='" & txtTipDoc & "'," & _
        "Monto=" & Var37(txtMonto.Text) & "," & _
        "CodMov='" & txtCodMov.Text & "'," & _
        "Observ='" & txtObserv.Text & "'," & _
        "IdCta='" & msCtaCte & "'," & _
        "CodContable='" & txtCodContable.Text & "' " & _
        "WHERE Orden='" & txtCodIngreso.Text & "'"
        
       ' carga la colección asiento
       'Orden,Monto,NumCtaBanc, fecha, observ,Proceso
        gcolAsiento.Add _
        Key:=txtCodIngreso, _
        Item:=txtCodIngreso & "¯" _
          & Var37(txtMonto) & "¯" _
          & msCtaCte & "¯" _
          & FechaAMD(mskFecTrab.Text) & "¯INGRESO A BANCO¯IN¯IB¯B"

        
Else
        ' Guardar los  datos
      sSQL = "UPDATE INGRESOS SET " & _
         "NumDoc='" & txtDocIngreso & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "Monto=" & Var37(txtMonto.Text) & "," & _
         "CodMov='" & txtCodMov.Text & "'," & _
         "Observ='" & txtObserv.Text & "'," & _
         "CodContable='" & txtCodContable.Text & "' " & _
         "WHERE Orden='" & txtCodIngreso.Text & "'"

         ' carga la colección asiento
         'Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ,Proceso
         gcolAsiento.Add _
         Key:=txtCodIngreso, _
         Item:=txtCodIngreso & "¯" _
           & Var37(txtMonto) & "¯" _
           & msCtaCte & "¯" _
           & FechaAMD(mskFecTrab.Text) & "¯INGRESO A CAJA¯IN¯IC¯C"

         
End If

'SI al ejecutar hay error se sale de la aplicación
modIngreCajaBanco.SQL = sSQL
If modIngreCajaBanco.Ejecutar = HAY_ERROR Then
 End
End If
  
'Se cierra el componente de mod
modIngreCajaBanco.Cerrar

'Modificar el Mov afectado
ModificarMovAfectado

' LLama a la modificación del asiento automático
Conta19

' Si esta relacionado con Egresos en dólares
If txtMontoPendiente.Visible = True Then
  ' Si se modificó monto guarda la diferencia en pendientes
  GuardarMontoPendiente sCodEgresoCtasDol, dblMontoAnt, Val(Var37(txtMonto)), sPendiente
End If

 'Actualiza la transaccion
 Var8 -1, Empty
 
 ' Msg Ok
 MsgBox "Operación efectuada correctamente", , "SGCCaijo-Ingreso a Caja-Bancos"
 
 ' Si existen la operación genera Pendientes, Pregunta SI desea Ingresarlo
 If sPendiente = "SI" Then
    ' Pregunta SI el Usuario quiere ingresar el Monto en soles generado por el Egreso en caja o bancos
    If MsgBox("Existe un Monto Pendiente relacionado con el egreso " & sCodEgresoCtasDol & " de Ctas en dólares" & Chr(13) & _
              "¿Desea realizar el ingreso a Caja-Bancos?", _
              vbQuestion + vbYesNo, "Caja-Bancos- Egresos") = vbYes Then

        ' Muestra el formulario Ingreso a Caja Bancos por Egreso en Cts Extranjeras
        gsCodEgreso = Trim(sCodEgresoCtasDol)
        frmCBINxEGMndExt.Show vbModal, Me
        
     End If
  End If

End Sub

Private Sub ModificarMovAfectado()
'----------------------------------------------------------------------------
'Propósito: Modifica  Egreso relacionado con Terceros o Personal
'Recibe:  sPersTercsAnt string que indica la SI el egreso estaba relacioado con
'         la tabla Terceros,Pln_Personal o ninguno
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el Click Aceptar al modificar el registro
Dim sSQL As String
Dim modMovAfectado As New clsBD3

'SI NO ha cambiado de Afecta, actualizar el registro Mov relacionado (Terceros o Personal)
If msAfecta = "Tercero" Then
    sSQL = "UPDATE MOV_TERCEROS SET IdTercero='" _
          & txtAfecta.Text & "' WHERE Orden='" & txtCodIngreso.Text & "'"
    modMovAfectado.SQL = sSQL
    
    If modMovAfectado.Ejecutar = HAY_ERROR Then
       End
    End If
    
    modMovAfectado.Cerrar
    
    ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                              & Var37(txtMonto), _
                         Key:=txtCodContable
                   
ElseIf msAfecta = "Persona" Then
    sSQL = "UPDATE MOV_PERSONAL SET IdPersona='" _
                & txtAfecta.Text & "' WHERE Orden='" & txtCodIngreso.Text & "'"
    modMovAfectado.SQL = sSQL
    
    If modMovAfectado.Ejecutar = HAY_ERROR Then
       End
    End If
    
    modMovAfectado.Cerrar
    
    ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                              & Var37(txtMonto), _
                         Key:=txtCodContable
                         
'Verifica si es proceso
ElseIf msAfecta = "Proceso" Then
    Select Case msProceso
    
    Case "DEVOLUCION_PRESTAMOS"
    
        'Actualizan los prestamos
        ActualizaPrestamos
    Case "DEVOLUCION_RENDIR"
        'Actualizan los prestamos
        ActualizaEntregas
    End Select
    
Else
   'NO afecta a mas tablas de la BD
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                           & Var37(txtMonto), _
                       Key:=txtCodContable
End If

End Sub
 
Private Sub GuardarMontoPendiente(ByVal sCodEgresoCtasDol As String, ByVal dblMontoAnt As Double, _
                                  ByVal dblMonto As Double, sPendiente As String)
'----------------------------------------------------------------------------
'Propósito: Modifica el Monto pendiente en la tabla EGRESO_CTAS_EXTR SI se cambio el
'           Monto del ingreso y este esta relacionado con un egreso de Ctas en dólares
'Recibe:    dblMontoAnt, dblMonto monto original y monto modificado del reg
'           sIdEgresoCtasDol string que identifica el codigo del egreso de Ctas Dol
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota: Llamado desde el Click Aceptar o click Anular
Dim sSQL As String
Dim modMontoPendienteCtasDol As New clsBD3

'Inicializamos suponiendo que existen pendientes
 sPendiente = "SI"

If gsTipoOperacionIngreso = "Modificar" Then     ' Guardar los Montos pendientes
    ' Cuando Monto Modificado es Igual a txtMontoPendiente del egreso de Ctas en dólares sPendiente es "NO"
    If Format(dblMontoAnt - dblMonto + Val(Var37(txtMontoPendiente.Text)), "###,###,##0.00") = "0.00" Then
        sPendiente = "NO"
    End If
    sSQL = "UPDATE EGRESO_CTAS_EXTR SET " & _
       "MontoSol=MontoSol+" & Format(dblMontoAnt - dblMonto, "########0.00") & ", " & _
       "IngrePendiente='" & sPendiente & "' " & _
       "WHERE IdEgreso='" & curRegIngresoCajaBanco.campo(1) & "'"
Else
    sSQL = "UPDATE EGRESO_CTAS_EXTR SET " & _
       "MontoSol=MontoSol+" & Format(dblMontoAnt, "########0.00") & ", " & _
       "IngrePendiente='" & sPendiente & "' " & _
       "WHERE IdEgreso='" & curRegIngresoCajaBanco.campo(1) & "'"
End If
                   
'SI al ejecutar hay error se sale de la aplicación
modMontoPendienteCtasDol.SQL = sSQL
If modMontoPendienteCtasDol.Ejecutar = HAY_ERROR Then End

'Se cierra la query
modMontoPendienteCtasDol.Cerrar
        
End Sub

 Private Sub GuardarIngreso()
'----------------------------------------------------------------------------
'Propósito: Guarda el ingreso en la tabla de Ingreso
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el Click Aceptar
Dim sSQL As String
Dim modIngreCajaBanco As New clsBD3

'Verifica SI es a Caja
If msCajaoBanco = "CA" Then
            
    ' Guardar los  datos a Caja cuando no es ingreso de prestamos
    sSQL = "INSERT INTO INGRESOS VALUES('" & txtCodIngreso & "','" _
            & txtDocIngreso.Text & "','" & txtTipDoc.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
            & Var37(txtMonto.Text) & ",'','','NO','" & txtObserv.Text & "','" _
            & txtCodContable.Text & "')"

    ' carga la colección asiento
    'Orden,Monto,NumCtaBanc, fecha, observ, RelacProc
    gcolAsiento.Add _
    Key:=txtCodIngreso.Text, _
    Item:=txtCodIngreso.Text & "¯" _
      & Var37(txtMonto.Text) _
      & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
      & "¯INGRESO A CAJA¯IN¯IC¯C"
          
Else
    
    'Se graba en Banco, cuando no es un ingreso de prestamos
    sSQL = "INSERT INTO INGRESOS VALUES('" & txtCodIngreso & "','" _
            & txtDocIngreso.Text & "','" & txtTipDoc.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
            & Var37(txtMonto.Text) & ",'','" _
            & msCtaCte & "','NO','" & txtObserv.Text & "','" _
            & txtCodContable.Text & "')"
        
    ' carga la colección asiento
    'Orden,Monto,NumCtaBanc, fecha, observ, RelacProc
    gcolAsiento.Add _
    Key:=txtCodIngreso.Text, _
    Item:=txtCodIngreso.Text & "¯" _
      & Var37(txtMonto.Text) _
      & "¯" & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
      & "¯INGRESO A BANCO¯IN¯IB¯B"
                     
End If
  
'SI al ejecutar hay error se sale de la aplicación
modIngreCajaBanco.SQL = sSQL
If modIngreCajaBanco.Ejecutar = HAY_ERROR Then
 End
End If
   
'Se cierra la query
modIngreCajaBanco.Cerrar

'SI afecta es T la relacion es con terceros
GuardarMovAfectado

' Genera el asiento automático
Conta13

If CancelarVenta = True Then
  PagarVenta
End If

 'Actualiza la transaccion
 Var8 -1, Empty
 
 ' Msg Ok
 MsgBox "Operación efectuada correctamente", , "SGCCaijo-Ingreso a Caja-Bancos"

End Sub

Private Sub GuardarMovAfectado()
'----------------------------------------------------------------------------
'Propósito  : Guarda el Ingreso relacionado con Terceros o Personal
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el Click Aceptar al hacer un nuevo Ingreso
Dim modAfectado As New clsBD3
Dim sSQL As String

If msAfecta = "Tercero" Then
  'Cargamos sentencia que guarda en BD MOV_TERCERO
  sSQL = "INSERT INTO MOV_TERCEROS VALUES('" _
          & txtCodIngreso.Text & "','" _
          & Trim(txtAfecta.Text) & "')"
  
  'Copia la sentencia SQL
  modAfectado.SQL = sSQL

  'Verifica si hay eror
  If modAfectado.Ejecutar = HAY_ERROR Then
     End 'Finaliza la aplicacion indicando el error en SQL
  End If
    
  'Cierra el cursor
  modAfectado.Cerrar
          
  ' Carga el código contable del movimiento en colDetAsiento
  gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
                       
ElseIf msAfecta = "Persona" Then
  'Cargamos sentencia que guarda en BD MOV_PERSONAL
    sSQL = "INSERT INTO MOV_PERSONAL VALUES('" _
          & txtCodIngreso.Text & "','" _
          & Trim(txtAfecta.Text) & "')"
    
    'Copia la sentencia SQL
    modAfectado.SQL = sSQL
    
    'Verifica si hay eror
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    
    'Cierra el cursor
    modAfectado.Cerrar

    ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
                       
ElseIf msAfecta = "Proceso" Then
    Select Case msProceso
    
    Case "DEVOLUCION_PRESTAMOS"
    
        'Actualiza los prestamos
        ActualizaPrestamos
    Case "DEVOLUCION_RENDIR"
    
        'Actualizan los prestamos
        ActualizaEntregas
    End Select
                       
Else
   'NO afecta a mas tablas de la BD
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                           & Var37(txtMonto), _
                       Key:=txtCodContable
End If

End Sub

Private Sub ActualizaPrestamos()
'---------------------------------------------------------------------
'Propósito  : Actualiza los datos del prestamo relacionados al ingreso
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------------
Dim sSQL As String
Dim ModPrestamos_Cuota As New clsBD3
Dim ModDevolucion_Prestamos As New clsBD3
Dim modPrestamos As New clsBD3
Dim i As Integer
Dim MiObjeto As Variant
Dim MiObjeto1 As Variant
Dim MiObjeto2 As Variant

For Each MiObjeto In gcolDetPrestamos

    ' Guardar los  datos
    sSQL = ""
    sSQL = "UPDATE PRESTAMOS_CUOTAS SET " & _
            "Amortizado=" & Var30(MiObjeto, 4) & ", " & _
            "Cancelado='SI' " & _
            "WHERE IdPersona='" & txtAfecta.Text & "' And IdConPL= " & _
            " '" & Var30(MiObjeto, 1) & "' And " & _
            "NumPrestamo= '" & Var30(MiObjeto, 2) & "' And " & _
            "AnioMes='" & Var30(MiObjeto, 3) & "' "
            
    'Copia la sentencia SQL
    ModPrestamos_Cuota.SQL = sSQL
    
    'Ejecuta la sentencia SQL
    If ModPrestamos_Cuota.Ejecutar = HAY_ERROR Then
        'Termina la ejecucion
        End
    End If
    
    'Cierra el cursor
    ModPrestamos_Cuota.Cerrar

'Mueve al siguiente registro de la modificación
Next


'Guarda en la Base de Datos, tabla Prestamos
For Each MiObjeto1 In gcolPrestamos

    ' Guardar los  datos
    sSQL = ""
    sSQL = "UPDATE PRESTAMOS SET " & _
            "Cancelado='SI' " & _
            "WHERE IdPersona='" & txtAfecta.Text & "' And IdConPL= " & _
            " '" & Var30(MiObjeto1, 1) & "' And " & _
            "NumPrestamo= '" & Var30(MiObjeto1, 2) & "' "
                        
    'Copia la sentencia SQL
    modPrestamos.SQL = sSQL
    
    'Ejecuta la sentencia SQL
    If modPrestamos.Ejecutar = HAY_ERROR Then
        'Termina la ejecucion
        End
    End If
    
    'Cierra el cursor
    modPrestamos.Cerrar

    ' Carga códigos contables del movimiento en detAsiento
    gcolAsientoDet.Add Item:=Var30(MiObjeto1, 3) & "¯" _
                           & Var37(txtMonto), _
                       Key:=Var30(MiObjeto1, 3)
                   
'Mueve al siguiente registro de la modificación
Next

'Verifica el tipo de operacion de ingresos
If gsTipoOperacionIngreso = "Nuevo" Then
  
    'Guarda en la Base de datos, tabla Devoluciones_Prestamo
    For Each MiObjeto2 In gcolDetPrestamos
    
        
        ' Guardar los  datos
        sSQL = ""
        sSQL = "INSERT INTO DEVOLUCION_PRESTAMOSCB VALUES ('" & _
                txtAfecta.Text & "', '" & Var30(MiObjeto2, 1) & _
                "' , '" & Var30(MiObjeto2, 2) & _
                "', '" & Var30(MiObjeto2, 3) & "', '" & txtCodIngreso.Text & "')"
    
        'Copia la sentencia SQL
        ModDevolucion_Prestamos.SQL = sSQL
        
        'Ejecuta la sentencia SQL
        If ModDevolucion_Prestamos.Ejecutar = HAY_ERROR Then
            'Termina la ejecucion
            End
        End If
        
        'Cierra el cursor
        ModDevolucion_Prestamos.Cerrar
        
        Exit For
    'Mueve al siguiente registro de la modificación
    Next
End If

'Vacia las colecciones
Set gcolPrestamos = Nothing
Set gcolDetPrestamos = Nothing

End Sub

Private Sub ActualizaEntregas()
'---------------------------------------------------------------------
'Propósito  : Actualiza los datos de las entregas a rendir
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------------
Dim sSQL As String
Dim modEntregas As New clsBD3
Dim curEntregas As New clsBD2

'Verifica el tipo de operacion de ingresos
If gsTipoOperacionIngreso = "Nuevo" Then

    ' Guardar los  datos
    sSQL = ""
    sSQL = "INSERT INTO MOV_ENTREG_RENDIR VALUES(" & _
           "'" & txtCodIngreso.Text & "', " & _
           "'" & txtAfecta.Text & "', " & _
           "'E',0, " & Var37(txtMonto.Text) & ", 'NO', '" & FechaAMD(mskFecTrab) & "')"
Else
    sSQL = ""
    sSQL = "UPDATE MOV_ENTREG_RENDIR SET " & _
            "IdPersona='" & txtAfecta.Text & "', " & _
            "Egreso=" & Var37(txtMonto.Text) & " " & _
            "WHERE Orden='" & txtCodIngreso.Text & "' "

End If

'Copia la sentencia SQL
modEntregas.SQL = sSQL

'Ejecuta la sentencia SQL
If modEntregas.Ejecutar = HAY_ERROR Then
    'Termina la ejecucion
    End
End If

'Cierra el cursor
modEntregas.Cerrar


' Guardar los  datos
sSQL = ""
sSQL = "SELECT CodContable " & _
       "FROM CTB_ENTREG_RENDIR " & _
       "WHERE IdCTBERendir='01'"
        
'Copia la sentencia SQL
curEntregas.SQL = sSQL

'Ejecuta la sentencia SQL
If curEntregas.Abrir = HAY_ERROR Then
    'Termina la ejecucion
    End
End If

' Carga códigos contables del movimiento en detAsiento
gcolAsientoDet.Add Item:=curEntregas.campo(0) & "¯" _
                         & Var37(txtMonto.Text), _
                        Key:=curEntregas.campo(0)
'Cierra el cursor
curEntregas.Cerrar

End Sub

Private Function VerificarDocExiste() As Boolean
'--------------------------------------------------------------------
'Propósito: Verifica SI el Doc ha sido ingresado en caja o bancos, SI NO
'Recibe:    Nada
'Devuelve:  false:No existe, True: Existe
'Nota:      llamado desde el evento click de Aceptar
'--------------------------------------------------------------------

Dim sSQL As String
Dim curDocIngresado As New clsBD2

VerificarDocExiste = False

'Verifica SI el doc ingresado sea el mismo del registro en modificacion
If gsTipoOperacionIngreso = "Modificar" Then
    If txtDocIngreso.Text = curRegIngresoCajaBanco.campo(0) Then ' es el mismo del registro, NO hace nada
        Exit Function 'Sale de la funcion
    End If
End If

'Verifica SI el Doc esta en Caja o en Banco de la tabla ingresos
'Se averigua SI existe algun documento con el mismo numero en Banco
sSQL = "SELECT Count(I.NumDoc) as NroDoc FROM INGRESOS I " & _
       "WHERE I.NumDoc = '" & txtDocIngreso.Text & "'"
curDocIngresado.SQL = sSQL
If curDocIngresado.Abrir = HAY_ERROR Then
  End
End If
'Se encontró en Banco
If curDocIngresado.campo(0) <> 0 Then
    curDocIngresado.Cerrar
    VerificarDocExiste = True
    Exit Function
End If

'Se cierra el cursor
curDocIngresado.Cerrar

End Function


Private Sub cmdAnular_Click()
Dim modAnularIngreCajaBanco As New clsBD3
Dim modAnularEntrega As New clsBD3
Dim sSQL, sCodEgresoCtasDol As String
Dim dblMontoOriginal As Double
Dim sPendiente As String
'Actualiza variable de operación
HabilitaDeshabilitaBotones ("Anular")
dblMontoOriginal = curRegIngresoCajaBanco.campo(4)
sCodEgresoCtasDol = curRegIngresoCajaBanco.campo(1)

'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    'Devuelve el resultado
    Exit Sub
End If

'Valida el ingreso de caja con egreso de caja
If CalcularTotalIngresosCB() - curRegIngresoCajaBanco.campo(4) < CalcularTotalEgresosCB() Then
   If msCajaoBanco = "CA" Then ' Ingreso a caja
     ' Mensaje ,saldo insuficiente
      MsgBox "No se puede modificar, saldo insuficiente en Caja", , "SGCcaijo-Ingresos de Caja-Bancos"
   
   Else ' Ingreso a bancos
     ' Mensaje que no se puede realizar la anulación del registro
      MsgBox "No se puede Anular, saldo insuficiente en la Cta-" & mcolCodDesCtaCte(curRegIngresoCajaBanco.campo(10)), , "SGCcaijo-Ingresos de CajaBancos"
   End If
Else

    'Preguntar SI desea Anular el registro de Ingreso a Banco
    'Mensaje de conformidad de los datos
    If MsgBox("¿Seguro que desea anular el registro de ingreso " & txtCodIngreso.Text & "?", _
                  vbQuestion + vbYesNo, "Caja-Bancos-Anulación de Ingresos") = vbYes Then
            'Actualiza la transaccion
             Var8 1, gsFormulario
                  
            If msCajaoBanco = "BA" Then
              'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
               sSQL = "UPDATE INGRESOS SET " & _
                    "Anulado='SI'" & _
                    "WHERE Orden='" & txtCodIngreso.Text & "'"
                
               ' carga la colección asiento
               'Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, RelacProc
               gcolAsiento.Add _
               Key:=txtCodIngreso.Text, _
               Item:=txtCodIngreso.Text & "¯" _
                 & Var37(txtMonto.Text) _
                 & "¯" & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
                 & "¯INGRESO A BANCO¯IN¯IB¯B"
            
            Else
                'Cambiar el campo Anulado de Ingresos a "SI", los demás campos a anulado y cero
                 sSQL = "UPDATE INGRESOS SET " & _
                    "Anulado='SI'" & _
                    "WHERE Orden='" & txtCodIngreso.Text & "'"
                
                ' carga la colección asiento
                'Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, RelacProc
                gcolAsiento.Add _
                Key:=txtCodIngreso.Text, _
                Item:=txtCodIngreso.Text & "¯" _
                  & Var37(txtMonto.Text) _
                  & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
                  & "¯INGRESO A CAJA¯IN¯IC¯C"
                  
            End If
                    
            'SI al ejecutar hay error se sale de la aplicación
             modAnularIngreCajaBanco.SQL = sSQL
             If modAnularIngreCajaBanco.Ejecutar = HAY_ERROR Then
              End
             End If
            
            If msProceso = "DEVOLUCION_PRESTAMOS" Then
            
                'Restaura los Prestamos asociados a ingresos
                RestauraPrestamos
            ElseIf msProceso = "DEVOLUCION_RENDIR" Then
                'Cambiar el campo Anulado de MOV_ENTREG_RENDI a "SI", los demás campos a anulado y cero
                 sSQL = "UPDATE MOV_ENTREG_RENDIR SET " & _
                    "Anulado='SI'" & _
                    "WHERE Orden='" & txtCodIngreso.Text & "'"
                    
                'SI al ejecutar hay error se sale de la aplicación
                 modAnularEntrega.SQL = sSQL
                 If modAnularEntrega.Ejecutar = HAY_ERROR Then
                  End
                 End If
                 'Cierra el cursor
                 modAnularEntrega.Cerrar
            End If
            
            ' Anula los asientos automáticos
            Conta22
                
             'Se cierra la query
             modAnularIngreCajaBanco.Cerrar
             
             ' Si esta relacionado con un egreso se guarda en pendientes
             If txtMontoPendiente.Visible = True Then GuardarMontoPendiente sCodEgresoCtasDol, dblMontoOriginal, 0, sPendiente
             
             'Actualiza la transaccion
             Var8 -1, Empty
             
             ' Msg Ok
             MsgBox "Operación efectuada correctamente", , "SGCCaijo-Ingreso a Caja-Bancos"
             
             ' Si existen la operación genera Pendientes, Pregunta SI desea Ingresarlo
             If sPendiente = "SI" Then
                ' Pregunta SI el Usuario quiere ingresar el Monto en soles generado por el Egreso en caja o bancos
                If MsgBox("Existe un Monto Pendiente relacionado con el egreso " & sCodEgresoCtasDol & " de Ctas en dólares" & Chr(13) & _
                          "¿Desea realizar el ingreso a Caja-Bancos?", _
                          vbQuestion + vbYesNo, "Caja-Bancos- Egresos") = vbYes Then
            
                    ' Muestra el formulario Ingreso a Caja Bancos por Egreso en Cts Extranjeras
                    gsCodEgreso = Trim(sCodEgresoCtasDol)
                    frmCBINxEGMndExt.Show vbModal, Me
                    
                 End If
              End If
             
            'Oculta controles pendientes
             txtMontoPendiente.Visible = False
             lblMontoPendiente.Visible = False
            
            ' Limpia la pantalla para una nueva operación, Prepara el formulario
            LimpiarFormulario
            
            ' cierra el control egreso
            If mbIngresoCargado Then
               curRegIngresoCajaBanco.Cerrar
               mbIngresoCargado = False
               mskFecTrab = "__/__/____"
            End If
    
            'Prepara  el formulario para una nueva modificacion
            ModificarIngreso
            txtCodIngreso.SetFocus

    End If
    
End If 'Fin de CalcularTotalIngresos() - curRegIngresoCajaBanco.campo(4) < CalcularTotalEgresos()

End Sub

Private Function CalcularTotalEgresosCB() As Double
'-----------------------------------------------------
'Propósito  : Determina la suma de montos de los Egresos
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim sSQL As String
Dim curTotalMonto As New clsBD2

'Verifica si el ingreso que se realiza es de caja
If Left(txtCodIngreso.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
           & "FROM EGRESOS E " _
           & "WHERE Left(E.Orden,2)='CA' and E.Anulado='NO' and E.Origen='C'"
       
Else
    If curRegIngresoCajaBanco.campo(10) = msCtaCte Then ' misma cuenta
     'El egreso que se calcula es de banco de la cta msCtaCte
        sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
        & "FROM EGRESOS E " _
        & "WHERE Left(E.Orden,2)='BA' And E.IdCta='" & msCtaCte & "' " _
        & "And E.Anulado='NO' and E.Origen='B'"
    Else ' cuentas diferentes
    'El egreso que se calcula es de banco de la cta original
        sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
        & "FROM EGRESOS E " _
        & "WHERE Left(E.Orden,2)='BA' And E.IdCta='" & curRegIngresoCajaBanco.campo(10) & "' " _
        & "And E.Anulado='NO' And E.Origen='B'"
    End If
End If

'Copia la sentencia SQL
curTotalMonto.SQL = sSQL

'Verifica si hay error en el proceso
If curTotalMonto.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si la consulta tiene valores nulos
If Not IsNull(curTotalMonto.campo(0)) Then
    
    'Copia el resultado de la consulta
    CalcularTotalEgresosCB = curTotalMonto.campo(0)
End If

'Cierra el cursor
curTotalMonto.Cerrar

End Function

Private Function CalcularTotalIngresosCB() As Double
'-----------------------------------------------------
'Propósito  : Determina la suma de montos de los ingresos
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim sSQL As String
Dim curTotalMonto As New clsBD2

'Sentencia SQL
sSQL = ""
'Verifica si se esta calculando el total de caja
If Left(txtCodIngreso.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='CA' And I.Anulado='NO'"
Else
    If curRegIngresoCajaBanco.campo(10) = msCtaCte Then ' la cuenta es la misma
        'El total que se calcula es de banco de la cuenta msCtaCte
        sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='BA' And I.IdCta= '" & msCtaCte & "' And I.Anulado='NO'"
    Else ' la cuenta es Plan28
        'El total que se calcula es de banco de la cuenta original
        sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='BA' And I.IdCta= '" & curRegIngresoCajaBanco.campo(10) & "' And I.Anulado='NO'"
    End If
End If

'Copia la sentencia SQL
curTotalMonto.SQL = sSQL

'Verifica si hay error en el proceso
If curTotalMonto.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si la consulta tiene valores nulos
If Not IsNull(curTotalMonto.campo(0)) Then
    
    'Copia el resultado de la consulta
    CalcularTotalIngresosCB = curTotalMonto.campo(0)
End If

'Cierra el cursor
curTotalMonto.Cerrar

End Function

Private Sub RestauraPrestamos()
'---------------------------------------------------------------
'Propósito  : Restaura los prestamos relacionados con los ingresos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim curRestPrestamos As New clsBD2
Dim modRestPrestamos As New clsBD3

'Sentencia SQL
sSQL = ""
sSQL = "SELECT  IdConPl, NumPrestamo, AñoMes " & _
       "FROM DEVOLUCION_PRESTAMOSCB " & _
       "WHERE Orden = '" & txtCodIngreso.Text & "' "
       
'Copia la sentencia SQL
curRestPrestamos.SQL = sSQL

'Verifica si hay error
If curRestPrestamos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Restaura la tabla Prestamos
RestauraTablaPrestamos curRestPrestamos.campo(0), curRestPrestamos.campo(1)

'Mientras no sea fin de registro
Do While Not curRestPrestamos.EOF

    'Restaura la tabla Prestamos_Cuotas
    RestauraPrestamosCuotas curRestPrestamos.campo(0), curRestPrestamos.campo(1), curRestPrestamos.campo(2)
      
    'Elimina la tabla Devolucion_PrestamosCB
    EliminaDevolucionPrestamos curRestPrestamos.campo(0), curRestPrestamos.campo(1), curRestPrestamos.campo(2)
    
    'Mueve al siguiente registro
    curRestPrestamos.MoverSiguiente

Loop

'Cierra el curRestPrestamos
curRestPrestamos.Cerrar
 
End Sub

Private Sub EliminaDevolucionPrestamos(sIdConPl As String, sNumPrestamo As String, sAnioMes As String)
'---------------------------------------------------------------
'Propósito  : Elimina la tabla Devolucion_Prestamos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim modPrestamosDevolucion As New clsBD3

'Sentencia SQL
sSQL = ""
sSQL = "DELETE * " & _
       "FROM DEVOLUCION_PRESTAMOSCB " & _
       "WHERE IdPersona='" & txtAfecta.Text & "' And IdConPL= '" & sIdConPl & "' And " & _
       " NumPrestamo= '" & sNumPrestamo & "' And AnioMes ='" & sAnioMes & "' "

'Copia la sentencia SQL
modPrestamosDevolucion.SQL = sSQL

'Ejecuta la sentencia SQL
If modPrestamosDevolucion.Ejecutar = HAY_ERROR Then
    'Termina la ejecucion
    End
End If

'Cierra el cursor
modPrestamosDevolucion.Cerrar

End Sub

Private Sub RestauraPrestamosCuotas(sIdConPl As String, sNumPrestamo As String, sAnioMes As String)
'---------------------------------------------------------------
'Propósito  : Actualiza la tabla Prestamos_Cuotas
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim modPrestamosCuotas As New clsBD3

'Sentencia SQL
sSQL = ""
sSQL = "UPDATE PRESTAMOS_CUOTAS SET " & _
       "Amortizado=0 , " & _
       "Cancelado='NO' " & _
       "WHERE IdPersona='" & txtAfecta.Text & "' And IdConPL= '" & sIdConPl & "' And " & _
       " NumPrestamo= '" & sNumPrestamo & "' And AnioMes ='" & sAnioMes & "' "

'Copia la sentencia SQL
modPrestamosCuotas.SQL = sSQL

'Ejecuta la sentencia SQL
If modPrestamosCuotas.Ejecutar = HAY_ERROR Then
    'Termina la ejecucion
    End
End If

'Cierra el cursor
modPrestamosCuotas.Cerrar

End Sub

Private Sub RestauraTablaPrestamos(sIdConPl As String, sNumPrestamos As String)
'---------------------------------------------------------------
'Propósito  : Actualiza la tabla prestamos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim modPrestamos As New clsBD3

'Sentencia SQL
sSQL = ""
sSQL = "UPDATE PRESTAMOS SET " & _
       "Cancelado='NO' " & _
       "WHERE IdPersona='" & txtAfecta.Text & "' And IdConPL= '" & sIdConPl & "' And " & _
       " NumPrestamo= '" & sNumPrestamos & "' "

'Copia la sentencia SQL
modPrestamos.SQL = sSQL

'Ejecuta la sentencia SQL
If modPrestamos.Ejecutar = HAY_ERROR Then
    'Termina la ejecucion
    End
End If

'Cierra el cursor
modPrestamos.Cerrar

End Sub

Private Sub ModificarIngreso()
'---------------------------------------------------------------
'Propósito : Realiza la operación de modificar en el formulario
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------

' Habilita el txtCodEgreso
  txtCodIngreso = Empty
  txtCodIngreso.Enabled = True
  txtCodIngreso.BackColor = Obligatorio

' Inicializa los optCB
  optCaja.Value = False
  optBanco.Value = False

' Deshabilita controles del formulario
  DeshabilitarHabilitarFormulario False

' Limpia las colecciones
   Set gcolPrestamos = Nothing
   Set gcolDetPrestamos = Nothing
   
   msAfecta = Empty
   msProceso = Empty

' Inicializa la variable codigo de cuenta
  msCtaCte = Empty

' Muestra los resumen
  txtMontoPendiente = "0.00"
  
' Maneja estado de los botones del formulario
  HabilitaDeshabilitaBotones "Modificar"

End Sub

Private Sub cmdBuscar_Click()
'Verifica si esta con datos el Moviemiento
If txtCodMov.Text = Empty Or cboCodMov.Text = Empty Then
    'Mensaje
    MsgBox "Ingrese el movimiento", vbOKOnly + vbInformation, "SGCcaijo-Egreso sin Afectación"
    'Descarga el formulario
    Exit Sub
End If

'Verifica a quien afecta
If msAfecta = "Persona" Then
    ' Carga los títulos del grid selección
      giNroColMNSel = 4
      aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Condición", "Activo")
      aTamañosColumnas = Array(1000, 4500, 1500, 600)
    ' Muestra el formulario de busqueda
      frmMNSeleccion.Show vbModal, Me
    
    ' Verifica si se eligió algun dato a modificar
      If gsCodigoMant <> Empty Then
        txtAfecta.Text = gsCodigoMant
        SendKeys "{tab}"
      Else ' No se eligió nada a modificar
        ' Verifica si txtcodigo es habilitado
        If txtAfecta.Enabled = True Then txtAfecta.SetFocus
      End If
Else
   
    'Determina si el proceso es ingreso es por prestamos
    msProceso = DeterminarProceso(txtCodMov)
         
    'msProceso es dovolucion a rendir
    If msProceso = "DEVOLUCION_RENDIR" Then
        'Muestra el formulario de Fondo a Rendir
        frmCBIFondo_Rendir.Show vbModal, Me
    Else
    
        ' Carga los títulos del grid selección
          giNroColMNSel = 4
          aTitulosColGrid = Array("IdTercero", "Descripción", "RUC", "Dirección")
          aTamañosColumnas = Array(1000, 4500, 2000, 4000)
        ' Muestra el formulario de busqueda
          frmMNSeleccion.Show vbModal, Me
        
        ' Verifica si se eligió algun dato a modificar
          If gsCodigoMant <> Empty Then
            txtAfecta.Text = gsCodigoMant
            SendKeys "{tab}"
          Else ' No se eligió nada a modificar
            ' Verifica si txtcodigo es habilitado
            If txtAfecta.Enabled = True Then txtAfecta.SetFocus
          End If
      End If
End If
  
End Sub

Private Sub cmdBuscarIngreso_Click()

' Define el tipo de selección del Orden
gsTipoSeleccionOrden = "Ingreso"

' Muestra el formulario para elegir el egreso
frmCBSelOrden.Show vbModal, Me

End Sub

Private Sub cmdCancelar_Click()
         
' Limpia el formulario y pone en blanco variables
LimpiarFormulario

' Verifica el tipo operación
If gsTipoOperacionIngreso = "Nuevo" Then
  ' Prepara el formulario
  NuevoIngreso
Else
  ' cierra el control egreso
    If mbIngresoCargado Then
        curRegIngresoCajaBanco.Cerrar
        mbIngresoCargado = False
        mskFecTrab = "__/__/____"
    End If
    
  ' Prepara el formulario
  ModificarIngreso
End If

End Sub

Private Sub LimpiarFormulario()
'----------------------------------------------------------------------------
'Propósito: Limpia la las cajas de texto del formulario
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
'Si la operación elegida en el menu es Modificar se limpia cod Ingreso

txtCodMov.Text = Empty
txtDocIngreso.Text = Empty
txtTipDoc.Text = Empty
cboTipDoc.ListIndex = -1
txtMonto.Text = Empty
txtObserv.Text = Empty
' Limpia controles banco
txtBanco = Empty

End Sub

Private Sub NuevoIngreso()
'--------------------------------------------------------------
'Propósito : Realiza la operación de Ingreso a Caja o Bancos
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
' Inicializa las variable de modulo
  msCtaCte = Empty
  msAfecta = Empty
  msProceso = Empty
  
' Pone por defecto el egreso de caja y calcula el Orden
If optCaja.Value = True Then
    'realiza el evento de optclick
    optCaja_Click
Else
    ' cambia el valor del optCaja.value
    optCaja.Value = True
End If

' Coloca la fecha del sistema
mskFecTrab.Text = gsFecTrabajo

' Limpia las colecciones
Set gcolPrestamos = Nothing
Set gcolDetPrestamos = Nothing

' deshabilita los botones del formulario
HabilitaDeshabilitaBotones ("Nuevo")

End Sub

Private Sub EstableceCamposObligatoriosCaja()
    'Coloca color a controles  del ingreso caja dependiendo del estado
    If txtDocIngreso.Enabled = True Then
        If txtDocIngreso.Text = "" Then txtDocIngreso.BackColor = Obligatorio
    Else
        txtDocIngreso.BackColor = vbWhite
    End If
    If txtTipDoc.Enabled = True Then
        If txtTipDoc.Text = "" Then txtTipDoc.BackColor = Obligatorio
    Else
        txtTipDoc.BackColor = vbWhite
    End If
    If txtMonto.Enabled = True Then
        If txtMonto.Text = "" Then txtMonto.BackColor = Obligatorio
    Else
        txtMonto.BackColor = vbWhite
    End If
    If txtCodMov.Enabled = True Then
        If txtCodMov.Text = "" Then txtCodMov.BackColor = Obligatorio
    Else
        txtCodMov.BackColor = vbWhite
    End If
    If txtAfecta.Enabled = True Then
        If txtAfecta.Text = "" Then txtAfecta.BackColor = Obligatorio
    Else
        txtAfecta.BackColor = vbWhite
    End If
    If txtCodContable.Enabled = True Then
        If txtCodContable.Text = "" Then txtCodContable.BackColor = Obligatorio
    Else
        txtCodContable.BackColor = vbWhite
    End If
    
End Sub


Private Sub DeshabilitarBotones(ByVal sBoton As String)
'----------------------------------------------------------------------------
'Propósito: Deshabilita botones de acuerdo al botón presionado
'Recibe:  sBoton string que indica el nombre del botón presionado
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el evnto click de un botón presionado en el frm
Select Case sBoton
    
    Case "Nuevo"
         cmdCancelar.Enabled = True
    Case "Modificar"
         cmdCancelar.Enabled = True
    Case "Cancelar"
         cmdAceptar.Enabled = False
         cmdAnular.Enabled = False
         cmdCancelar.Enabled = True
         
End Select

End Sub

Private Function CalcularSigOrden(ByVal sCAoBAMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Determina el ultimo registro e incrementa en 1 el campo orden
'Recibe: Fecha
'Devuelve: El string Orden incrementado el campo orden
'----------------------------------------------------------------------------

Dim sUltimoReg As String
Dim sCodigo As String
Dim sNroSecuencial As String
Dim sSQL As String
Dim curOrdenIngreso As New clsBD2
Dim curOrdenEgreso As New clsBD2
Dim iNumSec As Integer


'Concatenamos el codigo CAAñoMes
sCodigo = sCAoBAMov & Right(gsFecTrabajo, 2) & Mid(gsFecTrabajo, 4, 2)
'Se carga un string con el ultimo registro del campo orden
sSQL = ""

  sSQL = "SELECT Max(Orden)  FROM INGRESOS WHERE  Orden LIKE '" & sCodigo & "*'"
curOrdenIngreso.SQL = sSQL
' Averigua el ultimo orden de ingreso
If curOrdenIngreso.Abrir = HAY_ERROR Then
  Unload Me
End If
  
  
  sSQL = "SELECT Max(Orden)  FROM EGRESOS WHERE  Orden LIKE '" & sCodigo & "*'"
curOrdenEgreso.SQL = sSQL
' Averigua el ultimo orden de egreso
If curOrdenEgreso.Abrir = HAY_ERROR Then
  Unload Me
End If


'Separa los cuatro últimos caracteres del maximo orden en Ingreso e egreso
If IsNull(curOrdenIngreso.campo(0)) And IsNull(curOrdenEgreso.campo(0)) Then ' NO hay registros
  CalcularSigOrden = (sCodigo & "0001")
Else
  If Not IsNull(curOrdenIngreso.campo(0)) And Not IsNull(curOrdenEgreso.campo(0)) Then
  'Ambos Ingresos e Egresos tienen registros
        If curOrdenIngreso.campo(0) < curOrdenEgreso.campo(0) Then
            iNumSec = Val(Right(curOrdenEgreso.campo(0), 4))
            CalcularSigOrden = Left(curOrdenEgreso.campo(0), 6) & Format(CStr(iNumSec) + 1, "000#")
        Else
            iNumSec = Val(Right(curOrdenIngreso.campo(0), 4))
            CalcularSigOrden = Left(curOrdenIngreso.campo(0), 6) & Format(CStr(iNumSec) + 1, "000#")
        End If
  Else ' Alguno Ingreso o Egreso tiene registros
        If IsNull(curOrdenIngreso.campo(0)) Then
           iNumSec = Val(Right(curOrdenEgreso.campo(0), 4))
           CalcularSigOrden = Left(curOrdenEgreso.campo(0), 6) & Format(CStr(iNumSec) + 1, "000#")
        Else
           iNumSec = Val(Right(curOrdenIngreso.campo(0), 4))
           CalcularSigOrden = Left(curOrdenIngreso.campo(0), 6) & Format(CStr(iNumSec) + 1, "000#")
        End If
  End If
End If
'Cierra el cursor
curOrdenEgreso.Cerrar
curOrdenIngreso.Cerrar

End Function

Private Sub cmdPBanco_Click()

If cboBanco.Enabled Then
    ' alto
     cboBanco.Height = CBOALTO
    ' focus a cbo
    cboBanco.SetFocus
End If

End Sub


Private Sub cmdpCodContable_Click()

If cboCtaContable.Enabled Then
    ' alto
     cboCtaContable.Height = CBOALTO
    ' focus a cbo
    cboCtaContable.SetFocus
End If

End Sub

Private Sub cmdPCodMov_Click()

If cboCodMov.Enabled Then
    ' alto
     cboCodMov.Height = CBOALTO
    ' focus a cbo
    cboCodMov.SetFocus
End If

End Sub

Private Sub cmdPCtaCte_Click()

If cboCtaCte.Enabled Then
    ' alto
     cboCtaCte.Height = CBOALTO
    ' focus a cbo
    cboCtaCte.SetFocus
End If

End Sub

Private Sub cmdpTipDoc_Click()

If cboTipDoc.Enabled Then
    ' alto
     cboTipDoc.Height = CBOALTO
    ' focus a cbo
    cboTipDoc.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()
'Inicializa los parametro para ingreso
CancelarVenta = False
'Cierra el formulario
Unload Me

End Sub



Private Sub Form_Load()

Dim sSQL As String

'Carga el combo tipo movimiento y las colecciones de tipo_mov
CargarColTipo_Mov

'Se carga el combo de Tipo de Documento
sSQL = ""
sSQL = "SELECT idTipoDoc, DescTipoDoc FROM TIPO_DOCUM " & _
           "WHERE RelacProc = 'IN'  ORDER BY DescTipoDoc"
CD_CargarColsCbo cboTipDoc, sSQL, mcolCodTipDoc, mcolCodDesTipDoc

'Se carga el combo de Cta Cte
sSQL = ""
sSQL = "SELECT IdCta, DescCta FROM TIPO_CUENTASBANC " & _
           "WHERE IdMoneda= 'SOL'   ORDER BY DescCta"
CD_CargarColsCbo cboCtaCte, sSQL, mcolCodCtaCte, mcolCodDesCtaCte

'Se carga el combo Bancos (sólo con los bancos que de moneda nacional)
sSQL = ""
sSQL = "SELECT DISTINCT b.IdBanco,b.DescBanco FROM TIPO_BANCOS B , TIPO_CUENTASBANC C" _
       & " WHERE b.idbanco = c.idbanco And c.idmoneda = 'SOL'" _
       & " ORDER BY DescBanco"
CD_CargarColsCbo cboBanco, sSQL, mcolCodBanco, mcolCodDesBanco

'Se Limpia el Combo de Cts Corrientes en dólares
cboCtaCte.Clear

' Inhabilita botones al cargar el formulario
cmdAceptar.Enabled = False
cmdAnular.Enabled = False

'Oculta Monto Pendiente
txtMontoPendiente.Visible = False
lblMontoPendiente.Visible = False

'Establece campos obligatorios del formulario
EstableceCamposObligatorios

' Verifica el tipo de operación a realizar en el formulario
If gsTipoOperacionIngreso = "Modificar" Then
    
    fraIngreso.Enabled = False
    
    'Deshabilita el control de busqueda
    cmdBuscarIngreso.Enabled = True
    
    ' Deshabilita el movimiento
    txtCodMov.Enabled = False
    cboCodMov.Enabled = False
    
    'Inicializa la variable
    mbIngresoCargado = False
    
    'Coloca titulo al formulario
    Me.Caption = "Caja y Bancos - Modificación de Ingresos a Caja o Bancos"
    
    'Prepara campos para modificar algun registro de ingreso
    ModificarIngreso
    
Else 'Nuevo Ingreso a Caja o Bancos
    Me.Caption = "Caja y Bancos- Ingreso a Caja o Bancos"
    cmdBuscarIngreso.Enabled = False
    
   'Deshabilita txtCodIngreso
    txtCodIngreso.Enabled = False
        
    'Nuevo ingreso
    NuevoIngreso
    
    If CancelarVenta = True Then
      txtCodMov = CodigoVenta
      txtAfecta = CodigoTerceroVenta
      txtMonto = VentaSaldo
    End If
End If

End Sub

Private Sub HabilitaDeshabilitaBotones(sProceso As String)
'-----------------------------------------------------------------
' Proposito: Coloca la condición de los botones en el proceso
' Recibe: Nada
' Entrega: Nada
'-----------------------------------------------------------------
Select Case sProceso

' depende del proceso habilita y deshabilita botones
Case "Nuevo"
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    
Case "Modificar"
    cmdBuscarIngreso.Enabled = True
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
End Select

End Sub

Private Sub EstableceCamposObligatorios()

' Establece los campos obligatorios
 txtDocIngreso.BackColor = Obligatorio
 txtTipDoc.BackColor = Obligatorio
 txtMonto.BackColor = Obligatorio
 txtCodMov.BackColor = Obligatorio
 txtAfecta.BackColor = Obligatorio
 txtCodContable.BackColor = Obligatorio
 txtBanco.BackColor = Obligatorio
 cboCtaCte.BackColor = Obligatorio
  
End Sub

Private Sub CargarColPersonal()
'--------------------------------------------------------------
'Propósito : Carga la coleccion del Personal
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String

'Sentencia para cargar la colección
If gsTipoOperacionIngreso = "Nuevo" Then
  sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
      & " PP.Condicion, PP.Activo " _
      & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
      & " WHERE P.IdPersona=PP.IdPersona and PP.Activo='SI' " _
      & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"
Else
  sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
      & " PP.Condicion, PP.Activo " _
      & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
      & " WHERE P.IdPersona=PP.IdPersona " _
      & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"
End If

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

End Sub

Private Sub CargarColTerceros()
'--------------------------------------------------------------
'Propósito : Carga la coleccion del PlanContable
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String

'Sentencia para cargar la colección
sSQL = "SELECT IdTerc, DescTerc, DNI_RUC_Terc, Dir_Terc " _
       & "FROM TIPO_TERCEROS ORDER BY DescTerc"

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

End Sub

Private Sub CargarColTipo_Mov()
'--------------------------------------------------------------
'Propósito : Carga la coleccion de Tipo_Mov con sus diferentes campos
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curTipoMov As New clsBD2

If gsTipoOperacionIngreso = "Nuevo" Then
    'la sentencia para cargar el combo y las colecciones de tipo movimiento
    sSQL = ""
    sSQL = "SELECT IdConCB, DescConCB, CodCont, Afecta FROM Tipo_MovCB " & _
            "WHERE RelacProc = 'IN' ORDER BY DescConCB"
Else
    'la sentencia para cargar el combo y las colecciones de tipo movimiento
    sSQL = ""
    sSQL = "SELECT IdConCB, DescConCB, CodCont, Afecta FROM Tipo_MovCB " & _
            "WHERE RelacProc = 'IN' OR RelacProc = 'IE' ORDER BY DescConCB"
End If

'Se carga el combo
CD_CargarColsCbo cboCodMov, sSQL, mcolCodMov, mcolCodDesCodMov

'Carga la coleccion de IdCodconCB y CodCont, IdCodconCB y Afecta
curTipoMov.SQL = sSQL
If curTipoMov.Abrir = HAY_ERROR Then
  End
End If
'Inicializamos el Tamaño del codigo contable del tipoMovimiento
miTamañoCodCont = 0
Do While Not curTipoMov.EOF
    ' Se carga la colección de IdConCB + CodCont
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolCodCont.Add curTipoMov.campo(0)
    mcolDesCodCont.Add Item:=curTipoMov.campo(2), Key:=curTipoMov.campo(0)
    
    mcolCodAfecta.Add curTipoMov.campo(0)
    mcolDesCodAfecta.Add Item:=curTipoMov.campo(3), Key:=curTipoMov.campo(0)
    'Averigua el Mayor Tamaño de CodContable de Conceptos
    If Len(Trim(curTipoMov.campo(2))) > miTamañoCodCont Then miTamañoCodCont = Len(Trim(curTipoMov.campo(2)))
    
    ' Se avanza a la siguiente fila del cursor
    curTipoMov.MoverSiguiente
Loop
'Cierra el cursor de curTipoMov
curTipoMov.Cerrar

End Sub


Private Sub HabilitarBotonAceptar()
'----------------------------------------------------------------------------
'PROPÓSITO: *Se habilita "Aceptar del formulario " en Ingreso de un Nuevo registro
'               Si se han rellenado los campos obligatorios
'           *Se habilita "Aceptar" en Modificacion
'               Si se han rellenado los campos, y Si se realizo algun cambio al registro
'----------------------------------------------------------------------------
 
' Verifica si se a introducido los datos obligatorios generales
If txtCodMov.BackColor <> vbWhite _
       Or txtDocIngreso.BackColor <> vbWhite _
       Or txtTipDoc.BackColor <> vbWhite _
       Or txtAfecta.BackColor <> vbWhite _
       Or txtCodContable.BackColor <> vbWhite _
       Or txtMonto.BackColor <> vbWhite _
Then
   ' Algún obligatorio falta ser introducido
   ' Deshabilita el botón
   cmdAceptar.Enabled = False
   Exit Sub
Else
   ' Verifica que se haigan introducido los datos obligatorios de bancos
   If optBanco.Value = True And (txtBanco.BackColor <> vbWhite Or cboCtaCte.BackColor <> vbWhite) Then
     ' Algún obligatorio de banco falta ser introducido
     ' Deshabilita el botón
     cmdAceptar.Enabled = False
     Exit Sub
   End If
End If

' Verifica si se cambio algún dato
If gsTipoOperacionIngreso = "Modificar" Then
    ' Verifica si se cambio los datos generales
   If fbCambioDatos = False Then
        ' No se cambio ningún dato
        ' Deshabilita el boton
        cmdAceptar.Enabled = False
        Exit Sub
   End If
End If

' Habilita botón aceptar
cmdAceptar.Enabled = True

End Sub

Private Function fbCambioDatos() As Boolean
' --------------------------------------------------------------
' Propósito : Verifica si se cambió algún dato general del egreso
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------------
' Inicializa la función
    fbCambioDatos = False
' verifica los datos de caja
If txtCodMov.Text <> curRegIngresoCajaBanco.campo(5) _
   Or txtDocIngreso.Text <> curRegIngresoCajaBanco.campo(0) _
   Or txtTipDoc.Text <> curRegIngresoCajaBanco.campo(2) _
   Or txtAfecta.Text <> gsCodAfectaAnterior _
   Or txtCodContable.Text <> curRegIngresoCajaBanco.campo(6) _
   Or Val(Var37(txtMonto.Text)) <> curRegIngresoCajaBanco.campo(4) _
   Or txtObserv.Text <> curRegIngresoCajaBanco.campo(8) Then
    ' cambio datos generales
    fbCambioDatos = True
    Exit Function
End If

' Verifica si se cambio Cta corriente y número de cheque
If msCajaoBanco = "BA" Then
    If msCtaCte <> curRegIngresoCajaBanco.campo(10) Then
        ' Cambió cta corriente
            ' cambio datos generales
            fbCambioDatos = True
            Exit Function
    End If

End If

End Function

Private Sub LimpiarPantalla()
'----------------------------------------------------------------------------
'Propósito: Limpia la las cajas de texto del formulario
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
'SI la operación elegida en el menu es Modificar se limpia cod Ingreso
If gsTipoOperacionIngreso = "Modificar" Then
  'If cmdCancelar.Value = False Then
  txtCodIngreso.Text = Empty
  txtCodIngreso.BackColor = Obligatorio
  txtCodIngreso.Enabled = True
  mskFecTrab.Text = "__/__/____"

End If

txtCodMov.Text = Empty
txtDocIngreso.Text = Empty
txtTipDoc.Text = Empty
cboTipDoc.ListIndex = -1
txtMonto.Text = Empty
txtObserv.Text = Empty
txtBanco.Text = Empty
cboCtaCte.ListIndex = -1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Destruimos las Colecciones
Set mcolCodMov = Nothing
Set mcolCodDesCodMov = Nothing

Set mcolCodBanco = Nothing
Set mcolCodDesBanco = Nothing

Set mcolCodTipDoc = Nothing
Set mcolCodDesTipDoc = Nothing

Set mcolCodCtaCte = Nothing
Set mcolCodDesCtaCte = Nothing

Set mcolCodAfecta = Nothing
Set mcolDesCodAfecta = Nothing

Set mcolCodCont = Nothing
Set mcolDesCodCont = Nothing

Set mcolCodPlanCont = Nothing
Set mcolDesCodPlanCont = Nothing

'Vacia las colecciones
Set gcolPrestamos = Nothing
Set gcolDetPrestamos = Nothing

' Verifica SI esta habilitado controles de ingreso caja bancos
If txtCodMov.Enabled Then

    'Verifica SI la operación a cancelar es modificar
    If gsTipoOperacionIngreso = "Modificar" And mbIngresoCargado = True Then
      curRegIngresoCajaBanco.Cerrar ' Cierra el cursor del ingreso
    End If
End If

Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Inicializa los parametro para ingreso
CancelarVenta = False
End Sub

Private Sub Image1_Click()
'Carga la Var48
Var48
End Sub

Private Sub optBanco_Click()

' Realiza el cambio de opción a Caja
 CambiaroptCajaBancos
  
' Habilita el botón aceptar
 HabilitarBotonAceptar

End Sub

Private Sub optBanco_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub optCaja_Click()

' Realiza el cambio de opción a Caja
 CambiaroptCajaBancos
  
' Habilita el botón aceptar
 HabilitarBotonAceptar

End Sub

Private Function VerificaMovimiento() As Boolean
'--------------------------------------------------------
' Propósito: Verifica el movimiento con las opciones de origen del egreso
' Recibe: Nada
' Entrega: Nada
'--------------------------------------------------------
If msAfecta = "Proceso" Then
    If msProceso = "DEVOLUCION_RENDIR" Then
       ' Verifica que el movimiento sea de caja
       If optCaja.Value = False Then
            ' El movimiento solo es de Caja
            VerificaMovimiento = False
            Exit Function
       End If
    End If
End If

' Todo Ok
VerificaMovimiento = True

End Function

Private Sub CambiaroptCajaBancos()
'-------------------------------------------------------------------
'Propósito : Establece los controles de la primera parte del formulario _
             cuando se cambia de optCaja a optBancos bis
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If gsTipoOperacionIngreso = "Nuevo" Then
   If optCaja.Value = True Then
    'Calcula el sigiente orden de Caja y lo muestra en el txtCodIngreso
    txtCodIngreso.Text = Var22("CA")
    msCajaoBanco = "CA"
   Else
    'Calcula el sigiente orden de Banco y lo muestra en el txtCodIngreso
    txtCodIngreso.Text = Var22("BA")
    msCajaoBanco = "BA"
   End If
End If
  
' Verifica el movimiento
If VerificaMovimiento = False Then
    ' Movimiento no valido
    MsgBox "El Movimiento elegido solo es de Caja", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
    optCaja.Value = True
    ' Sale
    Exit Sub
End If
  
' De acuerdo a la elección realizada maneja los controles de banco
ManejaControlesBanco

End Sub

Private Sub ManejaControlesBanco()
'-------------------------------------------------------------------
'Propósito : Establece los controles de banco cuando se cambia de optCaja a optBancos bis
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If optCaja.Value = True Then
 'Limpia y oculta los controles de Banco
 txtBanco.Text = Empty
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False
 cmdPBanco.Visible = False: cmdPCtaCte.Visible = False

Else
 'Muestra los controles de banco
 lblBanco.Visible = True: txtBanco.Visible = True: cboBanco.Visible = True
 lblCtaCte.Visible = True: cboCtaCte.Visible = True
 cmdPBanco.Visible = True: cmdPCtaCte.Visible = True
End If

End Sub

Private Sub optCaja_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtAfecta_Change()
'Verifica si el tamaño del txt es Igual al tamaño definido
If Len(txtAfecta) = txtAfecta.MaxLength Then
    'Actualiza el txtDesc
    ActualizaDesc
Else
    'Limpia el txtDescAfecta
    txtDesc.Text = Empty
End If

' Verifica SI el campo esta vacio
If txtAfecta.Text <> Empty And txtDesc.Text <> Empty Then
   ' Los campos coloca a color blanco
   txtAfecta.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtAfecta.BackColor = Obligatorio
End If
    
' Habilita botón aceptar
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
txtDesc.Text = Var30(gcolTabla.Item(txtAfecta.Text), 2)

' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe", , "SGCcaijo-Ingresos"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub ActualizaDescCtaContable()
'--------------------------------------------------------------
'PROPÓSITO  : Actualiza la descripcion de la CtaContable
'Recive     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------
Dim sSQL As String
Dim curDescCtaContable As New clsBD2

'Sentencia SQL
sSQL = ""
sSQL = "SELECT DescCuenta " _
      & " FROM PLAN_CONTABLE P " _
      & " WHERE CodContable= '" & txtCodContable.Text & "' "
        
'Copia la sentencia SQL
curDescCtaContable.SQL = sSQL

'Verifica si hay error al ejecuta la sentencia SQL
If curDescCtaContable.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Verifica si tiene algun registro
Do While Not curDescCtaContable.EOF

    'Devuelve el proceso relaionado
    cboCtaContable.Text = curDescCtaContable.campo(0)
    
    'Mueve al siguiente registro
    curDescCtaContable.MoverSiguiente
Loop

'Cierra el cursor
curDescCtaContable.Cerrar

End Sub

Private Sub txtAfecta_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtBanco_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboBanco, txtBanco, mcolCodDesBanco

  ' Verifica SI el campo esta vacio
If txtBanco.Text <> "" And cboBanco.Text <> "" Then
   ' Los campos coloca a color blanco
   txtBanco.BackColor = vbWhite
   'Actualiza el cboCtaCte con las descripciones de las cuentas relacionadas a txtBanco
    ActualizarListcboCtaCte
Else
   'Marca los campos obligatorios, y limpia el combo
   txtBanco.BackColor = Obligatorio
   cboCtaCte.Clear
   cboCtaCte.BackColor = Obligatorio
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
End Sub

Private Sub txtBanco_LostFocus()

'Se pasa el texto a mayúsculas
txtBanco = UCase(txtBanco)

End Sub

Private Sub txtCodContable_Change()

' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboCtaContable, txtCodContable, mcolDesCodPlanCont

' Verifica SI el campo esta vacio
If txtCodContable.Text <> "" And cboCtaContable.Text <> "" Then
   ' Los campos coloca a color blanco
   txtCodContable.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtCodContable.BackColor = Obligatorio
End If

' Habilita botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtCodContable_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtCodIngreso_Change()

' Verifica el proceso que se realiza en el formulario
If gsTipoOperacionIngreso = "Modificar" Then
    
    ' Verifica si se ha introducido el tamaño de el código
      If Len(txtCodIngreso.Text) = txtCodIngreso.MaxLength Then
      
        ' Verifica mayúsculas
        If UCase(txtCodIngreso.Text) = txtCodIngreso.Text Then
            
            ' Verifica si el Ingreso existe y es sin afectación
            If fbCargarIngreso = True Then
            
                ' Sale y deshabilita el control
                SendKeys vbTab
             
                ' deshabilita el txtcod egreso y el botón buscar, _
                  habilita anular
                txtCodIngreso.Enabled = False
                cmdBuscarIngreso.Enabled = False
                cmdAnular.Enabled = True
             
          End If ' fin de cargar egreso
          
        Else
            'Vuelve a mayúsulas el txtCodIngreso
            txtCodIngreso = UCase(txtCodIngreso)
            
        End If ' fin verificar mayúsculas
        
      End If ' fin de verificar el tamaño del texto
 End If
 
 ' Maneja el color del control txtCodIngreso
 If txtCodIngreso = Empty Then
    ' coloca el color obligatorio al control
    txtCodIngreso.BackColor = Obligatorio
 Else
    ' coloca el color de edición
    txtCodIngreso.BackColor = vbWhite
 End If
 
End Sub

Private Function VerificarIngresos()
'--------------------------------------------------------------------------
'Propósito  : Determina si esta relacionado el ingreso con prestamos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------------------
Dim sSQL As String
Dim curDevPrestamos As New clsBD2

'Sentencia SQL
sSQL = ""
sSQL = " SELECT Distinct IdPersona " _
        & "FROM DEVOLUCION_PRESTAMOSCB " _
        & "WHERE Orden= '" & txtCodIngreso.Text & "' "

'Copia la sentencia SQL
curDevPrestamos.SQL = sSQL

'Verifica si hay error al ejecuta la sentencia SQL
If curDevPrestamos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Devuelve el proceso relacionado
VerificarIngresos = False
    
'Verifica si tiene algun registro
Do While Not curDevPrestamos.EOF

    'Devuelve el proceso relaionado
    VerificarIngresos = True
    
    'Mueve al siguiente registro
    curDevPrestamos.MoverSiguiente
Loop

'Cierra el cursor
curDevPrestamos.Cerrar

End Function

Private Sub txtCodIngreso_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtCodMov_Change()

If UCase(txtCodMov.Text) = txtCodMov.Text Then
    
    ' SI procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboCodMov, txtCodMov, mcolCodDesCodMov
    
     'Inicializa la variable Afecta (Terceros o Personal)
     msAfecta = Empty
    
    'Limpia el cboAfecta
     txtAfecta.Text = Empty
     txtAfecta.BackColor = Obligatorio
     txtDesc.Text = Empty
     
     'Limpia los campos txtCodContable cboCtaContable
     cboCtaContable.Clear
     txtCodContable.Text = Empty
     txtCodContable.BackColor = Obligatorio
     
    ' Verifica Si el campo esta vacio
    If txtCodMov.Text <> Empty And cboCodMov.Text <> Empty Then
        'Los campos coloca a color blanco
        txtCodMov.BackColor = vbWhite
        'Carga el combo Afecta dependiendo del código de afecta
        msAfecta = DeterminarAfecta(txtCodMov.Text)
        CargarCboAfecta msAfecta
        'Carga el cboCtaContable dependiendo del tipo de movimiento
        CargacboCtaContable DeterminarCodCont(txtCodMov.Text)
        ' Maneja estado de los controles dependiendo de msafecta
        EstableceEstadoAfectaMonto
        'Si el combo sólo tiene un elemento, se muestra en pantalla
        MostrarUnicoItem
    Else
      'Marca los campos obligatorios
       txtCodMov.BackColor = Obligatorio
    End If

Else
    If Len(txtCodMov.Text) = txtCodMov.MaxLength Then
        'comvertimos a mayuscula
        txtCodMov.Text = UCase(txtCodMov.Text)
    End If
End If

'Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub EstableceEstadoAfectaMonto()
'----------------------------------------------------------------
' Propósito : Establece el estado de los controles Afecta y el monto _
              de egreso
' Recibe    : Nada
' Entrega   : Nada
'----------------------------------------------------------------
If msAfecta = "Proceso" Then
    If msProceso = "DEVOLUCION_RENDIR" Then
        ' Inhabilita afecta y monto
        txtAfecta.Enabled = False: cmdBuscar.Enabled = True
        txtMonto.Enabled = False
    Else
        ' Inhabilita afecta y monto
        txtAfecta.Enabled = False: cmdBuscar.Enabled = False
        txtMonto.Enabled = False
    End If
Else
    ' Habilita afecta y monto
    txtAfecta.Enabled = True: cmdBuscar.Enabled = True
    txtMonto.Enabled = True
End If

End Sub

Private Function DeterminarProceso(sCodMov As String)
'--------------------------------------------------------------------------
'Propósito  : Determina si el proceso esta relacionado con pago de prestamos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------------------
Dim sSQL As String
Dim curProcesoPrestamos As New clsBD2

'Sentencia SQL
sSQL = ""
sSQL = " SELECT PC.Proceso " _
        & "FROM PROCESO_CONCEPTOCB PC " _
        & "WHERE PC.IdConCB= '" & sCodMov & "' "

'Copia la sentencia SQL
curProcesoPrestamos.SQL = sSQL

'Verifica si hay error al ejecuta la sentencia SQL
If curProcesoPrestamos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If
If curProcesoPrestamos.EOF Then

    'Devuelve vacio
    DeterminarProceso = Empty
    
Else
    'Devuelve el proceso relaionado
    DeterminarProceso = curProcesoPrestamos.campo(0)
    
End If

'Cierra el cursor
curProcesoPrestamos.Cerrar

End Function

Private Sub CargacboCtaContable(sCodCont As String)
'----------------------------------------------------------------------------
'Propósito: Carga el combo de la cuenta contable a partir del código contable
'           del tipo de movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String

'Vaciamos las colecciones
Set mcolCodPlanCont = Nothing
Set mcolDesCodPlanCont = Nothing

' Si no tiene codigo contable sale
If sCodCont = Empty Then
    txtCodContable.BackColor = vbWhite
    lblCodContable.Visible = False: txtCodContable.Visible = False: cboCtaContable.Visible = False
    cmdPCodContable.Visible = False
Else ' Carga las cuentas contables del egreso

    sSQL = "SELECT CodContable, CodContable & ' ' &Left(DescCuenta,55) FROM PLAN_CONTABLE " & _
           "WHERE CodContable LIKE '" & sCodCont & "*' And (len(CodContable)=" & miTamañoCodCont _
           & ") ORDER BY CodContable"
    CD_CargarColsCbo cboCtaContable, sSQL, mcolCodPlanCont, mcolDesCodPlanCont
    
    'Definimos el numero de caracteres del control txtCodMov(Conceptos)
    txtCodContable.MaxLength = miTamañoCodCont

End If

End Sub

Function DeterminarCodCont(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que codigo contable un determinado tipo de
'           movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------

'Muestra a que codigo contable afecta el campo seleccionado en el combo tipo mov
If msAfecta = "Tercero" Or msAfecta = "Persona" Then
    'Muestra a que codigo contable afecta el campo seleccionado en el combo tipo mov
    DeterminarCodCont = mcolDesCodCont.Item(Trim(sCodMov))
Else
    DeterminarCodCont = Empty
End If

End Function

Private Sub CargarCboAfecta(sCodRec As String)

'Verifica a que afecta Personal(P), Terceros(T), PlanContable(C o N)
    Select Case sCodRec
    Case "Persona"
        ' Ningun proceso
        msProceso = Empty
        'Asigana el tamaño al maxlength del txtAfecta
        txtAfecta.MaxLength = 4
        lblEtiqueta.Caption = "Personal:"
        'Se carga la colección de Personal
        CargarColPersonal
        'Habilita los campos afecta
        txtAfecta.Enabled = True
        cmdBuscar.Enabled = True
        txtAfecta.BackColor = Obligatorio
        txtAfecta.Visible = True
        txtDesc.Visible = True
                
    Case "Tercero"
        ' Ningun proceso
        msProceso = Empty
    
        'Asigana el tamaño al maxlength del txtAfecta
        txtAfecta.MaxLength = 2
        lblEtiqueta.Caption = "Terceros:"
        'Se carga la colección de terceros
        CargarColTerceros
        'Habilita los campos afecta
        txtAfecta.Enabled = True
        cmdBuscar.Enabled = True
        txtAfecta.BackColor = Obligatorio
        txtAfecta.Visible = True
        txtDesc.Visible = True
        cmdBuscar.Visible = True
        
    Case "Proceso"
        'Determina si el proceso es ingreso es por prestamos
         msProceso = DeterminarProceso(txtCodMov)
         
        ' Verifica si el movimiento es un proceso
        Select Case msProceso
        Case "DEVOLUCION_PRESTAMOS"
            'Se carga la colección de Personal
            txtAfecta.MaxLength = 4
            CargarColPersonal
             ' Llama al proceso pago de prestamos
            If gsTipoOperacionIngreso = "Nuevo" Then
                frmCBIDev_Prestamos.Show vbModal, Me
                  
                'Verifica si se tiene el CodMOv
                If txtCodMov.Text <> Empty Then
                    lblEtiqueta = "Personal:"
                Else
                    'Ubica el cursor en el txtCodMov
                    txtCodMov.SetFocus
                    
                    'Termina de ejecutar el procedimiento
                    Exit Sub
                End If
             
            End If
        Case "DEVOLUCION_RENDIR"
           
            'Se carga la colección de Personal
            txtAfecta.MaxLength = 4
            CargarColPersonal
             
              ' Verifica el movimiento
            If VerificaMovimiento = False Then
                ' Movimiento no valido
                MsgBox "El Movimiento elegido solo es de Caja", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
                optCaja.Value = True
                ' Sale
                Exit Sub
            End If
            
             ' Llama al proceso pago de prestamos
            If gsTipoOperacionIngreso = "Nuevo" Then
                frmCBIFondo_Rendir.Show vbModal, Me
                  
                'Verifica si se tiene el CodMOv
                If txtAfecta.Text <> Empty Then
                    lblEtiqueta = "Personal:"
                    
                Else
                    'Ubica el cursor en el txtCodMov
                    txtCodMov.SetFocus
                    
                    'Termina de ejecutar el procedimiento
                    Exit Sub
                End If
             
            End If
        Case Empty
             ' El movimiento no tiene proceso
            If gsTipoOperacionEgreso = "Nuevo" Then
                 MsgBox "Este movimiento no esta relacionado a los Procesos:" & Chr(13) _
                        , "SGCcaijo - Ingresos a Caja"
                 txtCodMov.SetFocus
            End If
        End Select
    End Select
End Sub

Function DeterminarAfecta(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que afecta Pln_Personal (P), Terceros (T), PlanContable (C)
'           un determinado tipo de   movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------

'Muestra a que afecta Personal (Personal), Terceros (Terceros), PlanContble (C) el campo seleccionado en el combo
DeterminarAfecta = mcolDesCodAfecta.Item(Trim(sCodMov))
txtAfecta.Enabled = True: cmdBuscar.Enabled = True
lblCodContable.Visible = True: txtCodContable.Visible = True: cboCtaContable.Visible = True
txtMonto.Enabled = True: cmdPCodContable.Visible = True
  
End Function

Private Sub txtCodMov_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtDocIngreso_Change()

' Verifica SI el campo esta vacio
If txtDocIngreso.Text <> "" And InStr(txtDocIngreso, "'") = 0 Then
' El campos coloca a color blanco
   txtDocIngreso.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtDocIngreso.BackColor = Obligatorio
End If

'habilita Boton Aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtDocIngreso_KeyPress(KeyAscii As Integer)
' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  
End Sub

Private Sub txtMonto_Change()

' Verifica SI el campo esta vacio
If txtMonto.Text <> "" And Val(txtMonto.Text) <> 0 Then
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

Private Sub txtObserv_Change()

' Si en la observación hay apostrofes vacío
If InStr(txtObserv, "'") > 0 Then
   txtObserv = Empty
End If

End Sub

Private Sub txtObserv_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtTipDoc_Change()

' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboTipDoc, txtTipDoc, mcolCodDesTipDoc

' Verifica SI el campo esta vacio
If txtTipDoc.Text <> "" And cboTipDoc.Text <> "" Then
' Los campos coloca a color blanco
   txtTipDoc.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtTipDoc.BackColor = Obligatorio
End If

'habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Function fbCargarIngreso() As Boolean
'----------------------------------------------------------------------------
'Propósito: Carga el registro de ingreso de acuerdo al código en la caja _
            de texto
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota el codigo de Registro de ingreso es CA o BA AAMM9999
Dim sSQL As String
Dim curMontoPendienteIngreso As New clsBD2

' Verifica SI el ingreso es a Caja o Bancos
msCajaoBanco = Left(txtCodIngreso.Text, 2)

' Carga la sentencia que consulta a la BD acerca del registo de ingreso en Caja o Bancos
If msCajaoBanco = "CA" Then 'Consulta a Caja
    sSQL = "SELECT NumDoc,IdEgreso,IdTipoDoc,FecMov,Monto,CodMov, CodContable, TM.Afecta, Observ " & _
           "FROM INGRESOS, Tipo_MovCB TM WHERE " & _
           "Orden=" & "'" & Trim(txtCodIngreso.Text) & "' and Anulado='NO' and " & _
           "CodMov=TM.IdConCB"
           
Else 'Consulta a Bancos
    If msCajaoBanco = "BA" Then
            sSQL = "SELECT I.NumDoc,I.IdEgreso,I.IdTipoDoc,I.FecMov, " & _
                           "I.Monto,I.CodMov, I.CodContable, TM.Afecta, I.Observ, CTA.IdBanco,I.IdCta " & _
                   "FROM INGRESOS I, TIPO_CUENTASBANC CTA, Tipo_MovCB TM WHERE " & _
                   "I.Orden=" & "'" & Trim(txtCodIngreso.Text) & "' and I.Anulado='NO' " & _
                   "and I.codMov=TM.IdconCB and I.IdCta=CTA.IdCta"
               
    Else 'Mensaje Cod Registro Ingreso  NO Valido
        MsgBox "El Código de Ingreso No válido, debe ser CA o BA AAMM9999", _
        vbExclamation + vbOKOnly, "Caja-Bancos- Ingresos"
        fbCargarIngreso = False
        Exit Function
    End If
End If

curRegIngresoCajaBanco.SQL = sSQL
' Abre el cursor SI hay  error sale indicando la causa del error
If curRegIngresoCajaBanco.Abrir = HAY_ERROR Then
    End
End If

' Cursor abierto
mbIngresoCargado = True

'Verifica la existencia del registro de ingreso
If curRegIngresoCajaBanco.EOF Then
    'Mensaje de registro de Ingreso a Caja o Bancos NO existe
    MsgBox "El Código de Ingreso que se digito No está registrado o está Anulado", _
      vbInformation + vbOKOnly, "Caja-Bancos- Ingresos"
    curRegIngresoCajaBanco.Cerrar
    ' Cursor abierto
    mbIngresoCargado = False
    Exit Function
    
Else
    'Carga los controles con datos del ingreso y Habilita los controles
    CargarControlesIngreso
    
    'Carga el monto pendiente de ingreso, Este Monto se Puede repartir en esta modificacion
    If curRegIngresoCajaBanco.campo(1) <> "" Then 'El ingreso esta relacionado a un egreso de Ctas en dólares
        ' Sentencia sql que averigua el monto pendiente a repartir
        sSQL = "SELECT MontoSol FROM EGRESO_CTAS_EXTR " & _
             "WHERE IdEgreso='" & curRegIngresoCajaBanco.campo(1) & "'"

        curMontoPendienteIngreso.SQL = sSQL
        
        If curMontoPendienteIngreso.Abrir = HAY_ERROR Then
          End
        End If
        
        If curMontoPendienteIngreso.EOF Then
            MsgBox "No Existe un Monto para este pendiente." & Chr(13) & "Consulte a su Administrador", _
              vbInformation + vbOKOnly, "Caja-Bancos- Ingresos"
            End
            
        Else 'Carga controles de Monto Pendiente, que se puede repartir en esta modificacion
            txtMontoPendiente.Text = Format(curMontoPendienteIngreso.campo(0), "###,###,##0.00")
            lblMontoPendiente.Visible = True
            txtMontoPendiente.Visible = True
        
        End If
        
        curMontoPendienteIngreso.Cerrar
    Else ' No esta realcionado a salidas de Ctas extrangeras
            txtMontoPendiente.Text = "0.00"
            lblMontoPendiente.Visible = False
            txtMontoPendiente.Visible = False
    End If

End If

' Todo Ok
fbCargarIngreso = True

End Function

Private Sub CargarControlesIngreso()
'----------------------------------------------------------------------------
'Propósito  : Cargar los controles refentes al ingreso que se desea modificar
'Recibe     :  Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el procedimiento Cargar Registro de Ingreso

'Deshabilita CodIngreso
txtCodIngreso.BackColor = vbWhite
txtCodIngreso.Enabled = False

' Habilita el formulario
DeshabilitarHabilitarFormulario True

' Carga el optCajaBanco
If msCajaoBanco = "CA" Then
    optCaja.Value = True
Else
    optBanco.Value = True
End If

'Rellena los controles de Caja
txtDocIngreso.Text = curRegIngresoCajaBanco.campo(0)
txtTipDoc.Text = curRegIngresoCajaBanco.campo(2)
mskFecTrab.Text = FechaDMA(Trim(Str(curRegIngresoCajaBanco.campo(3))))
txtMonto.Text = Format(curRegIngresoCajaBanco.campo(4), "###,###,##0.00")
txtCodMov.Text = curRegIngresoCajaBanco.campo(5)
txtObserv.Text = curRegIngresoCajaBanco.campo(8)
txtCodContable.Text = curRegIngresoCajaBanco.campo(6)

'Se carga cbo Afecta con el dato del cursor(Terceros o Pln_Personals)
CargarRegAfecta

If msCajaoBanco = "BA" Then 'Rellena los controles de Banco
    txtBanco.Text = curRegIngresoCajaBanco.campo(9)
    msCtaCte = curRegIngresoCajaBanco.campo(10) 'Actualiza variable de Código de CtaCte
    CD_ActVarCbo cboCtaCte, msCtaCte, mcolCodDesCtaCte
        
End If

'Habilita Botones cancelar,Anular Caja o Bancos
cmdAnular.Enabled = True
cmdCancelar.Enabled = True

End Sub

Private Sub DeshabilitarHabilitarFormulario(bBoleano As Boolean)
'---------------------------------------------------------------
'Propósito : Deshabilita controles editables del Formulario
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
txtAfecta.Enabled = bBoleano: cmdBuscar.Enabled = bBoleano
txtTipDoc.Enabled = bBoleano: cboTipDoc.Enabled = bBoleano
txtCodContable.Enabled = bBoleano: cboCtaContable.Enabled = bBoleano
txtDocIngreso.Enabled = bBoleano
txtObserv.Enabled = bBoleano
txtMonto.Enabled = bBoleano
txtBanco.Enabled = bBoleano: cboBanco.Enabled = bBoleano
cboCtaCte.Enabled = bBoleano:

End Sub


Private Sub CargarRegAfecta()
'---------------------------------------------------------------------
'Propósito  : Carga los registros Afectados
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------------
'Nota
Dim sSQL As String
Dim curAfecta As New clsBD2
Dim curDetPrestamos As New clsBD2

'Verifica a quien afecta el concepto(Tipo Mov)
If curRegIngresoCajaBanco.campo(7) = "Tercero" Then
    'SI el registro de egreso ezta realacionado a terceros
    sSQL = "SELECT IdTercero FROM MOV_TERCEROS WHERE " _
        & "Orden='" & txtCodIngreso & "'"
    'Ejecuta la sentencia de consulta
  curAfecta.SQL = sSQL
  If curAfecta.Abrir = HAY_ERROR Then End 'error se cierra la aplicacion
 'Actualiza el combo afecta
  gsCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
  txtAfecta.MaxLength = Len(gsCodAfectaAnterior)
  txtAfecta.Text = gsCodAfectaAnterior
   lblEtiqueta.Caption = "Tercero:"
     
 'Cierra la consulta
  curAfecta.Cerrar

ElseIf curRegIngresoCajaBanco.campo(7) = "Persona" Then
    'SI el registro de egreso esta relacionado a personal
    sSQL = "SELECT IdPersona FROM MOV_PERSONAL WHERE " _
        & "Orden='" & txtCodIngreso & "'"
    'Ejecuta la sentencia de consulta
     curAfecta.SQL = sSQL
     If curAfecta.Abrir = HAY_ERROR Then End 'error se cierra la aplicacion
    'Actualiza el combo afecta
     gsCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
     txtAfecta.MaxLength = Len(gsCodAfectaAnterior)
     txtAfecta.Text = gsCodAfectaAnterior
      lblEtiqueta.Caption = "Personal:"
     
    'Cierra la consulta
     curAfecta.Cerrar
  
ElseIf curRegIngresoCajaBanco.campo(7) = "Proceso" Then
    'Sentencia SQL donde se encuentra los datos del personal
    sSQL = "SELECT DISTINCT DR.IdPersona, ( PR.Apellidos & ', ' & PR.Nombre), " _
        & "DR.Egreso " _
        & "FROM MOV_ENTREG_RENDIR DR, PLN_PERSONAL PR " _
        & "WHERE DR.Orden='" & txtCodIngreso & "' and DR.IdPersona=PR.IdPersona"
        
    ' Ejecuta la sentencia
    curAfecta.SQL = sSQL
    
    'Verifica si hay error
    If curAfecta.Abrir = HAY_ERROR Then End
    If Not curAfecta.EOF Then ' EL egreso esta relacionado a pago de prestamos
       ' Carga Afecta Persona
       gsCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
       txtAfecta.MaxLength = Len(gsCodAfectaAnterior)
       txtAfecta.Text = gsCodAfectaAnterior
       gdblMontoAnterior = curAfecta.campo(2)
       lblEtiqueta.Caption = "Trabajador:"
  
    End If
    
    'Cierra el cursor
    curAfecta.Cerrar
    
    ' Verifica si se pagó algún prestamo
    sSQL = "SELECT DISTINCT DP.IdPersona, ( PR.Apellidos & ', ' & PR.Nombre), " _
        & "DP.IdConPL, DP.NumPrestamo, P.Monto, PC.CodContable " _
        & "FROM DEVOLUCION_PRESTAMOSCB DP , PRESTAMOS P, PLNCONCEPTOS_OTROS PC, PLN_PERSONAL PR " _
        & "WHERE DP.Orden='" & txtCodIngreso & "' and DP.IdPersona=P.IdPersona and " _
        & "DP.IdConPL=P.IdConPL and DP.NumPrestamo=P.NumPrestamo and DP.IdConPL=PC.IdConPL " _
        & "AND DP.IdPersona=PR.IdPersona"
        
    ' Ejecuta la sentencia
    curAfecta.SQL = sSQL
    If curAfecta.Abrir = HAY_ERROR Then End
    If Not curAfecta.EOF Then ' EL egreso esta relacionado a pago de prestamos
       ' Carga Afecta Persona
       gsCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
       txtAfecta.MaxLength = Len(gsCodAfectaAnterior)
       txtAfecta.Text = gsCodAfectaAnterior
       lblEtiqueta.Caption = "Trabajador:"
         
       'Carga coleccion con registro seleccionado en el grdPrestamo
        gcolPrestamos.Add Item:=curAfecta.campo(2) & "¯" & _
                          curAfecta.campo(3) & "¯" & _
                          curAfecta.campo(5) & "¯" & _
                          curAfecta.campo(4), _
                          Key:=curAfecta.campo(2)
                   
        'Carga en la coleccion mcolDetPrestamos
        'Sentencia sql
        sSQL = ""
        sSQL = "SELECT PC.IdConPl,PC.NumCuota, PC.Cuota, " _
              & "PC.Amortizado, PC.AnioMes, PC.NumPrestamo " _
              & "FROM PRESTAMOS_CUOTAS PC " _
              & "WHERE PC.IdPersona= '" & curAfecta.campo(0) & "' And " _
              & "PC.NumPrestamo='" & curAfecta.campo(3) & "' " _
              & " And PC.IdConPL='" & curAfecta.campo(2) & "' " _
              & " And PC.Cancelado='SI' ORDER BY PC.NumCuota "

        ' Ejecuta la sentencia
        curDetPrestamos.SQL = sSQL
        
        'Verifica si hay error
        If curDetPrestamos.Abrir = HAY_ERROR Then End
        Do While Not curDetPrestamos.EOF
        
            'Agrega a la coleccion los datos del grdDetPrestamos
            gcolDetPrestamos.Add Item:=curDetPrestamos.campo(0) & "¯" & _
                                    curDetPrestamos.campo(5) & "¯" & _
                                    curDetPrestamos.campo(4) & "¯" & _
                                    curDetPrestamos.campo(2), _
                                 Key:=curDetPrestamos.campo(0) & "¯" & _
                                    curDetPrestamos.campo(5) & "¯" & _
                                    curDetPrestamos.campo(4)

                                
            'Mueve al siguiente registro
            curDetPrestamos.MoverSiguiente
            
        Loop
        
        'Cierra el cursor
        curDetPrestamos.Cerrar
        
        ' Cierra el cursor
        curAfecta.Cerrar
        
     End If
End If

End Sub

Private Sub DeshabilitaHabilitaControlesCaja()
'-----------------------------------------------------------------------
'Propósito: Deshabilita, habilita los controles segun la condición de estos
'Recibe: Nada
'Devuelve: Nada
'-----------------------------------------------------------------------
' Nota Llamado desde el evento formLoad y despues ingresar el codigo de ingreso

'???
txtDocIngreso.Enabled = Not txtDocIngreso.Enabled
txtTipDoc.Enabled = Not txtTipDoc.Enabled
cboTipDoc.Enabled = Not cboTipDoc.Enabled
txtMonto.Enabled = Not txtMonto.Enabled
txtCodMov.Enabled = Not txtCodMov.Enabled
cboCodMov.Enabled = Not cboCodMov.Enabled
txtCodContable.Enabled = Not txtCodContable.Enabled
cboCtaContable.Enabled = Not cboCtaContable.Enabled
txtObserv.Enabled = Not txtObserv.Enabled

'Coloca color a controles  del ingreso caja dependiendo del estado
EstableceCamposObligatoriosCaja
    
End Sub

Private Sub DeshabilitaHabilitaControlesBanco()
'-----------------------------------------------------------------------
'Propósito: Deshabilita, habilita los controles segun la condicion de estos
'Recibe: Nada
'Devuelve: Nada
'-----------------------------------------------------------------------
' Nota Llamado desde el evento formLoad y despues ingresar el codigo de ingreso
'???
txtBanco.Enabled = Not txtBanco.Enabled
cboBanco.Enabled = Not cboBanco.Enabled
cboCtaCte.Enabled = Not cboCtaCte.Enabled

'Oculta,Coloca color y contenido a controles  del ingreso Banco dependiendo del estado de estos
 EstableceEstadoCamposBanco
    
End Sub

Private Sub EstableceEstadoCamposBanco()
'-----------------------------------------------------------------------
'Propósito: Oculta, Establece color de los controles referentes a banco
'           dependiendo de la  condicion de estos
'Recibe: Nada
'Devuelve: Nada
'-----------------------------------------------------------------------

If txtBanco.Enabled = True Then 'Hace visible el controles y muestra el contenido
     lblBanco.Visible = True
     txtBanco.Visible = True
     cboBanco.Visible = True
     cmdPBanco.Visible = True
     If txtBanco.Text = "" Then txtBanco.BackColor = Obligatorio 'Campo Obligatorio
 Else 'oculta controles y  establece color blanco a txtbanco
     lblBanco.Visible = False
     txtBanco.Visible = False
     cboBanco.Visible = False
     cmdPBanco.Visible = False
     txtBanco.Text = Empty
     txtBanco.BackColor = vbWhite
 End If
 If cboCtaCte.Enabled = True Then 'Hace visible el controles y muestra el contenido
     lblCtaCte.Visible = True
     cboCtaCte.Visible = True
     cmdPCtaCte.Visible = True
     If cboCtaCte.Text = "" Then cboCtaCte.BackColor = Obligatorio 'Campo obligatorio
 Else 'Oculta controles y establece color blanco a cboCtaCte
     lblCtaCte.Visible = False
     cboCtaCte.Visible = False
     cmdPCtaCte.Visible = False
     cboCtaCte.BackColor = vbWhite
 End If
    
End Sub

Private Sub MostrarUnicoItem()
'----------------------------------------------------------------------------
'Propósito: Mostrar el elemento del combo en caso de que este sea único
'Recibe:   Nada
'Devuelve: Nada
'----------------------------------------------------------------------------

If cboCtaContable.ListCount = 1 Then
  
  cboCtaContable.ListIndex = 0
  'Se lanzan los eventos
  cboCtaContable_Click
  cboCtaContable_LostFocus
  'Se mantiene el alto del combo
  cboCtaContable.Height = CBONORMAL
  
End If

End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)
' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
End Sub

Sub PagarVenta()
  Dim sSQL As String
  Dim modPagoVenta As New clsBD3
  
  sSQL = ""
  ' Guardar los  datos a Caja cuando no es ingreso de prestamos
  sSQL = "INSERT INTO VENTAS_PAGOS VALUES('" & OrdenVenta & "','" _
          & FechaAMD(mskFecTrab.Text) & "'," _
          & Var37(txtMonto.Text) & ")"
    
  'SI al ejecutar hay error se sale de la aplicación
  modPagoVenta.SQL = sSQL
  If modPagoVenta.Ejecutar = HAY_ERROR Then
   End
  End If
     
  'Se cierra la query
  modPagoVenta.Cerrar
  
  If Val(VentaTotal) = (Val(VentaPagada) + Val(txtMonto)) Then
    ActualizarCancelacionVenta
  End If
End Sub

Sub ActualizarCancelacionVenta()
  Dim sSQL As String
  Dim modCancelVenta As New clsBD3
  
  sSQL = "UPDATE VENTAS SET " & _
        "Cancelado='SI' " & _
        "WHERE Orden='" & OrdenVenta & "'"

  modCancelVenta.SQL = sSQL

  If modCancelVenta.Ejecutar = HAY_ERROR Then
    End
  End If

  modCancelVenta.Cerrar
End Sub

