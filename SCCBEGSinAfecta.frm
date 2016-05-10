VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBEGSinAfecta 
   Caption         =   "Caja y Bancos - Egresos sin afectación financiera"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   HelpContextID   =   66
   Icon            =   "SCCBEGSinAfecta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpCtaCte 
      Height          =   255
      Left            =   5940
      Picture         =   "SCCBEGSinAfecta.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdpBanco 
      Height          =   255
      Left            =   3450
      Picture         =   "SCCBEGSinAfecta.frx":0BA2
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.ComboBox cboCtaCte 
      Height          =   315
      Left            =   4560
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   930
      Width           =   1665
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   1065
      Style           =   1  'Simple Combo
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   930
      Width           =   2680
   End
   Begin VB.TextBox txtRinde 
      Height          =   315
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   11
      Top             =   960
      Width           =   675
   End
   Begin VB.CommandButton cmdpTipoMov 
      Height          =   255
      Left            =   7680
      Picture         =   "SCCBEGSinAfecta.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1590
      Width           =   255
   End
   Begin VB.ComboBox cboCodMov 
      Height          =   315
      Left            =   2040
      Style           =   1  'Simple Combo
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   5930
   End
   Begin VB.CommandButton cmdpCodContable 
      Height          =   255
      Left            =   7650
      Picture         =   "SCCBEGSinAfecta.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2790
      Width           =   255
   End
   Begin VB.ComboBox cboCtaContable 
      Height          =   315
      Left            =   2120
      Style           =   1  'Simple Combo
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2760
      Width           =   5820
   End
   Begin VB.CommandButton cmdpTipDoc 
      Height          =   255
      Left            =   5080
      Picture         =   "SCCBEGSinAfecta.frx":142A
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3390
      Width           =   255
   End
   Begin VB.ComboBox cboTipDoc 
      Height          =   315
      Left            =   1750
      Style           =   1  'Simple Combo
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3610
   End
   Begin VB.CommandButton cmdBuscarEgreso 
      Caption         =   "..."
      Height          =   255
      Left            =   2010
      TabIndex        =   1
      Top             =   390
      Width           =   255
   End
   Begin VB.TextBox txtBanco 
      Height          =   315
      Left            =   690
      MaxLength       =   2
      TabIndex        =   5
      Top             =   930
      Width           =   375
   End
   Begin VB.TextBox txtNumCh 
      Height          =   315
      Left            =   6960
      MaxLength       =   15
      TabIndex        =   10
      Top             =   930
      Width           =   1575
   End
   Begin VB.TextBox txtCodContable 
      Height          =   315
      Left            =   1335
      TabIndex        =   18
      Top             =   2760
      Width           =   780
   End
   Begin VB.TextBox txtSaldoCB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3900
      MaxLength       =   12
      TabIndex        =   26
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtAfecta 
      Height          =   315
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2160
      Width           =   675
   End
   Begin VB.TextBox txtCodEgreso 
      Height          =   315
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   1450
   End
   Begin VB.TextBox txtMonto 
      Height          =   315
      Left            =   1335
      MaxLength       =   12
      TabIndex        =   25
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtObserv 
      Height          =   315
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   27
      Top             =   4440
      Width           =   6600
   End
   Begin VB.TextBox txtTipDoc 
      Height          =   315
      Left            =   1335
      MaxLength       =   2
      TabIndex        =   21
      Top             =   3360
      Width           =   420
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   5520
      TabIndex        =   29
      Top             =   5260
      Width           =   1005
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "An&ular"
      Height          =   405
      Left            =   6600
      TabIndex        =   30
      Top             =   5260
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7680
      TabIndex        =   31
      ToolTipText     =   "Volver al Menú Principal"
      Top             =   5260
      Width           =   1005
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   4440
      TabIndex        =   28
      Top             =   5260
      Width           =   1005
   End
   Begin VB.TextBox txtCodMov 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   13
      Top             =   1560
      Width           =   675
   End
   Begin MSMask.MaskEdBox mskFecTrab 
      Height          =   315
      Left            =   7400
      TabIndex        =   32
      Top             =   300
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraCB 
      Caption         =   "Egreso de:"
      Height          =   615
      Left            =   2400
      TabIndex        =   33
      Top             =   140
      Width           =   3375
      Begin VB.OptionButton optRendir 
         Caption         =   "Cuenta a Rendir"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optBanco 
         Caption         =   "Banco"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCaja 
         Caption         =   "Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   8550
      Begin VB.TextBox txtDescRinde 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   960
         Width           =   5400
      End
      Begin VB.CommandButton cmdBuscaRinde 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7320
         Picture         =   "SCCBEGSinAfecta.frx":1702
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7320
         Picture         =   "SCCBEGSinAfecta.frx":1804
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5400
      End
      Begin VB.TextBox txtDocEgreso 
         Height          =   315
         Left            =   6330
         MaxLength       =   15
         TabIndex        =   24
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   5760
         Picture         =   "SCCBEGSinAfecta.frx":1906
         Stretch         =   -1  'True
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblRinde 
         Caption         =   "Cuenta a Rendir:"
         Height          =   375
         Left            =   165
         TabIndex        =   52
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblNumCh 
         Caption         =   "&Num. Cheque:"
         Height          =   375
         Left            =   6210
         TabIndex        =   49
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCtaCte 
         AutoSize        =   -1  'True
         Caption         =   "Nº &Cuenta:"
         Height          =   195
         Left            =   3660
         TabIndex        =   48
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblBanco 
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   60
         TabIndex        =   47
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblCodContable 
         Caption         =   "CtaContable:"
         Height          =   375
         Left            =   165
         TabIndex        =   46
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblSaldoCB 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   2760
         TabIndex        =   45
         Top             =   3915
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         Caption         =   "Mo&vimiento:"
         Height          =   255
         Left            =   165
         TabIndex        =   44
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   6600
         TabIndex        =   43
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Doc. Egreso:"
         Height          =   195
         Left            =   5280
         TabIndex        =   42
         Top             =   3390
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "&Monto (S/):"
         Height          =   255
         Left            =   165
         TabIndex        =   41
         Top             =   3870
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "&Observación:"
         Height          =   255
         Left            =   165
         TabIndex        =   40
         Top             =   4470
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "&Tipo Doc.:"
         Height          =   255
         Left            =   165
         TabIndex        =   39
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Egreso:"
         Height          =   195
         Left            =   165
         TabIndex        =   38
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Afecta"
         Height          =   375
         Left            =   165
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.Label Label18 
      Caption         =   "&Tipo Doc.:"
      Height          =   255
      Left            =   2820
      TabIndex        =   35
      Top             =   2430
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      Height          =   195
      Left            =   2805
      TabIndex        =   34
      Top             =   3000
      Width           =   405
   End
End
Attribute VB_Name = "frmCBEGSinAfecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Colecciones para la carga del combo de Movimientos y demas campos
Private mcolCodMov As New Collection
Private mcolCodDesCodMov As New Collection
Private mcolDesCodAfecta As New Collection
Private mcolCodAfecta As New Collection

'Colecciones para la carga de codigo contable y  el tipo de movimiento
Private mcolDesCodCont As New Collection
Private mcolCodCont As New Collection

'Colecciones para cargar el código contable y su descripción
Private mcolCodPlanCont As New Collection
Private mcolDesCodPlanCont As New Collection

'Colecciones para la carga del codigo y nombre del personal
Private mcolCodPersonal As New Collection
Private mcolDesCodPersonal As New Collection

'Colecciones para la carga del codigo y nombre del personal (Préstamo)
Private mcolCodPersonalPrest As New Collection
Private mcolDesCodPersonalPrest As New Collection

'Colecciones para la carga del codigo y descripción de terceros
Private mcolDesCodTerc As New Collection
Private mcolCodTerc As New Collection

'Colecciones para la carga del combo de Tipo de Documento
Private mcolCodTipDoc As New Collection
Private mcolCodDesTipDoc As New Collection

'Colecciones para la carga del combo de Bancos
Private mcolCodBanco As New Collection
Private mcolCodDesBanco As New Collection

'Colecciones para la carga del combo de Cta Cte
Private mcolCodCtaCte As New Collection
Private mcolCodDesCtaCte As New Collection

'Cursor que carga el registro de egreso para su modificacion
Private mcurRegEgresoCajaBanco As New clsBD2

'Variable donde se carga el codigo equivalente al combobox recuperado
Private msCtaCte As String

'variable que identifica SI el egreso es a caja o bancos
Private msCajaoBanco As String
Private msCaBaAnt As String

' Variable que indica si se cargó un egreso
Private mbEgresoCargado As Boolean

'Determina el maxlength del campo txtAfecta cuando es personal y terceros
Private miTamañoPer As Integer
Private miTamañoTer As Integer

'Variable que identifica a que Afecta el concepto(Tipo_Mov), Terceros o Personal
Private msAfecta As String '(Tercero, Persona o Proceso)
Public msCodAfectaAnterior As String ' (Código de Terceros,Personal o Planilla)
Private msRindeAnt As String ' Cuenta a rendir anterior
Private msProceso As String

'Variable que determina el Mayor Tamaño de CodCont en Conceptos(TipoMov de la BD)
Private miTamañoCodCont As Integer

Private Sub DeshabilitarHabilitarFormulario(bBoleano As Boolean)
'---------------------------------------------------------------
'Propósito : Deshabilita controles editables del Formulario
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
txtAfecta.Enabled = bBoleano: cmdBuscar.Enabled = bBoleano
txtTipDoc.Enabled = bBoleano: cboTipDoc.Enabled = bBoleano
txtCodContable.Enabled = bBoleano: cboCtaContable.Enabled = bBoleano
txtDocEgreso.Enabled = bBoleano
txtObserv.Enabled = bBoleano
txtMonto.Enabled = bBoleano
txtBanco.Enabled = bBoleano: cboBanco.Enabled = bBoleano
cboCtaCte.Enabled = bBoleano: txtNumCh.Enabled = bBoleano
txtRinde.Enabled = bBoleano: txtDescRinde.Enabled = bBoleano: cmdBuscaRinde.Enabled = bBoleano
End Sub

Private Sub cboBanco_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboBanco.ListIndex) = False And cboBanco.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboBanco_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboBanco) = True Then SendKeys "{down}"

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

Private Sub cboCtaContable_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaContable) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaContable_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaContable.ListIndex) = False And cboCtaContable.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse


End Sub

Private Sub cboCtaContable_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
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

Private Sub cboCtaCte_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaCte.ListIndex) = False And cboCtaCte.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCtaCte_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaCte) = True Then SendKeys "{down}"

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
  'NO se encuentra la CtaCte
  msCtaCte = ""
End If

'Carga el saldo de la CtaCte en Soles
CargarSaldo
       
'Cambia el alto del combo
 cboCtaCte.Height = CBONORMAL

'Habilita botón Aceptar
HabilitarBotonAceptar

End Sub

Private Sub cboTipDoc_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboTipDoc.ListIndex) = False And cboTipDoc.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboTipDoc_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboTipDoc) = True Then SendKeys "{down}"

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

Private Function VerificarMontoModificar() As Boolean
'---------------------------------------------------------------------------------------
'Propósito :Verificar SI se puede guardar el Registro con un monto Correcto de acuerdo
'           al saldo de Caja o Ctas de Bancos
'Recibe    : Nada
'Devuelve  :booleano que indica SI el monto esta conforme
'---------------------------------------------------------------------------------------
'Nota llamado desde el evento aceptar en Modificacion de Egreso /SA
Dim dblMonto, dblMontoAnt As Double
'Inicializamos la funcion asumiendo que el monto esta correcto
VerificarMontoModificar = True
dblMontoAnt = mcurRegEgresoCajaBanco.campo(3)
dblMonto = Val(Var37(txtMonto.Text))

'verifica SI es caja
If msCajaoBanco = "CA" Then
    If dblMontoAnt <> dblMonto Then 'SI se modifico el monto
      'Verifica SI monto modificado excede el saldo de Caja o Bancos
      If (dblMontoAnt - dblMonto) + Val(Var37(txtSaldoCB.Text)) < -0.0001 Then
          MsgBox "El monto ingresado excede el saldo disponible de Caja o Bancos ", _
              vbInformation, "Caja-Banco- Modificación de Egreso"
          txtMonto.SetFocus
          VerificarMontoModificar = False
          Exit Function
       End If
    End If
Else 'verifica SI es banco
    'verifica SI es la misma cuenta corriente de egreso Original
    If mcurRegEgresoCajaBanco.campo(9) = msCtaCte Then 'La misma CtaCte
      'Verifica SI monto modificado excede el saldo de Caja o Bancos
      If (dblMontoAnt - dblMonto) + Val(Var37(txtSaldoCB.Text)) < -0.0001 Then
          MsgBox "El Monto Ingresado excede el Saldo de Caja o Bancos ", _
               vbInformation, "Caja-Banco- Modificación de Egreso"
          txtMonto.SetFocus
          VerificarMontoModificar = False
          Exit Function
       End If
    Else 'Se Cambio de CtaCte
      'Verifica SI monto excede el saldo de Caja o Bancos
      If dblMonto > Val(Var37(txtSaldoCB.Text)) Then
          MsgBox "El monto ingresado excede el saldo disponible de Caja o Bancos ", _
              vbInformation, "Caja-Banco- Modificación de Egreso"
          txtMonto.SetFocus
          VerificarMontoModificar = False
          Exit Function
       End If
    End If
End If

End Function

Private Function fbVerificarDatosIntroducidos() As Boolean
' -------------------------------------------------------
' Propósito: Verifica que los datos introducidos sean correctos
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------
Dim dblMontoDet As Double
Dim MiObjeto As Variant
Dim iResult As Integer
' Verifica si todos los datos estan correctos
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
 If gsTipoOperacionEgreso = "Nuevo" Then
    If Val(Var37(txtMonto)) > Val(Var37(txtSaldoCB)) Then
          ' Mensaje ,saldo insuficiente
        MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso sin afectación"
        fbVerificarDatosIntroducidos = False
        If txtMonto.Enabled = True Then txtMonto.SetFocus
        Exit Function
     End If
 Else
     ' Verifica la conformidad con el saldo
    If msCajaoBanco = "CA" Then   'Caja
      If msCaBaAnt = "CA" Then ' Anterior caja
        If (Val(Var37(txtMonto)) - Val(mcurRegEgresoCajaBanco.campo(3))) > Val(Var37(txtSaldoCB)) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso sin afectación"
            fbVerificarDatosIntroducidos = False
            If txtMonto.Enabled = True Then txtMonto.SetFocus
            Exit Function
        End If
      Else ' Anterior rendir
        If Val(Var37(txtMonto)) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
              MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso sin afectación"
              fbVerificarDatosIntroducidos = False
              If txtMonto.Enabled = True Then txtMonto.SetFocus
              Exit Function
        End If
      End If
    ElseIf msCajaoBanco = "BA" Then
        If mcurRegEgresoCajaBanco.campo(9) = msCtaCte Then 'La misma CtaCte
            If (Val(Var37(txtMonto)) - Val(mcurRegEgresoCajaBanco.campo(3))) > Val(Var37(txtSaldoCB)) Then
               ' Mensaje ,saldo insuficiente
                MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso sin afectación"
                fbVerificarDatosIntroducidos = False
                If txtMonto.Enabled = True Then txtMonto.SetFocus
                Exit Function
            End If
        ElseIf mcurRegEgresoCajaBanco.campo(9) <> msCtaCte Then 'Se Cambio de CtaCte
    
            If Val(Var37(txtMonto)) > Val(Var37(txtSaldoCB)) Then
               ' Mensaje ,saldo insuficiente
               MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso sin afectación"
               fbVerificarDatosIntroducidos = False
               If txtMonto.Enabled = True Then txtMonto.SetFocus
              Exit Function
            End If
        End If
    ElseIf msCajaoBanco = "ER" Then
      If msCaBaAnt = "ER" Then ' anterior rendir
        If msRindeAnt = txtRinde Then ' Misma cuenta a rendir
            If (Val(Var37(txtMonto)) - Val(mcurRegEgresoCajaBanco.campo(3))) > Val(Var37(txtSaldoCB)) Then
               ' Mensaje ,saldo insuficiente
                MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso sin afectación"
                fbVerificarDatosIntroducidos = False
                If txtMonto.Enabled = True Then txtMonto.SetFocus
                Exit Function
            End If
        ElseIf msRindeAnt <> txtRinde Then
            If Val(Var37(txtMonto)) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
              MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso sin afectación"
              fbVerificarDatosIntroducidos = False
              If txtMonto.Enabled = True Then txtMonto.SetFocus
              Exit Function
            End If
        End If
      Else ' anterior caja
            If Val(Var37(txtMonto)) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
              MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso sin afectación"
              fbVerificarDatosIntroducidos = False
              If txtMonto.Enabled = True Then txtMonto.SetFocus
              Exit Function
            End If
      End If
    End If
    
    ' Verifica si el adelanto fué cancelado
    If msAfecta = "Proceso" Then
    
        ' Verifica el proceso
        Select Case msProceso
        
        Case "PAGO_PRESTAMOS"
     
        Case "PAGO_ADELANTOS"

            For Each MiObjeto In gcolDetMovCB
                ' Verifica si modificó montos
                If Val(Var37(txtMonto)) <> Val(Var30(MiObjeto, 3)) Then
                
                    ' Asigna el resultado de la verificación de las cuotas
                    iResult = Var10(Var30(MiObjeto, 3), txtAfecta, _
                              FechaAM(mskFecTrab), Val(Var37(txtMonto)))
                    Select Case iResult
                    Case 0 ' Seguir con el proceso
                    Case 1 ' Interrumpir el proceso
                        fbVerificarDatosIntroducidos = False
                        ' Sale de la función
                        Exit Function
                    Case 2 ' No seguir validando las otras cuotas
                    End Select
                    
                    ' Verifica si se puede dar adelantos
                    If Var11(mskFecTrab) = False Then
                        fbVerificarDatosIntroducidos = False
                        Exit Function
                    End If
                 End If ' No se modifica monto
                 
             Next MiObjeto
             
        
        Case "PAGO_PLANILLAS"
     
        End Select
    
    End If ' Fin de verificar Proceso

End If ' FIn de verificar si es nuevo o modificar

' Verificados los datos
fbVerificarDatosIntroducidos = True

End Function

Private Sub cmdAceptar_Click()

'Verifica si existe un documento duplicado
If VerificarDocExiste Then
  'El documento ya ha sido ingresado mandamos mensaje
  If MsgBox("El Número de Documento está duplicado, ¿desea continuar con este mismo número de documento? ", _
        vbQuestion + vbYesNo, _
        "Caja-Bancos- Egreso Sin Afectación Financiera") = vbNo Then
         txtDocEgreso.SetFocus
        Exit Sub
  End If
End If

' Verifica si los datos son correctos
If fbVerificarDatosIntroducidos = False Then
    ' algún dato es incorrecto
    Exit Sub
End If

If gsTipoOperacionEgreso = "Nuevo" Then
 ' Pregunta aceptación de los datos
   If MsgBox("¿Está conforme con los datos?", _
      vbQuestion + vbYesNo, "Caja-Bancos, Egreso sin afectación") = vbYes Then
     'Actualiza la transaccion
      Var8 1, gsFormulario
     
     ' Se guardan los datos del egreso
      GuardarEgreso
   Else: Exit Sub ' Sale
   End If
Else

 ' Mensaje de conformidad de los datos
   If MsgBox("¿Está conforme con las modificaciones realizadas en el Egreso " & txtCodEgreso.Text & "?", _
                  vbQuestion + vbYesNo, "Caja-Bancos, Modificación de egreso sin afectación") = vbYes Then
     'Actualiza la transaccion
      Var8 1, gsFormulario
     
     ' Se Modifican los datos del egreso
       GuardarModificacionesEgreso
   Else: Exit Sub ' Sale
   End If
End If

'Actualiza la transaccion
Var8 -1, Empty

' Mensaje Ok
MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso sin Afectación"

' Limpia la pantalla para una nueva operación, Prepara el formulario
   LimpiarFormulario
   
If gsTipoOperacionEgreso = "Nuevo" Then

  ' Habilita las opciones CB
  fraCB.Enabled = True

 ' Nuevo egreso
   NuevoEgreso
   
Else
 ' cierra el control egreso
   If mbEgresoCargado Then
    mcurRegEgresoCajaBanco.Cerrar
    mbEgresoCargado = False
    mskFecTrab = "__/__/____"
   End If

 ' Se Modifican los datos del egreso
   ModificarEgreso
End If

End Sub
 
Private Sub GuardarAdelantos()
'----------------------------------------------------------------------------
'Propósito: Guarda los datos en la tabla Adelantos
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String
Dim modAdelanto As New clsBD3

sSQL = "INSERT INTO Adelantos VALUES ('" _
         & txtAfecta.Text _
         & "','" & txtCodEgreso.Text _
         & "','" & FechaAMD(mskFecTrab.Text) _
         & "'," & Var37(txtMonto.Text) _
         & ",'NO')"
modAdelanto.SQL = sSQL
If modAdelanto.Ejecutar = HAY_ERROR Then
  End
End If
modAdelanto.Cerrar

End Sub

 Private Sub ModificarAdelanto()
'----------------------------------------------------------------------------
'Propósito: Modifica los datos en la tabla Adelantos
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String
Dim modAdelanto As New clsBD3

sSQL = "UPDATE Adelantos SET " _
         & "idPersona='" & txtAfecta.Text _
         & "', Fecha='" & FechaAMD(mskFecTrab.Text) _
         & "', Cantidad=" & txtMonto.Text _
         & ", Cancelado='NO' " _
         & " WHERE Orden='" & txtCodEgreso.Text & "'"

modAdelanto.SQL = sSQL
If modAdelanto.Ejecutar = HAY_ERROR Then
  End
End If
modAdelanto.Cerrar

End Sub

 Private Sub EliminarAdelanto()
'----------------------------------------------------------------------------
'Propósito: Elimina los datos en la tabla Adelantos
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String
Dim modAdelanto As New clsBD3

sSQL = "DELETE * FROM Adelantos " _
        & " WHERE Orden='" & txtCodEgreso.Text & "'"
         
modAdelanto.SQL = sSQL
If modAdelanto.Ejecutar = HAY_ERROR Then
  End
End If
modAdelanto.Cerrar

End Sub

 Private Sub GuardarModificacionesEgreso()
'----------------------------------------------------------------------------
'Propósito: Modifica los datos del registro Egreso a Caja o a Bancos
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String
Dim modEgresoCajaBanco As New clsBD3

'Carga la sentencia que modifica el registro de ingreso
If msCajaoBanco = "BA" Then
     ' Guardar los  datos
     sSQL = "UPDATE EGRESOS SET " & _
        "NumDoc='" & txtDocEgreso & "'," & _
        "IdTipoDoc='" & txtTipDoc & "'," & _
        "MontoCB=" & Var37(txtMonto.Text) & "," & _
        "CodMov='" & txtCodMov.Text & "'," & _
        "Observ='" & UCase(txtObserv.Text) & "'," & _
        "IdCta='" & msCtaCte & "'," & _
        "NumCheque='" & txtNumCh.Text & "'," & _
        "CodContable='" & txtCodContable.Text & "'," & _
        "Origen='B' " & _
        "WHERE Orden='" & txtCodEgreso.Text & "'"
       ' Carga la colección asiento
       ' Orden,Monto,NumCtaBanc, fecha, observ,Proceso,Origen,optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) & "¯" _
           & msCtaCte & "¯" _
           & FechaAMD(mskFecTrab.Text) & "¯EGRESO BANCO SIN AFECTACION¯SS¯EB¯B"
ElseIf msCajaoBanco = "CA" Then
        ' Guardar los  datos
      sSQL = "UPDATE EGRESOS SET " & _
         "NumDoc='" & txtDocEgreso & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoCB=" & Var37(txtMonto.Text) & "," & _
         "CodMov='" & txtCodMov.Text & "'," & _
         "Observ='" & UCase(txtObserv.Text) & "'," & _
         "CodContable='" & txtCodContable.Text & "'," & _
         "Origen='C' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
         ' Carga la colección asiento
         ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) & "¯" _
           & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯EGRESO CAJA SIN AFECTACION¯SS¯EC¯C"
ElseIf msCajaoBanco = "ER" Then
    ' Guarda los datos
      sSQL = "UPDATE EGRESOS SET " & _
         "NumDoc='" & txtDocEgreso & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoCB=" & Var37(txtMonto.Text) & "," & _
         "CodMov='" & txtCodMov.Text & "'," & _
         "Observ='" & UCase(txtObserv.Text) & "'," & _
         "CodContable='" & txtCodContable.Text & "'," & _
         "Origen='R' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
         ' carga la colección asiento
         'Orden,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) & "¯" _
           & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯EGRESO CAJA SIN AFECTACION¯SS¯EC¯R"
End If
           
' Ejecuta la sentencia
modEgresoCajaBanco.SQL = sSQL
If modEgresoCajaBanco.Ejecutar = HAY_ERROR Then
 End
End If
' Cierra la componente
modEgresoCajaBanco.Cerrar
   
' Modifica el Movimiento Afectado
ModificarMovAfectado

If msCajaoBanco = "ER" Then
    ' Verifica el origen anterior
    If msCaBaAnt = "CA" Then
        ' Crea el registro en entrega a rendir
        Var4 txtCodEgreso, txtRinde, "E", Var37(txtMonto), FechaAMD(mskFecTrab.Text), FechaAMD(mskFecTrab.Text)
    ElseIf msCaBaAnt = "ER" Then
        ' Modifica los datos de entregas a rendir
        Var3 txtCodEgreso, txtRinde, "E", Var37(txtMonto)
    End If
ElseIf msCajaoBanco = "CA" Then
    If msCaBaAnt = "ER" Then
        ' Elimina el movimiento de entregas a rendir
        Var2 txtCodEgreso
    End If
End If

'Realiza la modificación del asiento automático
Conta19

End Sub
 
 Private Function ComprobarAdelantoCancelado() As Boolean
'----------------------------------------------------------------------------
'Propósito: Comprueba si el adelanto a modificar a sido cancelado
'Recibe:  Nada
'Devuelve: Booleano (True: Cancelado; False: NO Cancelado)
'----------------------------------------------------------------------------

Dim curAdelanto As New clsBD2
Dim sSQL As String

sSQL = "SELECT Cancelado FROM Adelantos " _
       & " WHERE Orden='" & txtCodEgreso.Text & "' "
curAdelanto.SQL = sSQL
If curAdelanto.Abrir = HAY_ERROR Then
  End
End If

If UCase(curAdelanto.campo(0)) = "SI" Then
  ComprobarAdelantoCancelado = True
Else
  ComprobarAdelantoCancelado = False
End If

End Function
 
Private Sub GuardarEgreso()
'----------------------------------------------------------------------------
'Propósito: Guarda el Egreso en EGRESOS BD
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota llamado desde el Click Aceptar
Dim sSQL As String
Dim modEgresoCB As New clsBD3

'Verifica si es a Caja
If msCajaoBanco = "CA" Then
        ' Guardar los  datos A Caja
    sSQL = "INSERT INTO EGRESOS VALUES('" & txtCodEgreso & "','','','','','" _
            & txtDocEgreso.Text & "','" & txtTipDoc.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecTrab.Text) & "',''," _
            & Var37(txtMonto.Text) & ",'','','','NO','" & UCase(txtObserv.Text) & "','" _
            & txtCodContable & "','NO','C','')"
         ' Carga la colección asiento
         ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) _
           & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯EGRESO CAJA SIN AFECTACION¯SS¯EC¯C"
            
ElseIf msCajaoBanco = "BA" Then 'Se graba en Banco

    sSQL = "INSERT INTO EGRESOS VALUES('" & txtCodEgreso & "','','','','','" _
            & txtDocEgreso.Text & "','" & txtTipDoc.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecTrab.Text) & "',''," _
            & Var37(txtMonto.Text) & ",'','" _
            & msCtaCte & "','" & txtNumCh.Text & "','NO','" & UCase(txtObserv.Text) & "','" _
            & txtCodContable & "','NO','B','')"
         ' Carga la colección asiento
         ' Orden,Monto,NumCtaBanc, fecha, observ,Proceso, Origen, optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) & "¯" _
           & msCtaCte & "¯" _
           & FechaAMD(mskFecTrab.Text) _
           & "¯EGRESO BANCO SIN AFECTACION¯SS¯EB¯B"
           
ElseIf msCajaoBanco = "ER" Then
        ' Guardar los  datos A Caja
    sSQL = "INSERT INTO EGRESOS VALUES('" & txtCodEgreso & "','','','','','" _
            & txtDocEgreso.Text & "','" & txtTipDoc.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecTrab.Text) & "',''," _
            & Var37(txtMonto.Text) & ",'','','','NO','" & UCase(txtObserv.Text) & "','" _
            & txtCodContable & "','NO','R','')"
         ' Carga la colección asiento
         ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
         gcolAsiento.Add _
         Key:=txtCodEgreso, _
         Item:=txtCodEgreso & "¯" _
           & Var37(txtMonto) _
           & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯EGRESO CAJA SIN AFECTACION¯SS¯EC¯R"
End If
  
' Si al ejecutar hay error se sale de la aplicación
modEgresoCB.SQL = sSQL
If modEgresoCB.Ejecutar = HAY_ERROR Then
 End
End If

' Se cierra la query
modEgresoCB.Cerrar
   
' Guarda los movimientos afectados
GuardarMovAfectado

If msCajaoBanco = "ER" Then
' Guarda el movimiento de ERendir
Var4 txtCodEgreso, txtRinde, "E", Var37(txtMonto), FechaAMD(mskFecTrab.Text), FechaAMD(mskFecTrab.Text)
End If

' Realiza el asiento automático
Conta13


End Sub

Private Sub GuardarMovAfectado()
'----------------------------------------------------------------------------
'Propósito: Guarda el Egreso relacionado con Terceros, Personal o Proceso
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el Click Aceptar al hacer un nuevo egreso
Dim modAfectado As New clsBD3
Dim sSQL As String

If msAfecta = "Tercero" Then
  'Cargamos sentencia que guarda en BD MOV_TERCERO
  sSQL = "INSERT INTO MOV_TERCEROS VALUES('" _
          & txtCodEgreso.Text & "','" _
          & Trim(txtAfecta.Text) & "')"
    modAfectado.SQL = sSQL
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    modAfectado.Cerrar
    
  ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
          
ElseIf msAfecta = "Persona" Then
  'Cargamos sentencia que guarda en BD MOV_PERSONAL
    sSQL = "INSERT INTO MOV_PERSONAL VALUES('" _
          & txtCodEgreso.Text & "','" _
          & Trim(txtAfecta.Text) & "')"
    modAfectado.SQL = sSQL
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    modAfectado.Cerrar
          
  ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
   
ElseIf msAfecta = "Proceso" Then
    ' Verifica el proceso
    Select Case msProceso
    Case "PAGO_PLANILLAS"
        ' Guarda el movimiento afectado Planilas
        GuardaAfectaPlanillas
    Case "PAGO_ADELANTOS"
        ' Guarda el movimiento afectado adelantos
        GuardaAfectaAdelantos
    Case "PAGO_PRESTAMOS"
        ' Guarda el movimiento afecta prestamos
        GuardaAfectaPrestamos
    Case "ENTREGA_RENDIR"
        ' Guarda el movimiento afectado Entregas a rendir
        GuardaAfectaCuentasRendir
    End Select
Else 'NO afecta a mas tablas de la BD
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                           & Var37(txtMonto), _
                       Key:=txtCodContable
  
End If

End Sub

Private Sub GuardaAfectaPlanillas()
'----------------------------------------------------------------------
'Propósito: Guarda los datos relacionados al pago de planillas
'Recibe: Nada
'Enatrega: Nada
'----------------------------------------------------------------------
Dim sSQL As String
Dim curPlanillas As New clsBD3
Dim MiObjeto As Variant

For Each MiObjeto In gcolDetMovCB

    ' Carga códigos contables del movimiento en detAsiento
    gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                           & Var30(MiObjeto, 3), _
                       Key:=Var30(MiObjeto, 2)
                       
    ' Guarda en Pago_Planillas los montos pagados a las contracuentas comprometidas
    ' CodPlanilla,Orden,CtaCTb,Monto - PAGO_PLANILLAS
    ' CodPlanilla,CodContable,Monto -  gcolDetMovCB
    sSQL = "INSERT INTO PAGO_PLANILLAS VALUES('" & Var30(MiObjeto, 1) & "','" _
          & txtCodEgreso & "','" & Var30(MiObjeto, 2) & "'," _
          & Var30(MiObjeto, 3) & ")"
          
    ' Ejecuta la sentencia
    curPlanillas.SQL = sSQL
    If curPlanillas.Ejecutar = HAY_ERROR Then End
    curPlanillas.Cerrar
                    
Next MiObjeto

' Actualiza la tabla planillas
ActualizaCanceladoPlanilla

' Vacía la colección
Set gcolDetMovCB = Nothing

End Sub

Private Sub ActualizaCanceladoPlanilla()
' ---------------------------------------------------
' Propósito: Actualiza el campo cancelado de planillas si se ha _
             pagado enteramente la planilla
' Recibe : Nada
' Entrega : Nada
' ---------------------------------------------------
Dim sSQL As String
Dim curPlanilla As New clsBD2
Dim modPlanilla As New clsBD3
Dim dblMontoPlanilla, dblMontoPagado As Double

' Averigua el total de la planilla
sSQL = "SELECT sum(PC.Monto) FROM PLN_CTB_TOTALES PC " _
     & "WHERE PC.CodPlanilla='" & txtAfecta & "'"
curPlanilla.SQL = sSQL
If curPlanilla.Abrir = HAY_ERROR Then End
If curPlanilla.EOF Then ' Verifica si es vacio
    ' Error
    MsgBox "No se tiene montos totales para la planilla" & Chr(13) _
        & "Debe algún error en BD", , "SGCcaijo-Pago de planillas"
    ' Sale
    End
ElseIf IsNull(curPlanilla.campo(0)) Then
    ' Error
    MsgBox "No se tiene montos totales para la planilla" & Chr(13) _
        & "Debe algún error en BD", , "SGCcaijo-Pago de planillas"
    ' Sale
    End
Else ' Asigna el monto a la variable
    dblMontoPlanilla = curPlanilla.campo(0)
End If

' Cierra la componente
curPlanilla.Cerrar

' Averigua el total cancelado
sSQL = "SELECT sum(PP.Monto) FROM PAGO_PLANILLAS PP, EGRESOS E " _
     & "WHERE PP.Orden=E.Orden and E.Anulado='NO' and " _
     & "PP.CodPlanilla='" & txtAfecta & "'"
curPlanilla.SQL = sSQL
If curPlanilla.Abrir = HAY_ERROR Then End
If curPlanilla.EOF Then ' Verifica si es vacio
    ' No se pago nada
    dblMontoPagado = 0
ElseIf IsNull(curPlanilla.campo(0)) Then
    ' No se pago nada
    dblMontoPagado = 0
Else ' Asigna el monto a la variable
    dblMontoPagado = curPlanilla.campo(0)
End If

' Cierra la componente
curPlanilla.Cerrar

' Verifica si se ha cancelado la planilla
If Val(dblMontoPlanilla) = Val(dblMontoPagado) Then
  ' Se ha pagado toda la planilla, actualiza el campo pagado de la planilla
  sSQL = "UPDATE PLN_PLANILLAS SET PagadoCB='SI' " _
       & "WHERE CodPlanilla='" & txtAfecta & "'"
  ' Ejecuta la sentencia
  modPlanilla.SQL = sSQL
  If modPlanilla.Ejecutar = HAY_ERROR Then End
  ' Cierra la componente
  modPlanilla.Cerrar
End If

End Sub

Private Sub GuardaAfectaPrestamos()
'----------------------------------------------------------------------
'Propósito: Guarda los datos relacionados al pago de prestamos
'Recibe: Nada
'Enatrega: Nada
'----------------------------------------------------------------------
Dim sSQL As String
Dim curPrestamos As New clsBD3
Dim MiObjeto As Variant

For Each MiObjeto In gcolDetMovCB
' Actualiza la tabla prestamos
sSQL = "UPDATE PRESTAMOS SET PagadoCB='SI' " _
      & "WHERE IdPersona='" & txtAfecta & "' and " _
      & "IdConPl='" & Var30(MiObjeto, 1) & "' and " _
      & "NumPrestamo='" & Var30(MiObjeto, 2) & "'"
 ' Ejecuta la sentencia
curPrestamos.SQL = sSQL
If curPrestamos.Ejecutar = HAY_ERROR Then End
curPrestamos.Cerrar

' Guarda en Pago_Prestamos
sSQL = "INSERT INTO PAGO_PRESTAMOS VALUES('" & txtAfecta & "','" _
      & Var30(MiObjeto, 1) & "','" _
      & Var30(MiObjeto, 2) & "','" _
      & txtCodEgreso & "')"
 ' Ejecuta la sentencia
curPrestamos.SQL = sSQL
If curPrestamos.Ejecutar = HAY_ERROR Then End
curPrestamos.Cerrar

' Carga códigos contables del movimiento en detAsiento
gcolAsientoDet.Add Item:=Var30(MiObjeto, 3) & "¯" _
                       & Var30(MiObjeto, 4), _
                   Key:=Var30(MiObjeto, 3)
                    
Next MiObjeto

' Vacía la colección
Set gcolDetMovCB = Nothing

End Sub

Private Sub GuardaAfectaCuentasRendir()
'----------------------------------------------------------------------
'Propósito: Guarda los datos relacionados a las entregas a rendir
'Recibe: Nada
'Enatrega: Nada
'----------------------------------------------------------------------
Dim sSQL As String
Dim curAdelantos As New clsBD3
Dim MiObjeto As Variant
'col : IdPersona,CtaCtb,Monto
For Each MiObjeto In gcolDetMovCB
    ' Guarda en mov_entregas a rendir
    Var4 txtCodEgreso, Var30(MiObjeto, 1), "I", Var30(MiObjeto, 3), FechaAMD(mskFecTrab.Text), FechaAMD(mskFecTrab.Text)

    ' Carga códigos contables del movimiento en detAsiento ctactb,monto
    gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                           & Var30(MiObjeto, 3), _
                       Key:=Var30(MiObjeto, 2)
Next MiObjeto

' Vacía la colección detalle
Set gcolDetMovCB = Nothing

End Sub

Private Sub GuardaAfectaAdelantos()
'----------------------------------------------------------------------
'Propósito: Guarda los datos relacionados al pago de Adelantos
'Recibe: Nada
'Enatrega: Nada
'----------------------------------------------------------------------
Dim sSQL As String
Dim curAdelantos As New clsBD3
Dim MiObjeto As Variant
        
For Each MiObjeto In gcolDetMovCB
' Guarda en Adelantos
sSQL = "INSERT INTO ADELANTOS VALUES('" & txtAfecta & "','" _
      & Var30(MiObjeto, 1) & "','" _
      & txtCodEgreso & "','" _
      & FechaAMD(mskFecTrab) & "','" _
      & Var30(MiObjeto, 3) & "','NO')"
 ' Ejecuta la sentencia
curAdelantos.SQL = sSQL
If curAdelantos.Ejecutar = HAY_ERROR Then End
curAdelantos.Cerrar


' Carga códigos contables del movimiento en detAsiento
gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                       & Var30(MiObjeto, 3), _
                   Key:=Var30(MiObjeto, 2)
                    
Next MiObjeto

' Vacía la colección
Set gcolDetMovCB = Nothing

End Sub

Private Sub ModificarMovAfectado()
'----------------------------------------------------------------------------
'Propósito: Modifica  Egreso relacionado con Terceros o Personal
'Recibe:  sPersTercsAnt string que indica la SI el egreso estaba relacioado con
'         la tabla Terceros,Pln_Personal o ninguno
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim modAfectado As New clsBD3
Dim sSQL As String
Dim MiObjeto As Variant
If msAfecta = "Tercero" Then
  ' Cargamos sentencia que guarda en BD MOV_TERCERO
    sSQL = "UPDATE MOV_TERCEROS SET IdTercero='" & txtAfecta.Text _
         & "' WHERE Orden='" & txtCodEgreso.Text & "'"
    modAfectado.SQL = sSQL
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    modAfectado.Cerrar
    
  ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
          
ElseIf msAfecta = "Persona" Then
  'Cargamos sentencia que guarda en BD MOV_PERSONAL
    sSQL = "UPDATE MOV_PERSONAL SET IdPersona='" & txtAfecta.Text _
         & "' WHERE Orden='" & txtCodEgreso.Text & "'"
    modAfectado.SQL = sSQL
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    modAfectado.Cerrar
          
  ' Carga el código contable del movimiento en colDetAsiento
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable
   
ElseIf msAfecta = "Proceso" Then
    ' Verifica el proceso
    Select Case msProceso
    Case "ENTREGA_RENDIR"
    For Each MiObjeto In gcolDetMovCB
      ' col : IdPersona,CtaCtb,Monto
      ' Carga los códigos contables del movimiento  en det asiento
      Var3 txtCodEgreso, Var30(MiObjeto, 1), "I", Var30(MiObjeto, 3)
      ' Carga códigos contables del movimiento en detAsiento
      gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                         & Var30(MiObjeto, 3), _
                     Key:=Var30(MiObjeto, 2)
    Next MiObjeto
    
    Case "PAGO_PRESTAMOS"
    For Each MiObjeto In gcolDetMovCB
      ' Carga códigos contables del movimiento en detAsiento
      gcolAsientoDet.Add Item:=Var30(MiObjeto, 3) & "¯" _
                         & Var30(MiObjeto, 4), _
                     Key:=Var30(MiObjeto, 3)
    Next MiObjeto
    
    Case "PAGO_ADELANTOS"
    For Each MiObjeto In gcolDetMovCB
     ' Actualiza el monto del adelanto
    sSQL = "UPDATE ADELANTOS SET Monto='" & Var37(txtMonto) _
         & "' WHERE Orden='" & txtCodEgreso.Text & "'"
    modAfectado.SQL = sSQL
    If modAfectado.Ejecutar = HAY_ERROR Then
        End 'Finaliza la aplicacion indicando el error en SQL
    End If
    modAfectado.Cerrar
     ' Carga códigos contables del movimiento en detAsiento
       gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                              & Var37(txtMonto), _
                          Key:=Var30(MiObjeto, 2)
    Next MiObjeto
    
    Case "PAGO_PLANILLAS"
    For Each MiObjeto In gcolDetMovCB
    ' Carga códigos contables del movimiento en detAsiento
       gcolAsientoDet.Add Item:=Var30(MiObjeto, 2) & "¯" _
                              & Var30(MiObjeto, 3), _
                          Key:=Var30(MiObjeto, 2)
    Next MiObjeto
    
    End Select

Else 'No afecta a mas tablas de la BD
    gcolAsientoDet.Add Item:=txtCodContable & "¯" _
                            & Var37(txtMonto), _
                       Key:=txtCodContable

End If

End Sub

Private Function VerificarDocExiste() As Boolean
'--------------------------------------------------------------------
'Propósito: Verifica SI el Doc ha sido ingresado en caja o bancos, SI NO
'Recibe:    Nada
'Devuelve:  False:NO existe, True: Existe
'Nota:      llamado desde el evento click de Aceptar
'--------------------------------------------------------------------
Dim sSQL As String
Dim curDocIngresado As New clsBD2

VerificarDocExiste = False

'Verifica SI el doc ingresado sea el mismo del registro en modificacion
If gsTipoOperacionEgreso = "Modificar" Then
    If txtDocEgreso.Text = mcurRegEgresoCajaBanco.campo(0) Then ' es el mismo del registro, NO hace nada
        Exit Function 'Sale de la funcion
    End If
End If

'Verifica SI el Doc esta en Caja o en Banco de la tabla Egresos
'Se averigua SI existe algun documento con el mismo numero en Banco
sSQL = "SELECT Count(E.NumDoc) as NroDoc FROM EGRESOS E " & _
       "WHERE E.NumDoc = '" & txtDocEgreso.Text & "'"
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
Dim modAnularEgreCajaBanco As New clsBD3
Dim sSQL As String

'Verifica, si se puede anular el egreso CA
If fbOkAnular = False Then Exit Sub

'Preguntar si desea Anular el registro de Ingreso a Banco
'Mensaje de conformidad de los datos
If MsgBox("¿Seguro que desea anular el egreso " & txtCodEgreso & "?", _
              vbQuestion + vbYesNo, "Caja-Bancos- Egreso con Afectación") = vbYes Then
     'Actualiza la transaccion
      Var8 1, gsFormulario
   
    'Cambiar el campo Anulado de Ingresos a "SI"
     sSQL = "UPDATE EGRESOS SET " & _
        "Anulado='SI'" & _
        "WHERE Orden='" & txtCodEgreso & "'"
    
    'SI al ejecutar hay error se sale de la aplicación
     modAnularEgreCajaBanco.SQL = sSQL
     If modAnularEgreCajaBanco.Ejecutar = HAY_ERROR Then
      End
     End If
    'Se cierra la query
    modAnularEgreCajaBanco.Cerrar
     
     'Verifica si es egreso de Caja o Banco
     If msCaBaAnt = "CA" Then
        ' carga la colección asiento para anular
        ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
        gcolAsiento.Add _
        Key:=txtCodEgreso, _
        Item:=txtCodEgreso & "¯Nulo¯Nulo¯Nulo¯EGRESO CAJA SIN AFECTACION¯SS¯EC¯C"
    ElseIf msCaBaAnt = "BA" Then
        ' carga la colección asiento para anular
        ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
        gcolAsiento.Add _
        Key:=txtCodEgreso, _
        Item:=txtCodEgreso & "¯Nulo¯Nulo¯Nulo¯EGRESO BANCO SIN AFECTACION¯SS¯EB¯B"
    ElseIf msCaBaAnt = "ER" Then
        ' carga la colección asiento para anular
        ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
        gcolAsiento.Add _
        Key:=txtCodEgreso, _
        Item:=txtCodEgreso & "¯Nulo¯Nulo¯Nulo¯EGRESO BANCO SIN AFECTACION¯SS¯EB¯R"
    End If
    
    ' Anula el asiento automatico
    Conta22
    
    ' Anula la información relacionada con Terceros, Personal o Procesos
    AnularMovAfectado
    
    ' Verifica el origen anterior
    If msCaBaAnt = "ER" Then
        ' Modifica los datos de entregas a rendir
        Var1 txtCodEgreso, msRindeAnt
    End If
    
    ' Actualiza la transaccion
    Var8 -1, Empty
    
    ' Mensaje Ok
    MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso con Afectación"
    
    ' Limpia la pantalla para una nueva operación, Prepara el formulario
    LimpiarFormulario
       
    If gsTipoOperacionEgreso = "Nuevo" Then
     ' Nuevo egreso
       NuevoEgreso
    Else
     ' Cierra el control egreso
       If mbEgresoCargado Then
        mcurRegEgresoCajaBanco.Cerrar
        mbEgresoCargado = False
        mskFecTrab = "__/__/____"
       End If
    
     ' Se Modifican los datos del egreso
       ModificarEgreso
    End If

End If

End Sub

Private Function fbOkAnular() As Boolean
'----------------------------------------------------------------------------
'Propósito: Verifica si se puede anular el egreso
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim curAfectado As New clsBD2
Dim sSQL As String
Dim MiObjeto As Variant

'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    'Devuelve el resultado
    fbOkAnular = False
    Exit Function
End If

If msAfecta = "Proceso" Then
    
    ' Verifica el proceso
    Select Case msProceso
    
    Case "ENTREGA_RENDIR"
    ' Verifica si se puede anular la entrega a rendir
    If VerificarAnularRendir = False Then
        ' No se puede anular
        fbOkAnular = False
        Exit Function
    End If
    
    Case "PAGO_PRESTAMOS"
    
    For Each MiObjeto In gcolDetMovCB
        sSQL = "SELECT * FROM PRESTAMOS_CUOTAS " _
             & "WHERE IdPersona='" & txtAfecta & "' and " _
             & "IdConPL='" & Var30(MiObjeto, 1) & "' and " _
             & "NumPrestamo='" & Var30(MiObjeto, 2) & "' and " _
             & "Cancelado='SI'"
       ' Averigua si se ha pagado alguna cuota
       curAfectado.SQL = sSQL
       If curAfectado.Abrir = HAY_ERROR Then End
       ' Averigua si existe alguna cuoita cancelada
       
       If Not curAfectado.EOF Then ' Existe alguna cuota cancelada
           ' Mensaje
           MsgBox "No se puede anular el prestamo, se han cancelado algunas cuotas del Prestamo" _
                   , , "SGCcaijo-Egreso sin Afectación"
           ' Devuelve resultado de la función
             fbOkAnular = False
           ' Cierra la componente
             curAfectado.Cerrar
           ' Sale del formulario
             Exit Function
       End If
       curAfectado.Cerrar
    Next MiObjeto
    
    Case "PAGO_ADELANTOS"
    For Each MiObjeto In gcolDetMovCB
        sSQL = "SELECT * FROM ADELANTOS " _
             & "WHERE IdPersona='" & txtAfecta & "' and " _
             & "IdConPL='" & Var30(MiObjeto, 1) & "' and " _
             & "Orden='" & txtCodEgreso & "' and Cancelado='SI'"
       ' Averigua si se ha cancelado el adelanto
       curAfectado.SQL = sSQL
       If curAfectado.Abrir = HAY_ERROR Then End
       ' Averigua si existe alguna cuota cancelada
       
       If Not curAfectado.EOF Then ' Existe alguna cuota cancelada
           ' Mensaje
           MsgBox "No se puede anular el egreso, se ha cancelado el adelanto" _
                   , , "SGCcaijo-Egreso sin Afectación"
           ' Devuelve resultado de la función
             fbOkAnular = False
           ' Cierra la componente
             curAfectado.Cerrar
           ' Sale del formulario
             Exit Function
       End If
       curAfectado.Cerrar
    Next MiObjeto
    
    Case "PAGO_PLANILLAS"
    
       sSQL = "SELECT * FROM PAGO_PLANILLAS P, EGRESOS E " _
            & "WHERE P.Orden=E.Orden and E.Anulado='NO'" _
            & "and P.CodPlanilla>'" & txtAfecta & "'"
            
       ' Averigua si se ha cancelado el adelanto
       curAfectado.SQL = sSQL
       If curAfectado.Abrir = HAY_ERROR Then End
       ' Averigua si existe alguna cuota cancelada
       
       If Not curAfectado.EOF Then ' Existe alguna cuota cancelada
           ' Mensaje
           MsgBox "No se puede anular el egreso, se han pagado planillas posteriores" _
                   , vbInformation + vbOKOnly, "SGCcaijo-Egreso sin Afectación"

           ' Devuelve resultado de la función
             fbOkAnular = False
           ' Cierra la componente
             curAfectado.Cerrar
           ' Sale del formulario
             Exit Function
       End If
       curAfectado.Cerrar
  
  End Select
  
End If ' Fin de verificar procesos


' Devuelve la función
fbOkAnular = True

End Function

Private Function VerificarAnularRendir() As Boolean
'---------------------------------------------------------------------------------------
'Propósito : Verificar si se puede anular la entrega a rendir
'Recibe    : Nada
'Devuelve  : Booleano que indica si se puede anular
'---------------------------------------------------------------------------------------
Dim dblMontoAnt, dblSaldo As Double

' Inicializamos la funcion asumiendo que el monto esta correcto
VerificarAnularRendir = True
dblMontoAnt = Val(mcurRegEgresoCajaBanco.campo(3))
' Averigua el saldo de la cuenta a rendir
dblSaldo = Var6(gsCodAfectaAnterior)

' Verifica si el monto excede el saldo de la cuenta a rendir
If Val(dblMontoAnt) > Val(dblSaldo) Then
    MsgBox "No se puede anular la entrega a rendir." & Chr(13) & _
          "No existe saldo suficiente en la cuenta original del egreso", _
        vbInformation + vbOKOnly, "Caja-Banco- Anulación de entregas a rendir"
    VerificarAnularRendir = False
    Exit Function
 End If

End Function

Private Sub AnularMovAfectado()
'----------------------------------------------------------------------------
'Propósito: Anula, elimina o actualiza los movimientos relacionados al _
            egreso
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim modAfectado As New clsBD3
Dim sSQL As String
Dim MiObjeto As Variant

If msAfecta = "Tercero" Then
  'Cargamos sentencia que guarda en BD MOV_TERCERO
' sSQL = "DELETE * FROM MOV_TERCEROS " _
'      & "WHERE Orden='" & txtCodEgreso.Text & "'"
'    modAfectado.SQL = sSQL
'    If modAfectado.Ejecutar = HAY_ERROR Then
'        End 'Finaliza la aplicacion indicando el error en SQL
'    End If
'    modAfectado.Cerrar
    
         
ElseIf msAfecta = "Persona" Then
'  'Cargamos sentencia que guarda en BD MOV_PERSONAL
'    sSQL = "DELETE * FROM MOV_PERSONAL " _
'         & " WHERE Orden='" & txtCodEgreso.Text & "'"
'    modAfectado.SQL = sSQL
'    If modAfectado.Ejecutar = HAY_ERROR Then
'        End 'Finaliza la aplicacion indicando el error en SQL
'    End If
'    modAfectado.Cerrar
'
   
ElseIf msAfecta = "Proceso" Then
    ' Verifica el proceso
    Select Case msProceso
    
    Case "PAGO_PRESTAMOS"
    'Carga sentencia que Modifica PRESTAMOS
      sSQL = "UPDATE PRESTAMOS SET PagadoCB='NO' " _
           & " WHERE IdPersona='" & txtAfecta & "' and " _
           & "IdConPl='" & Var30(gcolDetMovCB.Item(1), 1) & "' and " _
           & "NumPrestamo='" & Var30(gcolDetMovCB.Item(1), 2) & "'"
      modAfectado.SQL = sSQL
      If modAfectado.Ejecutar = HAY_ERROR Then
          End 'Finaliza la aplicacion indicando el error en SQL
      End If
      modAfectado.Cerrar
    
'    'Cargamos sentencia que elimina de pago prestamos
'      sSQL = "DELETE * FROM PAGO_PRESTAMOS " _
'           & " WHERE Orden='" & txtCodEgreso.Text & "'"
'      modAfectado.SQL = sSQL
'      If modAfectado.Ejecutar = HAY_ERROR Then
'          End 'Finaliza la aplicacion indicando el error en SQL
'      End If
'      modAfectado.Cerrar
    
    Case "PAGO_ADELANTOS"
'    'Cargamos sentencia que elimina de adelantos
'      sSQL = "DELETE * FROM ADELANTOS " _
'           & " WHERE Orden='" & txtCodEgreso.Text & "'"
'      modAfectado.SQL = sSQL
'      If modAfectado.Ejecutar = HAY_ERROR Then
'          End 'Finaliza la aplicacion indicando el error en SQL
'      End If
'      modAfectado.Cerrar
    
    Case "PAGO_PLANILLAS"
     'Carga sentencia que Modifica PLN_PLANILLAS
      sSQL = "UPDATE PLN_PLANILLAS SET PagadoCB='NO' " _
           & " WHERE CodPlanilla='" & txtAfecta & "'"
      modAfectado.SQL = sSQL
      If modAfectado.Ejecutar = HAY_ERROR Then
          End 'Finaliza la aplicacion indicando el error en SQL
      End If
      modAfectado.Cerrar
   
    
'    'Cargamos sentencia que elimina de pago planillas
'      sSQL = "DELETE * FROM PAGO_PLANILLAS " _
'           & " WHERE CodPlanilla='" & txtAfecta.Text & "' and " _
'           & " Orden='" & txtCodEgreso.Text & "'"
'      modAfectado.SQL = sSQL
'      If modAfectado.Ejecutar = HAY_ERROR Then
'          End 'Finaliza la aplicacion indicando el error en SQL
'      End If
'      modAfectado.Cerrar
    Case "ENTREGA_RENDIR"
     'Carga sentencia que Modifica PLN_PLANILLAS
      sSQL = "UPDATE MOV_ENTREG_RENDIR SET Anulado='SI' " _
           & " WHERE Orden='" & txtCodEgreso & "'"
      modAfectado.SQL = sSQL
      If modAfectado.Ejecutar = HAY_ERROR Then
          End 'Finaliza la aplicacion indicando el error en SQL
      End If
      modAfectado.Cerrar
    
    End Select

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
txtDocEgreso.Text = Empty
txtTipDoc.Text = Empty
cboTipDoc.ListIndex = -1
txtMonto.Text = Empty
txtObserv.Text = Empty
' Limpia controles banco
txtBanco = Empty
txtNumCh = Empty
' Limpia controles a rendir
txtRinde = Empty
' limpia resumen
txtSaldoCB.Text = "0.00"
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
ElseIf msAfecta = "Tercero" Then
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
ElseIf msAfecta = "Proceso" Then
    ' Verifica si el proceso es entregas a rendir
    If msProceso = "ENTREGA_RENDIR" Then ' Entregas a rendir
        'Muestra el formulario de Entregas a rendir
        frmCBEGEntrega_Rendir.Show vbModal, Me
    End If
End If

End Sub

Private Sub cmdBuscarEgreso_Click()

' Define el tipo de selección del Orden
gsTipoSeleccionOrden = "EgresoSA"
gsOrden = Empty

' Muestra el formulario para elegir el egreso
frmCBSelOrden.Show vbModal, Me

' Verifica si se eligió algun dato a modificar
If gsOrden <> Empty Then
  txtCodEgreso = gsOrden
End If

End Sub

Private Sub cmdBuscaRinde_Click()
'Determina si existen cuentas a rendir del personal
If gcolTRendir.Count = 0 Then
    'Mensaje de nos hay cuentas a rendir del personal
    MsgBox "No hay cuentas a rendir del personal", vbOKOnly + vbInformation, "SGCcaijo-Ingresos, fondos a rendir"
    'Decarga el formulario
    Exit Sub
End If

' Carga los títulos del grid selección
  giNroColMNSel = 3
  aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Monto")
  aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Saldo a Rendir")
  aTamañosColumnas = Array(1000, 5000, 1400)
' Muestra el formulario de busqueda
  frmMNSelecERendir.Show vbModal, Me

' Verifica si se eligió algun dato a modificar
  If gsCodigoMant <> Empty Then
    txtRinde.Text = gsCodigoMant
    SendKeys "{tab}"
  End If
  
End Sub

Private Sub cmdCancelar_Click()

' Limpia el formulario y pone en blanco variables
LimpiarFormulario

' Verifica el tipo operación
If gsTipoOperacionEgreso = "Nuevo" Then
  ' Habilita las opciones CB
  fraCB.Enabled = True
  ' Prepara el formulario
  NuevoEgreso
Else
  ' cierra el control egreso
   If mbEgresoCargado Then
    mcurRegEgresoCajaBanco.Cerrar
    mbEgresoCargado = False
    mskFecTrab = "__/__/____"
   End If
  ' Prepara el formulario
  ModificarEgreso
End If
         
End Sub



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

Private Sub cmdpTipoMov_Click()

If cboCodMov.Enabled Then
    ' alto
     cboCodMov.Height = CBOALTO
    ' focus a cbo
     cboCodMov.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()

' Cierra el formulario
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

'Se cargan los combos
'Carga el combo tipo movimiento y las colecciones de tipo_mov
CargarColTipo_Mov

'Se carga el combo de Tipo de Documento
sSQL = "SELECT idTipoDoc, DescTipoDoc FROM TIPO_DOCUM " & _
           "WHERE RelacProc='SS' or RelacProc='SA' " _
           & " ORDER BY DescTipoDoc"
CD_CargarColsCbo cboTipDoc, sSQL, mcolCodTipDoc, mcolCodDesTipDoc

'Se carga el combo de Cta Cte
sSQL = "SELECT IdCta, DescCta FROM TIPO_CUENTASBANC " & _
           "WHERE IdMoneda= 'SOL'   ORDER BY DescCta"
CD_CargarColsCbo cboCtaCte, sSQL, mcolCodCtaCte, mcolCodDesCtaCte

'Se carga el combo de Bancos
sSQL = "SELECT DISTINCT b.IdBanco,b.DescBanco FROM TIPO_BANCOS B , TIPO_CUENTASBANC C" _
       & " WHERE b.idbanco = c.idbanco And c.idmoneda = 'SOL'" _
       & " ORDER BY DescBanco"

CD_CargarColsCbo cboBanco, sSQL, mcolCodBanco, mcolCodDesBanco

'Se Limpia el Combo de Cts Corrientes en dólares
cboCtaCte.Clear

'Se carga la colección de Personal
CargarColPersonal

'Se carga la colección de terceros
CargarColTerceros

'Establece campos obligatorios del formulario
EstableceCamposObligatorios

' Dependiendo de la operación a realizar prepara el formulario
If gsTipoOperacionEgreso = "Nuevo" Then
    ' Deshabilita el txtCodEgreso
    txtCodEgreso.Enabled = False
    
    ' Deshabilita el botón elegir
    cmdBuscarEgreso.Enabled = False
    
    ' Prepara el formulario para un nuevo egreso
    NuevoEgreso
Else
    
    ' Deshabilita el movimiento
    txtCodMov.Enabled = False
    cboCodMov.Enabled = False
    
    ' Inicializa la variable
    mbEgresoCargado = False
    
    'Prepara el formulario para modificar el egreso
    ModificarEgreso
End If

End Sub

Private Sub ModificarEgreso()
 '---------------------------------------------------------------
'Propósito : Prepara el formulario para modificar un egreso
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------

' Deshabilita optCajaBancos
  fraCB.Enabled = False

' Habilita el txtCodEgreso
  txtCodEgreso = Empty
  txtCodEgreso.Enabled = True
  txtCodEgreso.BackColor = Obligatorio

'Se carga la colección de E Rendir
  Var5

' Inicializa los optCB
  optCaja.Value = False
  optBanco.Value = False
  optRendir.Value = False
  
' Deshabilita controles del formulario
  DeshabilitarHabilitarFormulario False
  
' Oculta controles
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 lblNumCh.Visible = False: txtNumCh.Visible = False

 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False

' Limpia las colecciones
   Set gcolDetMovCB = Nothing
   msAfecta = Empty
   msProceso = Empty
   msCodAfectaAnterior = Empty
   msCaBaAnt = Empty
   gsOrden = Empty
   gdblMontoAnterior = 0
   msCajaoBanco = Empty

' Inicializa la variable codigo de cuenta
  msCtaCte = Empty

' Muestra los resumen
  txtSaldoCB = "0.00"
  
' Maneja estado de los botones del formulario
  HabilitaDeshabilitaBotones "Modificar"

End Sub

Private Sub NuevoEgreso()
'---------------------------------------------------------------
'Propósito : Prepara el formulario para un egreso con afectación
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------

' Inicializa las variable de modulo
  msCtaCte = Empty
  msAfecta = Empty
  msProceso = Empty
  msCajaoBanco = Empty
  
' Muestra los resumen
  txtSaldoCB = "0.00"

'Se carga la colección de E Rendir
  Var5

' Pone por defecto el egreso de caja y calcula el Orden
If optCaja.Value = True Then
    ' realiza el evento de optclick
    optCaja_Click
Else
    ' cambia el valor del optCaja.value
    optCaja.Value = True
End If

' Coloca la fecha del sistema
mskFecTrab.Text = gsFecTrabajo

' Limpia las colecciones
Set gcolDetMovCB = Nothing
   
' deshabilita los botones del formulario
HabilitaDeshabilitaBotones ("Nuevo")

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
    cmdBuscarEgreso.Enabled = True
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
End Select

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

'la sentencia para cargar el combo y las colecciones de tipo movimiento
sSQL = ""
sSQL = "SELECT IdConCB, DescConCB, CodCont, Afecta FROM Tipo_MovCB " & _
          " WHERE RelacProc = 'SS' " & _
          " ORDER BY DescConCB"
          
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

Private Sub CargarSaldo()
'-------------------------------------------------------------------------
'Propósito: Cargar el saldo de Caja o Bancos hasta el momento dehacer la operación.
'Recibe Nada
'Devuelve Nada
'-------------------------------------------------------------------------
Dim curSaldo As New clsBD2
Dim sSQL As String
Dim dblSaldo As Double
Dim TotalIngreso As Double
Dim curEmpresas As New clsBD2
Dim EmpresasExistentes As String
Dim InstrucEmpresas As String
Dim TotalEgresoProyectos As Double
Dim TotalEgresoEmpresasSinRH As Double
Dim TotalEgresoEmpresasSoloRHCB As Double
Dim TotalEgresos As Double

On Error GoTo mnjError

'*-*-*-*-*
'*-*-*-*-*  EMPRESAS EXISTENTES
'*-*-*-*-*
sSQL = "SELECT IdProy " _
     & " FROM PROYECTOS " _
     & " WHERE (PROYECTOS.Tipo = 'EMPR') ORDER BY IdProy "

' Ejecuta la sentencia
curEmpresas.SQL = sSQL
If curEmpresas.Abrir = HAY_ERROR Then End

EmpresasExistentes = ""
If Not curEmpresas.EOF Then
  Do While Not curEmpresas.EOF
    EmpresasExistentes = EmpresasExistentes & curEmpresas.campo(0) & "@"
    
    curEmpresas.MoverSiguiente
  Loop
End If

InstrucEmpresas = ""
Do While InStr(1, EmpresasExistentes, "@")
  InstrucEmpresas = InstrucEmpresas & "IDPROY <> '" & Left(EmpresasExistentes, 2) & "' AND "
  EmpresasExistentes = Mid(EmpresasExistentes, 4, Len(EmpresasExistentes))
Loop

InstrucEmpresas = Left(InstrucEmpresas, Len(InstrucEmpresas) - 4)

curEmpresas.Cerrar

dblSaldo = 0 'Inicializa variable saldo
If optBanco.Value = True Then
  If msCtaCte <> Empty Then ' Si se eligió la CtaCte
    'Averigua el ingreso de la Cta
    sSQL = "SELECT SUM(Monto) as Ingreso FROM INGRESOS " _
          & "WHERE IdCta='" & msCtaCte & "' and Anulado='NO'"
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If Not IsNull(curSaldo.campo(0)) Then TotalIngreso = curSaldo.campo(0)
    curSaldo.Cerrar
    
    'Averigua el egreso de la Cta
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL DE EGRESOS PARA PROYECTOS CON AFECTACION Y SIN AFECTACION
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) as Egreso FROM EGRESOS " _
          & "WHERE IdCta='" & msCtaCte & "' and Anulado='NO' and Origen='B' AND " & InstrucEmpresas
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoProyectos = 0
    Else
        TotalEgresoProyectos = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    'Averigua el egreso de la Cta
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SIN RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoAfectado) as Egreso FROM EGRESOS, PROYECTOS " _
          & "WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.IdCta='" & msCtaCte & "' and Anulado='NO' and Origen='B' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc<>'02') "
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoEmpresasSinRH = 0
    Else
        TotalEgresoEmpresasSinRH = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    'Averigua el egreso de la Cta
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SOLO RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) as Egreso FROM EGRESOS, PROYECTOS " _
          & "WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.IdCta='" & msCtaCte & "' and Anulado='NO' and Origen='B' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc='02') "
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoEmpresasSoloRHCB = 0
    Else
        TotalEgresoEmpresasSoloRHCB = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    TotalEgresos = TotalEgresoProyectos + TotalEgresoEmpresasSinRH + TotalEgresoEmpresasSoloRHCB
    
    dblSaldo = TotalIngreso - TotalEgresos
  End If
    ' Muestra el saldo de la Ctacte o 0.00 si todavía no se eligió
    txtSaldoCB = Format(dblSaldo, "###,###,##0.00")
    lblSaldoCB.Caption = "Saldo de Cuenta:"

ElseIf optCaja.Value = True Then
    
    'Averigua el ingreso de Caja
    sSQL = "SELECT SUM(Monto) as Ingreso FROM INGRESOS " _
          & "WHERE IdCta='' and Anulado='NO'"
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If Not IsNull(curSaldo.campo(0)) Then TotalIngreso = curSaldo.campo(0)
    curSaldo.Cerrar
    
    'Averigua el egreso de Caja
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL DE EGRESOS PARA PROYECTOS CON AFECTACION Y SIN AFECTACION
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) as Egreso FROM EGRESOS " _
          & "WHERE IdCta='' and Anulado='NO' and Origen='C' AND " & InstrucEmpresas
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoProyectos = 0
    Else
        TotalEgresoProyectos = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    'Averigua el egreso de Caja
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SIN RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoAfectado) as Egreso FROM EGRESOS, PROYECTOS " _
          & "WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.IdCta='' and Anulado='NO' and Origen='C' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc<>'02') "
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoEmpresasSinRH = 0
    Else
        TotalEgresoEmpresasSinRH = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    'Averigua el egreso de Caja
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SOLO RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) as Egreso FROM EGRESOS, PROYECTOS " _
          & "WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.IdCta='' and Anulado='NO' and Origen='C' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc='02') "
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        TotalEgresoEmpresasSoloRHCB = 0
    Else
        TotalEgresoEmpresasSoloRHCB = curSaldo.campo(0)
    End If
    curSaldo.Cerrar
    
    TotalEgresos = TotalEgresoProyectos + TotalEgresoEmpresasSinRH + TotalEgresoEmpresasSoloRHCB
    
    dblSaldo = TotalIngreso - TotalEgresos
    
    txtSaldoCB = Format(dblSaldo, "###,###,##0.00")
    lblSaldoCB.Caption = "Saldo de Caja:"
    
ElseIf optRendir.Value = True Then
    ' Verifica si el código esta lleno
    If txtRinde <> Empty Then
        dblSaldo = Val(Var30(gcolTRendir.Item(txtRinde), 3))
    End If
PostErrClaveCol:
    txtSaldoCB = Format(dblSaldo, "###,###,##0.00")
    lblSaldoCB.Caption = "Saldo a Rendir:"
End If

Exit Sub
' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        ' Sigue la ejecución
        Resume PostErrClaveCol:
    End If
End Sub

Private Sub EstableceCamposObligatorios()

' Establece los campos obligatorios
 txtDocEgreso.BackColor = Obligatorio
 txtTipDoc.BackColor = Obligatorio
 txtMonto.BackColor = Obligatorio
 txtCodMov.BackColor = Obligatorio
 txtAfecta.BackColor = Obligatorio
 txtCodContable.BackColor = Obligatorio
 txtBanco.BackColor = Obligatorio
 cboCtaCte.BackColor = Obligatorio
 txtNumCh.BackColor = Obligatorio
 txtRinde.BackColor = Obligatorio
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
         cmdCancelar.Enabled = False
         
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

Private Sub cboCodMov_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCodMov) = True Then SendKeys "{down}"

End Sub

Private Sub cbocodMov_Click()

' Habilita txtAfecta
txtAfecta.Enabled = True

' Verifica si el evento ha sido activado por el teclado o Mouse
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

Set mcolCodPersonalPrest = Nothing
Set mcolDesCodPersonalPrest = Nothing

Set mcolCodPlanCont = Nothing
Set mcolDesCodPlanCont = Nothing

Set gcolDetMovCB = Nothing
Set gcolTabla = Nothing

' Verifica si esta habilitado controles de egreso caja bancos
If gsTipoOperacionEgreso = "Modificar" And mbEgresoCargado = True Then
     mcurRegEgresoCajaBanco.Cerrar ' Cierra el cursor del ingreso
End If
End Sub

Private Sub Image1_Click()
'Carga la Var48
Var48
End Sub

Private Sub optBanco_Click()

' Realiza el cambio de opción a Bancos
 CambiaroptCajaBancos
    
'  Habilita el botón aceptar
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

Private Sub CambiaroptCajaBancos()
'-------------------------------------------------------------------
'Propósito : Establece los controles de la primera parte del formulario _
             cuando se cambia de optCaja, optBancos, optRendir
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If gsTipoOperacionEgreso = "Nuevo" Then
   If optCaja.Value = True Then
    'Calcula el sigiente orden de Caja y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("CA")
    msCajaoBanco = "CA"
   ElseIf optBanco.Value = True Then
    'Calcula el sigiente orden de Banco y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("BA")
    msCajaoBanco = "BA"
   ElseIf optRendir.Value = True Then
    'Calcula el sigiente orden de Caja y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("CA")
    msCajaoBanco = "ER"
   End If
Else
   If optCaja.Value = True Then
    'Cambia de origen
    msCajaoBanco = "CA"
   ElseIf optBanco.Value = True Then
    'Cambia de origen
    msCajaoBanco = "BA"
   ElseIf optRendir.Value = True Then
    'Cambia de origen
    msCajaoBanco = "ER"
   End If
End If

' Verifica el movimiento
If VerificaMovimiento = False Then
    ' Sale
    Exit Sub
End If
  
' Maneja los controles de Banco y Rendir
ManejaControlesBanco

' Cargar Saldo
CargarSaldo

End Sub

Private Function VerificaMovimiento() As Boolean
'--------------------------------------------------------
' Propósito: Verifica el movimiento con las opciones de origen del egreso
' Recibe: Nada
' Entrega: Nada
'--------------------------------------------------------
If msAfecta = "Proceso" Then
    If msProceso = "ENTREGA_RENDIR" Then
       ' Verifica que el movimiento sea de caja
       If optCaja.Value = False Then
            ' Movimiento no valido
            MsgBox "El Movimiento elegido solo es de Caja", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
            optCaja.Value = True
            ' El movimiento solo es de Caja
            VerificaMovimiento = False
            Exit Function
       End If
    ElseIf msProceso = "PAGO_PLANILLAS" Then
       ' Verifica que el movimiento sea de caja
       If optCaja.Value = False And optBanco.Value = False Then
            ' Movimiento no valido
            MsgBox "El Movimiento elegido solo es de Caja o Bancos", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
            optCaja.Value = True
            ' El movimiento solo es de Caja
            VerificaMovimiento = False
            Exit Function
       End If
    End If
End If
' Todo Ok
VerificaMovimiento = True

End Function

Private Sub ManejaControlesBanco()
'-------------------------------------------------------------------
'Propósito : Establece los controles de banco y rendir cuando se cambia de optCaja a optBancos bis
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If optCaja.Value = True Then
 'Limpia y oculta los controles de Banco y Rendir
 txtBanco.Text = Empty
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 txtNumCh.Text = Empty: lblNumCh.Visible = False: txtNumCh.Visible = False
 msCtaCte = Empty
 txtRinde.Text = Empty
 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False
ElseIf optBanco.Value = True Then
 'Muestra los controles de banco y oculta el resto
 lblBanco.Visible = True: txtBanco.Visible = True: cboBanco.Visible = True: cmdPBanco.Visible = True
 lblCtaCte.Visible = True: cboCtaCte.Visible = True: cmdPCtaCte.Visible = True
 lblNumCh.Visible = True: txtNumCh.Visible = True
 txtRinde.Text = Empty
 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False

ElseIf optRendir.Value = True Then
 'Muestra los controles a rendir y oculta el resto
 lblRinde.Visible = True: txtRinde.Visible = True: txtDescRinde.Visible = True: cmdBuscaRinde.Visible = True
 txtBanco.Text = Empty
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 txtNumCh.Text = Empty: lblNumCh.Visible = False: txtNumCh.Visible = False
 msCtaCte = Empty
End If

End Sub


Private Sub optCaja_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub optRendir_Click()

' Realiza el cambio de opción a Bancos
 CambiaroptCajaBancos
    
'  Habilita el botón aceptar
 HabilitarBotonAceptar
 
End Sub

Private Sub optRendir_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub txtAfecta_Change()

' Verifica si no es pago de planillas
If msProceso <> "PAGO_PLANILLAS" Then
    ' Verifica si el tamaño del txt es Igual al tamaño definido
    If Len(txtAfecta) = txtAfecta.MaxLength Then
        ' Actualiza el txtDesc
        ActualizaDesc
    Else
        ' Limpia el txtDescAfecta
        txtDesc.Text = Empty
    End If
End If

' Verifica SI el campo esta vacio
If txtAfecta.Text <> "" And txtDesc.Text <> "" Then
   ' Los campos coloca a color blanco
   txtAfecta.BackColor = vbWhite
Else
  ' Marca los campos obligatorios
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
        MsgBox "El código ingresado no existe ", , "SGCcaijo-Ingresos"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub ActualizaDescRendir()
'--------------------------------------------------------------
'PROPÓSITO  : Actualiza la descripcion de la cuenta a rendir
'Recive     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------
On Error GoTo mnjError
'Copia la descripción
txtDescRinde.Text = Var30(gcolTRendir.Item(txtRinde), 2)
Exit Sub
' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe ", , "SGCcaijo-Verifica Datos"
        'Limpia la descripción
        txtDescRinde.Text = Empty
    End If
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
      
   ' Actualiza el txtSaldo de la Cta
   txtSaldoCB.Text = "0.00"

   'Actualiza el cboCtaCte con las descripciones de las cuentas relacionadas a txtBanco
    ActualizarListcboCtaCte
    
Else
  ' Actualiza el txtSaldo de la Cta
   txtSaldoCB.Text = "0.00"

   'Marca los campos obligatorios, y limpia el combo
   txtBanco.BackColor = Obligatorio
   'vacia el cboCtaCte
   cboCtaCte.Clear
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

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
       Or txtDocEgreso.BackColor <> vbWhite _
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
   If optBanco.Value = True And (txtBanco.BackColor <> vbWhite Or cboCtaCte.BackColor <> vbWhite _
      Or txtNumCh.BackColor <> vbWhite) Then
     ' Algún obligatorio de banco falta ser introducido
     ' Deshabilita el botón
     cmdAceptar.Enabled = False
     Exit Sub
   ElseIf optRendir.Value = True And (txtRinde.BackColor <> vbWhite) Then
    ' Algún obligatorio falta rendir
     cmdAceptar.Enabled = False
     Exit Sub
   End If
End If

' Verifica si se cambio algún dato
If gsTipoOperacionEgreso = "Modificar" Then
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
' Verifica si se cambio el origen
If msCajaoBanco <> msCaBaAnt Then
   ' Cambio el origen de los datos
    fbCambioDatos = True
    Exit Function
End If
' verifica los datos de caja
If txtCodMov.Text <> mcurRegEgresoCajaBanco.campo(4) _
   Or txtDocEgreso.Text <> mcurRegEgresoCajaBanco.campo(0) _
   Or txtTipDoc.Text <> mcurRegEgresoCajaBanco.campo(1) _
   Or txtAfecta.Text <> msCodAfectaAnterior _
   Or txtCodContable.Text <> mcurRegEgresoCajaBanco.campo(5) _
   Or Val(Var37(txtMonto.Text)) <> mcurRegEgresoCajaBanco.campo(3) _
   Or txtObserv.Text <> mcurRegEgresoCajaBanco.campo(7) Then
    ' cambio datos generales
    fbCambioDatos = True
    Exit Function
End If
' Verifica si se cambio Cta corriente y número de cheque
If msCajaoBanco = "BA" Then
    If msCtaCte <> mcurRegEgresoCajaBanco.campo(9) _
    Or txtNumCh <> mcurRegEgresoCajaBanco.campo(10) Then
        ' Cambió cta corriente
            ' cambio datos generales
            fbCambioDatos = True
            Exit Function
    End If
End If
' Verifica si se cambio de cuenta a rendir
If msCajaoBanco = "ER" Then
    If txtRinde <> msRindeAnt Then
        ' cambio datos generales
        fbCambioDatos = True
        Exit Function
    End If
End If

End Function


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
  ' Carga la Sentencia para obtener las Ctas en dólares que pertenecen al txtBanco
  If gsTipoOperacionEgreso = "Nuevo" Then
    sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
       " WHERE c.idbanco = '" & txtBanco & "' " & _
       " AND  b.idBanco = c.IdBanco" & _
       " AND c.idmoneda= 'SOL' AND C.ANULADO='NO' ORDER BY c.DescCta"
  Else
    sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
       " WHERE c.idbanco = '" & txtBanco & "' " & _
       " AND  b.idBanco = c.IdBanco" & _
       " AND c.idmoneda= 'SOL' ORDER BY c.DescCta"
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

Private Function fbCargarEgreso() As Boolean
'----------------------------------------------------------------------------
'Propósito: Cargar el registro de Egreso de acuerdo al Código en la caja de texto
'Recibe: Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota el codigo de Registro de ingreso es CA o BA AAMM9999
Dim sSQL As String

' Verifica si el egreso es a Caja o Bancos
msCajaoBanco = Left(txtCodEgreso.Text, 2)

' Carga la sentencia que consulta a la BD acerca del registo de Egreso en Caja o Bancos
If msCajaoBanco = "CA" Then  'Consulta a Caja

  sSQL = "SELECT E.NumDoc,E.IdTipoDoc,E.FecMov,E.MontoCB,E.CodMov,E.CodContable," & _
         "TM.Afecta,E.Observ,E.Origen " & _
         "FROM EGRESOS E , Tipo_MovCB TM   WHERE " & _
         "E.Orden=" & "'" & Trim(txtCodEgreso.Text) & "' and E.Anulado='NO' and " & _
         "E.CodMov=TM.IdConCB and E.IdProy=''"
           
ElseIf msCajaoBanco = "BA" Then 'Consulta a Bancos
    ' Establece el origen
     msCaBaAnt = "BA"
    sSQL = "SELECT E.NumDoc,E.IdTipoDoc,E.FecMov," & _
           "E.MontoCB,E.CodMov,E.CodContable,TM.Afecta," & _
           "E.Observ,CTA.IdBanco,E.IdCta, E.NumCheque " & _
           "FROM EGRESOS E, Tipo_MovCB TM, TIPO_CUENTASBANC CTA WHERE " & _
           "E.Orden=" & "'" & Trim(txtCodEgreso.Text) & "' and E.Anulado='NO' " & _
           "and E.codMov=TM.IdconCB and E.IdCta=CTA.IdCta and E.IdProy=''"
               
Else 'Mensaje Cod Registro Ingreso  NO Valido
        MsgBox "El código de egreso NO válido, debe ser CA o BA AAMM9999", _
               vbExclamation, "Caja y Banco- Egreso sin afectación a financiera"
        fbCargarEgreso = False
        Exit Function
End If

mcurRegEgresoCajaBanco.SQL = sSQL
' Abre el cursor SI hay  error sale indicando la causa del error
If mcurRegEgresoCajaBanco.Abrir = HAY_ERROR Then
    End
End If
' Cursor abierto
mbEgresoCargado = True

'Verifica la existencia del registro de egreso
If mcurRegEgresoCajaBanco.EOF Then
    'Mensaje de registro de egreso a Caja o Bancos no existe
    MsgBox "El Código de Egreso que se digito no está registrado como Egreso sin afectación fianciera", _
         vbExclamation, "Caja-Banco- Egreso Sin afectación financiera"
 
    ' Cursor cerrado
    mcurRegEgresoCajaBanco.Cerrar
    mbEgresoCargado = False
    
    ' No se pudo cargar
    fbCargarEgreso = False
    
Else

 'Carga los controles con datos del egreso y Habilita los controles
  CargarControlesEgreso
 
 'Carga el saldo de Caja o Banco
 CargarSaldo
 
 ' todo ok
 fbCargarEgreso = True
    
End If


End Function

Private Sub CargarControlesEgreso()
'----------------------------------------------------------------------------
'Propósito: Cargar los controles refentes al Egreso que se desea modificar
'Recibe:  Nada
'Devuelve: Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el procedimiento Cargar Registro de Egreso

'Deshabilita CodEgreso
txtCodEgreso.BackColor = vbWhite
txtCodEgreso.Enabled = False

' Habilita el formulario
DeshabilitarHabilitarFormulario True

' Carga el optCajaBanco
If msCajaoBanco = "CA" Then
    ' Verifica si la operación es de caja o de cuentas a Rendir
    If mcurRegEgresoCajaBanco.campo(8) = "C" Then
        ' Establece el origen
        msCaBaAnt = "CA"
        ' Habilita la opción de Caja
        optCaja.Value = True
        ' Actualiza el origen
        msCajaoBanco = "CA"
    ElseIf mcurRegEgresoCajaBanco.campo(8) = "R" Then
        ' Establece el origen
        msCaBaAnt = "ER"
        ' Habilita la opción de Rendir
        optRendir.Value = True
        ' Actualiza el origen
        msCajaoBanco = "ER"
        ' Averigua la cuenta a rendir
        CargarCuentaRendir
    End If
Else ' Habilita la opción banco
    optBanco.Value = True
End If

'Rellena los controles de Caja
txtDocEgreso.Text = mcurRegEgresoCajaBanco.campo(0)
txtTipDoc.Text = mcurRegEgresoCajaBanco.campo(1)
mskFecTrab.Text = FechaDMA(Trim(Str(mcurRegEgresoCajaBanco.campo(2))))
txtMonto.Text = Format(mcurRegEgresoCajaBanco.campo(3), "###,###,##0.00")
txtCodMov.Text = mcurRegEgresoCajaBanco.campo(4)
txtObserv.Text = mcurRegEgresoCajaBanco.campo(7)
txtCodContable.Text = mcurRegEgresoCajaBanco.campo(5)

'Se carga cbo Afecta con el dato del cursor(Terceros o Pln_Personals)
CargarRegAfecta

If msCajaoBanco = "BA" Then 'Rellena los controles de Banco
    txtBanco.Text = mcurRegEgresoCajaBanco.campo(8)
    msCtaCte = mcurRegEgresoCajaBanco.campo(9) 'Actualiza variable de Código de CtaCte
    CD_ActVarCbo cboCtaCte, msCtaCte, mcolCodDesCtaCte
    txtNumCh.Text = mcurRegEgresoCajaBanco.campo(10)
'    'halbilita controles de ingreso a bancos
'    DeshabilitaHabilitaControlesBanco
End If

' Maneja los estados de las opciones
Manejaopciones

'Habilita Botones cancelar,Anular Caja o Bancos
cmdAnular.Enabled = True
cmdCancelar.Enabled = True

End Sub

Private Sub Manejaopciones()
' --------------------------------------------------------------------
' Propósito: Habilita o deshabilita los controles opt
'            de acuerdo a la opción caja o bancos o rendir
' --------------------------------------------------------------------
If msCajaoBanco = "CA" Or msCajaoBanco = "ER" Then
    ' Habilita optcaja y optrendir
    optCaja.Enabled = True: optBanco.Enabled = False: optRendir.Enabled = True
ElseIf msCajaoBanco = "BA" Then
    ' No habilita las opciones
    optCaja.Enabled = False: optBanco.Enabled = False: optRendir.Enabled = False
End If

End Sub

Private Sub CargarCuentaRendir()
'---------------------------------------------------------------------
'Propósito: Carga la cuenta a rendir
'Recibe : Nada
'Devuelve : Nada
'---------------------------------------------------------------------
Dim sSQL As String
Dim curRendir As New clsBD2
' Inicia
msRindeAnt = Empty
' Carga la consulta
sSQL = "SELECT R.IdPersona FROM MOV_ENTREG_RENDIR R " _
    & "WHERE R.Orden='" & txtCodEgreso & "'"
' Ejecuta la sentencia
curRendir.SQL = sSQL
If curRendir.Abrir = HAY_ERROR Then End
' Asigna la cuenta a la variable de modulo
msRindeAnt = curRendir.campo(0)
txtRinde = msRindeAnt

End Sub

Private Sub CargarRegAfecta()
'---------------------------------------------------------------------
'Propósito: Carga el combo afecta
'Recibe : Nada
'Devuelve : Nada
'---------------------------------------------------------------------
'Nota
Dim sSQL As String
Dim curAfecta As New clsBD2

'Verifica a quien afecta el concepto(Tipo Mov)
If mcurRegEgresoCajaBanco.campo(6) = "Tercero" Then
 'Si el registro de egreso ezta realacionado a terceros
     sSQL = "SELECT IdTercero FROM MOV_TERCEROS WHERE " _
         & "Orden='" & txtCodEgreso & "'"
 'Ejecuta la sentencia de consulta
  curAfecta.SQL = sSQL
  If curAfecta.Abrir = HAY_ERROR Then End 'error se cierra la aplicacion
 'Actualiza el combo afecta
  msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
  txtAfecta.MaxLength = Len(msCodAfectaAnterior)
  txtAfecta.Text = msCodAfectaAnterior
 'Cierra la consulta
  curAfecta.Cerrar

ElseIf mcurRegEgresoCajaBanco.campo(6) = "Persona" Then
'Si el registro de egreso esta relacionado a personal
    sSQL = "SELECT IdPersona FROM MOV_PERSONAL WHERE " _
        & "Orden='" & txtCodEgreso & "'"
'Ejecuta la sentencia de consulta
  curAfecta.SQL = sSQL
  If curAfecta.Abrir = HAY_ERROR Then End 'error se cierra la aplicacion
 'Actualiza el combo afecta
  msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
  txtAfecta.MaxLength = Len(msCodAfectaAnterior)
  txtAfecta.Text = msCodAfectaAnterior
 'Cierra la consulta
  curAfecta.Cerrar
    
ElseIf mcurRegEgresoCajaBanco.campo(6) = "Proceso" Then
' Verifica que el origen sea de caja
    If msCajaoBanco = "CA" Then
    ' Verifica si es una entrega a rendir
        sSQL = "SELECT P.IdPersona, ( P.Apellidos & ', ' & P.Nombre), ME.Ingreso " _
            & "FROM PLN_PERSONAL P, MOV_ENTREG_RENDIR ME WHERE " _
            & "ME.Orden='" & txtCodEgreso & "' and ME.IdPersona=P.IdPersona"
        ' Ejecuta la sentencia
        curAfecta.SQL = sSQL
        If curAfecta.Abrir = HAY_ERROR Then End
        If Not curAfecta.EOF Then ' EL egreso esta relacionado a entregas a rendir
           ' Carga Afecta Persona
              msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
              gsCodAfectaAnterior = msCodAfectaAnterior
              gdblMontoAnterior = Val(curAfecta.campo(2))
              txtAfecta.MaxLength = Len(msCodAfectaAnterior)
              txtDesc.Text = curAfecta.campo(1)
              txtAfecta.Text = msCodAfectaAnterior
              msProceso = "ENTREGA_RENDIR"
           ' Carga la coleccion global del detalle del movimieto
            gcolDetMovCB.Add Item:=msCodAfectaAnterior & "¯" _
                           & Conta25 & "¯" _
                           & Format(curAfecta.campo(2), "########0.00"), _
                        Key:=curAfecta.campo(0) & "¯" _
                           & curAfecta.campo(2)
            ' Sale de el procedimiento
            curAfecta.Cerrar
            Exit Sub
         End If
        ' Cierra la componente
        curAfecta.Cerrar
    End If
    
   ' Verifica si se pagó algún prestamo
    sSQL = "SELECT PP.IdPersona, ( PR.Apellidos & ', ' & PR.Nombre), PP.IdConPL, PP.NumPrestamo, P.Monto, PC.CodContable " _
        & "FROM PAGO_PRESTAMOS PP , PRESTAMOS P, PLNCONCEPTOS_OTROS PC, PLN_PERSONAL PR WHERE " _
        & "PP.Orden='" & txtCodEgreso & "' and PP.IdPersona=P.IdPersona and " _
        & "PP.IdConPL=P.IdConPL and PP.NumPrestamo=P.NumPrestamo and PP.IdConPL=PC.IdConPL AND " _
        & "PP.IdPersona=PR.IdPersona"
    ' Ejecuta la sentencia
    curAfecta.SQL = sSQL
    If curAfecta.Abrir = HAY_ERROR Then End
    If Not curAfecta.EOF Then ' EL egreso esta relacionado a pago de prestamos
       ' Carga Afecta Persona
          msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
          txtAfecta.MaxLength = Len(msCodAfectaAnterior)
          txtDesc.Text = curAfecta.campo(1)
          txtAfecta.Text = msCodAfectaAnterior
          msProceso = "PAGO_PRESTAMOS"
       ' Carga la coleccion global del detalle del movimieto
        gcolDetMovCB.Add Item:=curAfecta.campo(2) & "¯" _
                       & curAfecta.campo(3) & "¯" _
                       & curAfecta.campo(5) & "¯" _
                       & Format(curAfecta.campo(4), "##0.00"), _
                    Key:=curAfecta.campo(2) & "¯" _
                       & curAfecta.campo(3)
        ' Sale de el procedimiento
        curAfecta.Cerrar
        Exit Sub
     End If
     
    ' Cierra la componente
    curAfecta.Cerrar
    
' Verifica si se pagó algún adelanto
    sSQL = "SELECT AP.IdPersona, ( PR.Apellidos & ', ' & PR.Nombre), AP.IdConPL, AP.Monto, PC.CodContable " _
        & "FROM ADELANTOS AP , PLNCONCEPTOS_OTROS PC, PLN_PERSONAL PR WHERE " _
        & "AP.Orden='" & txtCodEgreso & "' and AP.IdConPL=PC.IdConPL AND " _
        & "AP.IdPersona=PR.IdPersona"
    ' Ejecuta la sentencia
    curAfecta.SQL = sSQL
    If curAfecta.Abrir = HAY_ERROR Then End
    If Not curAfecta.EOF Then ' EL egreso esta relacionado a pago de Adelantos
       ' Carga Afecta Persona
          msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
          txtAfecta.MaxLength = Len(msCodAfectaAnterior)
          txtDesc.Text = curAfecta.campo(1)
          txtAfecta.Text = msCodAfectaAnterior
          msProceso = "PAGO_ADELANTOS"
       ' Carga la coleccion global del detalle del movimieto
        gcolDetMovCB.Add Item:=curAfecta.campo(2) & "¯" _
                       & curAfecta.campo(4) & "¯" _
                       & Format(curAfecta.campo(3), "##0.00"), _
                    Key:=curAfecta.campo(2)
        ' Sale de el procedimiento
        curAfecta.Cerrar
        Exit Sub
     End If
     
    ' Cierra la componente
    curAfecta.Cerrar


' Verifica si se pagó alguna planilla
    sSQL = "SELECT PPL.CodPlanilla, PL.DescPlanilla, PPL.Monto, PPL.CodContable " _
        & "FROM PAGO_PLANILLAS PPL , PLN_PLANILLAS PL WHERE " _
        & "PPL.Orden='" & txtCodEgreso & "' and PPL.CodPlanilla=PL.CodPlanilla"
    ' Ejecuta la sentencia
    curAfecta.SQL = sSQL
    If curAfecta.Abrir = HAY_ERROR Then End
    If curAfecta.EOF Then
    Else
       ' Carga afecta a planillas
        msCodAfectaAnterior = curAfecta.campo(0) 'Actualiza Código Afecta(Terc o Pers)
        txtAfecta.MaxLength = Len(msCodAfectaAnterior)
        txtDesc = curAfecta.campo(1)
        txtAfecta.Text = msCodAfectaAnterior
        msProceso = "PAGO_PLANILLAS"
        Do While Not curAfecta.EOF   ' EL egreso esta relacionado a pago de prestamos
          
        ' Carga la coleccion global del detalle del movimieto
         gcolDetMovCB.Add Item:=curAfecta.campo(0) & "¯" _
                        & curAfecta.campo(3) & "¯" _
                        & Format(curAfecta.campo(2), "##0.00"), _
                     Key:=curAfecta.campo(0) & "¯" _
                        & curAfecta.campo(3)
        ' Siguiente elemento
         curAfecta.MoverSiguiente
        Loop
           
        ' Cierra la componente
        curAfecta.Cerrar
        Exit Sub
       
    End If
    ' Cierra la componente
    curAfecta.Cerrar
    
    ' Mensaje de error
    MsgBox "No se guardaron todos los datos de este egreso" & Chr(13) _
        & "Debe anular este egreso", , "SGCcaijo-Egreso sin afectación"

ElseIf mcurRegEgresoCajaBanco.campo(6) = "Ninguno" Then
' El registro de egreso no esta relacionado  a niguno
' Se inhabilitan los controles de Afecta(Terceros o Personal)
    txtAfecta.Text = Empty
    txtAfecta.Enabled = False
    cmdBuscar.Enabled = False
    lblEtiqueta.Caption = Empty
    msCodAfectaAnterior = Empty 'variable que guarda el Código de (Terce o Pers) al Modificar
    Exit Sub 'Sale del procedimiento
End If

End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If

End Sub

Private Sub txtCodContable_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub txtCodEgreso_Change()

' Verifica el proceso que se realiza en el formulario
If gsTipoOperacionEgreso = "Modificar" Then
    
    ' Verifica si se ha introducido el tamaño de el código
      If Len(txtCodEgreso) = txtCodEgreso.MaxLength Then
      
        ' Verifica mayúsculas
        If UCase(txtCodEgreso) = txtCodEgreso Then
            
          ' Verifica si el egreso existe y es sin afectación
          If fbCargarEgreso = True Then
             ' Sale y deshabilita el control
             SendKeys vbTab
             ' Deshabilita el txtcod egreso y el botón buscar, habilita anular
             txtCodEgreso.Enabled = False
             cmdBuscarEgreso.Enabled = False
             cmdAnular.Enabled = True
          End If ' fin de cargar egreso
          
        Else ' vuelve a mayúsulas el txtcodegreso
            txtCodEgreso = UCase(txtCodEgreso)
        End If ' fin verificar mayúsculas
        
      End If ' fin de verofocr el tamalo del texto
 End If
 
 ' Maneja el color del control txtcodegreso
 If txtCodEgreso = Empty Then
    ' coloca el color obligatorio al control
    txtCodEgreso.BackColor = Obligatorio
 Else
    ' coloca el color de edición
    txtCodEgreso.BackColor = vbWhite
 End If

End Sub

Private Sub txtCodEgreso_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub txtCodMov_Change()

' Si procede, se actualiza descripción correspondiente a código introducido
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
 
' Verifica si el campo esta vacio
If txtCodMov.Text <> Empty And cboCodMov.Text <> Empty Then
   ' Los campos coloca a color blanco
   txtCodMov.BackColor = vbWhite
   
   ' Carga el combo Afecta dependiendo del codigo de afecta
   msAfecta = DeterminarAfecta(txtCodMov.Text)
   CargarCboAfecta msAfecta
           
   ' Carga el cboCtaContable dependiendo del tipo de movimiento
   CargacboCtaContable DeterminarCodCont(txtCodMov.Text)
   
   ' Maneja estado de los controles dependiendo de msafecta
   EstableceEstadoAfectaMonto
   
   ' Maneja estado de las opciones Caja-Bancos.(De acuerdo al movimiento)
   EstableceEstadoOpcionesCB
        
   ' Si el combo sólo tiene un elemento, se muestra en pantalla
   MostrarUnicoItem
   
Else
  'Marca los campos obligatorios
   txtCodMov.BackColor = Obligatorio
End If

' Habilita botón aceptar
HabilitarBotonAceptar

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
If gsTipoOperacionEgreso = "Nuevo" Then
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



Private Sub EstableceEstadoAfectaMonto()
'----------------------------------------------------------------
' Propósito: Establece el estado de los controles Afecta y el monto _
             de egreso, tambien establece el estado de optcaja y optbanco
' Recibe: Nada
' Entrega: Nada
'----------------------------------------------------------------
If msAfecta = "Proceso" Then
   ' Verifica si el proceso es planillas
   If msProceso = "PAGO_PLANILLAS" Then
        ' Inhabilita caja-bancos
        fraCB.Enabled = False
   End If
   ' Inhabilita afecta y monto
   txtAfecta.Enabled = False: cmdBuscar.Enabled = False
   ' Verifica si el proceso es Pago_adelantos
   If msProceso = "PAGO_ADELANTOS" And gsTipoOperacionEgreso = "Modificar" Then
    ' Habilita el monto
    txtMonto.Enabled = True
   ElseIf msProceso = "ENTREGA_RENDIR" Then
    ' No puede cambiar de opcion origen
    If gsTipoOperacionEgreso = "Modificar" Then fraCB.Enabled = False
    ' Habilita cmdBuscar
    cmdBuscar.Enabled = True
    ' Inhabilita el monto
    txtMonto.Enabled = False
   Else
    ' Inhabilita el monto
    txtMonto.Enabled = False
   End If
Else
   fraCB.Enabled = True
   ' Habilita afecta y monto y las opt's caja-bancos
   txtAfecta.Enabled = True: cmdBuscar.Enabled = True
   txtMonto.Enabled = True
End If

End Sub

Private Sub EstableceEstadoOpcionesCB()
'----------------------------------------------------------------
' Propósito: Establece el estado de las opciones Caja-Bancos _
             del egreso
' Recibe: Nada
' Entrega: Nada
'----------------------------------------------------------------

' Verifica si el tipo de operación de egreso es nuevo
If gsTipoOperacionEgreso = "Nuevo" Then
    ' Habilita las opciones CB si es cualquier movimiento
    fraCB.Enabled = True
    If msAfecta = "Proceso" Then
       ' Verifica si el proceso es el pago de planillas
       If msProceso = "PAGO_PLANILLAS" Then
            ' Inhabilita caja-bancos
            fraCB.Enabled = False
       End If
    End If
End If

End Sub

Private Sub CargarCboAfecta(sCodRec As String)
Dim sSQL As String

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
        
Case "Proceso"
    msProceso = DeterminarProceso(txtCodMov)
    
   ' Verifica si el movimiento es un proceso
   Select Case msProceso
        
   Case "PAGO_PRESTAMOS"
        'Se carga la colección de Personal
       lblEtiqueta = "Personal:"
       txtAfecta.MaxLength = 4
       CargarColPersonal
        ' Llama al proceso pago de prestamos
       If gsTipoOperacionEgreso = "Nuevo" Then
        frmCBEGPago_Prestamos.Show vbModal, Me
       End If
       
   Case "PAGO_PLANILLAS"
          ' Verifica que el movimiento sea de caja
       If optCaja.Value = True Or optBanco.Value = True Then
            lblEtiqueta = "Planilla:"
             ' Llama al proceso pago de planillas
            If gsTipoOperacionEgreso = "Nuevo" Then
             frmCBEGPago_Planillas.Show vbModal, Me
            End If
       Else ' El movimiento solo es de Caja
            MsgBox "El Movimiento elegido solo es de Caja o Bancos", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
            txtCodMov = Empty
       End If
       
   Case "PAGO_ADELANTOS"
       txtAfecta.MaxLength = 4
       lblEtiqueta = "Personal:"
       CargarColPersonal
        ' Llama al proceso pago de adelantos
       If gsTipoOperacionEgreso = "Nuevo" Then
        frmCBEGPago_Adelantos.Show vbModal, Me
       End If
   
   Case "ENTREGA_RENDIR"
       ' Verifica que el movimiento sea de caja
       If optCaja.Value Then
            txtAfecta.MaxLength = 4
            lblEtiqueta = "Cuenta a Rendir:"
            CargarColPersonal
         ' Llama al proceso entregas a rendir
            If gsTipoOperacionEgreso = "Nuevo" Then
             frmCBEGEntrega_Rendir.Show vbModal, Me
            End If
       Else ' El movimiento solo es de Caja
            MsgBox "El Movimiento elegido solo es de Caja", vbCritical + vbOKOnly, "SGCcaijo-Verifica Movimiento"
            txtCodMov = Empty
       End If
   Case Empty
        ' El movimiento no tiene proceso
       If gsTipoOperacionEgreso = "Nuevo" Then
            MsgBox "Este movimiento no esta relacionado a los Procesos:" & Chr(13) _
                   & "Pago de Planillas,Adelantos o Prestamos", , "SGCcaijo - Egreso sin Afectación"
            lblEtiqueta = Empty
            txtCodMov.SetFocus
       End If
   End Select
    
End Select

End Sub

'Private Function CargarCboAfectaPrest() As String
''----------------------------------------------------------------------------
''Propósito: Carga el combo afecta con la planilla de la tabla Préstamos
''Recibe:   Nada
''Devuelve: Nada
''----------------------------------------------------------------------------
'Dim sSQL As String
'Dim curPrestamos As New clsBD2
'
'sSQL = "SELECT DISTINCT p.idpersona, (r.Apellidos + ', ' + r.Nombre)" _
'       & " FROM PRESTAMOS AS P, PLN_PERSONAL AS R" _
'       & " Where p.idpersona = r.idpersona AND Cancelado= 'NO'" _
'       & " ORDER BY (r.Apellidos + ', ' + r.Nombre)"
'
'CD_CargarColsCbo cboAfecta, sSQL, mcolCodPersonalPrest, mcolDesCodPersonalPrest
'
'End Function

Private Function DeterminarAfecta(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que afecta Pln_Personal (P), Terceros (T), PlanContable (C)
'           un determinado tipo de   movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sAfecta  As String

' Muestra a que afecta Personal (P), Terceros (T), o Procesos de Egreso como _
  Pago de planillas , Pago de prestamos, Pago de adelantos
  
  DeterminarAfecta = mcolDesCodAfecta.Item(Trim(sCodMov))
  txtAfecta.Enabled = True: cmdBuscar.Enabled = True
  lblCodContable.Visible = True: txtCodContable.Visible = True: cboCtaContable.Visible = True: cmdPCodContable.Visible = True
  txtMonto.Enabled = True
  
End Function

Private Function DeterminarProceso(sCodMov) As String
'--------------------------------------------------------------------------
'Propósito  : Determina si el proceso esta relacionado con pago de prestamos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------------------
Dim sSQL As String
Dim curProceso As New clsBD2

'Sentencia SQL
sSQL = ""
sSQL = " SELECT PC.Proceso " _
        & "FROM PROCESO_CONCEPTOCB PC " _
        & "WHERE PC.IdConCB= '" & sCodMov & "' "

'Copia la sentencia SQL
curProceso.SQL = sSQL

'Verifica si hay error al ejecuta la sentencia SQL
If curProceso.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If
If curProceso.EOF Then

    'Devuelve vacio
    DeterminarProceso = Empty
    
Else
    'Devuelve el proceso relacionado
    DeterminarProceso = curProceso.campo(0)
End If

'Cierra el cursor
curProceso.Cerrar

End Function

Private Sub CargacboCtaContable(sCodCont As String)
'----------------------------------------------------------------------------
'Propósito: Carga el combo de la cuenta contable a partir del código contable
'           del tipo de movimiento
'Recibe:   sCodCont (Código contable del movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String

'Vaciamos las colecciones
Set mcolCodPlanCont = Nothing
Set mcolDesCodPlanCont = Nothing

' Si no tiene codigo contable sale
If sCodCont = Empty Then
    txtCodContable.BackColor = vbWhite
    lblCodContable.Visible = False: txtCodContable.Visible = False: cboCtaContable.Visible = False: cmdPCodContable.Visible = False
Else ' Carga las cuentas contables del egreso

 sSQL = "SELECT CodContable, CodContable & ' ' & Left(DescCuenta,55) FROM PLAN_CONTABLE " & _
        "WHERE CodContable LIKE '" & sCodCont & "*' And (len(CodContable)=" & miTamañoCodCont _
      & ") ORDER BY CodContable"
 CD_CargarColsCbo cboCtaContable, sSQL, mcolCodPlanCont, mcolDesCodPlanCont

 'Definimos el numero de caracteres del control txtCodMov(Conceptos)
 txtCodContable.MaxLength = miTamañoCodCont

End If

End Sub

Private Function DeterminarCodCont(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que codigo contable un determinado tipo de
'           movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------
If msAfecta = "Tercero" Or msAfecta = "Persona" Or msAfecta = "Ninguno" Then
    'Muestra a que codigo contable afecta el campo seleccionado en el combo tipo mov
    DeterminarCodCont = mcolDesCodCont.Item(Trim(sCodMov))
Else
    DeterminarCodCont = Empty
End If

End Function

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

Private Sub txtCodMov_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  
End Sub



Private Sub txtDocEgreso_Change()

'Verifica SI el campo esta vacio
If txtDocEgreso.Text <> "" And InStr(txtDocEgreso, "'") = 0 Then
  'El campos coloca a color blanco
   txtDocEgreso.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtDocEgreso.BackColor = Obligatorio
End If

'Habilita el botón aceptar en caso de estar lleno todos los campos
HabilitarBotonAceptar

End Sub

Private Sub txtDocEgreso_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  
End Sub

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


Private Sub txtNumCh_Change()

'Verifica SI el campo esta vacio
If txtNumCh.Text <> "" And InStr(txtNumCh, "'") = 0 Then
  'El campos coloca a color blanco
   txtNumCh.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtNumCh.BackColor = Obligatorio
End If

'Habilita el botón aceptar en caso de estar lleno todos los campos
HabilitarBotonAceptar

End Sub

Private Sub txtNumCh_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If

End Sub

Private Sub txtObserv_Change()

' Si en la observación hay apostrofes vacío
If InStr(txtObserv, "'") > 0 Then
   txtObserv = Empty
End If

' Habilita el botón aceptar
 HabilitarBotonAceptar

End Sub

Private Sub txtObserv_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtRinde_Change()

' Verifica si el tamaño del txt es Igual al tamaño definido
If Len(txtRinde) = txtRinde.MaxLength Then
    ' Actualiza el txtDesc
    ActualizaDescRendir
Else
    ' Limpia el txtDescRendir
    txtDescRinde = Empty
End If

' Verifica SI el campo esta vacio
If txtRinde <> Empty And txtDescRinde <> Empty Then
   ' Los campos coloca a color blanco
   txtRinde.BackColor = vbWhite
Else
  ' Marca los campos obligatorios
   txtRinde.BackColor = Obligatorio
End If

' Carga el saldo de la cuenta a rendir
CargarSaldo

' Habilita botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtRinde_KeyPress(KeyAscii As Integer)

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
  End If
  
End Sub
