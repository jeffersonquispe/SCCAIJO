VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBINxEGMndExt 
   Caption         =   "Caja y Bancos- Ingresos- Pendientes a Caja o Bancos"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   HelpContextID   =   61
   Icon            =   "SCCBINxEGMndExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPMov 
      Height          =   255
      Left            =   6600
      Picture         =   "SCCBINxEGMndExt.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   905
      Width           =   220
   End
   Begin VB.ComboBox cboCodMov 
      Height          =   315
      Left            =   2040
      Style           =   1  'Simple Combo
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   875
      Width           =   4815
   End
   Begin VB.CommandButton cmdPCtaContable 
      Height          =   255
      Left            =   6600
      Picture         =   "SCCBINxEGMndExt.frx":0BA2
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1730
      Width           =   225
   End
   Begin VB.ComboBox cboCtaContable 
      Height          =   315
      Left            =   2040
      Style           =   1  'Simple Combo
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1700
      Width           =   4815
   End
   Begin VB.CommandButton cmdPTipoDoc 
      Height          =   255
      Left            =   3645
      Picture         =   "SCCBINxEGMndExt.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2145
      Width           =   220
   End
   Begin VB.ComboBox cboTipDoc 
      Height          =   315
      Left            =   1680
      Style           =   1  'Simple Combo
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2120
      Width           =   2220
   End
   Begin VB.CommandButton cmdPCtaCte 
      Height          =   255
      Left            =   6960
      Picture         =   "SCCBINxEGMndExt.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3090
      Width           =   220
   End
   Begin VB.ComboBox cboCtaCte 
      Height          =   315
      Left            =   5205
      Style           =   1  'Simple Combo
      TabIndex        =   24
      Top             =   3070
      Width           =   1980
   End
   Begin VB.CommandButton cmdPBanco 
      Height          =   255
      Left            =   4020
      Picture         =   "SCCBINxEGMndExt.frx":142A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3090
      Width           =   220
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   1620
      Style           =   1  'Simple Combo
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3070
      Width           =   2655
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      ToolTipText     =   "Vuelve al Menú Principal"
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "&Ingresar"
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      ToolTipText     =   "Graba los datos"
      Top             =   3740
      Width           =   1000
   End
   Begin VB.TextBox txtBanco 
      Height          =   315
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   20
      Top             =   3070
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtMontoEgreso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         MaxLength       =   18
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1365
      End
      Begin MSMask.MaskEdBox mskFecEgreso 
         Height          =   300
         Left            =   3120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         Caption         =   " &Monto en Soles Pendiente de Ingreso:"
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "&Fecha de Egreso de Ctas en Dólares:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   28
      ToolTipText     =   "Vuelve al Menú Principal"
      Top             =   6460
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar2 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      ToolTipText     =   "Graba los datos"
      Top             =   6460
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      ToolTipText     =   "Vuelve al Menú Principal"
      Top             =   6460
      Width           =   1000
   End
   Begin VB.TextBox txtMontoPorIngresar 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      MaxLength       =   22
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6460
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetIngresoCajaBanco 
      Height          =   2200
      Left            =   195
      TabIndex        =   29
      Top             =   4120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3863
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2800
      Left            =   120
      TabIndex        =   36
      Top             =   690
      Width           =   9015
      Begin VB.TextBox txtObserv 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1840
         Width           =   5655
      End
      Begin VB.TextBox txtCodMov 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtDoc 
         Height          =   315
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1420
         Width           =   1335
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   7350
         MaxLength       =   14
         TabIndex        =   18
         Top             =   1420
         Width           =   1455
      End
      Begin VB.TextBox txtTipDoc 
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1420
         Width           =   420
      End
      Begin VB.TextBox txtCodContable 
         Height          =   315
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1000
         Width           =   795
      End
      Begin VB.TextBox txtAfecta 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   8
         Top             =   580
         Width           =   795
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ingreso a:"
         Height          =   735
         Left            =   7005
         TabIndex        =   37
         Top             =   585
         Width           =   1815
         Begin VB.OptionButton optBanco 
            Caption         =   "Ba&nco"
            Height          =   240
            Left            =   960
            TabIndex        =   3
            Top             =   300
            Width           =   800
         End
         Begin VB.OptionButton optCaja 
            Caption         =   "Ca&ja"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   580
         Width           =   4320
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   6240
         Picture         =   "SCCBINxEGMndExt.frx":1702
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   580
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   8520
         Picture         =   "SCCBINxEGMndExt.frx":1804
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   585
         Width           =   45
      End
      Begin VB.Label lblCodMov 
         AutoSize        =   -1  'True
         Caption         =   "Mo&vimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   195
         Width           =   855
      End
      Begin VB.Label lblFecTrab 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha:"
         Height          =   195
         Left            =   7080
         TabIndex        =   46
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblObserv 
         AutoSize        =   -1  'True
         Caption         =   "&Observación:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label lblCtaCte 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cuen&ta:"
         Height          =   195
         Left            =   4250
         TabIndex        =   44
         Top             =   2420
         Width           =   780
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   2420
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo Doc.:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Monto (S/.):"
         Height          =   195
         Left            =   6480
         TabIndex        =   41
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Doc. Ingreso:"
         Height          =   195
         Left            =   3840
         TabIndex        =   40
         Top             =   1440
         Width           =   960
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   8160
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cta.Contab&le:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1020
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2895
      Left            =   120
      TabIndex        =   48
      Top             =   3480
      Width           =   9015
   End
   Begin VB.Label Label2 
      Caption         =   "Monto por ingresar:"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   6460
      Width           =   1455
   End
End
Attribute VB_Name = "frmCBINxEGMndExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Colecciones para la carga del combo de Movimientos
Private mcolCodMov As New Collection
Private mcolCodDesCodMov As New Collection

'Colecciones para la carga del combo Tipo de Documentos
Private mcolCodDoc As New Collection
Private mcolCodDesCodDoc As New Collection

'Colecciones para la carga del combo de Ctas Ctes
Private mcolCodCtaCte As New Collection
Private mcolCodDesCtaCte As New Collection

'Colecciones para la carga del combo de Bancos
Private mcolCodBanco As New Collection
Private mcolCodDesBanco As New Collection

'Colección para la carga de Código contable y código del tipo de movimiento
Private mcolCodCont As New Collection
Private mcolDesCodCont As New Collection

'Colección para la carga de Afecta y código del tipo de movimiento
Private mcolCodAfecta As New Collection
Private mcolDesCodAfecta As New Collection

'Colección para la carga del código contable referente al tipo de movimiento
Private mcolCodPlanCont As New Collection
Private mcolDesCodPlanCont As New Collection

'Cursor de egresos
Dim curEgresoPendiente As New clsBD2

'Variable que identifica la CtaCte en soles
Dim msCtaCte As String

'Determina el maxlength del campo txtAfecta cuando es personal y terceros
Private miTamañoPer As Integer

'Variable que identifica a que Afecta el concepto(Tipo_Mov), Terceros o Personal
Private msPersTerc As String '(T o P o Vacio)
Private msCodAfectaAnterior As String ' (Código de Terceros,Personal o Vacio)

'Variable que determina el Mayor Tamaño de CodCont en Conceptos(TipoMov de la BD)
Private miTamañoCodCont As Integer

'Variable que determina el Mayor Tamaño de teceros en Conceptos(Terceros de la BD)
Private miTamañoTer As Integer

'Variable que identifica el egreso de Ctas en Dolares
Dim msCodEgreso As String

'Variable para el manjo del grid
Dim ipos As Long

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
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
   CD_ActCod cboBanco.Text, txtBanco, mcolCodBanco, mcolCodDesBanco
Else '  Vaciar Controles enlazados al combo
    txtBanco.Text = Empty
End If

'Cambia el alto del combo
cboBanco.Height = CBONORMAL

End Sub

Private Sub cboCodEgreso_Change()

End Sub

Private Sub cboCodEgreso_Click()
End Sub

Private Sub cboCodEgreso_GotFocus()
End Sub

Private Sub cboCodEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub cboCodEgreso_LostFocus()

End Sub

Private Sub cboCodMov_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCodMov) = True Then SendKeys "{down}"

End Sub

Private Sub cbocodMov_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCodMov.ListIndex) = False And cboCodMov.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

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
    txtCodMov.Text = ""
End If

'Cambia el alto del combo
cboCodMov.Height = CBONORMAL

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
  ' Se actualiza código (TextBox) correspondiente a descripción introducida
   CD_ActCboVar cboCtaCte.Text, msCtaCte, mcolCodCtaCte, mcolCodDesCtaCte
Else '  Vaciar Controles enlazados al combo
  'NO se encuentra la CtaCte
  msCtaCte = ""
End If

'Cambia el alto del combo
cboCtaCte.Height = CBONORMAL

' Habilita el botón añadir
HabilitarBotonIngresar

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
    CD_ActCod cboTipDoc.Text, txtTipDoc, mcolCodDoc, mcolCodDesCodDoc
Else '  Vaciar Controles enlazados al combo
    txtTipDoc.Text = Empty
End If

'Cambia el alto del combo
cboTipDoc.Height = CBONORMAL

End Sub

Private Sub PrepararIngresodeEgresoCtasDol()
'--------------------------------------------------------------------
'Propósito: Prepara todos los componentes del formulario, para iniciar el ingreso
'           del Monto Pendiente de ingreso relacionado con un Egreso en Ctas Dol
'Recibe: Nada
'Devuelve: Nada
'--------------------------------------------------------------------
' Nota :  LLamado al actualizar el codigo del egresoDol en el Click de cboCodEgreso

'Los controles para el ingreso de caja Habilitados y banco Deshabilitados
DeshabilitaHabilitaCtrlsCaja

'Deshabilita botones Ingresar,Aceptar2 en la 2da parte del formulario, habilita botones eliminar
cmdIngresar.Enabled = False
cmdAceptar2.Enabled = False
cmdEliminar.Enabled = True
CmdCancelar2.Enabled = True

'Inicia el con el txtMontoPorIngresar con el valor de txtMontoEgreso
txtMontoPorIngresar.Text = txtMontoEgreso.Text

End Sub

Private Sub cmd_Click()

End Sub

Private Sub cmdAceptar2_Click()
'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    Exit Sub
End If

 ' Pregunta al usuario SI esta de acuderdo con los datos
 ' Mensaje de conformidad de los datos
If MsgBox("¿Está conforme con los datos?", _
            vbQuestion + vbYesNo, "Ingreso Pendiente") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario

    ' Vacia el grdDetIngrCajaBanco a la BD
     GuardarIngresoPendienteenBD
   
    'Actualiza la transaccion
     Var8 -1, Empty
   
   ' Limpia los CtrlsCaja y deshabilita bancos, el valor de optCajalo establece a True
   ' Limpiar los Ctrls de la segunda parte del formulario
    grdDetIngresoCajaBanco.Rows = 1
    txtMontoPorIngresar.Text = "0"
   ' Deshabilita botones de la segunda parte del formulario
    Deshabilita2daPartefrm
    
    ' Inicializa el grid
      ipos = 0
      gbCambioCelda = False
   
    'verifica SI el formulario fue llamado por Modificacion de Ingresos
    If gsTipoOperacionIngreso = "Modificar" Then
        
        'Cierra el formulario y vuelve a frmCBIngresos(Modificar Ingresos)
        Unload Me
    Else ' El formulario fue llamado desde el menu principal
       
       ' Habilita primera parte del fomulario, Actualiza los pendientes SI los hay
        txtMontoEgreso.Text = Empty
        mskFecEgreso.Text = "__/__/____"
        curEgresoPendiente.Cerrar
        CargarEgresosPendientes

       'Verifica SI existen egresos pendientes a txtBanco
        If curEgresoPendiente.EOF Then
        
            'No existen Ingresos pendientes
            MsgBox "No existen ingresos pendientes", _
                    vbInformation + vbOKOnly, "Caja-Bancos- Ingresos Pendientes"
            Unload Me
        End If
    End If
    
End If

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
If msPersTerc = "Persona" Then
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
End Sub

Private Sub cmdCancelar2_Click()

'Limpiar los Ctrls de la segunda parte del formulario
PrepararIngresoCtrlsCajaBanco
grdDetIngresoCajaBanco.Rows = 1
txtMontoPorIngresar.Text = txtMontoEgreso.Text
'Deshabilita botones de la segunda parte del formulario
cmdAceptar2.Enabled = False
' Inicializa el grid
  ipos = 0
  gbCambioCelda = False
End Sub

Private Sub cmdEliminar_Click()

' Elimina la fila selccionada del Grid
If grdDetIngresoCajaBanco.CellBackColor = vbDarkBlue And grdDetIngresoCajaBanco.Row > 0 Then
      ' elimina la fila seleccionada del grid
    If grdDetIngresoCajaBanco.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdDetIngresoCajaBanco.RemoveItem grdDetIngresoCajaBanco.Row
    Else
            ' estable vacío el grddetalle
            grdDetIngresoCajaBanco.Rows = 1
    End If
    
    ' Actualiza el ipos
    ipos = 0

End If

' Calcula la suma de los montos en soles del grd DetIngresoCajaBanco
 txtMontoPorIngresar.Text = Trim(Str(Val(Var37(txtMontoEgreso.Text)) - CalculaMontoTotalSolIngr))

'InHabilita Aceptar2 de la segunda parte del formulario
cmdAceptar2.Enabled = False
If Val(Var37(txtMontoPorIngresar.Text)) = 0 Then
   ' Habilita botón Aceptar de la 2da Parte del formulario
    cmdAceptar2.Enabled = True
End If

End Sub

Private Sub cmdIngresar_Click()
Dim sDonde As String

'Verifica SI el monto + montoTotalIngr <= montoEgreso
If (Val(Var37(txtMonto.Text)) <= _
    Val(Var37(txtMontoPorIngresar.Text))) Then
    
    ' verifica SI esta el Doc en el grd detalle o en la BD la respuesta True,False
    If VerificarDocExiste = False Then
        ' Ingresa al grdDetIngresoCajaBanco los datos ingresados en los controles referentes al ingreso en caja o banco
        IngresarengrdDetIngr
        
        ' Calcula la suma de los montos en soles del grd DetIngresoCajaBanco
        txtMontoPorIngresar.Text = Trim(Str(Val(Var37(txtMontoEgreso.Text)) - CalculaMontoTotalSolIngr))
        
        ' Limpia los controles para un nuevo ingreso
        PrepararIngresoCtrlsCajaBanco
        
        'Verifica SI el monto ingresado es Igual al monto pendiente de ingreso
        If Val(Var37(txtMontoPorIngresar.Text)) = 0 Then
           ' Habilita botón Aceptar de la 2da Parte del formulario
            cmdAceptar2.Enabled = True
        End If
        
        ' Asigna el focus a opción caja
        If optCaja.Enabled Then optCaja.SetFocus
    
    Else
        ' Envia mensaje
        MsgBox "El número de documento ya existe ", vbExclamation + vbOKOnly, _
           "Caja-Banco- Ingreso generado por egreso"
           
        ' Limpia el txtDoc para dar opcion a elegir
        txtDoc.SetFocus
        cmdIngresar.Enabled = False
    
    End If
Else
    ' Envia mensaje de monto exedido, limpia txtmonto
    MsgBox "El monto ingresado excede el monto pendiente de ingreso", , "Aviso"
    txtMonto.SetFocus

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

Private Sub cmdPCtaContable_Click()

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

Private Sub cmdPMov_Click()

If cboCodMov.Enabled Then
    ' alto
     cboCodMov.Height = CBOALTO
    ' focus a cbo
    cboCodMov.SetFocus
End If

End Sub

Private Sub cmdPTipoDoc_Click()

If cboTipDoc.Enabled Then
    ' alto
     cboTipDoc.Height = CBOALTO
    ' focus a cbo
    cboTipDoc.SetFocus
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSalir_Click()

'Descarga el formulario
Unload Me

End Sub

Private Sub Form_Activate()

'SI NO existen pendientes entonces se cierra el formulario
If curEgresoPendiente.EOF Then
  'NO existen Ingresos pendientes
  MsgBox "No existen ingresos pendientes", _
          vbInformation + vbOKOnly, "S.G.Ccaijo -Ingreso Pendientes"
  Unload Me
Else
  optCaja_Click
End If

End Sub

Private Sub Form_Load()
Dim sSQL As String

'Cargamos los combos del formulario
'Se carga el combo de Tipo de Documento
sSQL = "SELECT IdTipoDoc, DescTipoDoc FROM TIPO_DOCUM" _
       & " WHERE RelacProc= 'IN'  ORDER BY DescTipodoc"
CD_CargarColsCbo cboTipDoc, sSQL, mcolCodDoc, mcolCodDesCodDoc

'Carga el combo tipo movimiento y las colecciones de tipo_mov
CargarColTipo_Mov

'Se carga el combo Bancos (sólo con los bancos que de moneda nacional)
sSQL = "SELECT DISTINCT b.IdBanco,b.DescBanco FROM TIPO_BANCOS B , TIPO_CUENTASBANC C" _
       & " WHERE b.idbanco = c.idbanco And c.idmoneda = 'SOL'" _
       & " ORDER BY DescBanco"
CD_CargarColsCbo cboBanco, sSQL, mcolCodBanco, mcolCodDesBanco

'Se carga el combo de Ctas Ctes
sSQL = "SELECT c.idCta, c.DescCta FROM TIPO_CUENTASBANC  c WHERE " & _
       "c.idmoneda= 'SOL' ORDER BY c.DescCta"
CD_CargarColsCbo cboCtaCte, sSQL, mcolCodCtaCte, mcolCodDesCtaCte

'Se Limpia el Combo de Cts Corrientes en dólares
cboCtaCte.Clear

' Actualiza Fecha de Trabajo,  la variable CtaCte vacia
mskFecTrab.Text = gsFecTrabajo
msCtaCte = ""

'Deshabilita la segunda parte del formulario
Deshabilita2daPartefrm

' Carga los egresos pendientes
CargarEgresosPendientes

' Inicializa el grid
ipos = 0
gbCambioCelda = False

'Coloca el titulo al Grid
aTitulosColGrid = Array("Ingr", "Movimiento", "Tipo_Doc", "Doc_Ingreso", "Monto S/.", "Banco", "Cta-Cte", "Cod_Mov", "Id_TipDoc", "Id_Banco", "Id_CtaCte", "Observaciones", "Afecta", "CodAfectado", "CodCont")
aTamañosColumnas = Array(600, 3400, 1500, 1500, 1600, 3000, 1500, 0, 0, 0, 0, 0, 0, 0, 0)
CargarGridTitulos grdDetIngresoCajaBanco, aTitulosColGrid, aTamañosColumnas

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

'la sentencia para cargar el combo y las colecciones de tipo movimiento
sSQL = ""
sSQL = "SELECT IdConCB, DescConCB, CodCont, Afecta FROM Tipo_MovCB " & _
          " WHERE RelacProc = 'IE' " & _
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


Private Sub VerificarCampo(TxtDato As TextBox, CboDatos As ComboBox)
'----------------------------------------------------------------------------
'PROPÓSITO: Cambia a color amarillo SI el campo obligatorio esta sin dato
'           caso contrario a blanco
'----------------------------------------------------------------------------
 
 ' Verifica SI el campo esta vacio
If TxtDato.Text <> "" And CboDatos.Text <> "" Then
' Los campos coloca a color blanco
   TxtDato.BackColor = vbWhite
Else
'Marca los campos obligatorios
   TxtDato.BackColor = Obligatorio
End If

' Habilita botón Ingresar
HabilitarBotonIngresar

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' SI ya NO hay pendientes o el formulario fue llamado por Modificar Ingresos
If Not curEgresoPendiente.EOF Then
    'Muestra el mensaje
    If MsgBox("Hay ingresos pendientes, desea salir? ", vbOKCancel + vbInformation, "SGCCaijo-Ingreso a Caja") = vbCancel Then
       Cancel = 1
       Exit Sub
    End If
End If
    
    ' Se cierra el cusor Egreso
    curEgresoPendiente.Cerrar
    
    ' Se destruyen todas las colecciones
    Set mcolCodMov = Nothing
    Set mcolCodDesCodMov = Nothing
    
    Set mcolCodDoc = Nothing
    Set mcolCodDesCodDoc = Nothing
    
    Set mcolCodBanco = Nothing
    Set mcolCodDesBanco = Nothing
    
    Set mcolCodCtaCte = Nothing
    Set mcolCodDesCtaCte = Nothing
    
    Set mcolCodAfecta = Nothing
    Set mcolDesCodAfecta = Nothing
    
    Set mcolCodCont = Nothing
    Set mcolDesCodCont = Nothing
    
    Set mcolCodPlanCont = Nothing
    Set mcolDesCodPlanCont = Nothing
    
    ' Vacía la colección de datos
    Set gcolTabla = Nothing

End Sub

Private Sub grdDetIngresoCajaBanco_Click()

If grdDetIngresoCajaBanco.Row > 0 And grdDetIngresoCajaBanco.Row < grdDetIngresoCajaBanco.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdDetIngresoCajaBanco, vbWhite, vbDarkBlue
End If

End Sub

Private Sub grdDetIngresoCajaBanco_EnterCell()

If ipos <> grdDetIngresoCajaBanco.Row Then
    '  Verifica si es la última fila
    If grdDetIngresoCajaBanco.Row > 0 And grdDetIngresoCajaBanco.Row < grdDetIngresoCajaBanco.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdDetIngresoCajaBanco, ipos
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdDetIngresoCajaBanco.Row
End If

End Sub

Private Sub grdDetIngresoCajaBanco_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Llama al proceso que cambia la verificación de un producto
    grdDetIngresoCajaBanco_DblClick
 End If
 
End Sub

Private Sub grdDetIngresoCajaBanco_DblClick()

' Selecciona toda la iFila
If grdDetIngresoCajaBanco.Rows > 1 Then
    ' Verifica si esta seleccionado
    If grdDetIngresoCajaBanco.CellBackColor <> vbDarkBlue Then
       MarcarFilaGRID grdDetIngresoCajaBanco, vbWhite, vbDarkBlue
       Exit Sub
    End If

    'Verifica SI se habilita control de banco y ctacte
    If grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 0) = "CA" Then
        optCaja.Value = True
    Else
        optBanco.Value = True
        'Recupera el codigo del banco
        txtBanco.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 9)
        'Recupera en el msCtaCte la ctacte del grid seleccionado
        msCtaCte = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 10)
        
        'Recupera la cuenta corriente del msCtaCte seleccionado
        CD_ActVarCbo cboCtaCte, msCtaCte, mcolCodDesCtaCte
    End If
    
'"Ingr", "Movimiento", "Tipo_Doc", "Doc_Ingreso", "Monto S/.", "Banco", "Cta-Cte", "Cod_Mov", "Id_TipDoc", "Id_Banco", "Id_CtaCte", _
"Observaciones", "Afecta", "CodAfectado", "CodCont"
    'Recupera los combos  del grid seleccionado
    txtCodMov.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 7)
    txtDoc.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 3)
    txtTipDoc.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 8)
    txtObserv.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 11)
    txtMonto.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 4)
    txtCodContable.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 14)
    If grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 12) = "Tercero" Or _
       grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 12) = "Persona" Then
       txtAfecta.Text = grdDetIngresoCajaBanco.TextMatrix(grdDetIngresoCajaBanco.RowSel, 13)
    End If
    
    'Elimina la fila seleccionados en el grid
    cmdEliminar_Click
    
    'Da el focus a monto
    If txtMonto.Enabled Then txtMonto.SetFocus
    
End If

End Sub

Private Sub Image1_Click()
'Carga la Var48
Var48
End Sub

Private Sub mskFecTrab_Change()
' Se valida que la fecha sea correcta
'If ValidarFecha(mskFecTrab) Then
' mskFecTrab.BackColor = vbWhite
'End If
'HabilitarBotonAceptar
End Sub


Private Sub optBanco_Click()

txtBanco.Enabled = True
'Habilita Deshabilita controles relacionados con el Banco y lso limpia
HabilitarDeshabilitarLimpiarCtrlsBanco

'Habilita el botón aceptar
HabilitarBotonIngresar

' De acuerdo a la elección realizada maneja los controles de banco
ManejaControlesBanco
End Sub

Private Sub optBanco_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub optCaja_Click()

' Deshabilita controles para banco y los limpia
HabilitarDeshabilitarLimpiarCtrlsBanco

' Habilita botón ingresar
HabilitarBotonIngresar

' De acuerdo a la elección realizada maneja los controles de banco
ManejaControlesBanco
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

'habilita el botón ingresar
HabilitarBotonIngresar

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

Private Sub txtAfecta_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtBanco_Change()

If UCase(txtBanco.Text) = txtBanco.Text Then
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
    End If
    
    ' Habilita botón aceptar
    HabilitarBotonIngresar
Else
    If Len(txtBanco.Text) = txtBanco.MaxLength Then
        'comvertimos a mayuscula
        txtBanco.Text = UCase(txtBanco.Text)
    End If
End If

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

'habilita el botón ingresar
HabilitarBotonIngresar


End Sub

Private Sub txtCodContable_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub txtCodMov_Change()

If UCase(txtCodMov.Text) = txtCodMov.Text Then
    ' SI procede, se actualiza descripción correspondiente a código introducido
    CD_ActDesc cboCodMov, txtCodMov, mcolCodDesCodMov
    
    'Limpia el cboAfecta
     txtAfecta.Text = Empty
     txtAfecta.BackColor = Obligatorio
     txtDesc.Text = Empty
     
     'Lipia los campos txtCodContable cboCtaContable
     cboCtaContable.Clear
     txtCodContable.Text = Empty
     txtCodContable.BackColor = Obligatorio
            
    ' Verifica SI el campo esta vacio
    If txtCodMov.Text <> Empty And cboCodMov.Text <> Empty Then
      ' Los campos coloca a color blanco
      txtCodMov.BackColor = vbWhite
        
      'Carga el combo Afecta dependiendo del codigo de afecta
      msPersTerc = DeterminarAfecta(txtCodMov.Text)
      CargarCboAfecta msPersTerc
           
      'Carga el cboCtaContable dependiendo del tipo de movimiento
      CargacboCtaContable DeterminarCodCont(txtCodMov.Text)
    
      'SI el combo sólo tiene un elemento, se muestra en pantalla
      MostrarUnicoItem
    
    Else
      'Marca los campos obligatorios
       txtCodMov.BackColor = Obligatorio
    End If
    
    'habilita el botón ingresar
    HabilitarBotonIngresar
Else
    If Len(txtCodMov.Text) = txtCodMov.MaxLength Then
        'comvertimos a mayuscula
        txtCodMov.Text = UCase(txtCodMov.Text)
    End If
End If

End Sub

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

sSQL = ""
sSQL = "SELECT CodContable, Left(DescCuenta,55) & ' ' & CodContable FROM PLAN_CONTABLE " & _
        "WHERE CodContable LIKE '" & sCodCont & "*' And (len(CodContable)=" & miTamañoCodCont _
         & ") ORDER BY CodContable"
CD_CargarColsCbo cboCtaContable, sSQL, mcolCodPlanCont, mcolDesCodPlanCont

'Definimos el numero de caracteres del control txtCodMov(Conceptos)
txtCodContable.MaxLength = miTamañoCodCont

End Sub

Function DeterminarCodCont(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que codigo contable un determinado tipo de
'           movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------

'Muestra a que codigo contable afecta el campo seleccionado en el combo tipo mov
DeterminarCodCont = mcolDesCodCont.Item(Trim(sCodMov))

End Function

Private Sub CargarCboAfecta(sCodRec As String)

'Verifica a que afecta Personal(P), Terceros(T), PlanContable(C o N)
Select Case sCodRec
Case "Persona"
    'Asigana el tamaño al maxlength del txtAfecta
    txtAfecta.MaxLength = 4
    lblEtiqueta.Caption = "Personal:"
    'Carga la colección
    CargarColPersonal
                   
Case "Tercero"
    'Asigana el tamaño al maxlength del txtAfecta
    txtAfecta.MaxLength = 2
    lblEtiqueta.Caption = "Terceros:"
    'Carga la colección
    CargarColTerceros
End Select
End Sub

Function DeterminarAfecta(sCodMov As String) As String
'----------------------------------------------------------------------------
'Propósito: Detemina a que afecta Pln_Personal (P), Terceros (T), PlanContable (C)
'           un determinado tipo de   movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'----------------------------------------------------------------------------

'Muestra a que afecta Personal (P), Terceros (T), PlanContble (C) el campo seleccionado en el combo
DeterminarAfecta = mcolDesCodAfecta.Item(Trim(sCodMov))

End Function

Private Sub txtCodMov_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub TxtDoc_Change()

' Verifica SI el campo esta vacio
If txtDoc.Text <> "" And InStr(txtDoc, "'") = 0 Then
' El campos coloca a color blanco
   txtDoc.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtDoc.BackColor = Obligatorio
End If

'Habilita el botón aceptar en caso de estar lleno todos los campos
HabilitarBotonIngresar

End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)

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
If txtMonto.Text <> "" Then
  'El campos coloca a color blanco
   txtMonto.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtMonto.BackColor = Obligatorio
End If

'Habilita el botón aceptar en caso de estar lleno todos los campos
HabilitarBotonIngresar

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

Private Sub DeshabilitaHabilitaCtrlsCaja()
'--------------------------------------------------------------------
'Propósito: Deshabilita, Habilita los controles referentes al ingreso a Caja.
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento click de cmdAceptar1,cmdCancelar2
'--------------------------------------------------------------------

' Deshabilita, Habilita controles de la segunda parte del formulario, referentes a  un ingreso Caja
txtCodMov.Enabled = Not txtCodMov.Enabled
cboCodMov.Enabled = Not cboCodMov.Enabled
txtAfecta.Enabled = Not txtAfecta.Enabled
cmdBuscar.Enabled = Not cmdBuscar.Enabled
txtDoc.Enabled = Not txtDoc.Enabled
txtTipDoc.Enabled = Not txtTipDoc.Enabled
cboTipDoc.Enabled = Not cboTipDoc.Enabled
txtMonto.Enabled = Not txtMonto.Enabled
txtObserv.Enabled = Not txtObserv.Enabled
optCaja.Enabled = Not optCaja.Enabled
optBanco.Enabled = Not optBanco.Enabled
txtCodContable.Enabled = Not txtCodContable.Enabled
cboCtaContable.Enabled = Not cboCtaContable.Enabled
'Marcamos los datos obligatorios
If txtCodMov.Enabled Then
    txtCodMov.BackColor = Obligatorio
Else
    txtCodMov.BackColor = vbWhite
End If
If txtDoc.Enabled Then
    txtDoc.BackColor = Obligatorio
Else
   txtDoc.BackColor = vbWhite
End If

If txtAfecta.Enabled Then
    txtAfecta.BackColor = Obligatorio
Else
   txtAfecta.BackColor = vbWhite
End If
If txtMonto.Enabled Then
    txtMonto.BackColor = Obligatorio
Else
    txtMonto.BackColor = vbWhite
End If
If txtTipDoc.Enabled Then
    txtTipDoc.BackColor = Obligatorio
Else
    txtTipDoc.BackColor = vbWhite
End If
If txtCodContable.Enabled Then
    txtCodContable.BackColor = Obligatorio
Else
    txtCodContable.BackColor = vbWhite
End If

End Sub

Public Sub ActualizarListcboCtaCte()
'--------------------------------------------------------------------
'Propósito: Actualizar la lista del combo CtaCte de acuerdo al txtBanco
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento lostfocus de cboBanco y Change de txtBanco
'--------------------------------------------------------------------

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

Public Sub CargarEgresosPendientes()
'--------------------------------------------------------------------
'Propósito: Carga un cursor con el IdEgreso, Fecha, Monto del egreso
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento formload del frm , y Click de aceptar2 en la segundapare del formulario
'--------------------------------------------------------------------

Dim sSQL As String

' Carga la Sentencia para obtener los Egresos pendientes de ingreso a caja o a banco
sSQL = "SELECT IdEgreso,Fecha,MontoSol FROM EGRESO_CTAS_EXTR " & _
       "WHERE IngrePendiente = 'SI' AND Anulado = 'NO' ORDER BY IdEgreso"
' Carga el cursor
curEgresoPendiente.SQL = sSQL
If curEgresoPendiente.Abrir = HAY_ERROR Then
  End
End If
    
'Verifica SI existen egresos pendientes a de ingreso a caja o banco
If Not curEgresoPendiente.EOF Then
  'Carga el 1er pendiente de ingreso de acuerdo a la cronologia del _
   egreso de Ctas Extrangeras
    ' Se actualiza la fecha y el monto de el egreso
    ActualizarCtrlsEgreso
    ' Deshabilita controles refentes al egreso de CtasDol, Habilita los demas del formulario
    PrepararIngresodeEgresoCtasDol
End If

End Sub

Public Sub Deshabilita2daPartefrm()
'--------------------------------------------------------------------
'Propósito: Deshabilita los controles de la 2da parte del formulario
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento formload del frm , y Click de Cancelar2
'--------------------------------------------------------------------

'Deshabilita controles de la segunda parte del formulario
txtCodMov.Enabled = False
cboCodMov.Enabled = False
txtAfecta.Enabled = False
cmdBuscar.Enabled = False
txtDoc.Enabled = False
txtTipDoc.Enabled = False
cboTipDoc.Enabled = False
txtMonto.Enabled = False
txtObserv.Enabled = False
optCaja.Enabled = False
optBanco.Enabled = False
txtBanco.Enabled = False
cboBanco.Enabled = False
cboCtaCte.Enabled = False
cmdIngresar.Enabled = False
cmdEliminar.Enabled = False
cmdAceptar2.Enabled = False
CmdCancelar2.Enabled = False
txtCodContable.Enabled = False
cboCtaContable.Enabled = False

'Marcamos los datos obligatorios
txtCodMov.BackColor = vbWhite
txtAfecta.BackColor = vbWhite
txtDoc.BackColor = vbWhite
txtMonto.BackColor = vbWhite
txtTipDoc.BackColor = vbWhite

End Sub

Private Sub ActualizarCtrlsEgreso()
'--------------------------------------------------------------------
'Propósito: Actualiza los controles referentes a item seleccionado en cboCodEgreso
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el lostfocus de cboCodEgreso
'--------------------------------------------------------------------
If gsTipoOperacionIngreso = "Modificar" Then
  ' Recorre el cursor de pendientes para encontrar el pendiente generado por _
    una modificacion en algun registro de ingreso
  Do While Not curEgresoPendiente.EOF
    If gsCodEgreso = curEgresoPendiente.campo(0) Then
        msCodEgreso = curEgresoPendiente.campo(0)
        mskFecEgreso.Text = FechaDMA(curEgresoPendiente.campo(1))
        txtMontoEgreso.Text = Format(Str(curEgresoPendiente.campo(2)), "###,###,###,##0.00")
    End If
    curEgresoPendiente.MoverSiguiente
  Loop
  curEgresoPendiente.MoverPrimero
Else
    msCodEgreso = curEgresoPendiente.campo(0)
    mskFecEgreso.Text = FechaDMA(curEgresoPendiente.campo(1))
    txtMontoEgreso.Text = Format(Str(curEgresoPendiente.campo(2)), "###,###,###,##0.00")
End If

End Sub



Private Sub txtMontoPorIngresar_Change()

'Da formato de numero a txtMontoTotal
txtMontoPorIngresar.Text = Format(Val(Var37(txtMontoPorIngresar.Text)), "###,###,###,##0.00")

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
CD_ActDesc cboTipDoc, txtTipDoc, mcolCodDesCodDoc

'Verifica  que NO tenga campos en blanco
VerificarCampo txtTipDoc, cboTipDoc

End Sub

Private Sub HabilitarDeshabilitarLimpiarCtrlsBanco()
'--------------------------------------------------------------------
'Propósito: Habilita SI optBanco=true ,Deshabilita y Limpia SI optCaja=true
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento click de optCaja y optBanco
'--------------------------------------------------------------------

'Verifica SI los controles estan habiliatados para limpiarlos
If txtBanco.Enabled = True Then
    txtBanco.Text = Empty
    msCtaCte = Empty
    txtBanco.BackColor = Obligatorio
    cboCtaCte.BackColor = Obligatorio
Else
    txtBanco.BackColor = vbWhite
    cboCtaCte.BackColor = vbWhite
End If

'Habilita deshabilita los controles de Banco
'txtBanco.Enabled = Not txtBanco.Enabled
'cboBanco.Enabled = Not cboBanco.Enabled
'cboCtaCte.Enabled = Not cboCtaCte.Enabled

End Sub

Private Sub HabilitarBotonIngresar()
'--------------------------------------------------------------------
'Propósito: Habilita el botón ingresar SI los controles obligatorios
'            están rellenos
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento change de las cajas de texto y
'            lostfocus de cboCtaCTe
'--------------------------------------------------------------------

If optBanco.Value = True Then 'Verifica SI el ingreso es a Bancos
    If txtCodMov.BackColor <> Obligatorio _
       And txtDoc.BackColor <> Obligatorio _
       And txtAfecta.BackColor <> Obligatorio _
       And txtCodContable.BackColor <> Obligatorio _
       And txtTipDoc.BackColor <> Obligatorio _
       And txtMonto.BackColor <> Obligatorio _
       And txtBanco.BackColor <> Obligatorio _
       And cboCtaCte.BackColor <> Obligatorio Then
       'Habilita el botón aceptar
       cmdIngresar.Enabled = True
    Else
      'Desabilita el botón aceptar
       cmdIngresar.Enabled = False
    End If

Else ' el ingreso es de Caja
    If txtCodMov.BackColor <> Obligatorio _
       And txtDoc.BackColor <> Obligatorio _
       And txtTipDoc.BackColor <> Obligatorio _
       And txtAfecta.BackColor <> Obligatorio _
       And txtCodContable.BackColor <> Obligatorio _
       And txtMonto.BackColor <> Obligatorio Then
        'Habilita el botón aceptar
        cmdIngresar.Enabled = True
    Else
      'Desabilita el botón aceptar
       cmdIngresar.Enabled = False
    End If

End If


End Sub

Private Function VerificarDocExiste() As Boolean
'--------------------------------------------------------------------
'Propósito: Verifica SI el Doc ha sido ingresado
'Recibe:    Nada
'Devuelve:  False: NO existe, True: Existe
'Nota:      llamado desde el evento click de cmdañadir en al 2da parte del formulario
'--------------------------------------------------------------------

Dim j As Integer
Dim sSQL As String
Dim curDocIngresado As New clsBD2

VerificarDocExiste = False

'Verifica el Doc esta en el grdDetIngrCajaBanco
For j = 1 To grdDetIngresoCajaBanco.Rows - 1    'recorremos las filas del grdDetIngrCajaBanco
 If grdDetIngresoCajaBanco.TextMatrix(j, 3) = txtDoc.Text Then
    
    VerificarDocExiste = True
    Exit Function
    
 End If
Next j


    'Verifica SI el Doc esta en Caja o en Banco de la tabla ingresos
    'Se averigua SI existe algun documento con el mismo numero en Banco
    sSQL = "SELECT Count(I.NumDoc) as NroDoc FROM INGRESOS I " & _
           "WHERE I.NumDoc = '" & txtDoc.Text & "'"
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

Private Sub IngresarengrdDetIngr()
'--------------------------------------------------------------------
'Propósito: Añadir al grd det una fila con los datos del ingreso a caja o a bancos
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento click de cmdañadir en al 2da parte del formulario
'--------------------------------------------------------------------

'  Añade el nuevo registro al grid
Dim sIngr As String
If optCaja = True Then
    sIngr = "CA"
Else: sIngr = "BA"
End If
 '("Ingr", "Movimiento", "Tipo_Doc", "Doc_Ingreso", "Monto S/.", "Banco", "Cta-Cte", "Cod_Mov", "Id_TipDoc", "Id_Banco", "Id_CtaCte", "Observaciones", "Afecta", "CodAfectado", "CodCont")
grdDetIngresoCajaBanco.AddItem (sIngr & vbTab & cboCodMov.Text & vbTab & cboTipDoc.Text _
                    & vbTab & txtDoc.Text & vbTab & txtMonto.Text & vbTab & cboBanco.Text _
                    & vbTab & cboCtaCte.Text & vbTab & txtCodMov.Text & vbTab & txtTipDoc.Text _
                    & vbTab & txtBanco.Text & vbTab & msCtaCte & vbTab & txtObserv.Text _
                    & vbTab & msPersTerc & vbTab & txtAfecta.Text & vbTab & txtCodContable.Text)

End Sub

Private Function CalculaMontoTotalSolIngr() As Double
'--------------------------------------------------------------------
'Propósito: Calcula el monto ingresado a Caja y a Banco en grdDetIngrCajaBanco
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento click de cmdañadir en al 2da parte del formulario
'--------------------------------------------------------------------

Dim j As Integer
Dim dAcuml As Double

'Inicializamos a funcion asumiendo que la Cta NO esta en el grdDetEgresoctadol
dAcuml = 0

' recorremos el grid detalle de Egreso verificando la existencia de msCtaCte
For j = 1 To grdDetIngresoCajaBanco.Rows - 1
  dAcuml = dAcuml + Val(Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4)))
  Next j

'devuelve el valor a la funcion
CalculaMontoTotalSolIngr = dAcuml

End Function

Private Sub PrepararIngresoCtrlsCajaBanco()
'--------------------------------------------------------------------
'Propósito: Limpia los controles de Caja y Banco para un nuevo ingreso
'           El valor de optCaja lo pone Verdadero(Defecto), para el ingreso
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------------
'Nota:      llamado desde el evento click de cmdañadir en cmdCancelar2 en 2da parte del formulario
' Limpia los controles de  ingreso caja y Banco para un nuevo ingreso
txtCodMov.Text = Empty
txtDoc.Text = Empty
txtTipDoc.Text = Empty
txtMonto.Text = Empty
txtObserv.Text = Empty
optCaja.Value = True

End Sub

Private Sub GuardarIngresoPendienteenBD()
'------------------------------------------------------
'Propósito: Guarda los ingresos en CAJA o BANCO de acuerdo a la columna 0 "Ingr" del gridDetIngrCajaBanco
'Recibe:    Nada
'Devuelve:  Nada
'Nota:      llamado desde el evento click de cmdAceptar2 en 2da parte del formulario
'------------------------------------------------------
Dim j As Integer
Dim sSQL, sOrden As String
Dim modIngrCajaBanco As New clsBD3
    
' Recorre el grid y lo almacena en la BD
For j = 1 To grdDetIngresoCajaBanco.Rows - 1 'recorre las Filas (Ctas en dólares de las cuales se saco dinero)
'("Ingr", "Movimiento", "Tipo_Doc", "Doc_Ingreso", "Monto S/.", "Banco", "Cta-Cte", _
 "Cod_Mov", "Id_TipDoc", "Id_Banco", "Id_CtaCte", "Observaciones")
    'Guarda los datos introducidos en la tabla Ingresos segun Columna 0 del grd
    If grdDetIngresoCajaBanco.TextMatrix(j, 0) = "CA" Then    ' Guarda en Caja
           ' Guardar los  datos
         sOrden = CalcularSigOrden("CA") 'Código de Ingreso
         sSQL = "INSERT INTO INGRESOS VALUES('" & sOrden & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 3) & "','" & grdDetIngresoCajaBanco.TextMatrix(j, 8) & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 7) & "','" & FechaAMD(mskFecTrab.Text) & "'," _
           & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4)) & ",'" & msCodEgreso & "','','NO','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 11) & "','" & grdDetIngresoCajaBanco.TextMatrix(j, 14) & "')"
           
           ' carga la colección asiento
           'Orden,Monto,NumCtaBanc, fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=sOrden, _
         Item:=sOrden & "¯" _
           & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4)) _
           & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) & "¯" _
           & "INGRESO A CAJA" & "¯IN¯IC¯C"
           
        gcolAsientoDet.Add _
        Key:=grdDetIngresoCajaBanco.TextMatrix(j, 14), _
        Item:=grdDetIngresoCajaBanco.TextMatrix(j, 14) & "¯" & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4))
        
        
    Else    'Guarda en Banco
         sOrden = CalcularSigOrden("BA") 'Código de Ingreso
         sSQL = "INSERT INTO INGRESOS VALUES('" & sOrden & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 3) & "','" & grdDetIngresoCajaBanco.TextMatrix(j, 8) & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 7) & "','" & FechaAMD(mskFecTrab.Text) & "'," _
           & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4)) & ",'" & msCodEgreso & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 10) & "','NO','" & grdDetIngresoCajaBanco.TextMatrix(j, 11) & "','" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 14) & "')"
           
         ' carga la colección asiento
         'Orden,Monto,NumCtaBanc, fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=sOrden, _
         Item:=sOrden & "¯" _
           & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4)) & "¯" _
           & grdDetIngresoCajaBanco.TextMatrix(j, 10) & "¯" _
           & FechaAMD(mskFecTrab.Text) & "¯" _
           & "INGRESO A BANCO¯IN¯IB¯B"
           
        gcolAsientoDet.Add _
        Key:=grdDetIngresoCajaBanco.TextMatrix(j, 14), _
        Item:=grdDetIngresoCajaBanco.TextMatrix(j, 14) & "¯" & Var37(grdDetIngresoCajaBanco.TextMatrix(j, 4))

     End If
        
        'SI al ejecutar hay error se sale de la aplicación
        modIngrCajaBanco.SQL = sSQL
        If modIngrCajaBanco.Ejecutar = HAY_ERROR Then
          End
        End If
                
        'Se cierra la query
        modIngrCajaBanco.Cerrar
    
        'Guardar Mov Afectado SI Afecta es terceros o Persoanl(Concepto afecta a T o P)
        If grdDetIngresoCajaBanco.TextMatrix(j, 12) = "Tercero" Or _
           grdDetIngresoCajaBanco.TextMatrix(j, 12) = "Persona" Then
             'Verifica SI el Mov Afectado es terceros
             If grdDetIngresoCajaBanco.TextMatrix(j, 12) = "Tercero" Then
                'Cargamos sentencia que guarda en BD MOV_TERCERO
                sSQL = "INSERT INTO MOV_TERCEROS VALUES('" _
                  & sOrden & "','" _
                  & Trim(grdDetIngresoCajaBanco.TextMatrix(j, 13)) & "')"
             End If
             'Verifica SI el Mov Afectado es Personal
             If grdDetIngresoCajaBanco.TextMatrix(j, 12) = "Persona" Then
                'Cargamos sentencia que guarda en BD MOV_PERSONAL
                sSQL = "INSERT INTO MOV_PERSONAL VALUES('" _
                  & sOrden & "','" _
                  & Trim(grdDetIngresoCajaBanco.TextMatrix(j, 13)) & "')"
             End If
            
            'SI al ejecutar hay error se sale de la aplicación
            modIngrCajaBanco.SQL = sSQL
            If modIngrCajaBanco.Ejecutar = HAY_ERROR Then
              End
            End If
                    
            'Se cierra la query
            modIngrCajaBanco.Cerrar

        End If
        
        ' Genera el asiento automático
        Conta13

Next j

' Modificamos el campo pendiente de la tabla EGRESO_CTAS_EXTR a "NO" (el Egreso en  Ctas dólares ha sido ingresado a Caja o Banco )
sSQL = "UPDATE EGRESO_CTAS_EXTR " _
        & "SET IngrePendiente = 'NO',MontoSol=0 " _
        & "WHERE IdEgreso=" & "'" & msCodEgreso & "'"
        
'SI al ejecutar hay error se sale de la aplicación
modIngrCajaBanco.SQL = sSQL
If modIngrCajaBanco.Ejecutar = HAY_ERROR Then
  End
End If
        
'Se cierra la query
modIngrCajaBanco.Cerrar

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
 cboBanco.Enabled = False: cboCtaCte.Enabled = False
 txtBanco.Enabled = False

Else
 'Muestra los controles de banco
 lblBanco.Visible = True: txtBanco.Visible = True: cboBanco.Visible = True
 lblCtaCte.Visible = True: cboCtaCte.Visible = True
 cmdPBanco.Visible = True: cmdPCtaCte.Visible = True
 cboBanco.Enabled = True: cboCtaCte.Enabled = True
 txtBanco.Enabled = True
End If

End Sub
