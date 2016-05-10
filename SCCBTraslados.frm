VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBTraslado 
   Caption         =   "Caja - Banco Traslados"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   HelpContextID   =   56
   Icon            =   "SCCBTraslados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPTipoDocDestino 
      Height          =   255
      Left            =   4290
      Picture         =   "SCCBTraslados.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5670
      Width           =   220
   End
   Begin VB.ComboBox cboTipoDocDestino 
      Height          =   315
      Left            =   1560
      Style           =   1  'Simple Combo
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2985
   End
   Begin VB.CommandButton cmdPCtaCteDestino 
      Height          =   255
      Left            =   6900
      Picture         =   "SCCBTraslados.frx":0BA2
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6270
      Width           =   220
   End
   Begin VB.ComboBox cboCtaCteDestino 
      Height          =   315
      Left            =   5520
      Style           =   1  'Simple Combo
      TabIndex        =   34
      Top             =   6240
      Width           =   1620
   End
   Begin VB.CommandButton cmdPBancoDestino 
      Height          =   255
      Left            =   4080
      Picture         =   "SCCBTraslados.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6275
      Width           =   225
   End
   Begin VB.ComboBox cboBancoDestino 
      Height          =   315
      Left            =   1560
      Style           =   1  'Simple Combo
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2775
   End
   Begin VB.CommandButton cmdPTipDoc 
      Height          =   255
      Left            =   4290
      Picture         =   "SCCBTraslados.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3270
      Width           =   225
   End
   Begin VB.ComboBox cboTipDocOrigen 
      Height          =   315
      Left            =   1560
      Style           =   1  'Simple Combo
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2985
   End
   Begin VB.CommandButton cmdPCtaCteOrigen 
      Height          =   255
      Left            =   6930
      Picture         =   "SCCBTraslados.frx":142A
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3750
      Width           =   220
   End
   Begin VB.ComboBox cboCtaCteOrigen 
      Height          =   315
      Left            =   5520
      Style           =   1  'Simple Combo
      TabIndex        =   20
      Top             =   3720
      Width           =   1665
   End
   Begin VB.CommandButton cmdPBancoOrigen 
      Height          =   255
      Left            =   4080
      Picture         =   "SCCBTraslados.frx":1702
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3755
      Width           =   225
   End
   Begin VB.ComboBox cboBancoOrigen 
      Height          =   315
      Left            =   1560
      Style           =   1  'Simple Combo
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   3120
      TabIndex        =   36
      ToolTipText     =   "Graba los datos"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6360
      TabIndex        =   39
      ToolTipText     =   "Volver al Menú Principal"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "An&ular"
      Height          =   405
      Left            =   5280
      TabIndex        =   38
      ToolTipText     =   "Graba los datos"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   4200
      TabIndex        =   37
      ToolTipText     =   "Vuelve al Menú Principal"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   120
      TabIndex        =   42
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cboPMovimiento 
         Height          =   255
         Left            =   6720
         Picture         =   "SCCBTraslados.frx":19DA
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   220
      End
      Begin VB.ComboBox cboMovimiento 
         Height          =   315
         Left            =   1750
         Style           =   1  'Simple Combo
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   5220
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6480
         Picture         =   "SCCBTraslados.frx":1CB2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1750
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   4710
      End
      Begin VB.TextBox txtCodPersonal 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtCodMov 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtObservacion 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1680
         Width           =   5880
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   5800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha:"
         Height          =   195
         Left            =   5160
         TabIndex        =   59
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "&Observación:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Monto:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Personal:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mo&vimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origen de Transferencia"
      Height          =   2295
      Left            =   120
      TabIndex        =   40
      Top             =   2400
      Width           =   7215
      Begin VB.TextBox txtNumCh 
         Height          =   315
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtSaldoOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1620
      End
      Begin VB.TextBox txtTipoDocOrigen 
         Height          =   315
         Left            =   960
         MaxLength       =   2
         TabIndex        =   12
         Top             =   840
         Width           =   420
      End
      Begin VB.TextBox txtDocOrigen 
         Height          =   315
         Left            =   5520
         MaxLength       =   15
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBancoOrigen 
         Height          =   315
         Left            =   960
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1320
         Width           =   450
      End
      Begin VB.Frame fraIngreso 
         Caption         =   "&Egreso de "
         Height          =   615
         Left            =   2760
         TabIndex        =   43
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optBancoOrigen 
            Caption         =   "Ba&nco"
            Height          =   255
            Left            =   960
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCajaOrigen 
            Caption         =   "Ca&ja"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtOrdenOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   10
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Cheque:"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   1875
         Width           =   1425
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Origen:"
         Height          =   195
         Left            =   4320
         TabIndex        =   60
         Top             =   1845
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "T&ipo Doc.:"
         Height          =   255
         Left            =   135
         TabIndex        =   48
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Documento:"
         Height          =   195
         Left            =   4560
         TabIndex        =   47
         Top             =   840
         Width           =   870
      End
      Begin VB.Label lblCtaCteOrigen 
         Caption         =   "Nº Cuen&ta:"
         Height          =   195
         Left            =   4440
         TabIndex        =   46
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblBancoOrigen 
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1365
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destino de Transferencia"
      Height          =   2415
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   7215
      Begin VB.TextBox txtSaldoDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1965
         Width           =   1575
      End
      Begin VB.TextBox txtTipoDocDestino 
         Height          =   315
         Left            =   960
         MaxLength       =   2
         TabIndex        =   26
         Top             =   840
         Width           =   420
      End
      Begin VB.TextBox txtDocDestino 
         Height          =   315
         Left            =   5535
         MaxLength       =   15
         TabIndex        =   29
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBancoDestino 
         Height          =   315
         Left            =   960
         MaxLength       =   2
         TabIndex        =   30
         Top             =   1440
         Width           =   450
      End
      Begin VB.Frame fraEgreso 
         Caption         =   "Ingreso a:"
         Height          =   615
         Left            =   3600
         TabIndex        =   44
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optBancoDestino 
            Caption         =   "Ba&nco"
            Height          =   255
            Left            =   960
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCajaDestino 
            Caption         =   "Ca&ja"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtOrdenDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   10
         TabIndex        =   23
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Destino:"
         Height          =   195
         Left            =   4250
         TabIndex        =   62
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "T&ipo Doc.:"
         Height          =   255
         Left            =   135
         TabIndex        =   52
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Documento:"
         Height          =   195
         Left            =   4560
         TabIndex        =   51
         Top             =   840
         Width           =   870
      End
      Begin VB.Label lblCtaCteDestino 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cuen&ta:"
         Height          =   195
         Left            =   4560
         TabIndex        =   50
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label lblBancoDestino 
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1440
         Width           =   600
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7320
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmCBTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Coleccion para cargar los movimientos
Private mcolCodMov As New Collection
Private mcolCodDesCodMov As New Collection

'Coleccion para cargar el personal
Private mcolCodPersonal As New Collection
Private mcolCodDesCodPersonal As New Collection

'Coleccion para la carga de bancos
Private mcolCodBanco As New Collection
Private mcolCodDesBanco As New Collection

'Colección para la carga de ctasctes
Private mcolCodCtaCte As New Collection
Private mcolCodDesCtaCte As New Collection

'Colección para la carga de tipo de documentos
Private mcolCodTipDocOrigen As New Collection
Private mcolCodDesTipDocOrigen As New Collection
Private mcolCodTipDocDestino As New Collection
Private mcolCodDesTipDocDestino As New Collection

'Variable para cargar las ctas corrientes msCtaCtes
Private msCtaCteOrigen As String
Private msCtaCteDestino As String

'Variable para determinar el movimiento de caja o banco
Private msCajaoBancoOrigen As String
Private msCajaoBancoDestino As String

'Determina el proceso del tipo de movimiento
Private mcolProceso As New Collection

'Cursor para guargar los traslados
Private mcurRegTrasladosOrigen As New clsBD2
Private mcurRegTrasladosDestino As New clsBD2

'Variable para cargar formulario
Private mbCargadoOrigen As Boolean
Private mbCargadoDestino As Boolean


Private Sub cboBancoDestino_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboBancoDestino) = True Then SendKeys "{down}"

End Sub

Private Sub cboBancoDestino_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboBancoDestino.ListIndex) = False And cboBancoDestino.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
   
End Sub

Private Sub cboBancoDestino_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboBancoDestino_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboBancoDestino, vbWhite) = True Then

  'Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboBancoDestino.Text, txtBancoDestino, mcolCodBanco, mcolCodDesBanco
Else
  txtBancoDestino.Text = Empty
End If

'Cambia el alto del combo
 cboBancoDestino.Height = CBONORMAL

End Sub
Private Sub cboBancoOrigen_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboBancoOrigen) = True Then SendKeys "{down}"

End Sub

Private Sub cboBancoOrigen_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboBancoOrigen.ListIndex) = False And cboBancoOrigen.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
   
End Sub

Private Sub cboBancoOrigen_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboBancoOrigen_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboBancoOrigen, vbWhite) = True Then

  'Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboBancoOrigen.Text, txtBancoOrigen, mcolCodBanco, mcolCodDesBanco
Else
  txtBancoOrigen.Text = Empty
End If

'Cambia el alto del combo
 cboBancoOrigen.Height = CBONORMAL

End Sub



Private Sub cboCtaCteDestino_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaCteDestino) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaCteDestino_Click()

' Deshabilita el botón
cmdAceptar.Enabled = True

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaCteDestino.ListIndex) = False And cboCtaCteDestino.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCtaCteDestino_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCtaCteDestino_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCtaCteDestino, Obligatorio) = True Then
    
    'Se actualiza código (String) correspondiente a descripción introducida
    CD_ActCboVar cboCtaCteDestino.Text, msCtaCteDestino, mcolCodCtaCte, mcolCodDesCtaCte
      
    ' Verifica SI el campo esta vacio
    If cboCtaCteDestino.Text <> Empty Then
       ' Los campos coloca a color blanco
       cboCtaCteDestino.BackColor = vbWhite
       txtBancoDestino.BackColor = vbWhite
       
       'Carga el saldo de destino
       CargarSaldoDestino
    Else
       'Marca los campos obligatorios
       cboCtaCteDestino.BackColor = Obligatorio
    End If
Else
  'No se encuentra la CtaCte
  msCtaCteDestino = Empty
  txtSaldoDestino = "0.00"
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

'Cambia el alto del combo
cboCtaCteDestino.Height = CBONORMAL

End Sub

Private Sub cboCtaCteOrigen_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaCteOrigen) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaCteOrigen_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaCteOrigen.ListIndex) = False And cboCtaCteOrigen.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboCtaCteOrigen_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCtaCteOrigen_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCtaCteOrigen, Obligatorio) = True Then
    
    'Se actualiza código (String) correspondiente a descripción introducida
    CD_ActCboVar cboCtaCteOrigen.Text, msCtaCteOrigen, mcolCodCtaCte, mcolCodDesCtaCte
      
    ' Verifica SI el campo esta vacio
    If cboCtaCteOrigen.Text <> "" Then
       ' Los campos coloca a color blanco
       cboCtaCteOrigen.BackColor = vbWhite
       txtBancoOrigen.BackColor = vbWhite
       
       'Carga el saldo
       CargarSaldoOrigen
       
    Else
       'Marca los campos obligatorios
       cboCtaCteOrigen.BackColor = Obligatorio
    End If
Else

  'No se encuentra la CtaCte
  msCtaCteOrigen = Empty
  txtSaldoOrigen = "0.00"
  
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

'Cambia el alto del combo
cboCtaCteOrigen.Height = CBONORMAL

End Sub


Private Sub cboMovimiento_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboMovimiento) = True Then SendKeys "{down}"

End Sub

Private Sub cboMovimiento_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboMovimiento.ListIndex) = False And cboMovimiento.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboMovimiento_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboMovimiento, vbWhite) = True Then
  ' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboMovimiento.Text, txtCodMov, mcolCodMov, mcolCodDesCodMov

Else '  Vaciar Controles enlazados al combo
    txtCodMov.Text = Empty
End If

'Cambia el alto del combo
cboMovimiento.Height = CBONORMAL

End Sub

Private Sub cboPMovimiento_Click()

If cboMovimiento.Enabled Then
    ' alto
     cboMovimiento.Height = CBOALTO
    ' focus a cbo
    cboMovimiento.SetFocus
End If

End Sub



Private Sub cboTipDocOrigen_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboTipDocOrigen) = True Then SendKeys "{down}"

End Sub

Private Sub cboTipDocOrigen_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboTipDocOrigen.ListIndex) = False And cboTipDocOrigen.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboTipDocOrigen_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboTipDocOrigen_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboTipDocOrigen, vbWhite) = True Then

  ' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboTipDocOrigen.Text, txtTipoDocOrigen, mcolCodTipDocOrigen, mcolCodDesTipDocOrigen

Else '  Vaciar Controles enlazados al combo
    txtTipoDocOrigen.Text = Empty
End If

'Cambia el alto del combo
cboTipDocOrigen.Height = CBONORMAL

End Sub

Private Sub cboTipoDocDestino_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboTipoDocDestino) = True Then SendKeys "{down}"

End Sub

Private Sub cboTipoDocDestino_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboTipoDocDestino.ListIndex) = False And cboTipoDocDestino.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboTipoDocDestino_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboTipoDocDestino_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboTipoDocDestino, vbWhite) = True Then

  ' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboTipoDocDestino.Text, txtTipoDocDestino, mcolCodTipDocDestino, mcolCodDesTipDocDestino

Else '  Vaciar Controles enlazados al combo
    txtTipoDocDestino.Text = Empty
End If

'Cambia el alto del combo
cboTipoDocDestino.Height = CBONORMAL

End Sub

Private Sub cmdAceptar_Click()
Dim dblMonto, dblMontoAnt As Double

'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operaciónes
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    Exit Sub
End If

'Verifica que los documentos sean distintos
If msCtaCteOrigen = msCtaCteDestino Then
      'El documento ya ha sido ingresado mandamos mensaje
        MsgBox "¡El número de la cuenta de ingreso y egreso es la misma! ", _
        vbExclamation + vbOKOnly, _
        "Caja-Bancos- Traslados"
        
        'Ubica el cursor en el cboCtaCteDestino
        cboCtaCteDestino.SetFocus
        Exit Sub
  
End If

'Verifica si existe un documento duplicado en el egreso
If VerificarDocExisteEgreso Then

  'El documento ya ha sido ingresado mandamos mensaje
  If MsgBox("El número de documento del egreso esta duplicado, ¿desea continuar con este mismo número de documento? ", _
        vbQuestion + vbYesNo, _
        "Caja-Bancos- Traslados") = vbNo Then
        
        txtDocOrigen.SetFocus
        Exit Sub
  End If
End If

'Verifica si existe un documento duplicado en el ingreso
If VerificarDocExisteIngreso Then

  'El documento ya ha sido ingresado mandamos mensaje
  If MsgBox("El número de documento del ingreso esta duplicado, ¿desea continuar con este mismo número de documento? ", _
        vbQuestion + vbYesNo, _
        "Caja-Bancos- Traslados") = vbNo Then
        
        txtDocDestino.SetFocus
        Exit Sub
  End If
End If


' Verifica si los datos son correctos
If fbVerificarDatosIntroducidosEgreso = False Then
    ' Algún dato es incorrecto
    Exit Sub
End If

' Verifica si los datos son correctos
If fbVerificarDatosIntroducidosIngreso = False Then
    ' Algún dato es incorrecto
    Exit Sub
End If

If gsTipoOperacionTraslado = "Nuevo" Then
     ' Pregunta aceptación de los datos
   If MsgBox("¿Está conforme con los datos?", _
      vbQuestion + vbYesNo, "Caja-Bancos, Traslados") = vbYes Then
      'Actualiza la transaccion
       Var8 1, gsFormulario
      
       ' Se guardan los datos del Traslado
       GuardarTraslado
       
   Else: Exit Sub ' Sale
   End If
Else

    ' Mensaje de conformidad de los datos
      If MsgBox("¿Está conforme con las modificaciones realizadas en el Traslado ?", _
                  vbQuestion + vbYesNo, "Caja-Bancos, Modificación de Traslados") = vbYes Then
           'Actualiza la transaccion
           Var8 1, gsFormulario
          
           ' Se Modifican los datos del egreso
           GuardarModificarTraslados
           
      Else: Exit Sub ' Sale
      End If

End If

'Actualiza la transaccion
 Var8 -1, Empty

' Mensaje Ok
MsgBox "Operación efectuada correctamente", , "SGCCaijo-Traslados"

'Limpia la pantalla para una nueva operación, Prepara el formulario
LimpiarFormulario
   
If gsTipoOperacionTraslado = "Nuevo" Then
    'Nuevo egreso
    NuevoTraslado
   
Else

  ' Se Modifican los datos del egreso
   ModificarTraslado
    
  'Descarga el formulario
  Unload Me

End If

End Sub

Private Sub GuardarModificarTraslados()
'---------------------------------------------------------------
'Propósito  : Realiza la operación de modificar en el formulario
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
        
  'Modifica los traslados
  ModificarRegistroTraslados
  
  ' Limpia la las cajas de texto del formulario
  LimpiarFormulario
  
 ' cierra el control egreso
 mcurRegTrasladosOrigen.Cerrar
 mcurRegTrasladosDestino.Cerrar

End Sub

Private Sub ModificarRegistroTraslados()
'----------------------------------------------------------------------------
'Propósito  : Guarda los traslados en EGRESOS  e INGRESOS BD
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota llamado desde el Click Aceptar
Dim sSQL As String
Dim modTrasladoOrigen As New clsBD3
Dim modTrasladoDestino As New clsBD3

'Verifica si es a Caja el origen del movimiento
If msCajaoBancoOrigen = "CA" Then
     ' Guardar los  datos
      sSQL = "UPDATE EGRESOS SET " & _
         "NumDoc='" & txtDocOrigen & "'," & _
         "IdTipoDoc='" & txtTipoDocOrigen & "'," & _
         "MontoCB=" & Var37(txtMonto.Text) & "," & _
         "Observ='" & txtObservacion.Text & "' " & _
         "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenOrigen.Text & "'"
            
    ' Si al ejecutar hay error se sale de la aplicación
    modTrasladoOrigen.SQL = sSQL
    If modTrasladoOrigen.Ejecutar = HAY_ERROR Then
     End
    End If
    
    ' Se cierra la query
    modTrasladoOrigen.Cerrar
        
    'Verifica si es caja el destino del movimiento
    If msCajaoBancoDestino = "CA" Then
       
        ' Guardar los  datos
        sSQL = "UPDATE INGRESOS SET " & _
           "NumDoc='" & txtDocDestino & "'," & _
           "IdTipoDoc='" & txtTipoDocDestino & "'," & _
           "Monto=" & Var37(txtMonto.Text) & "," & _
           "Observ='" & txtObservacion.Text & "' " & _
           "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenDestino.Text & "'"
          
        ' Carga la colección asiento
        ' OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona , Monto, fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯Nulo¯" & "Nulo¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TC"
              
    Else
        
        ' Guardar los  datos
        sSQL = "UPDATE INGRESOS SET " & _
           "NumDoc='" & txtDocDestino.Text & "'," & _
           "IdTipoDoc='" & txtTipoDocDestino & "'," & _
           "Monto=" & Var37(txtMonto.Text) & "," & _
           "Observ='" & txtObservacion.Text & "'," & _
           "IdCta='" & msCtaCteDestino & "' " & _
           "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenDestino.Text & "'"

         'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯Nulo¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TB"
                         
    End If
    
    'SI al ejecutar hay error se sale de la aplicación
    modTrasladoDestino.SQL = sSQL
    If modTrasladoDestino.Ejecutar = HAY_ERROR Then
     End
    End If
       
    'Se cierra la query
    modTrasladoDestino.Cerrar

Else    'El origen del movimiento es Banco
    sSQL = ""
    sSQL = "UPDATE EGRESOS SET " & _
       "NumDoc='" & txtDocOrigen.Text & "'," & _
       "IdTipoDoc='" & txtTipoDocOrigen & "'," & _
       "MontoCB=" & Var37(txtMonto.Text) & "," & _
       "Observ='" & txtObservacion.Text & "'," & _
       "IdCta='" & msCtaCteOrigen & "'," & _
       "NumCheque='" & txtNumCh.Text & "' " & _
       "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenOrigen.Text & "'"
      
    ' Si al ejecutar hay error se sale de la aplicación
    modTrasladoOrigen.SQL = sSQL
    If modTrasladoOrigen.Ejecutar = HAY_ERROR Then
     End
    End If
    
    ' Se cierra la query
    modTrasladoOrigen.Cerrar

    'Verifica si el destino del movimiento es caja
    If msCajaoBancoDestino = "CA" Then
          
        ' Guardar los  datos
        sSQL = "UPDATE INGRESOS SET " & _
           "NumDoc='" & txtDocDestino & "'," & _
           "IdTipoDoc='" & txtTipoDocDestino & "'," & _
           "Monto=" & Var37(txtMonto.Text) & "," & _
           "Observ='" & txtObservacion.Text & "' " & _
           "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenDestino.Text & "'"

         'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯" & msCtaCteOrigen & "¯Nulo¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TC"
    Else
        
        ' Guardar los  datos
        sSQL = "UPDATE INGRESOS SET " & _
           "NumDoc='" & txtDocDestino.Text & "'," & _
           "IdTipoDoc='" & txtTipoDocDestino & "'," & _
           "Monto=" & Var37(txtMonto.Text) & "," & _
           "Observ='" & txtObservacion.Text & "'," & _
           "IdCta='" & msCtaCteDestino & "' " & _
           "WHERE CodMov='" & txtCodMov.Text & "' And Orden='" & txtOrdenDestino.Text & "'"
               
        'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯" & msCtaCteOrigen & "¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
         & "¯" & cboMovimiento.Text & "¯TR¯TB"
                 
    End If
    
    'SI al ejecutar hay error se sale de la aplicación
    modTrasladoDestino.SQL = sSQL
    If modTrasladoDestino.Ejecutar = HAY_ERROR Then
     End
    End If
       
    'Se cierra la query
    modTrasladoDestino.Cerrar
           
End If
  
' Realiza el asiento automático
Conta20

End Sub

Private Function CalcularTotalIngresos() As Double
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
If Left(txtOrdenDestino.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='CA' And I.Anulado='NO'"
Else
    'El total que se esta calculande es de banco de la cuenta msCtaCte
    sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='BA' And I.IdCta= '" & msCtaCteDestino & "' And I.Anulado='NO'"

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
    CalcularTotalIngresos = curTotalMonto.campo(0)
End If

'Cierra el cursor
curTotalMonto.Cerrar

End Function


Private Function CalcularTotalEgresos() As Double
'-----------------------------------------------------
'Propósito  : Determina la suma de montos de los egresos
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim sSQL As String
Dim curTotalMonto As New clsBD2

'Sentencia SQL
sSQL = ""
'Verifica si se esta calculando el total de caja
If Left(txtOrdenDestino.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
           & "FROM EGRESOS E " _
           & "WHERE Left(E.Orden,2)='CA' And E.Anulado='NO' And E.Origen='C'"
Else
    'El total que se esta calculande es de banco de la cuenta msCtaCte
    sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
           & "FROM EGRESOS E " _
           & "WHERE Left(E.Orden,2)='BA' And E.IdCta= '" & msCtaCteDestino & "' And E.Anulado='NO' And E.Origen='B'"

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
    CalcularTotalEgresos = curTotalMonto.campo(0)
End If

'Cierra el cursor
curTotalMonto.Cerrar

End Function

Private Sub GuardarTraslado()
'----------------------------------------------------------------------------
'Propósito  : Guarda los traslados en EGRESOS  e INGRESOS BD
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota llamado desde el Click Aceptar
Dim sSQL As String
Dim modTrasladoOrigen As New clsBD3
Dim modTrasladoDestino As New clsBD3

'Verifica si es a Caja
If msCajaoBancoOrigen = "CA" Then
        ' Guardar los  datos A Caja
    sSQL = "INSERT INTO EGRESOS VALUES('" & txtOrdenOrigen.Text & "','','','','','" _
            & txtDocOrigen.Text & "','" & txtTipoDocOrigen.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecTrab.Text) & "',''," _
            & Var37(txtMonto.Text) & ",'','','','NO','" & txtObservacion.Text & "','','NO','C','')"
            
    ' Si al ejecutar hay error se sale de la aplicación
    modTrasladoOrigen.SQL = sSQL
    If modTrasladoOrigen.Ejecutar = HAY_ERROR Then
     End
    End If
    
    ' Se cierra la query
    modTrasladoOrigen.Cerrar
        
    'Verifica si el destino es caja o banco
    If msCajaoBancoDestino = "CA" Then
       
        ' Guardar los  datos a Caja cuando no es ingreso de prestamos
        sSQL = ""
        sSQL = "INSERT INTO INGRESOS VALUES('" & txtOrdenDestino & "','" _
                & txtDocDestino.Text & "','" & txtTipoDocDestino.Text & "','" _
                & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
                & Var37(txtMonto.Text) & ",'','','NO','" & txtObservacion.Text & "','')"
        
        ' Carga la colección asiento
        ' OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona , Monto, fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯Nulo¯" & "Nulo¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TC"
              
    Else
        
        'Se graba en Banco, cuando no es un ingreso de prestamos
        sSQL = ""
        sSQL = "INSERT INTO INGRESOS VALUES('" & txtOrdenDestino & "','" _
                & txtDocDestino.Text & "','" & txtTipoDocDestino.Text & "','" _
                & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
                & Var37(txtMonto.Text) & ",'','" _
                & msCtaCteDestino & "','NO','" & txtObservacion.Text & "','')"

         'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯Nulo¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TC"
                         
    End If
    
    'SI al ejecutar hay error se sale de la aplicación
    modTrasladoDestino.SQL = sSQL
    If modTrasladoDestino.Ejecutar = HAY_ERROR Then
     End
    End If
       
    'Se cierra la query
    modTrasladoDestino.Cerrar

Else    'Se graba en Banco
    sSQL = ""
    sSQL = "INSERT INTO EGRESOS VALUES('" & txtOrdenOrigen.Text & "','','','','','" _
            & txtDocOrigen.Text & "','" & txtTipoDocOrigen.Text & "','" _
            & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecTrab.Text) & "',''," _
            & Var37(txtMonto.Text) & ",'','" _
            & msCtaCteOrigen & "','" & txtNumCh.Text & "','NO','" & txtObservacion.Text & "','','NO','B','')"
       
    ' Si al ejecutar hay error se sale de la aplicación
    modTrasladoOrigen.SQL = sSQL
    If modTrasladoOrigen.Ejecutar = HAY_ERROR Then
     End
    End If
    
    ' Se cierra la query
    modTrasladoOrigen.Cerrar

    'Verifica SI es a Caja o banco el destino
    If msCajaoBancoDestino = "CA" Then
          
        ' Guardar los  datos a Caja cuando no es ingreso de prestamos
        sSQL = "INSERT INTO INGRESOS VALUES('" & txtOrdenDestino & "','" _
                & txtDocDestino.Text & "','" & txtTipoDocDestino.Text & "','" _
                & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
                & Var37(txtMonto.Text) & ",'','','NO','" & txtObservacion.Text & "','')"

         'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯" & msCtaCteOrigen & "¯Nulo¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & cboMovimiento.Text & "¯TR¯TB"
    Else
        
        'Se graba en Banco, cuando no es un ingreso de prestamos
        sSQL = "INSERT INTO INGRESOS VALUES('" & txtOrdenDestino & "','" _
                & txtDocDestino.Text & "','" & txtTipoDocDestino.Text & "','" _
                & txtCodMov.Text & "','" & FechaAMD(mskFecTrab.Text) & "'," _
                & Var37(txtMonto.Text) & ",'','" _
                & msCtaCteDestino & "','NO','" & txtObservacion.Text & "','')"
            
        'Carga la colección asiento
         'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
          idPersona,Monto,fecha, observ, Proceso
         gcolAsiento.Add _
         Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
         Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
           & "¯" & msCtaCteOrigen & "¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
           & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
         & "¯" & cboMovimiento.Text & "¯TR¯TB"
                 
    End If
    
    'SI al ejecutar hay error se sale de la aplicación
    modTrasladoDestino.SQL = sSQL
    If modTrasladoDestino.Ejecutar = HAY_ERROR Then
     End
    End If
       
    'Se cierra la query
    modTrasladoDestino.Cerrar
           
End If
  
' Realiza el asiento automático
Conta12

End Sub

Private Function CalcularTotalEgresosCB() As Double
'-----------------------------------------------------
'Propósito  : Determina la suma de montos de los Egresos
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------
Dim sSQL As String
Dim curTotalMonto As New clsBD2

'Sentencia SQL
sSQL = ""

'Verifica si el ingreso que se realiza es de caja
If Left(txtOrdenDestino.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
           & "FROM EGRESOS E " _
           & "WHERE Left(E.Orden,2)='CA' and E.Anulado='NO' And E.Origen='C'"
       
Else
    If mcurRegTrasladosDestino.campo(4) = msCtaCteDestino Then ' misma cuenta
     'El egreso que se calcula es de banco de la cta msCtaCte
        sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
        & "FROM EGRESOS E " _
        & "WHERE Left(E.Orden,2)='BA' And E.IdCta='" & msCtaCteDestino & "' And E.Anulado='NO'And E.Origen='B'"
    Else ' cuentas diferentes
    'El egreso que se calcula es de banco de la cta original
        sSQL = "SELECT SUM(E.MontoCB) as MontoTotal " _
        & "FROM EGRESOS E " _
        & "WHERE Left(E.Orden,2)='BA' And E.IdCta='" & mcurRegTrasladosDestino.campo(4) & "' And E.Anulado='NO'And E.Origen='B'"
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

Private Function fbVerificarDatosIntroducidosIngreso()
' -------------------------------------------------------
' Propósito: Verifica que los datos introducidos sean correctos
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------
Dim dblMontoDet As Double
' Verifica que lo que sale de caja-bancos sea Menor que el saldo de Caja-bancos
 If gsTipoOperacionTraslado = "Modificar" Then
 
    ' verifica la conformidad con el saldo
    If msCajaoBancoDestino = "CA" Then   'Caja
    
        If Val(Var37(txtSaldoDestino.Text)) < (Val(mcurRegTrasladosOrigen.campo(2)) - Val(Var37(txtMonto))) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Traslados Caja Bancos"
            fbVerificarDatosIntroducidosIngreso = False
            txtMonto.SetFocus
            Exit Function
        End If
    ElseIf (msCajaoBancoDestino = "BA" And (mcurRegTrasladosDestino.campo(4) = msCtaCteDestino)) Then 'La misma CtaCte
            
            If Val(Var37(txtSaldoDestino.Text)) < (Val(mcurRegTrasladosOrigen.campo(2)) - Val(Var37(txtMonto.Text))) Then
                ' Mensaje ,saldo insuficiente
                 MsgBox "El monto de egreso excede al ingreso en el destino de la Transferencia", , "SGCcaijo-Modificación de Traslados"
                 fbVerificarDatosIntroducidosIngreso = False
                 txtMonto.SetFocus
                Exit Function
            End If
    ElseIf (msCajaoBancoDestino = "BA" And (mcurRegTrasladosDestino.campo(4) <> msCtaCteDestino)) Then 'Se Cambio de CtaCte
    
        If CalcularTotalIngresosCB - Val(mcurRegTrasladosOrigen.campo(2)) < CalcularTotalEgresosCB() Then
            ' Mensaje ,saldo insuficiente
            MsgBox "No se puede cambiar de Cta.Cte. en el destino de la transferencia, el egreso es Mayor al ingreso en la Cta.Cte. original. ", vbExclamation + vbOKOnly, "SGCcaijo-Modificación de Traslados"
            fbVerificarDatosIntroducidosIngreso = False
            txtMonto.SetFocus
          Exit Function
        End If
    
    End If
End If

' Verificados los datos
fbVerificarDatosIntroducidosIngreso = True

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
If Left(txtOrdenDestino.Text, 2) = "CA" Then
    sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='CA' And I.Anulado='NO'"
Else
    If mcurRegTrasladosDestino.campo(4) = msCtaCteDestino Then ' la cuenta es la misma
        'El total que se calcula es de banco de la cuenta msCtaCte
        sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='BA' And I.IdCta= '" & msCtaCteDestino & "' And I.Anulado='NO'"
    Else ' la cuenta es Plan28
        'El total que se calcula es de banco de la cuenta original
        sSQL = "SELECT SUM(I.Monto) as MontoTotal " _
           & "FROM INGRESOS I " _
           & "WHERE Left(I.Orden,2)='BA' And I.IdCta= '" & mcurRegTrasladosDestino.campo(4) & "' And I.Anulado='NO'"
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

Private Function fbVerificarDatosIntroducidosEgreso()
' -------------------------------------------------------
' Propósito: Verifica que los datos introducidos sean correctos
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------
Dim dblMontoDet As Double

' Verifica que lo que sale de caja-bancos sea Menor que el saldo de Caja-bancos
 If gsTipoOperacionTraslado = "Nuevo" Then
    'Verifica que sea Egreso
    If optBancoOrigen.Value Or optCajaOrigen.Value Then
        
        'Verifica que el monto ingresado es Mayor al saldo
        If Val(Var37(txtMonto.Text)) > Val(Var37(txtSaldoOrigen.Text)) Then
              ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo CAJA - BANCOS"
            fbVerificarDatosIntroducidosEgreso = False
            txtMonto.SetFocus
            Exit Function
         End If
    End If
 Else
    ' verifica la conformidad con el saldo
    If msCajaoBancoOrigen = "CA" Then   'Caja

        If Val(Var37(txtSaldoOrigen.Text)) < (Val(Var37(txtMonto.Text)) - Val(mcurRegTrasladosOrigen.campo(2))) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Traslados Caja Bancos"
            fbVerificarDatosIntroducidosEgreso = False
            If txtMonto.Enabled = True Then txtMonto.SetFocus
            Exit Function
        End If
    ElseIf (msCajaoBancoOrigen = "BA" And (mcurRegTrasladosOrigen.campo(9) = msCtaCteOrigen)) Then 'La misma CtaCte
            If Val(Var37(txtSaldoOrigen.Text)) < (Val(Var37(txtMonto.Text)) - Val(mcurRegTrasladosOrigen.campo(2))) Then
                ' Mensaje ,saldo insuficiente
                 MsgBox "El monto de egreso excede el saldo del origen de la Transferencia", , "SGCcaijo-Modificación de Traslados"
                 fbVerificarDatosIntroducidosEgreso = False
                 If txtMonto.Enabled = True Then txtMonto.SetFocus
                Exit Function
            End If
    ElseIf (msCajaoBancoOrigen = "BA" And (mcurRegTrasladosOrigen.campo(9) <> msCtaCteOrigen)) Then 'Se Cambio de CtaCte

        If Val(Var37(txtMonto.Text)) > Val(Var37(txtSaldoOrigen.Text)) Then
            ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Modificación de Traslados"
            fbVerificarDatosIntroducidosEgreso = False
            If txtMonto.Enabled = True Then txtMonto.SetFocus
          Exit Function
        End If

    End If
   
 End If

' Verificados los datos
fbVerificarDatosIntroducidosEgreso = True

End Function

Private Function VerificarDocExisteEgreso() As Boolean
'--------------------------------------------------------------------
'Propósito  : Verifica si el Doc ha sido ingresado en caja o bancos, SI NO
'Recibe     : Nada
'Devuelve   : False:NO existe, True: Existe
'Nota       : Llamado desde el evento click de Aceptar
'--------------------------------------------------------------------
Dim sSQL As String
Dim curDocIngresado As New clsBD2

VerificarDocExisteEgreso = False

'Verifica SI el doc ingresado sea el mismo del registro en modificacion
If gsTipoOperacionTraslado = "Modificar" Then
    If txtDocOrigen.Text = mcurRegTrasladosOrigen.campo(5) Then ' es el mismo del registro, NO hace nada
        Exit Function 'Sale de la funcion
    End If
End If

'Verifica SI el Doc esta en Caja o en Banco de la tabla Egresos
'Se averigua SI existe algun documento con el mismo numero en Banco
sSQL = "SELECT Count(E.NumDoc) as NroDoc FROM EGRESOS E " & _
       "WHERE E.NumDoc = '" & txtDocOrigen.Text & "'"
curDocIngresado.SQL = sSQL
If curDocIngresado.Abrir = HAY_ERROR Then
  End
End If
'Se encontró en Banco
If curDocIngresado.campo(0) <> 0 Then
    curDocIngresado.Cerrar
    VerificarDocExisteEgreso = True
    Exit Function
    
End If
'Se cierra el cursor
curDocIngresado.Cerrar

End Function

Private Function VerificarDocExisteIngreso() As Boolean
'--------------------------------------------------------------------
'Propósito  : Verifica si el Doc ha sido ingresado en caja o bancos, SI NO
'Recibe     : Nada
'Devuelve   : False:NO existe, True: Existe
'Nota       : Llamado desde el evento click de Aceptar
'--------------------------------------------------------------------
Dim sSQL As String
Dim curDocIngresado As New clsBD2

VerificarDocExisteIngreso = False

'Verifica SI el doc ingresado sea el mismo del registro en modificacion
If gsTipoOperacionTraslado = "Modificar" Then
    If txtDocDestino.Text = mcurRegTrasladosDestino.campo(0) Then ' es el mismo del registro, NO hace nada
        Exit Function 'Sale de la funcion
    End If
End If

'Verifica SI el Doc esta en Caja o en Banco de la tabla Egresos
'Se averigua SI existe algun documento con el mismo numero en Banco
sSQL = "SELECT Count(I.NumDoc) as NroDoc FROM INGRESOS I " & _
       "WHERE I.NumDoc = '" & txtDocDestino.Text & "'"
curDocIngresado.SQL = sSQL
If curDocIngresado.Abrir = HAY_ERROR Then
  End
End If
'Se encontró en Banco
If curDocIngresado.campo(0) <> 0 Then
    curDocIngresado.Cerrar
    VerificarDocExisteIngreso = True
    Exit Function
    
End If
'Se cierra el cursor
curDocIngresado.Cerrar

End Function

Private Sub cmdAnular_Click()
Dim modAnularTrasladoOrigen As New clsBD3
Dim modAnularTrasladoDestino As New clsBD3
Dim sSQL As String

'Actualiza variable de operación
HabilitaDeshabilitaBotones ("Anular")

'Verifica si el año esta cerrado
If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    Exit Sub
End If

'Valida el ingreso de caja con egreso de caja
If CalcularTotalIngresos - mcurRegTrasladosOrigen.campo(2) < CalcularTotalEgresos Then
    'Mensaje que no se puede realizar la anulación del registro
    MsgBox "¡No se puede anular el registro de traslados!, egresos Mayor que ingresos", _
    vbExclamation + vbOKOnly, "Caja-Bancos-Anulación de traslados"
    Exit Sub
Else

    'Preguntar SI desea Anular el registro de Ingreso a Banco
    'Mensaje de conformidad de los datos
    If MsgBox("¿Seguro que desea anular el registro de traslado?", _
                  vbQuestion + vbYesNo, "Caja-Bancos-Anulación de Traslados") = vbYes Then
        'Actualiza la transaccion
        Var8 1, gsFormulario
             
        'Verifica si el movimiento es con caja o banco
        If msCajaoBancoOrigen = "BA" Then
        
              'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
               sSQL = "UPDATE EGRESOS E SET " & _
                    "E.Anulado='SI'" & _
                    "WHERE E.Orden='" & txtOrdenOrigen.Text & "'"
                
                'SI al ejecutar hay error se sale de la aplicación
                modAnularTrasladoOrigen.SQL = sSQL
                If modAnularTrasladoOrigen.Ejecutar = HAY_ERROR Then
                 End
                End If
                
                'Cierra el modAnularTraslado
                modAnularTrasladoOrigen.Cerrar
                    
               If msCajaoBancoDestino = "BA" Then
                    'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
                    sSQL = "UPDATE INGRESOS I SET " & _
                         "I.Anulado='SI'" & _
                         "WHERE I.Orden='" & txtOrdenDestino.Text & "'"
                         
                   'Carga la colección asiento
                   'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
                    idPersona,Monto,fecha, observ, Proceso
                   gcolAsiento.Add _
                   Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
                   Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
                     & "¯" & msCtaCteOrigen & "¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
                     & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
                   & "¯" & cboMovimiento.Text & "¯TR¯TB"

               Else
                    'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
                    sSQL = "UPDATE INGRESOS I SET " & _
                         "I.Anulado='SI'" & _
                         "WHERE I.Orden='" & txtOrdenDestino.Text & "'"
                         
                    'Carga la colección asiento
                    'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
                     idPersona,Monto,fecha, observ, Proceso
                    gcolAsiento.Add _
                    Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
                    Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
                      & "¯" & msCtaCteOrigen & "¯Nulo¯" & txtCodPersonal.Text & "¯" _
                      & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
                      & "¯" & cboMovimiento.Text & "¯TR¯TB"
               End If
               
                                           'SI al ejecutar hay error se sale de la aplicación
                modAnularTrasladoDestino.SQL = sSQL
                If modAnularTrasladoDestino.Ejecutar = HAY_ERROR Then
                 End
                End If
                
                'Cierra el modAnularTraslado
                modAnularTrasladoDestino.Cerrar
               
        Else
            'Cambiar el campo Anulado de Ingresos a "SI", los demás campos a anulado y cero
             sSQL = "UPDATE EGRESOS E SET " & _
                "E.Anulado='SI'" & _
                "WHERE E.Orden='" & txtOrdenOrigen.Text & "'"
                
            'SI al ejecutar hay error se sale de la aplicación
            modAnularTrasladoOrigen.SQL = sSQL
            If modAnularTrasladoOrigen.Ejecutar = HAY_ERROR Then
             End
            End If
            
            'Cierra el modAnularTraslado
            modAnularTrasladoOrigen.Cerrar
         
            If msCajaoBancoDestino = "BA" Then
                    'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
                    sSQL = "UPDATE INGRESOS I SET " & _
                         "I.Anulado='SI'" & _
                         "WHERE I.Orden='" & txtOrdenDestino.Text & "'"
                         
                    'Carga la colección asiento
                    'OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
                     idPersona,Monto,fecha, observ, Proceso
                    gcolAsiento.Add _
                    Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
                    Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
                      & "¯Nulo¯" & msCtaCteDestino & "¯" & txtCodPersonal.Text & "¯" _
                      & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
                      & "¯" & cboMovimiento.Text & "¯TR¯TC"
                      
           Else
                'Cambiar el campo Anulado de Ingresos a "SI, Los demas campos a anulado y cero"
                sSQL = "UPDATE INGRESOS I SET " & _
                     "I.Anulado='SI'" & _
                     "WHERE I.Orden='" & txtOrdenDestino.Text & "'"
                     
                ' Carga la colección asiento
                ' OrdenOrigen,OrdenDestino,NumCtaBancOrigen,NumCtaBancDestino, _
                  idPersona , Monto, fecha, observ, Proceso
                 gcolAsiento.Add _
                 Key:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text, _
                 Item:=txtOrdenOrigen.Text & "¯" & txtOrdenDestino.Text _
                   & "¯Nulo¯" & "Nulo¯" & txtCodPersonal.Text & "¯" _
                   & Var37(txtMonto) & "¯" & FechaAMD(mskFecTrab.Text) _
                   & "¯" & cboMovimiento.Text & "¯TR¯TC"


           End If
           
            'SI al ejecutar hay error se sale de la aplicación
            modAnularTrasladoDestino.SQL = sSQL
            If modAnularTrasladoDestino.Ejecutar = HAY_ERROR Then
                End
            End If
            
            'Cierra el modAnularTraslados
            modAnularTrasladoDestino.Cerrar
                                 
        End If
    Else
        'Termina la ejecución del procedimiento
        Exit Sub
    End If
End If 'Fin de CalcularTotalIngresos() - curRegIngresoCajaBanco.campo(4) < CalcularTotalEgresos()

'Anula el asiento contable
Conta23

'Cierra los cursores
mcurRegTrasladosDestino.Cerrar
mcurRegTrasladosOrigen.Cerrar

'Actualiza la transaccion
Var8 -1, gsFormulario

'Mensaje
MsgBox "Operación realizada correctamente", vbOKOnly + vbInformation, "SGCcaijo-Traslados"

'Descarga el formulario
Unload Me

End Sub

Private Sub cmdBuscar_Click()

' Carga los títulos del grid selección
  giNroColMNSel = 4
  aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Condición", "Activo")
  aTamañosColumnas = Array(1000, 4500, 1500, 600)
' Muestra el formulario de busqueda
  frmMNSeleccion.Show vbModal, Me

' Verifica si se eligió algun dato a modificar
  If gsCodigoMant <> Empty Then
    txtCodPersonal.Text = gsCodigoMant
    SendKeys "{tab}"
  Else ' No se eligió nada a modificar
    ' Verifica si txtcodigo es habilitado
    If txtCodPersonal.Enabled = True Then txtCodPersonal.SetFocus
  End If

End Sub

Private Sub cmdCancelar_Click()

' Verifica el tipo operación
If gsTipoOperacionTraslado = "Nuevo" Then

    ' Limpia el formulario y pone en blanco variables
    LimpiarFormulario

    'Prepara el formulario
    NuevoTraslado
    
Else

    'Coloca los datos originales de origen
    CargarControlesOrigen
    
    'Carga los controles originales de destino
    CargarControlesDestino
    
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
cboMovimiento.ListIndex = -1
txtCodPersonal.Text = Empty
txtDesc.Text = Empty
txtDocOrigen.Text = Empty
txtDocDestino.Text = Empty
txtTipoDocOrigen.Text = Empty
cboTipDocOrigen.ListIndex = -1
txtTipoDocDestino.Text = Empty
cboTipoDocDestino.ListIndex = -1
txtMonto.Text = Empty
txtObservacion.Text = Empty
txtSaldoDestino.Text = Empty

'Limpia controles banco
txtBancoOrigen.Text = Empty
txtNumCh.Text = Empty
txtBancoDestino.Text = Empty
'Habilita los opts
HabilitarOps

'Deshabilita los fra
fraEgreso.Enabled = False
fraIngreso.Enabled = False

End Sub

Private Sub cmdPBancoDestino_Click()
If cboBancoDestino.Enabled Then
    ' alto
     cboBancoDestino.Height = CBOALTO
    ' focus a cbo
    cboBancoDestino.SetFocus
End If
End Sub

Private Sub cmdPBancoOrigen_Click()

If cboBancoOrigen.Enabled Then
    ' alto
     cboBancoOrigen.Height = CBOALTO
    ' focus a cbo
    cboBancoOrigen.SetFocus
End If

End Sub


Private Sub cmdPCtaCteDestino_Click()

If cboCtaCteDestino.Enabled Then
    ' alto
     cboCtaCteDestino.Height = CBOALTO
    ' focus a cbo
    cboCtaCteDestino.SetFocus
End If

End Sub

Private Sub cmdPCtaCteOrigen_Click()

If cboCtaCteOrigen.Enabled Then
    ' alto
     cboCtaCteOrigen.Height = CBOALTO
    ' focus a cbo
    cboCtaCteOrigen.SetFocus
End If

End Sub

Private Sub cmdpTipDoc_Click()

If cboTipDocOrigen.Enabled Then
    ' alto
     cboTipDocOrigen.Height = CBOALTO
    ' focus a cbo
    cboTipDocOrigen.SetFocus
End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdPTipoDocDestino_Click()

If cboTipoDocDestino.Enabled Then
    ' alto
     cboTipoDocDestino.Height = CBOALTO
    ' focus a cbo
     cboTipoDocDestino.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()

If gsTipoOperacionTraslado = "Modificar" Then

    'Cierra los cursores
    mcurRegTrasladosDestino.Cerrar
    mcurRegTrasladosOrigen.Cerrar
    
End If

'Termina la ejecución
Unload Me

End Sub

Private Sub Form_Load()

'Carga los moviemientos de traslados
CargarMovTraslados

'Cargar tipo de documentos
CargarTipoDocOrigen

'Cargar tipo de documentos
CargarTipoDocDestino

'Carga el personal
Conta35

'Carga los bancos
CargarBancosOrigen

'Cargar Cuentas ctes origen
CargarCtaCtes

'Limpia el cboCtaCte
cboCtaCteOrigen.Clear

'Carga el cboBancoDestino
CargarCboCols cboBancoDestino, mcolCodDesBanco

'Establece campos obligatorios del formulario
EstablecerCamposObligatorios

'Deshabilita los Controles
DesHabilitarControles

'Deshabilita los fra
fraIngreso.Enabled = False
fraEgreso.Enabled = False
    
' Verifica el tipo de operación a realizar en el formulario
If gsTipoOperacionTraslado = "Modificar" Then
    
    'Coloca titulo al formulario
    Me.Caption = "Caja y Bancos - Modificación de Traslados"
    
Else 'Nuevo Traslado
    Me.Caption = "Caja y Bancos- Traslados"
    
   'Deshabilita txtCodIngreso
    txtOrdenOrigen.Enabled = False
    txtOrdenDestino.Enabled = False
    
    'Nuevo ingreso
    NuevoTraslado
    
End If

End Sub

Private Sub CargarCtaCtes()
'Carga las cuentas corrientes
Dim sSQL As String
'Se carga el combo de Cta Cte
sSQL = ""
sSQL = "SELECT IdCta, DescCta FROM TIPO_CUENTASBANC " & _
           "WHERE IdMoneda= 'SOL'   ORDER BY DescCta"
CD_CargarColsCbo cboCtaCteOrigen, sSQL, mcolCodCtaCte, mcolCodDesCtaCte

End Sub

Private Sub NuevoTraslado()
'--------------------------------------------------------------
'Propósito : Realiza la operación de Ingreso a Caja o Bancos
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
' Limpia el txtOrden
txtOrdenOrigen.Text = Empty
txtOrdenDestino.Text = Empty

'Inicializa las variable de modulo
msCtaCteOrigen = Empty
msCtaCteDestino = Empty
txtSaldoOrigen.Text = "0.00"

'Coloca la fecha del sistema
mskFecTrab.Text = gsFecTrabajo

'Deshabilita los botones del formulario
HabilitaDeshabilitaBotones ("Nuevo")

End Sub

Private Sub ModificarTraslado()
'---------------------------------------------------------------
'Propósito  : Realiza la operación de modificar en el formulario
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------

' Limpia el txtOrden
  txtOrdenOrigen.Text = Empty
  txtOrdenDestino.Text = Empty
    
' Inicializa la variable codigo de cuenta
  msCtaCteOrigen = Empty
  msCtaCteDestino = Empty

  
' Maneja estado de los botones del formulario
  HabilitaDeshabilitaBotones "Modificar"

End Sub

Private Sub HabilitaDeshabilitaBotones(sProceso As String)
'-----------------------------------------------------------------
' Proposito: Coloca la condición de los botones en el proceso
' Recibe: Nada
' Entrega: Nada
'-----------------------------------------------------------------
Select Case sProceso

' depende del proceso habilita y deshabilita botones
Case "Nuevo", "Modificar"
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    
End Select

End Sub

Private Sub DesHabilitarControles()
' Inhabilita botones al cargar el formulario
cmdAceptar.Enabled = False
cmdAnular.Enabled = False
End Sub

Private Sub CargarBancosOrigen()
'--------------------------------------------------------
'Propósito  : Carga los bancos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String

'Se carga el combo Bancos (sólo con los bancos que de moneda nacional)
sSQL = ""
sSQL = "SELECT DISTINCT b.IdBanco,b.DescBanco FROM TIPO_BANCOS B , TIPO_CUENTASBANC C" _
       & " WHERE b.idbanco = c.idbanco And c.idmoneda = 'SOL'" _
       & " ORDER BY DescBanco"
CD_CargarColsCbo cboBancoOrigen, sSQL, mcolCodBanco, mcolCodDesBanco

End Sub

Private Sub CargarTipoDocOrigen()
'--------------------------------------------------------
'Propósito  : Carga los documentos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String

'Se carga el combo de Tipo de Documento
sSQL = ""
sSQL = "SELECT idTipoDoc, DescTipoDoc FROM TIPO_DOCUM " & _
           "WHERE RelacProc = 'SS'  ORDER BY DescTipoDoc"
CD_CargarColsCbo cboTipDocOrigen, sSQL, mcolCodTipDocOrigen, mcolCodDesTipDocOrigen

End Sub

Private Sub CargarTipoDocDestino()
'--------------------------------------------------------
'Propósito  : Carga los documentos
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String

'Se carga el combo de Tipo de Documento
sSQL = ""
sSQL = "SELECT idTipoDoc, DescTipoDoc FROM TIPO_DOCUM " & _
           "WHERE RelacProc = 'IN'  ORDER BY DescTipoDoc"
CD_CargarColsCbo cboTipoDocDestino, sSQL, mcolCodTipDocDestino, mcolCodDesTipDocDestino

End Sub

Private Sub Conta35()
'--------------------------------------------------------------
'Propósito  : Carga la coleccion de Tipo_Mov con sus diferentes campos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
If gsTipoOperacionTraslado = "Modificar" Then
    'Sentencia SQL
    sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
          & " PP.Condicion, PP.Activo " _
          & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
          & " WHERE P.IdPersona=PP.IdPersona " _
          & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"

ElseIf gsTipoOperacionTraslado = "Nuevo" Then
    'Sentencia SQL
    sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
          & " PP.Condicion, PP.Activo " _
          & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
          & " WHERE P.IdPersona=PP.IdPersona and PP.Activo='SI' " _
          & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"
End If

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

End Sub

Private Sub EstablecerCamposObligatorios()
'---------------------------------------------
'Propósito  : Coloca a obligatorio los campos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------
 txtDocOrigen.BackColor = Obligatorio
 txtDocDestino.BackColor = Obligatorio
 txtTipoDocOrigen.BackColor = Obligatorio
 txtTipoDocDestino.BackColor = Obligatorio
 txtCodPersonal.BackColor = Obligatorio
 txtMonto.BackColor = Obligatorio
 txtCodMov.BackColor = Obligatorio
 txtBancoOrigen.BackColor = Obligatorio
 txtBancoDestino.BackColor = Obligatorio
 cboCtaCteOrigen.BackColor = Obligatorio
 cboCtaCteDestino.BackColor = Obligatorio
 txtNumCh.BackColor = Obligatorio
 
End Sub

Private Sub CargarMovTraslados()
'--------------------------------------------------------------
'Propósito  : Carga la coleccion de Tipo_Mov con sus diferentes campos
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curProceso As New clsBD2

'la sentencia para cargar el combo y las colecciones de tipo movimiento
sSQL = ""
sSQL = "SELECT TM.IdConCB, TM.DescConCB,PC.Proceso FROM Tipo_MovCB TM, PROCESO_CONCEPTOCB PC " & _
        "WHERE TM.IdConCB LIKE  'T*' And  TM.Afecta='Proceso' And " & _
        "TM.IdConCB= PC.IdConCB ORDER BY DescConCB"
        
'Se carga el combo
CD_CargarColsCbo cboMovimiento, sSQL, mcolCodMov, mcolCodDesCodMov

'Carga la coleccion de Proceso
curProceso.SQL = sSQL
If curProceso.Abrir = HAY_ERROR Then
  End
End If

Do While Not curProceso.EOF
    'Añade a la Colección
    mcolProceso.Add Item:=curProceso.campo(2), Key:=curProceso.campo(0)
    
    ' Se avanza a la siguiente fila del cursor
    curProceso.MoverSiguiente
Loop

'Cierra el cursor de curProceso
curProceso.Cerrar

End Sub

Private Sub CambiaroptCajaBancos()
'-------------------------------------------------------------------
'Propósito : Establece los controles de la primera parte del formulario _
             cuando se cambia de optCaja a optBancos bis
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If gsTipoOperacionTraslado = "Nuevo" Then
    'Incializa las variables
    msCajaoBancoOrigen = Empty
    msCajaoBancoDestino = Empty
    
   'Verifica si se selecciono los opts Origen
   If optCajaOrigen.Value Then
        'Calcula el sigiente orden de Caja y lo muestra en el txtCodIngreso
        txtOrdenOrigen.Text = Var22("CA")
        msCajaoBancoOrigen = "CA"
        
        'Carga los saldos
        CargarSaldoOrigen

   ElseIf optBancoOrigen.Value Then
        'Calcula el sigiente orden de Banco y lo muestra en el txtCodIngreso
        txtOrdenOrigen.Text = Var22("BA")
        msCajaoBancoOrigen = "BA"
        
        'Carga los saldos
        CargarSaldoOrigen

   End If
   
   'Verifica si se selecciono los opts Destino
   If optCajaDestino.Value Then
        'Calcula el sigiente orden de Caja y lo muestra en el txtCodIngreso
        txtOrdenDestino.Text = Var22("CA")
        msCajaoBancoDestino = "CA"
        
        'Carga los saldos
        CargarSaldoDestino
        
   ElseIf optBancoDestino.Value Then
        If optBancoOrigen.Value Then
            'Calcula el sigiente orden de Banco y lo muestra en el txtCodIngreso
            txtOrdenDestino.Text = Left(txtOrdenOrigen.Text, 6) & Format(CStr(CInt(Right(txtOrdenOrigen.Text, 4)) + 1), "000#")
            msCajaoBancoDestino = "BA"
        Else
            'Calcula el sigiente orden de Banco y lo muestra en el txtCodIngreso
            txtOrdenDestino.Text = Var22("BA")
            msCajaoBancoDestino = "BA"
        End If
        
        'Carga los saldos
        CargarSaldoDestino
        
   End If
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
If optCajaOrigen.Value Then
    'Limpia y oculta los controles de Banco
    txtBancoOrigen.Text = Empty
    lblBancoOrigen.Visible = False: txtBancoOrigen.Visible = False: cboBancoOrigen.Visible = False
    lblCtaCteOrigen.Visible = False: cboCtaCteOrigen.Visible = False
    cmdPBancoOrigen.Visible = False: cmdPCtaCteOrigen.Visible = False
    txtNumCh.Visible = False: txtNumCh.Visible = False
    lblCheque.Visible = False
Else
    'Muestra los controles de banco
    lblBancoOrigen.Visible = True: txtBancoOrigen.Visible = True: cboBancoOrigen.Visible = True
    lblCtaCteOrigen.Visible = True: cboCtaCteOrigen.Visible = True
    cmdPBancoOrigen.Visible = True: cmdPCtaCteOrigen.Visible = True
    txtNumCh.Visible = True: txtNumCh.Visible = True
    lblCheque.Visible = True
End If

If optCajaDestino.Value Then
    'Limpia y oculta los controles de Banco
    txtBancoDestino.Text = Empty
    lblBancoDestino.Visible = False: txtBancoDestino.Visible = False: cboBancoDestino.Visible = False
    lblCtaCteDestino.Visible = False: cboCtaCteDestino.Visible = False
    cmdPBancoDestino.Visible = False: cmdPCtaCteDestino.Visible = False

Else
    'Muestra los controles de banco
    lblBancoDestino.Visible = True: txtBancoDestino.Visible = True: cboBancoDestino.Visible = True
    lblCtaCteDestino.Visible = True: cboCtaCteDestino.Visible = True
    cmdPBancoDestino.Visible = True: cmdPCtaCteDestino.Visible = True

End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Destruye las colecciones
Set mcolCodMov = Nothing
Set mcolCodDesCodMov = Nothing
Set mcolCodBanco = Nothing
Set mcolCodDesBanco = Nothing
Set mcolCodCtaCte = Nothing
Set mcolCodDesCtaCte = Nothing
Set mcolCodTipDocOrigen = Nothing
Set mcolCodDesTipDocOrigen = Nothing
Set mcolCodTipDocDestino = Nothing
Set mcolCodDesTipDocDestino = Nothing
Set mcolProceso = Nothing
Set gcolTabla = Nothing

End Sub

Private Sub CargarSaldoOrigen()
'-------------------------------------------------------------------------
'Propósito  : Cargar el saldo de Caja o Bancos hasta el momento dehacer la operación.
'Recibe     : Nada
'Devuelve   : Nada
'-------------------------------------------------------------------------

Dim sSQL As String
Dim curCuentas As New clsBD2
Dim curIngreso As New clsBD2
Dim curEgreso As New clsBD2

Dim iCol As Integer
Dim curEmpresas As New clsBD2
Dim EmpresasExistentes As String
Dim InstrucEmpresas As String
Dim TotalEgresoProyectos As Double
Dim TotalEgresoEmpresasSinRH As Double
Dim TotalEgresoEmpresasSoloRHCB As Double
Dim TotalEgresos As Double
Dim TotalIngresos As Double


Dim curSaldo As New clsBD2
Dim dblSaldo As Double

Dim iSaldo As String

dblSaldo = 0 'Inicializa variable saldo
If optBancoOrigen.Value = True Then
  If msCtaCteOrigen <> Empty Then ' Si se eligió la CtaCte
    'Averigua el ingreso de la Cta
    sSQL = "SELECT SUM(Monto) as Ingreso FROM INGRESOS " _
          & "WHERE IdCta='" & msCtaCteOrigen & "' and Anulado='NO'"
          
      curSaldo.SQL = sSQL
      If curSaldo.Abrir = HAY_ERROR Then End
      If Not IsNull(curSaldo.campo(0)) Then dblSaldo = curSaldo.campo(0)
      curSaldo.Cerrar

      'Averigua el egreso de la Cta
      sSQL = "SELECT SUM(MontoCB) as Egresos FROM EGRESOS " _
            & "WHERE IdCta='" & msCtaCteOrigen & "' and Anulado='NO'" _
            & "and Origen='B'"

      curSaldo.SQL = sSQL
      If curSaldo.Abrir = HAY_ERROR Then End
      If IsNull(curSaldo.campo(0)) Then
          dblSaldo = dblSaldo
      Else
          'Determina el saldo
          dblSaldo = dblSaldo - curSaldo.campo(0)
      End If
      curSaldo.Cerrar
    End If

      ' Muestra el saldo de la Ctacte o 0.00 si todavía no se eligió
      txtSaldoOrigen.Text = Format(dblSaldo, "###,###,##0.00")

  Else  'Hallar Saldo para Caja
     txtSaldoOrigen.Text = Format(CargarSaldoCaja, "###,###,##0.00")
 End If

End Sub

Private Sub CargarSaldoDestino()
'-------------------------------------------------------------------------
'Propósito  : Cargar el saldo de Caja o Bancos hasta el momento dehacer la operación.
'Recibe     : Nada
'Devuelve   : Nada
'-------------------------------------------------------------------------
Dim curSaldo As New clsBD2
Dim sSQL As String
Dim dblSaldo As Double

dblSaldo = 0 'Inicializa variable saldo
If optBancoDestino.Value = True Then
  If msCtaCteDestino <> Empty Then ' Si se eligió la CtaCte
    'Averigua el ingreso de la Cta
    sSQL = "SELECT SUM(Monto) as Ingreso FROM INGRESOS " _
          & "WHERE IdCta='" & msCtaCteDestino & "' And Anulado='NO'"
          
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If Not IsNull(curSaldo.campo(0)) Then dblSaldo = curSaldo.campo(0)
    curSaldo.Cerrar
    
    'Averigua el egreso de la Cta
    sSQL = "SELECT SUM(MontoCB) as Egresos FROM EGRESOS " _
          & "WHERE IdCta='" & msCtaCteDestino & "' And Anulado='NO'" _
          & "And Origen='B'"
          
    curSaldo.SQL = sSQL
    If curSaldo.Abrir = HAY_ERROR Then End
    If IsNull(curSaldo.campo(0)) Then
        dblSaldo = dblSaldo
    Else
        'Determina el saldo
        dblSaldo = dblSaldo - curSaldo.campo(0)
    End If
    curSaldo.Cerrar
  End If
  
    ' Muestra el saldo de la Ctacte o 0.00 si todavía no se eligió
    txtSaldoDestino.Text = Format(dblSaldo, "###,###,##0.00")

Else
 'Hallar Saldo Caja
    txtSaldoDestino.Text = Format(CargarSaldoCaja, "###,###,##0.00")
   
End If

End Sub

Private Sub optBancoDestino_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optBancoOrigen_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub



Private Sub optCajaDestino_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optCajaOrigen_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtBancoDestino_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtBancoOrigen_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboBancoOrigen, txtBancoOrigen, mcolCodDesBanco

  ' Verifica SI el campo esta vacio
If txtBancoOrigen.Text <> "" And cboBancoOrigen.Text <> "" Then
   ' Los campos coloca a color blanco
   txtBancoOrigen.BackColor = vbWhite
   
   'Actualiza el cboCtaCte con las descripciones de las cuentas relacionadas a txtBanco
    ActualizarListcboCtaCte txtBancoOrigen, cboCtaCteOrigen
Else
   'Marca los campos obligatorios, y limpia el combo
   txtBancoOrigen.BackColor = Obligatorio
   cboCtaCteOrigen.Clear
   cboCtaCteOrigen.BackColor = Obligatorio
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtBancoDestino_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboBancoDestino, txtBancoDestino, mcolCodDesBanco

  ' Verifica SI el campo esta vacio
If txtBancoDestino.Text <> "" And cboBancoDestino.Text <> "" Then
   ' Los campos coloca a color blanco
   txtBancoDestino.BackColor = vbWhite
   
   'Actualiza el cboCtaCte con las descripciones de las cuentas relacionadas a txtBanco
    ActualizarListcboCtaCte txtBancoDestino, cboCtaCteDestino
Else
   'Marca los campos obligatorios, y limpia el combo
   txtBancoDestino.BackColor = Obligatorio
   cboCtaCteDestino.Clear
   cboCtaCteDestino.BackColor = Obligatorio
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

End Sub

Public Sub ActualizarListcboCtaCte(txtBanco As TextBox, cboCtaCte As ComboBox)
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
    If gsTipoOperacionTraslado = "Modificar" Then
      ' VADICK MODIFICACION CARGA TODAS LAS CUENTAS ANULADAS O NO PARA MODIFICACION
      sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
           " WHERE c.idbanco = '" & txtBanco & "' " & _
           " AND  b.idBanco = c.IdBanco" & _
           " AND c.idmoneda= 'SOL' ORDER BY c.DescCta"
    Else
      ' VADICK MODIFICACION SOLO CARGA LAS CUENTAS QUE NO ESTAN ANULADAS PARA NUEVO TRASLADO
      sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
           " WHERE c.idbanco = '" & txtBanco & "' " & _
           " AND  b.idBanco = c.IdBanco" & _
           " AND c.idmoneda= 'SOL' AND C.ANULADO='NO' ORDER BY c.DescCta"
           
      ' VADICK CONSULTA PARA EL CASO DE QUE A UNA CUENTA ANULADA SE LE QUIERA AGREGAR UN MOVIMIENTO ANTES O EN LA FECHA DE ANULACION
      ' CAMBIANDO LA FECHA DE TRABAJO LOCAL
      'sSQL = "SELECT c.desccta FROM TIPO_BANCOS B, TIPO_CUENTASBANC C " & _
           " WHERE c.idbanco = '" & txtBanco & "' " & _
           " AND  b.idBanco = c.IdBanco" & _
           " AND c.idmoneda= 'SOL'" & _
           " AND (C.ANULADO='NO' OR C.FECHAANULADO >= '" & FechaAMD(gsFecTrabajo) & "' OR C.FECHAANULADO = NULL ) " & _
           " ORDER BY c.DescCta "
    End If
    
    curCtaCte.SQL = sSQL
    If curCtaCte.Abrir = HAY_ERROR Then
      End
    End If
      
    'Verifica SI existen cuentas asociadas a txtBanco
    If curCtaCte.EOF Then
      'NO existe cuentas asociadas
      MsgBox "No existen cuentas en el banco seleccionado. Consulte al administrador", _
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

Private Sub txtBancoOrigen_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtCodMov_Change()

    'Verifica si esta en mayusculas
    If UCase(txtCodMov.Text) = txtCodMov.Text Then
    
        ' SI procede, se actualiza descripción correspondiente a código introducido
        CD_ActDesc cboMovimiento, txtCodMov, mcolCodDesCodMov
         
        ' Verifica Si el campo esta vacio
        If txtCodMov.Text <> Empty And cboMovimiento.Text <> Empty Then
            'Vacia el msCtaCteOrigen
            msCtaCteOrigen = Empty
            msCtaCteDestino = Empty
            
            'Los campos coloca a color blanco
            txtCodMov.BackColor = vbWhite
            
            'Habilita los fra
            fraEgreso.Enabled = True
            fraIngreso.Enabled = True
            
            'Habilita los opts
            HabilitarOps
            
            'Determina el proceso
            SeleccionarOpts DeterminarProceso
            
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

Private Sub HabilitarOps()
'-----------------------------------
'Propósito  : Habilita los opts
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------
'Coloca su valor a falso
optCajaOrigen.Value = False
optCajaDestino.Value = False
optBancoOrigen.Value = False
optBancoDestino.Value = False

'Habilita los opts
optCajaOrigen.Enabled = True
optCajaDestino.Enabled = True
optBancoOrigen.Enabled = True
optBancoDestino.Enabled = True

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
    Or txtCodPersonal.BackColor <> vbWhite _
    Or txtMonto.BackColor <> vbWhite _
    Or txtDocOrigen.BackColor <> vbWhite _
    Or txtDocDestino.BackColor <> vbWhite _
    Or txtTipoDocOrigen.BackColor <> vbWhite _
    Or txtTipoDocDestino.BackColor <> vbWhite _
Then
    ' Deshabilita el botón
    cmdAceptar.Enabled = False

    'Termina la ejecución del procedimientos
   Exit Sub
Else
   ' Verifica que se haigan introducido los datos obligatorios de bancos
   If optBancoOrigen.Value = True Then
        If txtBancoOrigen.BackColor <> vbWhite _
            Or cboCtaCteOrigen.BackColor <> vbWhite Or txtNumCh.BackColor <> vbWhite Then
            ' Deshabilita el botón
            cmdAceptar.Enabled = False

            Exit Sub
        End If
   End If
   
   If optBancoDestino.Value Then
        If txtBancoDestino.BackColor <> vbWhite _
            Or cboCtaCteDestino.BackColor <> vbWhite Then
            ' Deshabilita el botón
            cmdAceptar.Enabled = False

            ' Algún obligatorio de banco falta ser introducido
            Exit Sub
        End If
   End If
End If

' Verifica si se cambio algún dato
If gsTipoOperacionTraslado = "Modificar" Then
    If fbCambioDatosGrales = False Then
        ' Deshabilita el botón
        cmdAceptar.Enabled = False

        'Termina la ejecución del procedimiento
        Exit Sub
    End If
End If

' Habilita botón aceptar
cmdAceptar.Enabled = True

End Sub

Private Function fbCambioDatosGrales() As Boolean
' --------------------------------------------------------------
' Propósito : Verifica si se cambió algún dato general del egreso
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------------
' Inicializa la función
fbCambioDatosGrales = False
If txtCodPersonal.Text <> mcurRegTrasladosOrigen.campo(1) _
       Or Val(Var37(txtMonto.Text)) <> mcurRegTrasladosOrigen.campo(2) _
       Or txtTipoDocOrigen.Text <> mcurRegTrasladosOrigen.campo(6) _
       Or txtTipoDocDestino.Text <> mcurRegTrasladosDestino.campo(1) _
       Or txtDocOrigen.Text <> mcurRegTrasladosOrigen.campo(5) _
       Or txtDocDestino.Text <> mcurRegTrasladosDestino.campo(0) _
       Or txtObservacion.Text <> mcurRegTrasladosOrigen.campo(4) _
    Then
        ' cambio datos generales
        fbCambioDatosGrales = True
        Exit Function
    Else
        ' Verifica que se hayan introducido los datos obligatorios de bancos
       If optBancoOrigen.Value = True Then
          If (txtBancoOrigen.Text <> mcurRegTrasladosOrigen.campo(8) _
              Or msCtaCteOrigen <> mcurRegTrasladosOrigen.campo(9) _
              Or txtNumCh.Text <> mcurRegTrasladosOrigen.campo(10)) Then
           
              ' cambio datos generales
            fbCambioDatosGrales = True
            Exit Function
          End If
       End If
       
       'Verifica si se selecciono optBancoDestino
       If optBancoDestino.Value = True Then
          If txtBancoDestino.Text <> mcurRegTrasladosDestino.campo(3) _
             Or msCtaCteDestino <> mcurRegTrasladosDestino.campo(4) Then
             ' cambio datos generales
             fbCambioDatosGrales = True
             Exit Function
           End If
       End If
End If

End Function

Private Sub SeleccionarOpts(strProceso As String)
'--------------------------------------------------------------
'Propósito  : Selecciona los opts correspondientes al movimiento
'Recibe     : strProceso, Proceso relacionado al movimiento
'Devuelve   : Nada
'--------------------------------------------------------------
'Selecciona el valor
Select Case strProceso
Case "CAJA_CTACTE"
    'Selecciona el optCajaOrigen
    optCajaOrigen.Value = True
    optBancoOrigen.Enabled = False
    
    'Selecciona el optBancoDestino
    optBancoDestino.Value = True
    optCajaDestino.Enabled = False
    
    'Limpia los controles
    txtBancoDestino.Text = Empty
    txtBancoDestino.BackColor = Obligatorio
    
    ' Realiza el cambio de opción a Caja
    CambiaroptCajaBancos
      
Case "CTACTE_CAJA"
    'Selecciona el optCajaOrigen
    optBancoOrigen.Value = True
    optCajaOrigen.Enabled = False
    
    'Limpia los controles
    txtBancoOrigen.Text = Empty
    txtBancoOrigen.BackColor = Obligatorio
    
    'Selecciona el optBancoDestino
    optCajaDestino.Value = True
    optBancoDestino.Enabled = False
    
    ' Realiza el cambio de opción a Caja
    CambiaroptCajaBancos

Case "CTACTE_CTACTE"
     'Selecciona el optCajaOrigen
    optBancoOrigen.Value = True
    optCajaOrigen.Enabled = False
    
    'Selecciona el optBancoDestino
    optBancoDestino.Value = True
    optCajaDestino.Enabled = False
    
    'Limpia los controles
    txtBancoOrigen.Text = Empty
    txtBancoOrigen.BackColor = Obligatorio
    txtBancoDestino.Text = Empty
    txtBancoDestino.BackColor = Obligatorio
    
    ' Realiza el cambio de opción a Caja
    CambiaroptCajaBancos
    
End Select

End Sub

Function DeterminarProceso() As String
'--------------------------------------------------------------
'Propósito: Detemina el proceso relacionado con el movimiento
'Recibe:   sCodMov (Código del Movimiento)
'Devuelve: Nada
'--------------------------------------------------------------

'Muestra a que Proceso corresponde
DeterminarProceso = mcolProceso.Item(Trim(txtCodMov.Text))
  
End Function

Private Sub txtCodMov_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtCodPersonal_Change()

If Len(txtCodPersonal.Text) = txtCodPersonal.MaxLength Then
    'Actualiza el txtDesc
    ActualizaDesc
Else
    'Limpia el txtDescAfecta
    txtDesc.Text = Empty
End If

' Verifica Si el campo esta vacio
If txtCodPersonal.Text <> Empty And txtDesc.Text <> Empty Then
    'Los campos coloca a color blanco
    txtCodPersonal.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtCodPersonal.BackColor = Obligatorio
End If

'Habilita el botón aceptar
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
txtDesc.Text = Var30(gcolTabla.Item(txtCodPersonal.Text), 2)

' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe", , "SGCcaijo-Traslados Caja_Banco"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub txtCodPersonal_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtDocDestino_Change()
' Verifica SI el campo esta vacio
If txtDocDestino.Text <> Empty And InStr(txtDocDestino, "'") = 0 Then

    ' El campos coloca a color blanco
    txtDocDestino.BackColor = vbWhite
Else

    'Marca los campos obligatorios
    txtDocDestino.BackColor = Obligatorio
End If

'Habilita Boton Aceptar
HabilitarBotonAceptar
End Sub

Private Sub txtDocDestino_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtDocOrigen_Change()

' Verifica SI el campo esta vacio
If txtDocOrigen.Text <> Empty And InStr(txtDocOrigen, "'") = 0 Then

    ' El campos coloca a color blanco
    txtDocOrigen.BackColor = vbWhite
Else

    'Marca los campos obligatorios
    txtDocOrigen.BackColor = Obligatorio
End If

'Habilita Boton Aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtDocOrigen_KeyPress(KeyAscii As Integer)
' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtMonto_Change()

' Verifica SI el campo esta vacio
If txtMonto.Text <> Empty And Val(txtMonto.Text) <> 0 Then
    ' El campos coloca a color blanco
    txtMonto.BackColor = vbWhite
Else
    'Marca los campos obligatorios
    txtMonto.BackColor = Obligatorio
End If

'habilita Boton Aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtMonto_GotFocus()

'Coloca el tamaño a 12
txtMonto.MaxLength = 12
'Da formato de moneda
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

'Aumenta el tamaño del txtMonto
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
If txtNumCh.Text <> Empty And InStr(txtNumCh, "'") = 0 Then
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
' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtObservacion_Change()

' Si en la observación hay apostrofes vacío
If InStr(txtObservacion, "'") > 0 Then
   txtObservacion = Empty
End If

End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtOrdenDestino_Change()
' Verifica el proceso que se realiza en el formulario
If gsTipoOperacionTraslado = "Modificar" Then
    
    ' Verifica si se ha introducido el tamaño de el código
      If Len(txtOrdenDestino.Text) = txtOrdenDestino.MaxLength Then
      
            ' Verifica si el Ingreso existe y es sin afectación
            If fbCargarTrasladoDestino = True Then

                ' Sale y deshabilita el control
                SendKeys vbTab

                'habilita anular
                cmdAnular.Enabled = True
                cmdAceptar.Enabled = False
                

            End If ' fin de cargar egreso
               
      End If ' fin de verificar el tamaño del texto
 End If
End Sub

Private Sub txtOrdenOrigen_Change()

' Verifica el proceso que se realiza en el formulario
If gsTipoOperacionTraslado = "Modificar" Then
    
    ' Verifica si se ha introducido el tamaño de el código
      If Len(txtOrdenOrigen.Text) = txtOrdenOrigen.MaxLength Then
      
            ' Verifica si el Ingreso existe y es sin afectación
            If fbCargarTrasladoOrigen = True Then
            
                ' Sale y deshabilita el control
                SendKeys vbTab
             
                ' deshabilita el txtcod egreso y el botón buscar, _
                  habilita anular
                cmdAnular.Enabled = True
             
          End If ' fin de cargar egreso
               
      End If ' fin de verificar el tamaño del texto
 End If
 
End Sub

Private Function fbCargarTrasladoDestino() As Boolean
'----------------------------------------------------------------------------
'Propósito  : Carga el registro de ingreso de acuerdo al código en la caja _
              de texto
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota el codigo de Registro de ingreso es CA o BA AAMM9999
Dim sSQL As String

' Verifica SI el ingreso es a Caja o Bancos
msCajaoBancoDestino = Left(txtOrdenDestino.Text, 2)

' Carga la sentencia que consulta a la BD acerca del registo de ingreso en Caja o Bancos
If msCajaoBancoDestino = "CA" Then 'Consulta a Caja
    sSQL = "SELECT I.NumDoc,I.IdTipoDoc, CT.OrdenEgreso " & _
           "FROM INGRESOS I, CTB_TRASLADOCAJABANCOS CT WHERE " & _
           "CT.OrdenIngreso=" & "'" & Trim(txtOrdenDestino.Text) & "' " & _
           "And I.Orden= CT.OrdenIngreso And I.Anulado='NO'"
           
Else 'Consulta a Bancos
    If msCajaoBancoDestino = "BA" Then
            sSQL = "SELECT I.NumDoc,I.IdTipoDoc,CT.OrdenEgreso, CTA.IdBanco, I.IdCta " & _
                   "FROM CTB_TRASLADOCAJABANCOS CT,INGRESOS I, TIPO_CUENTASBANC CTA WHERE " & _
                   "CT.OrdenIngreso=" & "'" & Trim(txtOrdenDestino.Text) & "' And " & _
                   "I.Orden= CT.OrdenIngreso and I.Anulado='NO' and I.IdCta=CTA.IdCta"
               
    Else 'Mensaje Cod Registro Ingreso  NO Valido
        MsgBox "El Código de traslado no válido, debe ser CA o BA AAMM9999", _
        vbExclamation + vbOKOnly, "Caja-Bancos Traslado"
        fbCargarTrasladoDestino = False
        Exit Function
    End If
End If

'Copia la sentencia SQL
mcurRegTrasladosDestino.SQL = sSQL

' Abre el cursor SI hay  error sale indicando la causa del error
If mcurRegTrasladosDestino.Abrir = HAY_ERROR Then
    End
End If

' Cursor abierto
mbCargadoOrigen = True

'Verifica la existencia del registro de ingreso
If mcurRegTrasladosDestino.EOF Then
    'Mensaje de registro de Ingreso a Caja o Bancos NO existe
    MsgBox "El Código de Ingreso que se digito No está registrado o está Anulado", _
      vbInformation + vbOKOnly, "Caja-Bancos- Traslados Modificación"
    mcurRegTrasladosDestino.Cerrar
    ' Cursor abierto
    mbCargadoOrigen = False
    Exit Function
    
Else
    'Carga los controles con datos del ingreso y Habilita los controles
    CargarControlesDestino
    
End If

' Todo Ok
fbCargarTrasladoDestino = True

End Function
Private Function CargarSaldoCaja() As Double
 
 'Hallar Saldo para Caja
Dim sSQL As String
Dim curCuentas As New clsBD2
Dim curIngreso As New clsBD2
Dim curEgreso As New clsBD2

Dim iCol As Integer
Dim curEmpresas As New clsBD2
Dim EmpresasExistentes As String
Dim InstrucEmpresas As String
Dim TotalEgresoProyectos As Double
Dim TotalEgresoEmpresasSinRH As Double
Dim TotalEgresoEmpresasSoloRHCB As Double
Dim TotalEgresos As Double
Dim TotalIngresos As Double


Dim curSaldo As New clsBD2
Dim dblSaldo As Double

Dim iSaldo As String
 'Hallar el Ingreso
    sSQL = "SELECT sum(monto)From Ingresos " & _
            "WHERE  LEFT(Orden,2)= 'CA' AND Anulado='NO'"
              
      curIngreso.SQL = sSQL
      If curIngreso.Abrir = HAY_ERROR Then
        End
      End If
          
    'Se muestran los Ingresos
    If IsNull(curIngreso.campo(0)) Then
      TotalIngresos = "0"
    Else
      TotalIngresos = Format(curIngreso.campo(0), "###,###,##0.00")
    End If
    
    'Hallar el Egreso
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
    
    ' Carga la sentencia
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL DE EGRESOS PARA PROYECTOS CON AFECTACION Y SIN AFECTACION
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) " _
         & " FROM EGRESOS " _
         & " WHERE Origen='C' " _
         & " and Anulado='NO' and Orden like 'CA*' AND " & InstrucEmpresas
    
    ' Ejecuta la sentencia
    curEgreso.SQL = sSQL
    If curEgreso.Abrir = HAY_ERROR Then End
    
    ' Verifica si es vacío
    If curEgreso.EOF Then
       ' Envía 0.00 como resultado
       TotalEgresoProyectos = 0
    Else
      If IsNull(curEgreso.campo(0)) Then
         ' Envía 0.00 como resultado
         TotalEgresoProyectos = 0
      Else
        ' Envía la suma de los ingresos
        TotalEgresoProyectos = curEgreso.campo(0)
      End If
    End If
    
    ' Cierra el cursor
    curEgreso.Cerrar
    
    ' Carga la sentencia
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SIN RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoAfectado) " _
         & " FROM EGRESOS, PROYECTOS " _
         & " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And Origen='C'" _
         & " and Anulado='NO' and Orden like 'CA*' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc<>'02') "
    
    ' Ejecuta la sentencia
    curEgreso.SQL = sSQL
    If curEgreso.Abrir = HAY_ERROR Then End
    
    ' Verifica si es vacío
    If curEgreso.EOF Then
       ' Envía 0.00 como resultado
       TotalEgresoEmpresasSinRH = 0
    Else
      If IsNull(curEgreso.campo(0)) Then
         ' Envía 0.00 como resultado
         TotalEgresoEmpresasSinRH = 0
      Else
        ' Envía la suma de los ingresos
        TotalEgresoEmpresasSinRH = curEgreso.campo(0)
      End If
    End If
    
    ' Cierra el cursor
    curEgreso.Cerrar
    
    ' Carga la sentencia
    '*-*-*-*-*
    '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SOLO RH
    '*-*-*-*-*
    sSQL = "SELECT SUM(MontoCB) " _
         & " FROM EGRESOS, PROYECTOS " _
         & " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And Origen='C'" _
         & " and Anulado='NO' and Orden like 'CA*' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc='02') "
    
    ' Ejecuta la sentencia
    curEgreso.SQL = sSQL
    If curEgreso.Abrir = HAY_ERROR Then End
    
    ' Verifica si es vacío
    If curEgreso.EOF Then
       ' Envía 0.00 como resultado
       TotalEgresoEmpresasSoloRHCB = 0
    Else
      If IsNull(curEgreso.campo(0)) Then
         ' Envía 0.00 como resultado
         TotalEgresoEmpresasSoloRHCB = 0
      Else
        ' Envía la suma de los ingresos
        TotalEgresoEmpresasSoloRHCB = curEgreso.campo(0)
      End If
    End If
    
    ' Cierra el cursor
    curEgreso.Cerrar
    
    TotalEgresos = TotalEgresoProyectos + TotalEgresoEmpresasSinRH + TotalEgresoEmpresasSoloRHCB
    
    'txtEgresos.Text = TotalEgreso
    TotalEgresos = Format(TotalEgresos, "###,###,##0.00")
    iSaldo = Val(Var37(TotalIngresos)) - Val(Var37(TotalEgresos))
    CargarSaldoCaja = Format(iSaldo, "###,###,##0.00")

End Function

Private Function fbCargarTrasladoOrigen() As Boolean
'----------------------------------------------------------------------------
'Propósito  : Carga el registro de ingreso de acuerdo al código en la caja _
              de texto
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota el codigo de Registro de ingreso es CA o BA AAMM9999
Dim sSQL As String

'Deshabilita los movimientos
txtCodMov.Enabled = False
cboMovimiento.Enabled = False

' Verifica SI el ingreso es a Caja o Bancos
msCajaoBancoOrigen = Left(txtOrdenOrigen.Text, 2)

' Carga la sentencia que consulta a la BD acerca del registo de ingreso en Caja o Bancos
If msCajaoBancoOrigen = "CA" Then 'Consulta a Caja
    sSQL = "SELECT E.CodMov,CT.IdPersona, E.MontoCB,E.FecMov, E.Observ,E.NumDoc,E.IdTipoDoc, CT.OrdenIngreso " & _
           "FROM CTB_TRASLADOCAJABANCOS CT, EGRESOS E WHERE " & _
           "CT.OrdenEgreso='" & Trim(txtOrdenOrigen.Text) & "' And CT.OrdenEgreso=  E.Orden " & _
           "And Anulado='NO' "
           
Else 'Consulta a Bancos
    If msCajaoBancoOrigen = "BA" Then
    sSQL = "SELECT E.CodMov,CT.IdPersona, E.MontoCB,E.FecMov, E.Observ,E.NumDoc,E.IdTipoDoc, CT.OrdenIngreso , " & _
           "CTA.IdBanco,E.IdCta, E.NumCheque FROM TIPO_CUENTASBANC CTA,CTB_TRASLADOCAJABANCOS CT, EGRESOS E WHERE " & _
           "CT.OrdenEgreso='" & Trim(txtOrdenOrigen.Text) & "' And E.Orden=CT.OrdenEgreso " & _
           "And E.Anulado='NO' And E.IdCta=CTA.IdCta"
               
    Else 'Mensaje Cod Registro Ingreso  NO Valido
        MsgBox "El Código de traslado no válido, debe ser CA o BA AAMM9999", _
        vbExclamation + vbOKOnly, "Caja-Bancos Traslado"
        fbCargarTrasladoOrigen = False
        Exit Function
    End If
End If

'Copia la sentencia SQL
mcurRegTrasladosOrigen.SQL = sSQL

' Abre el cursor SI hay  error sale indicando la causa del error
If mcurRegTrasladosOrigen.Abrir = HAY_ERROR Then
    End
End If

' Cursor abierto
mbCargadoOrigen = True

'Verifica la existencia del registro de ingreso
If mcurRegTrasladosOrigen.EOF Then
    'Mensaje de registro de Ingreso a Caja o Bancos NO existe
    MsgBox "El Código de Ingreso que se digito No está registrado o está Anulado", _
      vbInformation + vbOKOnly, "Caja-Bancos- Traslados Modificación"
    mcurRegTrasladosOrigen.Cerrar
    ' Cursor abierto
    mbCargadoOrigen = False
    Exit Function
    
Else
    'Carga los controles con datos del ingreso y Habilita los controles
    CargarControlesOrigen
    
End If

' Todo Ok
fbCargarTrasladoOrigen = True

End Function

Private Sub CargarControlesDestino()
'----------------------------------------------------------------------------
'Propósito  : Cargar los controles refentes al ingreso que se desea modificar
'Recibe     :  Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el procedimiento Cargar Registro de Ingreso

'Rellena los controles de Caja
txtDocDestino.Text = mcurRegTrasladosDestino.campo(0)
txtTipoDocDestino.Text = mcurRegTrasladosDestino.campo(1)
txtOrdenOrigen.Text = mcurRegTrasladosDestino.campo(2)

'Verifica si es caja o banco
If msCajaoBancoDestino = "BA" Then 'Rellena los controles de Banco
    txtBancoDestino.Text = mcurRegTrasladosDestino.campo(3)
    msCtaCteDestino = mcurRegTrasladosDestino.campo(4) 'Actualiza variable de Código de CtaCte
    CD_ActVarCbo cboCtaCteDestino, msCtaCteDestino, mcolCodDesCtaCte
        
End If

'Carga el saldo del CtaCte
CargarSaldoDestino

'Habilita Botones cancelar,Anular Caja o Bancos
cmdAnular.Enabled = True
cmdCancelar.Enabled = True

End Sub

Private Sub CargarControlesOrigen()
'----------------------------------------------------------------------------
'Propósito  : Cargar los controles refentes al ingreso que se desea modificar
'Recibe     :  Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Nota Llamado desde el procedimiento Cargar Registro de Ingreso

'Rellena los controles de Caja
txtCodMov.Text = mcurRegTrasladosOrigen.campo(0)
txtCodPersonal.Text = mcurRegTrasladosOrigen.campo(1)
txtMonto.Text = Format(mcurRegTrasladosOrigen.campo(2), "###,###,##0.00")
mskFecTrab.Text = FechaDMA(Trim(Str(mcurRegTrasladosOrigen.campo(3))))
txtObservacion.Text = mcurRegTrasladosOrigen.campo(4)
txtDocOrigen.Text = mcurRegTrasladosOrigen.campo(5)
txtTipoDocOrigen.Text = mcurRegTrasladosOrigen.campo(6)
txtOrdenDestino.Text = mcurRegTrasladosOrigen.campo(7)

'Verifica si es caja o banco
If msCajaoBancoOrigen = "BA" Then 'Rellena los controles de Banco
    txtBancoOrigen.Text = mcurRegTrasladosOrigen.campo(8)
    msCtaCteOrigen = mcurRegTrasladosOrigen.campo(9) 'Actualiza variable de Código de CtaCte
    txtNumCh.Text = mcurRegTrasladosOrigen.campo(10)
    CD_ActVarCbo cboCtaCteOrigen, msCtaCteOrigen, mcolCodDesCtaCte
         
End If

'Carga el saldo del CtaCte
CargarSaldoOrigen

'Habilita Botones cancelar,Anular Caja o Bancos
cmdAnular.Enabled = True
cmdCancelar.Enabled = True

End Sub

Private Sub txtTipoDocDestino_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtTipoDocOrigen_Change()

' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboTipDocOrigen, txtTipoDocOrigen, mcolCodDesTipDocOrigen

' Verifica SI el campo esta vacio
If txtTipoDocOrigen.Text <> "" And cboTipDocOrigen.Text <> "" Then
' Los campos coloca a color blanco
   txtTipoDocOrigen.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtTipoDocOrigen.BackColor = Obligatorio
End If

'Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtTipoDocDestino_Change()
' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboTipoDocDestino, txtTipoDocDestino, mcolCodDesTipDocDestino

' Verifica SI el campo esta vacio
If txtTipoDocDestino.Text <> "" And cboTipoDocDestino.Text <> "" Then
' Los campos coloca a color blanco
   txtTipoDocDestino.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtTipoDocDestino.BackColor = Obligatorio
End If

'Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtTipoDocOrigen_KeyPress(KeyAscii As Integer)

' Si se presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub
