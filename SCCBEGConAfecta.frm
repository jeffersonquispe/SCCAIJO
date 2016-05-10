VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBEGConAfecta 
   Caption         =   "Caja y Banco -  Egreso con afectación"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   63
   Icon            =   "SCCBEGConAfecta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPCtaCte 
      Height          =   255
      Left            =   7440
      Picture         =   "SCCBEGConAfecta.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   750
      Width           =   225
   End
   Begin VB.ComboBox cboCtaCte 
      Height          =   315
      Left            =   5760
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdPBanco 
      Height          =   255
      Left            =   4095
      Picture         =   "SCCBEGConAfecta.frx":0BA2
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   750
      Width           =   225
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   2910
   End
   Begin VB.CommandButton cmdPProy 
      Height          =   255
      Left            =   5460
      Picture         =   "SCCBEGConAfecta.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1170
      Width           =   220
   End
   Begin VB.ComboBox cboProy 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1140
      Width           =   4260
   End
   Begin VB.CommandButton cmdPProg 
      Height          =   255
      Left            =   5460
      Picture         =   "SCCBEGConAfecta.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1575
      Width           =   225
   End
   Begin VB.ComboBox cboProg 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1545
      Width           =   4260
   End
   Begin VB.CommandButton cmdPLinea 
      Height          =   255
      Left            =   5460
      Picture         =   "SCCBEGConAfecta.frx":142A
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1965
      Width           =   225
   End
   Begin VB.ComboBox cboLinea 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1935
      Width           =   4260
   End
   Begin VB.CommandButton cmdPActiv 
      Height          =   255
      Left            =   5460
      Picture         =   "SCCBEGConAfecta.frx":1702
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2370
      Width           =   225
   End
   Begin VB.ComboBox cboActiv 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2340
      Width           =   4260
   End
   Begin VB.CommandButton cmdPCategoriaGasto 
      Height          =   255
      Left            =   5460
      Picture         =   "SCCBEGConAfecta.frx":19DA
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2760
      Width           =   225
   End
   Begin VB.ComboBox cboCategoriaGasto 
      Height          =   315
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2730
      Width           =   4260
   End
   Begin VB.CommandButton cmdBuscarEgreso 
      Caption         =   "..."
      Height          =   255
      Left            =   2790
      TabIndex        =   1
      Top             =   270
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   4080
      Left            =   80
      TabIndex        =   63
      Top             =   3220
      Width           =   11700
      Begin VB.CommandButton CmdAgregarProdServ 
         Caption         =   "Agregar Producto/Servicio"
         Height          =   495
         Left            =   9240
         TabIndex        =   97
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdSelImpuestos 
         Caption         =   "Aplicar &Impuestos"
         Height          =   400
         Left            =   9240
         TabIndex        =   40
         Top             =   240
         Width           =   2175
      End
      Begin VB.Frame fraPagar 
         Caption         =   "Pagar:"
         Height          =   615
         Left            =   1320
         TabIndex        =   89
         Top             =   120
         Width           =   2895
         Begin VB.OptionButton optProducto 
            Caption         =   "Productos"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optServicio 
            Caption         =   "Servicios"
            Height          =   195
            Left            =   1680
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraTipoGiro 
         Caption         =   "Tipo de giro del Doc.:"
         Height          =   615
         Left            =   4680
         TabIndex        =   88
         Top             =   120
         Width           =   4215
         Begin VB.OptionButton optGiroConImpuestos 
            Caption         =   "Girado con impuestos"
            Height          =   195
            Left            =   2160
            TabIndex        =   44
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optGiroSinImpuestos 
            Caption         =   "Girado sin impuestos"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdPProdServ 
         Height          =   255
         Left            =   5835
         Picture         =   "SCCBEGConAfecta.frx":1CB2
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   825
         Width           =   220
      End
      Begin VB.ComboBox cboProdServ 
         Height          =   315
         ItemData        =   "SCCBEGConAfecta.frx":1F8A
         Left            =   1155
         List            =   "SCCBEGConAfecta.frx":1F8C
         Style           =   1  'Simple Combo
         TabIndex        =   46
         Top             =   795
         Width           =   4935
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   9360
         TabIndex        =   81
         Top             =   1935
         Width           =   2055
         Begin VB.TextBox txtSaldoCB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            MaxLength       =   14
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox txtMontoImpuesto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            MaxLength       =   14
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   370
            Width           =   1335
         End
         Begin VB.TextBox txtMontoCB 
            Enabled         =   0   'False
            Height          =   315
            Left            =   360
            MaxLength       =   14
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   950
            Width           =   1335
         End
         Begin VB.Label lblSaldoCB 
            AutoSize        =   -1  'True
            Caption         =   "Saldo de Caja "
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   1275
            Width           =   1035
         End
         Begin VB.Label Label7 
            Caption         =   "Monto de Impuestos"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   150
            Width           =   1455
         End
         Begin VB.Label lblMontoPagado 
            Caption         =   "Monto a Pagar"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   735
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "A&ñadir"
         Height          =   400
         Left            =   9240
         TabIndex        =   53
         Top             =   1440
         Width           =   1000
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   10395
         TabIndex        =   59
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox txtPrecioUniCompra 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1335
      End
      Begin VB.TextBox txtValorCompra 
         Height          =   315
         Left            =   3840
         TabIndex        =   51
         Top             =   1545
         Width           =   1815
      End
      Begin VB.TextBox txtPrecioUniVenta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtCant 
         Height          =   315
         Left            =   1155
         MaxLength       =   11
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtMedida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   795
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetalle 
         Height          =   2100
         Left            =   360
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1900
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3704
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         ForeColorSel    =   16777215
         AllowBigSelection=   -1  'True
         HighLight       =   0
         FillStyle       =   1
      End
      Begin VB.TextBox txtValorVenta 
         Height          =   315
         Left            =   3840
         TabIndex        =   49
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "&Medida:"
         Height          =   255
         Left            =   6240
         TabIndex        =   85
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrecioUniCompra 
         Caption         =   "Precio unit. Total"
         Height          =   255
         Left            =   6240
         TabIndex        =   69
         Top             =   1545
         Width           =   1215
      End
      Begin VB.Label lblValorCompra 
         Caption         =   "Costo Total"
         Height          =   255
         Left            =   2880
         TabIndex        =   68
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label lblPrecioUniVenta 
         Caption         =   "Precio unit. Neto"
         Height          =   255
         Left            =   6240
         TabIndex        =   67
         Top             =   1185
         Width           =   1215
      End
      Begin VB.Label lblCant 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   1185
         Width           =   855
      End
      Begin VB.Label lblProdServ 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblValorVenta 
         Caption         =   "Costo Neto"
         Height          =   255
         Left            =   2880
         TabIndex        =   64
         Top             =   1185
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "A&nular"
      Height          =   400
      Left            =   9720
      TabIndex        =   56
      Top             =   7350
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10800
      TabIndex        =   57
      Top             =   7350
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   8640
      TabIndex        =   55
      Top             =   7350
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   7560
      TabIndex        =   54
      Top             =   7350
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   3360
      Left            =   75
      TabIndex        =   70
      Top             =   0
      Width           =   11700
      Begin VB.TextBox TxtCategoriaGasto 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   26
         Top             =   2730
         Width           =   540
      End
      Begin VB.TextBox txtRinde 
         Height          =   315
         Left            =   855
         MaxLength       =   4
         TabIndex        =   11
         Top             =   705
         Width           =   675
      End
      Begin VB.CommandButton cmdBuscaRinde 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6930
         Picture         =   "SCCBEGConAfecta.frx":1F8E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox txtDescRinde 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1530
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   705
         Width           =   5400
      End
      Begin VB.Frame fraCB 
         Caption         =   "Egreso de:"
         Height          =   550
         Left            =   3120
         TabIndex        =   94
         Top             =   120
         Width           =   3375
         Begin VB.OptionButton optCaja 
            Caption         =   "Caja"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   210
            Width           =   630
         End
         Begin VB.OptionButton optBanco 
            Caption         =   "Banco"
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   200
            Width           =   780
         End
         Begin VB.OptionButton optRendir 
            Caption         =   "Cuenta a Rendir"
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   200
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdPTipoDoc 
         Height          =   255
         Left            =   11250
         Picture         =   "SCCBEGConAfecta.frx":2090
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1575
         Width           =   220
      End
      Begin VB.ComboBox cboTipDoc 
         Height          =   315
         Left            =   6900
         Style           =   1  'Simple Combo
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1545
         Width           =   4600
      End
      Begin VB.TextBox txtNumCheque 
         Height          =   315
         Left            =   9180
         MaxLength       =   15
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtBanco 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   5
         Top             =   720
         Width           =   540
      End
      Begin VB.TextBox txtDocEgreso 
         Height          =   315
         Left            =   6480
         MaxLength       =   15
         TabIndex        =   35
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDoc 
         Height          =   315
         Left            =   9540
         MaxLength       =   14
         TabIndex        =   38
         Top             =   2340
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   310
         Left            =   11040
         Picture         =   "SCCBEGConAfecta.frx":2368
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1920
         Width           =   470
      End
      Begin VB.TextBox txtProy 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1140
         Width           =   540
      End
      Begin VB.TextBox txtProg 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1545
         Width           =   540
      End
      Begin VB.TextBox txtLinea 
         Height          =   315
         Left            =   840
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1935
         Width           =   540
      End
      Begin VB.TextBox txtTipDoc 
         Height          =   315
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   29
         Top             =   1545
         Width           =   420
      End
      Begin VB.TextBox txtRUCDNI 
         Height          =   315
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   32
         Top             =   1935
         Width           =   1215
      End
      Begin VB.TextBox txtNombrProv 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7720
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1935
         Width           =   3280
      End
      Begin VB.TextBox txtActiv 
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2340
         Width           =   540
      End
      Begin VB.TextBox txtCodEgreso 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1440
      End
      Begin VB.TextBox txtFinan 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1140
         Width           =   5000
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   10380
         TabIndex        =   71
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtObserv 
         Height          =   315
         Left            =   6480
         MaxLength       =   100
         TabIndex        =   39
         Top             =   2760
         Width           =   4995
      End
      Begin MSMask.MaskEdBox mskFecDoc 
         Height          =   315
         Left            =   7780
         TabIndex        =   13
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha del Doc.:"
         Height          =   255
         Left            =   6600
         TabIndex        =   98
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label Label10 
         Caption         =   "CatGasto:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   2760
         Width           =   670
      End
      Begin VB.Label lblRinde 
         Caption         =   "Cuenta a Rendir:"
         Height          =   435
         Left            =   90
         TabIndex        =   95
         Top             =   660
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   11160
         Picture         =   "SCCBEGConAfecta.frx":246A
         Stretch         =   -1  'True
         Top             =   680
         Width           =   360
      End
      Begin VB.Label lblBanco 
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblNumCheque 
         Caption         =   "Num. Cheque:"
         Height          =   255
         Left            =   7920
         TabIndex        =   92
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblCtaCte 
         Caption         =   "Num Cuen&ta:"
         Height          =   255
         Left            =   4560
         TabIndex        =   91
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Num.Doc.:"
         Height          =   195
         Left            =   5685
         TabIndex        =   90
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label15 
         Caption         =   "Observac:"
         Height          =   255
         Left            =   5685
         TabIndex        =   87
         Top             =   2775
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Importe Total:"
         Height          =   255
         Left            =   8445
         TabIndex        =   86
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label lblProy 
         Caption         =   "Proyecto:"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Programa:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Línea:"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1935
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Left            =   5685
         TabIndex        =   77
         Top             =   1530
         Width           =   750
      End
      Begin VB.Label lblRuc 
         AutoSize        =   -1  'True
         Caption         =   "RUC/DNI:"
         Height          =   195
         Left            =   5685
         TabIndex        =   76
         Top             =   1935
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Código del Egreso :"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblFinan 
         AutoSize        =   -1  'True
         Caption         =   "Financiera:"
         Height          =   195
         Left            =   5685
         TabIndex        =   73
         Top             =   1140
         Width           =   780
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha de trabajo:"
         Height          =   255
         Left            =   9100
         TabIndex        =   72
         Top             =   255
         Width           =   1250
      End
   End
End
Attribute VB_Name = "frmCBEGConAfecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Colecciones para la carga del combo de Proyectos
Private mcolCodProy As New Collection
Private mcolCodDesProy As New Collection

' Colecciones para la carga del combo de Programas
Private mcolCodProg As New Collection
Private mcolCodDesProg As New Collection

'Colecciones para la carga del combo de Lineas
Private mcolCodLinea As New Collection
Private mcolCodDesLinea As New Collection

'Colecciones para la carga del combo de Actividades
Private mcolCodActiv As New Collection
Private mcolCodDesActiv As New Collection

Private mcolCodBanco As New Collection
Private mcolCodDesBanco As New Collection
Private mcolCodCtaCte As New Collection
Private mcolCodDesCtaCte As New Collection
Private msCtaCte As String

Private mcolCodCatGasto As New Collection
Private mcolCodDesCatGasto As New Collection

Private msCaBaAnt As String
Private msRindeAnt As String ' Cuenta a rendir anterior

'Colecciones para la carga del combo de Productos
Private mcolidprod As New Collection
Private mcolCodDesProd As New Collection
Private mcolDesMedidaContProd As New Collection

'Colecciones para la carga del combo de Servicios
Private mcolCodServ As New Collection
Private mcolCodDesServ As New Collection
Private mcolDesMedidaContServ As New Collection

'Colecciones para la carga del combo de tipo de documento
Private mcolCodTipDoc As New Collection
Private mcolCodDesTipDoc As New Collection
Private mcolCodRetencionPaga As New Collection

' Colección para la carga de detalle, impuestos
Private mcolEgreDet  As New Collection
Private mcolEgreImpt As New Collection

'Variables que son Código del producto o servicio y su cta contable
Private msIdProdServ As String
Private msCodCont As String

'Variable indica si se cancela las operaciones entre optsProServ
Private mbCancelaOptClick As Boolean
Private mbCancelaCambioTipDoc As Boolean
Private mbCancelagrid As Boolean

' Variable que indica si ha sido encontrado el proveedor
Private mbEncontradoProv As Boolean
' variable que indica si se está calculando los valores de compra y de venta
Private mbCalculando As Boolean
' Variable que indica si se cargó un egreso
Private mbEgresoCargado As Boolean

' Variables de modulo relacionadas al impuesto aplicado
Dim mdblMontoTotalImpt As Double
Dim mdblSumImpt As Double

' Variable que almacena el tipo de documento anterior
Dim msAntTipDoc As String

' Variable que almacena el código con el cual se guardó el egreso y del proveedor
Dim msOrden As String

' Cursores que cargan los datos del egreso para la modificación
Dim mcurEgreso As New clsBD2
Dim mcurDetalleEgreso As New clsBD2
Dim mcurImpuestos As New clsBD2

Dim mcurProyectos As New clsBD2

' Variable para el manejo del grid
Dim ipos As Long

Dim DetalleSubido As Boolean
Dim CodigoCategoriaGasto As String

Private Sub EstableceCamposObligatorios1raParte()
  txtProy.BackColor = Obligatorio
  txtProg.BackColor = Obligatorio
  txtLinea.BackColor = Obligatorio
  txtActiv.BackColor = Obligatorio
  txtTipDoc.BackColor = Obligatorio
  txtRUCDNI.BackColor = Obligatorio
  txtDocEgreso.BackColor = Obligatorio
  txtTotalDoc.BackColor = Obligatorio
  txtBanco.BackColor = Obligatorio
  cboCtaCte.BackColor = Obligatorio
  txtNumCheque.BackColor = Obligatorio
  txtRinde.BackColor = Obligatorio
  mskFecDoc.BackColor = Obligatorio
End Sub

Private Sub EstableceCamposObligatorios2daParte()
  cboProdServ.BackColor = Obligatorio
  txtCant.BackColor = Obligatorio
  txtValorVenta.BackColor = Obligatorio
  txtValorCompra.BackColor = Obligatorio
End Sub

Private Sub cboCategoriaGasto_Change()
  ' verifica SI lo ingresado esta en la lista del combo
  If VerificarTextoEnLista(cboCategoriaGasto) = True Then SendKeys "{down}"
End Sub

Private Sub cboCategoriaGasto_Click()
  ' Verifica SI el evento ha sido activado por el teclado o Mouse
  If VerificarClick(cboCategoriaGasto.ListIndex) = False And cboCategoriaGasto.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
End Sub

Private Sub cboCategoriaGasto_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Verifica SI es enter para salir o flechas para recorrer
  VerificaKeyDowncbo (KeyCode)
End Sub

Private Sub cboCategoriaGasto_LostFocus()
  ' sale del combo y acualiza datos enlazados
  If ValidarDatoCbo(cboCategoriaGasto, vbWhite) = True Then
    
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboCategoriaGasto.Text, TxtCategoriaGasto, mcolCodCatGasto, mcolCodDesCatGasto
    
  Else '  Vaciar Controles enlazados al combo
    TxtCategoriaGasto.Text = Empty
  End If
  
  'Cambia el alto del combo
  cboCategoriaGasto.Height = CBONORMAL
End Sub

Private Sub cboProdServ_Change()
  ' verifica SI lo ingresado esta en la lista del combo
  If VerificarTextoEnLista(cboProdServ) = True Then SendKeys "{down}"
End Sub

Private Sub cboProdServ_Click()
  ' Verifica SI el evento ha sido activado por el teclado o Mouse
  If VerificarClick(cboProdServ.ListIndex) = False And cboProdServ.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse
End Sub

Private Sub cboProdServ_GotFocus()
  ' Verificar si se a introducido un tipo de documento y si se seleccionó _
    impuestos
  If fbOperarDetalle = False Then
    ' sale del procedimiento
    Exit Sub
  End If
End Sub

Private Function fbOperarDetalle() As Boolean
  fbOperarDetalle = False
  
  ' Verifica si se ha introducido algún tipo de documento, monto y impuestos
  If txtTipDoc <> Empty And cboTipDoc <> Empty _
      And gbImpuestos = True And txtTotalDoc <> Empty Then
      ' los datos para el ingreso del detalle están completos
      fbOperarDetalle = True
  Else
    ' verifica si se introdujo algún tipo de documento
    If txtTipDoc <> Empty And cboTipDoc <> Empty Then
      If txtTotalDoc.Text <> Empty Then ' verifica si se introdujo monto
         ' Mensaje de Información, no se selccionó impuestos
          MsgBox "No se seleccionó los impuestos para este documento", , "SGCcaijo - Egreso con afectación"
         ' ya se ingreso el monto del documento, muestra el formulario de Impuestos
          cmdSelImpuestos.SetFocus
          cmdSelImpuestos_Click
      Else ' no se ingresó el monto del documento
          txtTotalDoc.SetFocus
      End If
    Else
      ' Mensaje de Información
      MsgBox "No se seleccionó un documento", , "SGCcaijo - Egreso con afectación"
      ' pone el focus al cbo tipo de documento
      cboTipDoc.SetFocus
    End If
  End If
End Function

Private Sub cboProdServ_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Verifica SI es enter para salir o flechas para recorrer
  VerificaKeyDowncbo (KeyCode)
End Sub

Private Sub cboProdServ_LostFocus()
  ' sale del combo y acualiza datos enlazados
  If ValidarDatoCbo(cboProdServ, Obligatorio) = True Then
    If optProducto.Value = True Then
      ' Se actualiza código de la variable correspondiente a descripción introducida
      'CD_ActCboVar cboProdServ.Text, msIdProdServ, mcolidprod, mcolCodDesProd
      ActualizarInfoProdServ cboProdServ.Text, msIdProdServ, mcolidprod, mcolCodDesProd
    Else
      ' Se actualiza código de la variable correspondiente a descripción introducida
      'CD_ActCboVar cboProdServ.Text, msIdProdServ, mcolCodServ, mcolCodDesServ
      ActualizarInfoProdServ cboProdServ.Text, msIdProdServ, mcolCodServ, mcolCodDesServ
    End If
     
     'Actualiza la medida correspondiente al producto seleccionada y su codcont
     MostrarMedidaCont
  Else
     'no se eligió un product o serv
     msIdProdServ = ""
  End If
  
  'Cambia el alto del combo
  cboProdServ.Height = CBONORMAL
  
  'habilitar el boton añadir
  HabilitaBotonAñadir
End Sub

Private Sub MostrarMedidaCont()
  If optProducto.Value = True Then
    'Muestra la medida del producto seleccionado en el combo
    txtMedida.Text = Var30(mcolDesMedidaContProd.Item(msIdProdServ), 1)
    msCodCont = Var30(mcolDesMedidaContProd.Item(msIdProdServ), 2)
  Else
    'Muestra la medida del producto seleccionado en el combo
    txtMedida.Text = Var30(mcolDesMedidaContServ.Item(msIdProdServ), 1)
    msCodCont = Var30(mcolDesMedidaContServ.Item(msIdProdServ), 2)
  End If
End Sub

Private Sub cboTipDoc_Change()
  ' verifica SI lo ingresado esta en la lista del combo
  If VerificarTextoEnLista(cboTipDoc) = True Then
    SendKeys "{down}"
  End If
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

Private Function fsEstaenDetalle(sProdServ As Variant) As String
  Dim j As Integer
  
  'Inicializamos a funcion asumiendo que Procuto NO esta en el grddetalle 1
  fsEstaenDetalle = Empty
  
  ' recorremos el grid detalle de Producto verificando la existencia de txtProd
  For j = 1 To grdDetalle.Rows - 1
   If grdDetalle.TextMatrix(j, 4) = sProdServ Then
  ' Carga registro orignal, "codConcepto", "cantidad", "Monto"
  '"Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "CodCont")
      fsEstaenDetalle = grdDetalle.TextMatrix(j, 4) & "¯" & grdDetalle.TextMatrix(j, 1) _
               & "¯" & grdDetalle.TextMatrix(j, 3)
      Exit Function
   End If
  Next j
End Function

Private Sub cmdAceptar_Click()
  ' Verifica si los datos sson correctos
  If fbVerificarDatosIntroducidos = False Then
    ' algún dato es incorrecto
    Exit Sub
  End If
  
  If gsTipoOperacionEgreso = "Nuevo" Then
   ' Pregunta aceptación de los datos
     If MsgBox("¿Está conforme con los datos?", _
        vbQuestion + vbYesNo, "Caja-Bancos, Egreso con afectación") = vbYes Then
       'Actualiza la transaccion
        Var8 1, gsFormulario
       
       ' Se guardan los datos del egreso
        GuardarEgreso
     Else: Exit Sub ' sale
     End If
  Else
   ' Controla si algún producto ha sido verificado en almacén
   If fbOkProductosAlmacen("Modificar") = False Then
      ' algún dato incorrecto
      Exit Sub
   End If
  
   ' Mensaje de conformidad de los datos
     If MsgBox("¿Está conforme con las modificaciones realizadas en el Egreso " & txtCodEgreso.Text & "?", _
                    vbQuestion + vbYesNo, "Caja-Bancos, Modificación de egreso con afectación") = vbYes Then
       'Actualiza la transaccion
        Var8 1, gsFormulario
                    
       ' Se Modifican los datos del egreso
         GuardarModificacionesEgreso
     Else: Exit Sub ' sale
     End If
  End If
  
  'Actualiza la transaccion
   Var8 -1, Empty
  
  ' Mensaje Ok
  MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso con Afectación"
  
  ' Limpia la pantalla para una nueva operación, Prepara el formulario
  ' Limpia la pantalla
     LimpiarFormulario
     
  If gsTipoOperacionEgreso = "Nuevo" Then
   ' Nuevo egreso
     NuevoEgreso
  Else
   ' cierra el control egreso
     If mbEgresoCargado Then
        mcurEgreso.Cerrar
        mbEgresoCargado = False
        mskFecTrab = "__/__/____"
     End If
   ' Se Modifican los datos del egreso
     ModificarEgreso
  End If
End Sub

Private Function fbOkProductosAlmacen(sOperacion As String) As Boolean
  Dim curVerificados As New clsBD2
  Dim sSQL As String
  Dim sProd As Variant
  Dim colProdenAlmacen As New Collection
  
  ' Inicializa la función asumiendo que no se ha ingresado productos
  fbOkProductosAlmacen = True
  
  ' Averigua los productos verificados en almacén
  sSQL = "SELECT idProd FROM ALMACEN_VERIFICACION " _
       & "WHERE Orden='" & msOrden & "' and Verificado='SI'"
  
  ' Ejecuta la sentencia
  curVerificados.SQL = sSQL
  If curVerificados.Abrir = HAY_ERROR Then End
  
  ' verifica si tiene algún producto verificado
  If curVerificados.EOF Then
    ' No hay productos ingresados en almacén por este egreso
    Exit Function
  Else
    ' Existen productos ingresados a almacén por este egreso
    ' Carga la colección de los productos ingresados a almacén
    Do While Not curVerificados.EOF
      '
      colProdenAlmacen.Add Item:=curVerificados.campo(0), _
                            Key:=curVerificados.campo(0)
      ' Mueve al siguiente registro
      curVerificados.MoverSiguiente
    Loop
    curVerificados.Cerrar
    
    ' Si hay productos en almacén , no se puede anular
    If colProdenAlmacen.Count > 0 And sOperacion = "Anular" Then
        MsgBox "No se puede Anular por que algunos productos " & Chr(13) _
             & "se han verificado en almacén. Consulte al administrador", , "SGCcaijo- Egreso con Afectación"
        ' Devuelve el resultado de la verificación
        fbOkProductosAlmacen = False
        Set colProdenAlmacen = Nothing
        Exit Function
    End If
    
    ' Si hay productos verificados en almacén, y se cambia la opción _
      a pagar servicios, no se permite modificar.
    If colProdenAlmacen.Count > 0 And sOperacion = "Modificar" And optServicio.Value = True Then
        MsgBox "No se puede guardar los cambios por que se pagó productos, " _
                & Chr(13) & "y algunos se han verificado en almacén. Consulte al administrador", , "SGCcaijo- Egreso con Afectación"
        ' Devuelve el resultado de la verificación
        fbOkProductosAlmacen = False
        Set colProdenAlmacen = Nothing
        Exit Function
    End If
  End If
  
  ' Verifica si se cambio los datos de los productos verificados _
  en almacén
  If sOperacion = "Modificar" Then
    For Each sProd In colProdenAlmacen
      If mcolEgreDet.Item(sProd) <> fsEstaenDetalle(sProd) Then
          MsgBox "No se puede guardar los cambios por que el producto : " _
                  & Chr(13) & mcolCodDesProd(sProd) _
                  & Chr(13) & "Se ha verificado en almacén. Consulte al administrador", , "SGCcaijo- Egreso con Afectación"
          ' Devuelve el resultado de la verificación
          fbOkProductosAlmacen = False
          Set colProdenAlmacen = Nothing
          Exit Function
      End If
    Next sProd
  End If
  ' Vacía la colección de impuestos
  Set colProdenAlmacen = Nothing
End Function

Private Sub GuardarEgreso()
  ' Asigna el orden
   msOrden = txtCodEgreso
  
  'Guarda el registro general del EgresoCA en Egreso
  GrabarEgresoGeneral
                 
  'Guarda los impuestos aplicados
  GrabarImpuestos
  
  'Guarda el detalle de moviento en Gastos
  GrabarDetalleEnGastosAlmacen
                 
  'Realiza el asiento automatico
  Conta11
End Sub

Private Sub GrabarImpuestos()
  Dim modRetenciones As New clsBD3
  Dim MiObjeto As Variant
  Dim sSQL As String
  Dim CodDocumento As String
    
  CodDocumento = ""
  If TipoEgreso = "EMPR" Then
    If (txtTipDoc = "01") Or (txtTipDoc = "04") Or (txtTipDoc = "05") Or (txtTipDoc = "06") Or (txtTipDoc = "12") Or (txtTipDoc = "13") Or (txtTipDoc = "14") Then
      CodDocumento = "40"
    Else
      CodDocumento = txtTipDoc
    End If
  Else
    CodDocumento = txtTipDoc
  End If
  
  For Each MiObjeto In gcolImpSel 'recorre las retenciones elegidas en la collection
    sSQL = "INSERT INTO MOV_IMPUESTOS VALUES('" _
           & msOrden & "','" _
           & Var30(MiObjeto, 1) & "'," _
           & Var37(Var30(MiObjeto, 3)) & "," _
           & Var37(Var30(MiObjeto, 2)) & ",'" _
           & Var30(mcolCodRetencionPaga.Item(CodDocumento), 1) & "','" _
           & Trim(DeterminarCtaContable(Var30(MiObjeto, 4))) & "','" _
           & Trim(Var30(MiObjeto, 5)) & "')"
    ' ejecuta la sentencia que guarda los datos en la bd
    modRetenciones.SQL = sSQL
    If modRetenciones.Ejecutar = HAY_ERROR Then End
    ' cierra la componente
    modRetenciones.Cerrar
       
     gcolAsientoDetImp.Add _
          Key:=Var30(MiObjeto, 1), _
          Item:=Trim(DeterminarCtaContable(Var30(MiObjeto, 4))) & "¯" & _
                Var37(Var30(MiObjeto, 3))
  Next MiObjeto
End Sub

Private Function DeterminarCtaContable(ValorCta As String) As String
  If (TipoEgreso = "EMPR") And (txtTipDoc = "01" Or txtTipDoc = "04" Or txtTipDoc = "05" Or txtTipDoc = "06" Or txtTipDoc = "12" Or txtTipDoc = "13" Or txtTipDoc = "14") And (ValorCta = "40114") Then
    DeterminarCtaContable = "40111"
  Else
    DeterminarCtaContable = ValorCta
  End If
End Function

Private Sub GrabarEgresoGeneral()
  Dim modEgresoCajaBanco As New clsBD3
  Dim sSQL As String
  Dim MontoSinIGV As Double
    
  If TxtCategoriaGasto.Text <> "" Then
    CodigoCategoriaGasto = Trim(TxtCategoriaGasto.Text)
  Else
    CodigoCategoriaGasto = ""
  End If
  
  MontoSinIGV = Val(Var37(txtMontoCB.Text)) - Val(Var37(txtMontoImpuesto.Text))
  If TipoEgreso = "PROY" Then
    'Carga la sentencia que inserta un registro en Egresos
    If optCaja.Value = True Then 'Carga la sentencia sql que guarda el egreso de caja
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(txtMontoCB) & ",'" _
      & gsIdProv & "','','','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','C','" & CodigoCategoriaGasto & "') "
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) _
          & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯C"
        
    ElseIf optBanco.Value = True Then 'Carga la sentencia sql que guarda el egreso de Banco
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(txtMontoCB) & ",'" _
      & gsIdProv & "','" & msCtaCte & "','" & Trim(txtNumCheque.Text) _
      & "','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','B','" & CodigoCategoriaGasto & "')"
       
  ' carga la colección asiento
  ' Orden,Tip_Doc,Monto,NumCtaBanc,fecha,Observ,Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) & "¯" _
          & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO BANCO CON AFECTACION" & "¯SC¯EB¯B"
    
    ElseIf optRendir.Value = True Then 'Carga la sentencia sql que guarda el egreso de rendir
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(txtMontoCB) & ",'" _
      & gsIdProv & "','','','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','R','" & CodigoCategoriaGasto & "')"
  ' carga la colección asiento
  ' Orden,Tip_Doc,Monto,NumCtaBanc,fecha,Observ,Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) _
          & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯R"
    End If
  Else
    '/*/*/*/*/*/*
    '/*/*/*/*/*/* TIPO DE EGRESO PARA EMPRESA
    '/*/*/*/*/*/*
    'Carga la sentencia que inserta un registro en Egresos
    If optCaja.Value = True Then 'Carga la sentencia sql que guarda el egreso de caja
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(MontoSinIGV) & ",'" _
      & gsIdProv & "','','','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','C','" & CodigoCategoriaGasto & "') "
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) _
          & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯C"
        
    ElseIf optBanco.Value = True Then 'Carga la sentencia sql que guarda el egreso de Banco
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(MontoSinIGV) & ",'" _
      & gsIdProv & "','" & msCtaCte & "','" & Trim(txtNumCheque.Text) _
      & "','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','B','" & CodigoCategoriaGasto & "')"
       
  ' carga la colección asiento
  ' Orden,Tip_Doc,Monto,NumCtaBanc,fecha,Observ,Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) & "¯" _
          & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO BANCO CON AFECTACION" & "¯SC¯EB¯B"
    
    ElseIf optRendir.Value = True Then 'Carga la sentencia sql que guarda el egreso de rendir
      sSQL = "INSERT INTO EGRESOS VALUES('" & msOrden & "','" & txtProy _
      & "','" & txtProg.Text & "','" & txtLinea.Text & "','" _
      & txtActiv.Text & "','" & Trim(txtDocEgreso.Text) & "','" _
      & txtTipDoc.Text & "','','" & FechaAMD(mskFecTrab.Text) & "','" & FechaAMD(mskFecDoc.Text) & "'," _
      & Var37(txtTotalDoc) & "," & Var37(MontoSinIGV) & ",'" _
      & gsIdProv & "','','','NO','" & Trim(txtObserv.Text) & "','','" _
      & fsTipoGiro & "','R','" & CodigoCategoriaGasto & "')"
  ' carga la colección asiento
  ' Orden,Tip_Doc,Monto,NumCtaBanc,fecha,Observ,Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) _
          & "¯Nulo¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯R"
    End If
  End If
      
  ' Ejecuta la sentencia que inserta un registro a Egresos
  ' si al ejecutar hay error se sale de la aplicación
  modEgresoCajaBanco.SQL = sSQL
  If modEgresoCajaBanco.Ejecutar = HAY_ERROR Then
   ' cierra error
    End
  End If
  
  'Se cierra la query
  modEgresoCajaBanco.Cerrar
  
  If optRendir.Value = True Then
  ' Guarda el movimiento de ERendir
    Var4 msOrden, txtRinde, "E", Var37(txtMontoCB), FechaAMD(mskFecTrab.Text), FechaAMD(mskFecDoc.Text)
  End If
End Sub

Private Function DefinirProdServ() As String
  'De acuerdo al option Buton optProducto o optServicio define el concepto del gasto
  If optProducto.Value = True Then 'Gasto en Productos
    DefinirProdServ = "P"
  Else 'Gasto en servicios
    DefinirProdServ = "S"
  End If
End Function

Public Sub GrabarDetalleEnGastosAlmacen()
  Dim i As Integer
  Dim sSQL, sProdServ  As String
  Dim modGastos As New clsBD3
  
  ' Asigna el tipo de conceptos pagados "Productos" o "Servicios"
  sProdServ = DefinirProdServ
  
  ' Recorre el grid y lo almacena en la BD
  For i = 1 To grdDetalle.Rows - 1
    sSQL = "INSERT INTO GASTOS VALUES('" & msOrden & "','" & sProdServ & "','" _
    & grdDetalle.TextMatrix(i, 4) & "'," & Var37(grdDetalle.TextMatrix(i, 1)) & "," _
    & Var37(grdDetalle.TextMatrix(i, 3)) & ")"
    
    gcolAsientoDet.Add _
         Key:=grdDetalle.TextMatrix(i, 4), _
         Item:=grdDetalle.TextMatrix(i, 6) & "¯" & _
               Var37(grdDetalle.TextMatrix(i, 3))
    
    modGastos.SQL = sSQL
    If modGastos.Ejecutar = HAY_ERROR Then
      End
    End If
        
    'Se cierra la query
    modGastos.Cerrar
    
      'Guarda el ingreso de productos en almacen, para la verificación
    If sProdServ = "P" Then GuardarVerif grdDetalle.TextMatrix(i, 4)
  Next i
End Sub

Private Sub GuardarVerif(sIdProdRec As String)
  '------------------------------------------------------------
  'Propósito Guarda los productos en BD a la tabla almacen_verificación
  'Recibe sIdProdRec Identificador del producto
  'Devuelve nada
  '------------------------------------------------------------
  Dim sSQL As String
  
  '///// ELLOS CON ANTERIOR CLASE Dim modAlmacen As New clsBDModificacion
  '///// MODIFICACION VADICK CAMBIANDO clsBDModificacion POR clsBD3
  Dim modAlmacen As New clsBD3
  ' carga la sentencia que guarda el producto y la cantidad en BD
     sSQL = "INSERT INTO ALMACEN_VERIFICACION VALUES('" & msOrden & "','" _
     & sIdProdRec & "','NO')"
  modAlmacen.SQL = sSQL
  'ejecuta la sentencia SQL que inserta registro en Almacen
  If modAlmacen.Ejecutar = HAY_ERROR Then End
  'cierra el query de modifcacion
  modAlmacen.Cerrar
End Sub

Private Function fsTipoGiro() As String
  If optGiroConImpuestos.Value = True Then
    fsTipoGiro = "SI"
  Else
    fsTipoGiro = "NO"
  End If
End Function

Private Sub GuardarModificacionesEgreso()
  ModificarEgresoGeneral
                 
  'Guarda los impuestos aplicados
  ModificarImpuestos
  
  'Guarda el detalle de moviento en Gastos
  ModificarDetalleEnGastosAlmacen
                 
  'Realiza el asiento automatico
  Conta18
End Sub

Private Sub ModificarEgresoGeneral()
  ' -----------------------------------------------
  'Propósito: Modificar el egreso general en la bd
  'Recibe : Nada
  'Enatrega : Nada
  ' -----------------------------------------------
  Dim sSQL As String
  Dim modEgreso As New clsBD3
  Dim MontoSinIGV As Double
  
  ' carga la sentencia que modifica el egreso
  If TxtCategoriaGasto.Text <> "" Then
    CodigoCategoriaGasto = Trim(TxtCategoriaGasto.Text)
  Else
    CodigoCategoriaGasto = ""
  End If
  
  MontoSinIGV = Val(Var37(txtMontoCB.Text)) - Val(Var37(txtMontoImpuesto.Text))
  
  If txtTipDoc.Text = "02" Or txtTipDoc.Text = "04" Then
    txtMontoCB.Text = Val(Var37(txtMontoCB.Text)) - Val(Var37(txtMontoImpuesto.Text))
  End If
  
  If TipoEgreso = "PROY" Then
    ' Si es banco añade los datos de banco
    If optBanco.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(txtMontoCB) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "IdCta='" & msCtaCte & "'," & _
         "NumCheque='" & Trim(txtNumCheque.Text) & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='B', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      
      ' carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
          & Var37(txtMontoCB) & "¯" _
          & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
          & "¯" & "EGRESO BANCO CON AFECTACION" & "¯SC¯EB¯B"
    
    ElseIf optCaja.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(txtMontoCB) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='C', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      ' Carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
           & Var37(txtMontoCB) & "¯" _
           & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯C"
    
    ElseIf optRendir.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(txtMontoCB) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='R', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      ' Carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
           & Var37(txtMontoCB) & "¯" _
           & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯R"
    End If
  Else
    '/*/*/*/*/*/*
    '/*/*/*/*/*/* TIPO DE EGRESO PARA EMPRESA
    '/*/*/*/*/*/*
      ' Si es banco añade los datos de banco
    If optBanco.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(MontoSinIGV) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "IdCta='" & msCtaCte & "'," & _
         "NumCheque='" & Trim(txtNumCheque.Text) & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='B', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      
      ' carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        If txtTipDoc = "02" Then
          gcolAsiento.Add _
          Key:=msOrden, _
          Item:=msOrden & "¯" & txtTipDoc & "¯" _
            & Var37(MontoSinIGV) & "¯" _
            & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
            & "¯" & "EGRESO BANCO CON AFECTACION" & "¯SC¯EB¯B"
        Else
          gcolAsiento.Add _
          Key:=msOrden, _
          Item:=msOrden & "¯" & txtTipDoc & "¯" _
            & Var37(txtMontoCB) & "¯" _
            & msCtaCte & "¯" & FechaAMD(mskFecTrab.Text) _
            & "¯" & "EGRESO BANCO CON AFECTACION" & "¯SC¯EB¯B"
        End If
    
    ElseIf optCaja.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(MontoSinIGV) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='C', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      ' Carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        If txtTipDoc = "02" Then
          gcolAsiento.Add _
          Key:=msOrden, _
          Item:=msOrden & "¯" & txtTipDoc & "¯" _
             & Var37(MontoSinIGV) & "¯" _
             & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
             & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯C"
        Else
          gcolAsiento.Add _
          Key:=msOrden, _
          Item:=msOrden & "¯" & txtTipDoc & "¯" _
             & Var37(txtMontoCB) & "¯" _
             & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
             & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯C"
        End If
    
    ElseIf optRendir.Value = True Then
       ' Carga la sentencia que modifica el egreso general
      sSQL = "UPDATE EGRESOS SET " & _
         "IdProy='" & txtProy.Text & "'," & _
         "IdProg='" & txtProg.Text & "'," & _
         "IdLinea='" & txtLinea.Text & "'," & _
         "IdActiv='" & txtActiv.Text & "'," & _
         "NumDoc='" & Trim(txtDocEgreso) & "'," & _
         "IdTipoDoc='" & txtTipDoc & "'," & _
         "MontoAfectado=" & Var37(txtTotalDoc) & "," & _
         "MontoCB=" & Var37(MontoSinIGV) & "," & _
         "IdProveedor='" & gsIdProv & "'," & _
         "Observ='" & Trim(txtObserv.Text) & "'," & _
         "GiradoConImpuestos='" & fsTipoGiro & "'," & _
         "Origen='R', " & _
         "CODCATGASTO='" & CodigoCategoriaGasto & "', " & _
         "FecDoc='" & FechaAMD(mskFecDoc.Text) & "' " & _
         "WHERE Orden='" & txtCodEgreso.Text & "'"
      ' Carga la colección asiento
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
        gcolAsiento.Add _
        Key:=msOrden, _
        Item:=msOrden & "¯" & txtTipDoc & "¯" _
           & Var37(txtMontoCB) & "¯" _
           & "Nulo¯" & FechaAMD(mskFecTrab.Text) _
           & "¯" & "EGRESO CAJA CON AFECTACION" & "¯SC¯EC¯R"
    End If
  End If
  
  ' Ejecuta la sentencia que modifica el egreso
  modEgreso.SQL = sSQL
  If modEgreso.Ejecutar = HAY_ERROR Then End
  'Cierra la componente
  modEgreso.Cerrar
  
  If fsAveriguaOrigen = "ER" Then
    ' Verifica el origen anterior
    If msCaBaAnt = "CA" Then
      ' Crea el registro en entrega a rendir
      Var4 txtCodEgreso, txtRinde, "E", Var37(txtMontoCB), FechaAMD(mskFecTrab.Text), FechaAMD(mskFecDoc.Text)
    ElseIf msCaBaAnt = "ER" Then
      ' Modifica los datos de entregas a rendir
      Var3 txtCodEgreso, txtRinde, "E", Var37(txtMontoCB)
    End If
  ElseIf fsAveriguaOrigen = "CA" Then
    If msCaBaAnt = "ER" Then
      ' Elimina el movimiento de entregas a rendir
      Var2 txtCodEgreso
    End If
  End If
End Sub

Private Sub ModificarImpuestos()
  Dim modImpuestos As New clsBD3
  Dim sSQL As String
  
  ' Carga la sentencia que elimina de mov_impuestos los impuestos anteriores
  sSQL = "DELETE *  FROM MOV_IMPUESTOS" _
                & " WHERE Orden = '" & msOrden & "'"
  ' Ejecuta la sentencia
  modImpuestos.SQL = sSQL
  If modImpuestos.Ejecutar = HAY_ERROR Then End
  modImpuestos.Cerrar
  
  ' Guarda los impuestos
  GrabarImpuestos
End Sub

Private Sub ModificarDetalleEnGastosAlmacen()
  '------------------------------------------------------------
  ' Propósito:  Modifica los registros detalle en las tablas _
                gastos y en Almacen_Verificacion
  ' Recibe : Nada
  ' Entrega: Nada
  '------------------------------------------------------------
  Dim sSQL, sRegDetalle, sProdServ As String
  Dim modDetEgresoCajaBanco As New clsBD3
  Dim i As Integer
  
  On Error GoTo ErrClaveCol
  
  ' Define si lo que se guarda es producto o servicio
  sProdServ = DefinirProdServ
  
  'recorre el grid
  For i = 1 To grdDetalle.Rows - 1
    'carga el reg detalle para compararlo con el registro cargado en la colección
    '"idproducto", "Cantidad", "Monto",
    sRegDetalle = grdDetalle.TextMatrix(i, 4) & "¯" & grdDetalle.TextMatrix(i, 1) _
        & "¯" & grdDetalle.TextMatrix(i, 3)
    'verifica SI el registro se encuentra en la colección y SI se modificó
    ' SI NO se encuentra se inserta en ErrClaveCol: error 5
    If mcolEgreDet.Item(grdDetalle.TextMatrix(i, 4)) <> sRegDetalle Then
      ' se modifico el registro, entonces se actualiza la BD
      sSQL = "UPDATE GASTOS SET " & _
      "Cantidad=" & grdDetalle.TextMatrix(i, 1) & "," & _
      "Monto=" & Var37(grdDetalle.TextMatrix(i, 3)) & _
      " WHERE Orden='" & msOrden & "' and " _
           & "CodConcepto='" & grdDetalle.TextMatrix(i, 4) & "'"
      
      modDetEgresoCajaBanco.SQL = sSQL
      'ejecuta la sentencia que modifica el registro en gastos
      If modDetEgresoCajaBanco.Ejecutar = HAY_ERROR Then End
      modDetEgresoCajaBanco.Cerrar
      'eliminar el elmento modificado de la colección
      mcolEgreDet.Remove (grdDetalle.TextMatrix(i, 4))
    Else 'registro NO se modifico
      'Solo se elimina de la colección para seguir con los demas registros
      mcolEgreDet.Remove (grdDetalle.TextMatrix(i, 4))
    End If
  
PostErrClaveCol:
    
  ' carga la colección de asiento detalle Egreso
  ' Codcont, monto
  gcolAsientoDet.Add _
       Key:=grdDetalle.TextMatrix(i, 4), _
       Item:=grdDetalle.TextMatrix(i, 6) & "¯" & _
             Var37(grdDetalle.TextMatrix(i, 3))
     
  Next i
  
  'eliminar los que quedan en la colección
   ElimnarRegsDetEliminados
  '-------------------------------------------------------------------
ErrClaveCol:
  
      If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
          'el registro NO existe en el egreso con afectacion original
          'carga la sentencia que inserta el registro detalle en la base de datos
          sSQL = "INSERT INTO GASTOS VALUES('" & msOrden & "','" & sProdServ & "','" _
              & grdDetalle.TextMatrix(i, 4) & "'," & grdDetalle.TextMatrix(i, 1) & "," _
              & Var37(grdDetalle.TextMatrix(i, 3)) & ")"
          modDetEgresoCajaBanco.SQL = sSQL
          'ejecuta la sentencia que añade registro  a Gastos
          If modDetEgresoCajaBanco.Ejecutar = HAY_ERROR Then End
          modDetEgresoCajaBanco.Cerrar
          'VADICK QUITARON ESTA LINEA DE CODIGO Guarda el ingreso de productos en almacen, para la verificación
          If sProdServ = "P" Then GuardarVerif grdDetalle.TextMatrix(i, 4)
          Resume PostErrClaveCol ' La ejecución sigue por aquí
      End If
End Sub

Private Sub ElimnarRegsDetEliminados()
  Dim sSQL As String
  Dim modRegDetEgreso As New clsBD3
  Dim MiObjeto As Variant ' Variables de información.
    
  For Each MiObjeto In mcolEgreDet  ' Recorre los elementos que quedanen la colección
    'Elimina los registros de gastos
    sSQL = "DELETE * FROM GASTOS " _
         & "WHERE Orden ='" & msOrden & "'" _
         & " and CodConcepto='" & Var30(MiObjeto, 1) & "'"
    modRegDetEgreso.SQL = sSQL
    'ejecuta la sentencia que elimina los registros eliminados del egreso
    If modRegDetEgreso.Ejecutar = HAY_ERROR Then End
    modRegDetEgreso.Cerrar
    
    'Elimina los registros de almacen
    sSQL = "DELETE * FROM ALMACEN_VERIFICACION " _
         & "WHERE Orden ='" & msOrden & "'" _
         & " and IdProd='" & Var30(MiObjeto, 1) & "'"
    modRegDetEgreso.SQL = sSQL
    'ejecuta la sentencia que elimina los registros eliminados del egreso
    If modRegDetEgreso.Ejecutar = HAY_ERROR Then End
    modRegDetEgreso.Cerrar
  Next MiObjeto
End Sub

Private Sub CmdAgregarProdServ_Click()
  If (optProducto.Value = False) And (optServicio.Value = False) Then
    MsgBox "Debe Seleccionar Productos o Servicios.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
  Else
    If optProducto.Value = True Then
'      Guarda al formulario en la variable
'      gsFormulario = "078"
'
'      Carga en la BD el formulario en uso y usuario
'      Var42 ("078")
'
'      Verifica si es exclusivo y esta en uso
'      If Var47("078") Then
'        'Elimna la sesionn de la BD
'        Var43 "078"
'        'Termina la ejecucion del procedimiento
'        Exit Sub
'      End If
      
      'Muestra la pantalla de Mto. de Productos
      frmMNProducto.Show vbModal, Me
      
'      'Elimna la sesionn de la BD
'      Var43 "078"
      
      'Colecciones para la carga del combo de Productos
      Set mcolidprod = Nothing
      Set mcolCodDesProd = Nothing
      Set mcolDesMedidaContProd = Nothing
      
      CargarColProducto
      'Cambia la etiqueta del cboprocServ y a la columna 0 del grdDetalle
      lblProdServ = "Producto"
      grdDetalle.TextMatrix(0, 0) = "Producto"
      'Carga el cboProducto de acuerdo a la relación
      CargarCboCols cboProdServ, mcolCodDesProd
    Else
'      Guarda al formulario en la variable
'      gsFormulario = "061"
'
'      Carga en la BD el formulario en uso y usuario
'      Var42 ("061")
'
'      Verifica si es exclusivo y esta en uso
'      If Var47("061") Then
'        Elimna la sesionn de la BD
'        Var43 "061"
'        Termina la ejecucion del procedimiento
'        Exit Sub
'      End If
      
      'Muestra la pantalla de Mto. de Productos
      frmMNServicio.Show vbModal, Me
      
'      'Elimna la sesionn de la BD
'      Var43 "061"
      
      'Colecciones para la carga del combo de Servicios
      Set mcolCodServ = Nothing
      Set mcolCodDesServ = Nothing
      Set mcolDesMedidaContServ = Nothing
  
      CargarColServicio
      'Cambiamos la etiqueta del cboprocServ
      lblProdServ = "Servicio"
      grdDetalle.TextMatrix(0, 0) = "Servicio"
      'carga los servicios en el cbo
      CargarCboCols cboProdServ, mcolCodDesServ
    End If
  End If
End Sub

Private Sub cmdAnular_Click()
  Dim modAnularEgreCajaBanco As New clsBD3
  Dim sSQL As String
  
  'Verificar en almacén, si se puede anular el egreso CA
  If fbOkProductosAlmacen("Anular") = False Then Exit Sub
  
  'Verifica si el año esta cerrado y se puede modificar el codigo presupuestal
  'If Conta52(Right(mskFecTrab.Text, 4)) = True And OkCierreContableMod = False Then
  If Conta52(Right(mskFecTrab.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
    ' Sale
    Exit Sub
  End If
  
  'Preguntar si desea Anular el registro de Ingreso a Banco
  'Mensaje de conformidad de los datos
  If MsgBox("¿Seguro que desea anular el egreso " & msOrden & "?", _
                vbQuestion + vbYesNo, "Caja-Bancos- Egreso con Afectación") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario
    
    'Cambiar el campo Anulado de Ingresos a "SI"
    sSQL = "UPDATE EGRESOS SET " & _
       "Anulado='SI'" & _
       "WHERE Orden='" & msOrden & "'"
    
    'SI al ejecutar hay error se sale de la aplicación
    modAnularEgreCajaBanco.SQL = sSQL
    If modAnularEgreCajaBanco.Ejecutar = HAY_ERROR Then
     End
    End If
    'Se cierra la query
    modAnularEgreCajaBanco.Cerrar
     
    ' carga la colección asiento para anular
    If msCaBaAnt = "CA" Then
      ' carga la colección asiento para anular
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
      gcolAsiento.Add _
      Key:=msOrden, _
      Item:=msOrden & "¯Nulo¯Nulo¯Nulo¯Nulo¯EGRESO CAJA-BANCO CON AFECTACION¯SC¯EC¯C"
    ElseIf msCaBaAnt = "BA" Then
      ' carga la colección asiento para anular
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
      gcolAsiento.Add _
      Key:=msOrden, _
      Item:=msOrden & "¯Nulo¯Nulo¯Nulo¯Nulo¯EGRESO CAJA-BANCO CON AFECTACION¯SC¯EB¯B"
    ElseIf msCaBaAnt = "ER" Then
      ' carga la colección asiento para anular
      ' Orden,Monto,NumCtaBanc, fecha, observ, Proceso, Origen, optEgre
      ' Orden,Tip_Doc,Monto,NumCtaBanc, fecha, observ, Proceso,Origen,optEgre
      gcolAsiento.Add _
      Key:=msOrden, _
      Item:=msOrden & "¯Nulo¯Nulo¯Nulo¯Nulo¯EGRESO CAJA-BANCO CON AFECTACION¯SC¯EC¯R"
    End If
     
    'anula el asiento automatico
    Conta21
        
    ' Verifica el origen anterior
    If msCaBaAnt = "ER" Then
      ' Modifica los datos de entregas a rendir
      Var1 msOrden, msRindeAnt
    End If
        
    'Eliminar los productos en almacen_verificacion
    sSQL = "DELETE * FROM ALMACEN_VERIFICACION " _
         & "WHERE Orden ='" & msOrden & "'"
    
    ' Ejecuta la sentencia
    modAnularEgreCajaBanco.SQL = sSQL
    'ejecuta la sentencia que elimina los registros eliminados del egreso
    If modAnularEgreCajaBanco.Ejecutar = HAY_ERROR Then End
    modAnularEgreCajaBanco.Cerrar
     
    'Actualiza la transaccion
    Var8 -1, Empty
    
    ' Mensaje Ok
    MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso con Afectación"
    
    ' Limpia la pantalla para una nueva operación, Prepara el formulario
    ' Limpia la pantalla
    LimpiarFormulario
       
    If gsTipoOperacionEgreso = "Nuevo" Then
      ' Nuevo egreso
      NuevoEgreso
    Else
     ' cierra el control egreso
      If mbEgresoCargado Then
        mcurEgreso.Cerrar
        mbEgresoCargado = False
        mskFecTrab = "__/__/____"
      End If
    
     ' Se Modifican los datos del egreso
      ModificarEgreso
    End If
  End If
End Sub

Private Sub cmdAñadir_Click()
  '"Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "CodCont"

  If fsEstaenDetalle(msIdProdServ) = Empty Then
    ' Verifica el tipo de giro del documento
    If optGiroConImpuestos Then
      ' Añade un registro al grd tomando el valor de venta _
      y el precio unitario de venta
      grdDetalle.AddItem cboProdServ.Text & vbTab & txtCant & vbTab & _
                          txtPrecioUniVenta & vbTab & txtValorVenta & vbTab & _
                          msIdProdServ & vbTab & txtMedida & vbTab & msCodCont
    Else
      ' Añade un elemento al grd detalle tomando el valor _
      de compra y el precio unitario de compra
      
      If TipoEgreso = "EMPR" Then
        If txtTipDoc.Text <> "02" Then
          grdDetalle.AddItem cboProdServ.Text & vbTab & txtCant & vbTab & _
                          txtPrecioUniVenta & vbTab & txtValorVenta & vbTab & _
                          msIdProdServ & vbTab & txtMedida & vbTab & msCodCont
        Else
          grdDetalle.AddItem cboProdServ.Text & vbTab & txtCant & vbTab & _
                          txtPrecioUniCompra & vbTab & txtValorCompra & vbTab & _
                          msIdProdServ & vbTab & txtMedida & vbTab & msCodCont
        End If
      Else
        grdDetalle.AddItem cboProdServ.Text & vbTab & txtCant & vbTab & _
                          txtPrecioUniCompra & vbTab & txtValorCompra & vbTab & _
                          msIdProdServ & vbTab & txtMedida & vbTab & msCodCont
      End If
    End If

    ' Vaciar los controles de ingreso del detalle
    LimpiaControlesIngreDet
    
    ' Da el control al cboProdServ
    cboProdServ.SetFocus
    DetalleSubido = False
    DesmarcarFilaGRID grdDetalle
  Else
    ' Envia mensaje
    MsgBox cboProdServ & " , ha sido anteriormente ingresado" & Chr(13) & _
           "Debe elegir nuevamente", _
            , "SGCcaijo- Egreso con afectación "
    ' limpia el cbo cuenta para  dar opcion a elegir
    cboProdServ.SetFocus
    cmdAñadir.Enabled = False
  End If
  
  ' Habilita el boton aceptar
  HabilitarBotonAceptar
End Sub

Private Sub cmdAñadir_GotFocus()
  ' Verificar si se a introducido un tipo de documento
  If fbOperarDetalle = False Then
      Exit Sub
  End If
End Sub

Private Sub cmdBuscar_Click()
  ' Carga los títulos del grid selección
  giNroColMNSel = 8
  aTitulosColGrid = Array("Nro.RUC/DNI", " IdProv.", "Descripción", "RUC/DNI", "Dirección", "Teléfono", "Fax", "Representante")
  aTamañosColumnas = Array(1500, 1000, 4500, 800, 3500, 2000, 2000, 3000)
  txtRUCDNI.Text = Empty
  
  frmMNSelecProvCaja.OrigenBusqueda = "CON AFECTACION"
  'Muestra el formulario de busqueda de proveedores
  frmMNSelecProvCaja.Show vbModal, Me
  txtDocEgreso.SetFocus
End Sub

Private Sub cmdBuscarEgreso_Click()
  ' Define el tipo de selección del Orden
  gsTipoSeleccionOrden = "EgresoCA"
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
    ' Prepara el formulario
    NuevoEgreso
  Else
    ' cierra el control egreso
    If mbEgresoCargado Then
      mcurEgreso.Cerrar
      mbEgresoCargado = False
      mskFecTrab = "__/__/____"
    End If
    ' Prepara el formulario
    ModificarEgreso
  End If
End Sub

Private Sub cmdEliminar_Click()
  ' Elimina la fila selccionada del Grid
  If grdDetalle.CellBackColor = vbDarkBlue And grdDetalle.Row > 0 Then
    If DetalleSubido Then
      MsgBox "Se tiene un detalle subido para editar, No se puede eliminar el detalle.", , "SGCcaijo - Egreso con afectación"
    Else
      ' elimina la fila seleccionada del grid
      If grdDetalle.Rows > 2 Then
        ' elimina la fila seleccionada del grid
        grdDetalle.RemoveItem grdDetalle.Row
      Else
        ' estable vacío el grddetalle
        grdDetalle.Rows = 1
      End If
    End If
  End If
  
  If grdDetalle.Rows < 2 Then
    cmdEliminar.Enabled = False
  End If
  
  ' Actualiza la posición del grid
  ipos = 0
  
  ' Habilita el botón aceptar
  HabilitarBotonAceptar
End Sub

Private Sub cmdPActiv_Click()
  If cboActiv.Enabled Then
    ' alto
     cboActiv.Height = CBOALTO
    ' focus a cbo
    cboActiv.SetFocus
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

Private Sub cmdPCategoriaGasto_Click()
  If cboCategoriaGasto.Enabled Then
    ' alto
     cboCategoriaGasto.Height = CBOALTO
    ' focus a cbo
    cboCategoriaGasto.SetFocus
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

Private Sub cmdPLinea_Click()
  If cboLinea.Enabled Then
    ' alto
     cboLinea.Height = CBOALTO
    ' focus a cbo
    cboLinea.SetFocus
  End If
End Sub

Private Sub cmdPProdServ_Click()
  If cboProdServ.Enabled Then
    ' alto
     cboProdServ.Height = CBOALTO
    ' focus a cbo
    cboProdServ.SetFocus
  End If
End Sub

Private Sub cmdPProg_Click()
  If cboProg.Enabled Then
    ' alto
     cboProg.Height = CBOALTO
    ' focus a cbo
    cboProg.SetFocus
  End If
End Sub

Private Sub cmdPProy_Click()
  If cboProy.Enabled Then
    ' alto
     cboProy.Height = CBOALTO
    ' focus a cbo
    cboProy.SetFocus
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

Private Sub cmdSalir_Click()
  'cierra el formulario
  Unload Me
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

Private Sub cmdSelImpuestos_Click()

' Copia el valor de txtTotalDoc a gdblMontoTotal
gdblMontoTotal = Val(Var37(txtTotalDoc.Text))

'If TipoEgreso = "PROY" Then
'  ' Define el tipo de tributo documento
'  gsRelacTributo = Var30(mcolCodRetencionPaga.Item(txtTipDoc.Text), 1)
'ElseIf (TipoEgreso = "EMPR") Then
'  If ((txtTipDoc.Text = "01") Or (txtTipDoc.Text = "04") Or (txtTipDoc.Text = "12") Or (txtTipDoc.Text = "16")) Then
'    gsRelacTributo = "Retiene"
'  Else
'    ' Define el tipo de tributo documento
'    gsRelacTributo = Var30(mcolCodRetencionPaga.Item(txtTipDoc.Text), 1)
'  End If
'End If

' Define el tipo de tributo documento
  gsRelacTributo = Var30(mcolCodRetencionPaga.Item(txtTipDoc.Text), 1)
  
' borra los controles del ingreso a detalle en la segunda parte _
 del formulario, tambien reset al resumen
 LimpiaControlesIngreDet
 LimpiaCtrlsResumen

' Muestra el formulario para elegir las retenciones
frmCBEGSelImp.Show vbModal, Me


' Averigua el Monto total de impuestos relacionados a el documento
mdblMontoTotalImpt = SumaMontosImpuestos

' Averigua la suma de los porcentajes aplicados a el documento
mdblSumImpt = SumaImpuestosAplicados

' Muestra el monto relacionado a los impuestos
txtMontoImpuesto.Text = Format(mdblMontoTotalImpt, "###,###,##0.00")

' Muestra el monto que saldrá de caja o de bancos cuando se cambie los impuestos
txtMontoCB.Text = Format(CalculaMontoEgresoCB, "###,###,##0.00")

' Si al entrar cambió los impuestos aplicados
If gbImpuestos = True Then SendKeys vbTab

' habilitar botón aceptar
HabilitarBotonAceptar

End Sub

Private Function CalculaMontoEgresoCB() As Double
'-----------------------------------------------------------------
' Propósito: Calcula el monto real de egreso de Caja-Bancos
' Recibe: Nada
' Entrega: Nada
'-----------------------------------------------------------------
Dim sRelacImpuestos  As String

'  If TipoEgreso = "PROY" Then
'    ' Asigna a la variable la relación del documento con el impuesto
'    sRelacImpuestos = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1)
'  ElseIf (TipoEgreso = "EMPR") Then
'    If ((txtTipDoc.Text = "01") Or (txtTipDoc.Text = "04") Or (txtTipDoc.Text = "12") Or (txtTipDoc.Text = "16")) Then
'      sRelacImpuestos = "Retiene"
'    Else
'      ' Asigna a la variable la relación del documento con el impuesto
'      sRelacImpuestos = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1)
'    End If
'  End If
  
  ' Asigna a la variable la relación del documento con el impuesto
  sRelacImpuestos = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1)
    
  ' De acuerdo con el impuesto asigna el egreso de Caja-Bancos
  Select Case sRelacImpuestos
    Case "Paga", "Retiene"
      CalculaMontoEgresoCB = Val(Var37(txtTotalDoc) - mdblMontoTotalImpt)
    Case "Registra"
      CalculaMontoEgresoCB = Val(Var37(txtTotalDoc))
  End Select
End Function

Private Function SumaMontosImpuestos() As Double
Dim varItem As Variant
Dim dblMonto As Double

' Inicializa la variable que acumula los montos
dblMonto = 0

' Recorre la colección que almacena los impuestos aplicados
For Each varItem In gcolImpSel
  ' Acumula los montos de los impuestos
    dblMonto = dblMonto + Val(Var30(varItem, 3))
Next varItem

' Devuelve el resultado
SumaMontosImpuestos = dblMonto

End Function

Private Function SumaImpuestosAplicados() As Double
Dim varItem As Variant
Dim dblPorcentajes As Double

' Inicializa la variable que acumula los montos
dblPorcentajes = 0

' Recorre la colección que almacena los impuestos aplicados
For Each varItem In gcolImpSel
  ' Acumula los montos de los impuestos
    dblPorcentajes = dblPorcentajes + (Val(Var30(varItem, 2)) / 100)
Next varItem

' Devuelve el resultado
SumaImpuestosAplicados = dblPorcentajes

End Function

Private Sub Form_Load()
Dim sSQL As String
 
TipoEgreso = ""
DetalleSubido = False
'Se carga el combo de Bancos
 sSQL = ""
sSQL = "SELECT DISTINCT b.IdBanco,b.DescBanco FROM TIPO_BANCOS B , TIPO_CUENTASBANC C" _
       & " WHERE b.idbanco = c.idbanco And c.idmoneda = 'SOL'" _
       & " ORDER BY DescBanco"
 CD_CargarColsCbo cboBanco, sSQL, mcolCodBanco, mcolCodDesBanco
 
 'Se carga el combo de Cta Cte
 sSQL = ""
 sSQL = "SELECT IdCta, DescCta FROM TIPO_CUENTASBANC " & _
        "WHERE IdMoneda= 'SOL'   ORDER BY DescCta"
 CD_CargarColsCbo cboCtaCte, sSQL, mcolCodCtaCte, mcolCodDesCtaCte
 'Limpia el combo
 cboCtaCte.Clear

' Carga Tipos de documentos
CargarTipoDocumentos

'Carga la colección de producto
CargarColProducto

'Carga la colección de servicio
CargarColServicio

'Carga la colección con datos de los proveedores
CargaColDatos

'Se carga el combo de Categorias de Gasto
sSQL = ""
sSQL = "SELECT CODCATGASTO, CODCATGASTO + '   ' + DESCRIPCIONGASTO " _
        & " FROM CATEG_GASTO " _
        & " ORDER BY CODCATGASTO "
CD_CargarColsCbo cboCategoriaGasto, sSQL, mcolCodCatGasto, mcolCodDesCatGasto

'Coloca el titulo al Grid
aTitulosColGrid = Array("Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "CodCont")
aTamañosColumnas = Array(3900, 1400, 1400, 1400, 0, 0, 0)
CargarGridTitulos grdDetalle, aTitulosColGrid, aTamañosColumnas

'Establece campos obligatorios de la primera parte del formulario
EstableceCamposObligatorios1raParte

'Establece campos obligatorios de la segunda parte del formulario
EstableceCamposObligatorios2daParte

' Dependiendo de la operación a realizar prepara el formulario
If gsTipoOperacionEgreso = "Nuevo" Then
    ' Deshabilita el txtCodEgreso
    txtCodEgreso.Enabled = False
    ' Deshabilita el botón elegir
    cmdBuscarEgreso.Enabled = False
    
    ' Coloca la fecha del sistema
    mskFecTrab.Text = gsFecTrabajo
    
    ' Carga las colecciones necesarias para el manejo del formulario
    ' Carga Proyectos en el nuevo
    sSQL = "SELECT idproy, IdProy + '   ' + descproy  " & _
           " FROM Proyectos WHERE idproy IN " & _
           "(SELECT idproy FROM Presupuesto_proy ) And " & _
           "'" & FechaAMD(mskFecTrab.Text) & "' BETWEEN FecInicio And FecFin " & _
           "ORDER BY idproy "
    CD_CargarColsCbo cboProy, sSQL, mcolCodProy, mcolCodDesProy

    ' Prepara el formulario para un nuevo egreso
    NuevoEgreso
Else
    ' Inicializa la variables
    mbEgresoCargado = False
    
    'Prepara el formulario para modificar el egreso
    ModificarEgreso
End If

End Sub

Private Sub HabilitarBotonAceptar()
cmdAceptar.Enabled = False
 
' Verifica si se a introducido los datos obligatorios generales
If txtProy.BackColor <> vbWhite Or txtProg.BackColor <> vbWhite Or _
   txtLinea.BackColor <> vbWhite Or txtActiv.BackColor <> vbWhite Or _
   txtTipDoc.BackColor <> vbWhite Or txtRUCDNI.BackColor <> vbWhite Or _
   txtDocEgreso.BackColor <> vbWhite Or txtTotalDoc.BackColor <> vbWhite Or _
   grdDetalle.Rows <= 1 Or gbImpuestos = False Or gsIdProv = Empty Or TxtCategoriaGasto.BackColor <> vbWhite Or mskFecDoc.BackColor <> vbWhite Then
   ' algún obligatorio falta ser introducido
   cmdAceptar.Enabled = False
   Exit Sub
Else
   ' Verifica que se haigan introducido los datos obligatorios de bancos
   If optBanco.Value = True And (txtBanco.BackColor <> vbWhite Or cboCtaCte.BackColor <> vbWhite _
      Or txtNumCheque.BackColor <> vbWhite) Then
     ' algún obligatorio de banco falta ser introducido
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
  If TxtCategoriaGasto.Visible = True Then
    If optCaja.Value = True Or optRendir.Value = True Then
      If TxtCategoriaGasto.Text <> mcurEgreso.campo(17) Then
        ' Habilita botón aceptar
        cmdAceptar.Enabled = True
      End If
    End If
    
    If optBanco.Value = True Then
      If TxtCategoriaGasto.Text <> mcurEgreso.campo(19) Then
        ' Habilita botón aceptar
        cmdAceptar.Enabled = True
      End If
    End If
  End If
  
    ' Verifica si se cambio los datos generales
   If fbCambioDatosGenerales = False Then
        If fbCambioImpuestos = False Then
            If fbCambioDetalle = False Then
                ' No se cambio ningún dato
                Exit Sub
            End If
        End If
   End If
End If

' Habilita botón aceptar
cmdAceptar.Enabled = True

End Sub

Private Function fsAveriguaOrigen() As String
'----------------------------------------------------------------
' Propósito : Averigua el origen del egreso
' Recibe : Nada
' Entrega : Nada
'----------------------------------------------------------------
' Verifica las opciones origen del egreso
If optBanco.Value = True Then
    ' devuelve el origen del egreso
    fsAveriguaOrigen = "BA"
ElseIf optCaja.Value = True Then
    ' devuelve el origen del egreso
    fsAveriguaOrigen = "CA"
ElseIf optRendir.Value = True Then
    ' devuelve el origen del egreso
    fsAveriguaOrigen = "ER"
Else
    ' devuelve el origen del egreso
    fsAveriguaOrigen = Empty
End If

End Function

Private Function fbCambioDatosGenerales() As Boolean
    fbCambioDatosGenerales = False
    
' Verifica si se cambio el origen
If fsAveriguaOrigen <> msCaBaAnt Then
   ' Cambio el origen de los datos
    fbCambioDatosGenerales = True
    Exit Function
End If

' verifica los datos de caja
If txtProy.Text <> mcurEgreso.campo(1) _
Or txtProg.Text <> mcurEgreso.campo(2) _
Or txtLinea.Text <> mcurEgreso.campo(3) _
Or txtActiv.Text <> mcurEgreso.campo(4) _
Or Trim(txtDocEgreso.Text) <> mcurEgreso.campo(5) _
Or txtTipDoc.Text <> mcurEgreso.campo(6) _
Or Val(Var37(txtTotalDoc.Text)) <> mcurEgreso.campo(8) _
Or Val(Var37(txtMontoCB.Text)) <> mcurEgreso.campo(9) _
Or gsIdProv <> mcurEgreso.campo(12) _
Or Trim(txtObserv.Text) <> mcurEgreso.campo(14) _
Or fsTipoGiro <> mcurEgreso.campo(15) Or mskFecDoc.Text <> mcurEgreso.campo(18) Then
    ' cambio datos generales
    fbCambioDatosGenerales = True
    Exit Function
End If

' Verifica si se cambio Cta corriente y número de cheque
If fsAveriguaOrigen = "BA" Then
    If msCtaCte <> mcurEgreso.campo(17) _
    Or Trim(txtNumCheque) <> mcurEgreso.campo(18) Or mskFecDoc.Text <> mcurEgreso.campo(20) Then
        ' Cambió cta corriente
            ' cambio datos generales
            fbCambioDatosGenerales = True
            Exit Function
    End If
End If

' Verifica si se cambio de cuenta a rendir
If fsAveriguaOrigen = "ER" Then
    If txtRinde <> msRindeAnt Or mskFecDoc.Text <> mcurEgreso.campo(18) Then
        ' cambio datos generales
        fbCambioDatosGenerales = True
        Exit Function
    End If
End If

End Function

Private Function fbCambioImpuestos() As Boolean
On Error GoTo mnjError:

Dim sReg As Variant
' Inicializa la función asumiendo que no se modificó impuestos
fbCambioImpuestos = False

' Verifica las cantidades de los impuestos aplicados
If gcolImpSel.Count <> mcolEgreImpt.Count Then
    ' Se ha cambiado los impuestos
    fbCambioImpuestos = True
    Exit Function
Else
    ' Verifica si se cambió los impuestos uno a uno, recorre el cursor
    For Each sReg In mcolEgreImpt
      ' Compara los registros de impuestos
        If sReg <> gcolImpSel.Item(Var30(sReg, 1)) Then
            ' Se ha cambiado los impuestos
            fbCambioImpuestos = True
        End If
    Next sReg

End If

mnjError:
'-------------------------------------------------------------------

    If Err.Number = 5 Then ' Error elemento no existe
            ' Se ha cambiado los impuestos
            fbCambioImpuestos = True
            Exit Function
    End If

End Function

Private Function fbCambioDetalle() As Boolean
' --------------------------------------------------------------
' Propósito : Verifica si se cambió algún dato del detalle
' Recibe : Nada
' Entrega : Nada
' --------------------------------------------------------------
On Error GoTo mnjError:

Dim sReg As String
Dim i As Integer

' Inicializa la función asumiendo que no se modificó detalle
fbCambioDetalle = False

' Verifica las cantidades del detalle
If grdDetalle.Rows - 1 <> mcolEgreDet.Count Then
    ' Se ha cambiado los impuestos
    fbCambioDetalle = True
    Exit Function
Else
    ' Verifica si se cambió el detalle, recorre el cursor
  If mcolEgreDet.Count = 0 Then
    ' Sale de la función
     Exit Function
  Else
    ' recorre el grid detalle
    For i = 1 To grdDetalle.Rows - 1
      ' Compara los registros de detalle
      ' Carga registro orignal, "codConcepto", "cantidad", "Monto"
'"Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "CodCont")
        sReg = grdDetalle.TextMatrix(i, 4) & "¯" & grdDetalle.TextMatrix(i, 1) _
             & "¯" & grdDetalle.TextMatrix(i, 3)
        If sReg <> mcolEgreDet.Item(grdDetalle.TextMatrix(i, 4)) Then
            ' Se ha cambiado el detalle
            fbCambioDetalle = True
        End If
    Next i

  End If
End If
'-----------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' error , elemento no encontrado
        ' Se ha cambiado el detalle
         fbCambioDetalle = True
         Exit Function
    End If
End Function

Private Function OkCierreContableMod() As Boolean
If fsAveriguaOrigen <> msCaBaAnt Then
   ' Cambio el origen de los datos
    OkCierreContableMod = False
    Exit Function
End If

' verifica los datos de caja
If Trim(txtDocEgreso.Text) <> mcurEgreso.campo(5) _
Or txtTipDoc.Text <> mcurEgreso.campo(6) _
Or Val(Var37(txtTotalDoc.Text)) <> mcurEgreso.campo(8) _
Or Val(Var37(txtMontoCB.Text)) <> mcurEgreso.campo(9) _
Or gsIdProv <> mcurEgreso.campo(12) _
Or Trim(txtObserv.Text) <> mcurEgreso.campo(14) _
Or fsTipoGiro <> mcurEgreso.campo(15) Then
    ' cambio datos generales
    OkCierreContableMod = False
    Exit Function
End If

' Verifica si se cambio el detalle
If fbCambioImpuestos = True Then
    ' cambio datos Impuestos
    OkCierreContableMod = False
    Exit Function
End If


' Verifica si se cambio el detalle
If fbCambioDetalle = True Then
    ' cambio datos del detalle
    OkCierreContableMod = False
    Exit Function
End If

' TOdo Ok
OkCierreContableMod = True

End Function

Private Function fbVerificarDatosIntroducidos()
Dim dblMontoDet As Double
Dim RestaTotalImpuesto As Double
Dim Temp1 As Double
Dim Temp2 As Double

' Verifica si todos los datos estan correctos
If gsTipoOperacionEgreso = "Nuevo" Then
    'Verifica si el año esta cerrado
    If Conta52(Right(mskFecTrab.Text, 4)) = True Then
        ' No se puede realizar la operación
        MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
        vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
        'Devuelve el resultado
        fbVerificarDatosIntroducidos = False
        Exit Function
    End If
ElseIf gsTipoOperacionEgreso = "Modificar" Then
    'Verifica si el año esta cerrado y se puede modificar el codigo presupuestal
    If Conta52(Right(mskFecTrab.Text, 4)) = True And OkCierreContableMod = False Then
        ' No se puede realizar la operación
        MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
        vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
        'Devuelve el resultado
        fbVerificarDatosIntroducidos = False
        Exit Function
    End If
End If

' Verifica que lo que sale de caja-bancos sea Menor que el saldo de Caja-bancos
 If gsTipoOperacionEgreso = "Nuevo" Then
    If Val(Var37(txtMontoCB)) > Val(Var37(txtSaldoCB)) Then
          ' Mensaje ,saldo insuficiente
        MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso con afectación"
        fbVerificarDatosIntroducidos = False
        If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
        Exit Function
     End If
 Else
    ' verifica la conformidad con el saldo
    If optCaja.Value = True Then   'Caja
      If msCaBaAnt = "CA" Then ' Caja anterior
        If (Val(Var37(txtMontoCB)) - Val(mcurEgreso.campo(9))) > Val(Var37(txtSaldoCB)) Then
           ' Mensaje ,saldo insuficiente
            MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso con afectación"
            fbVerificarDatosIntroducidos = False
            If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
            Exit Function
        End If
      Else ' Rendir anterior
            If Val(Var37(txtMontoCB)) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
              MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso con afectación"
              fbVerificarDatosIntroducidos = False
              If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
              Exit Function
            End If
      End If
    ElseIf optBanco.Value = True Then
        If mcurEgreso.campo(17) = msCtaCte Then  'La misma CtaCte
            If (Val(Var37(txtMontoCB)) - Val(mcurEgreso.campo(9))) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
               MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso con afectación"
               fbVerificarDatosIntroducidos = False
               If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
               Exit Function
            End If
        ElseIf mcurEgreso.campo(17) <> msCtaCte Then 'Se Cambio de CtaCte
            If Val(Var37(txtMontoCB)) > Val(Var37(txtSaldoCB)) Then
               ' Mensaje ,saldo insuficiente
               MsgBox "El monto de egreso excede el saldo de Caja-Bancos", , "SGCcaijo-Egreso con afectación"
               fbVerificarDatosIntroducidos = False
               If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
              Exit Function
            End If
        End If
      ElseIf optRendir.Value = True Then ' Verifica cuentas a rendir
       If msCaBaAnt = "ER" Then ' Anterior rendir
            If msRindeAnt = txtRinde Then ' Misma cuenta a rendir
                If (Val(Var37(txtMontoCB)) - Val(mcurEgreso.campo(9))) > Val(Var37(txtSaldoCB)) Then
                   ' Mensaje ,saldo insuficiente
                    MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso con afectación"
                    fbVerificarDatosIntroducidos = False
                    If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
                    Exit Function
                End If
            ElseIf msRindeAnt <> txtRinde Then
                If Val(Var37(txtMontoCB)) > Val(Var37(txtSaldoCB)) Then
                  ' Mensaje ,saldo insuficiente
                  MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso con afectación"
                  fbVerificarDatosIntroducidos = False
                  If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
                  Exit Function
                End If
            End If
       Else ' Anterior Caja
            If Val(Var37(txtMontoCB)) > Val(Var37(txtSaldoCB)) Then
              ' Mensaje ,saldo insuficiente
              MsgBox "El monto de egreso excede el saldo de Cuenta a rendir", , "SGCcaijo-Egreso con afectación"
              fbVerificarDatosIntroducidos = False
              If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
              Exit Function
            End If
       End If
    End If ' Fin de verificar el origen
 End If ' Fin de verifica nuevo o modificar

If TipoEgreso = "EMPR" Then
  If txtTipDoc.Text <> "02" Then
    ' Verifica que el total del detalle sea Igual al total del documento
    If Val(SumarMontoDetalle) <> (Val(Var37(txtTotalDoc)) - Val(Var37(txtMontoImpuesto))) Then
      If MsgBox("¿Se va a Corregir la diferencia en los SubTotales y Precio Unitario, Esta de Acuerdo?", _
          vbQuestion + vbYesNo, "Caja-Bancos, Egreso con afectación") = vbYes Then
              grdDetalle.TextMatrix(grdDetalle.Rows - 1, 3) = Format(Val(Var37(grdDetalle.TextMatrix(grdDetalle.Rows - 1, 3))) + ((Val(Var37(txtTotalDoc)) - Val(Var37(txtMontoImpuesto))) - Val(Var37(SumarMontoDetalle))), "###,###,##0.00")
              grdDetalle.TextMatrix(grdDetalle.Rows - 1, 2) = Format(Val(Var37(grdDetalle.TextMatrix(grdDetalle.Rows - 1, 3))) / Val(Var37(grdDetalle.TextMatrix(grdDetalle.Rows - 1, 1))), "###,###,##0.00")
      Else
        Exit Function
      End If
    End If
  Else
    ' Verifica que el total del detalle sea Igual al total del documento
    If Val(SumarMontoDetalle) <> Val(Var37(txtTotalDoc)) Then
      ' Mensaje , Montos diferentes
         MsgBox "El monto total del doc de egreso: " & txtTotalDoc.Text & " debe ser Igual a" & Chr(13) _
                & "la suma de los montos de los productos o servicios: " & Format(SumarMontoDetalle, "###,###,##0.00") _
                , , "SGCcaijo-Egreso con afectación"
         fbVerificarDatosIntroducidos = False
         If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
         Exit Function
    End If
  End If
Else
  ' Verifica que el total del detalle sea Igual al total del documento
  If Val(SumarMontoDetalle) <> Val(Var37(txtTotalDoc)) Then
      ' Mensaje , Montos diferentes
         MsgBox "El monto total del doc de egreso: " & txtTotalDoc.Text & " debe ser Igual a" & Chr(13) _
                & "la suma de los montos de los productos o servicios: " & Format(SumarMontoDetalle, "###,###,##0.00") _
                , , "SGCcaijo-Egreso con afectación"
         fbVerificarDatosIntroducidos = False
         If txtMontoCB.Enabled = True Then txtMontoCB.SetFocus
        Exit Function
  End If
End If

' Verificados los datos
fbVerificarDatosIntroducidos = True

End Function


Private Function SumarMontoDetalle() As Double
Dim i As Integer
Dim dblSuma As Double

' inicializa la suma
dblSuma = 0

' recorre el grd
For i = 1 To grdDetalle.Rows - 1
  dblSuma = dblSuma + Val(Var37(grdDetalle.TextMatrix(i, 3)))
Next i

' Asigna la función
SumarMontoDetalle = dblSuma

End Function

Private Sub NuevoEgreso()
'---------------------------------------------------------------
'Propósito : Prepara el formulario para un egreso con afectación
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
Dim sSQL As String

' Inicializa la variable codigo de cuenta
  msCtaCte = Empty
  msOrden = Empty

' Inicializa el grid
  ipos = 0
  gbCambioCelda = False

' Muestra los resumen
  LimpiaCtrlsResumen
  txtSaldoCB = "0.00"

'Se carga la colección de E Rendir
  Var5

' Pone por defecto el egreso de caja y calcula el Orden
If optCaja.Value = True Then
    'realiza el evento de optclick
    optCaja_Click
Else
    ' cambia el valor del optCaja.value
    optCaja.Value = True
End If

' Pone por defecto el  opt Girado sin impuestos
If optGiroSinImpuestos.Value = True Then
    'realiza el evento de optclick
    optGiroSinImpuestos_Click
Else
    ' cambia el valor del optGiroSinImpuestos
    optGiroSinImpuestos.Value = True
End If

'' Inicializa optPaga
  optProducto.Value = True
  optServicio.Value = False

' Limpia las colecciones
   Set gcolImpSel = Nothing
   gbImpuestos = False

' deshabilita los botones del formulario
HabilitaDeshabilitaBotones ("Nuevo")

End Sub

Private Sub HabilitaDeshabilitaBotones(sProceso As String)
Select Case sProceso

' depende del proceso habilita y deshabilita botones
Case "Nuevo"
    cmdSelImpuestos.Enabled = False
    cmdAñadir.Enabled = False
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    cmdEliminar.Enabled = False
    
Case "Modificar"
    cmdBuscarEgreso.Enabled = True
    cmdSelImpuestos.Enabled = False
    cmdAñadir.Enabled = False
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    cmdEliminar.Enabled = False
    
End Select

End Sub

Private Sub ModificarEgreso()
'---------------------------------------------------------------
'Propósito : Prepara el formulario para modificar un egreso
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------

' Habilita el txtCodEgreso
  txtCodEgreso = Empty
  txtCodEgreso.Enabled = True
  txtCodEgreso.BackColor = Obligatorio
  
' Inicializa el grid
  ipos = 0
  gbCambioCelda = False

'Se carga la colección de E Rendir
  Var5

' Inicializa los optCB
  optCaja.Value = False
  optBanco.Value = False
  optRendir.Value = False
  optCaja.Enabled = False
  optBanco.Enabled = False
  optRendir.Enabled = False
  
' Inicializa optPaga
  optProducto.Value = False
  optServicio.Value = False

' Pone por defecto el  opt Girado sin impuestos
 If optGiroSinImpuestos.Value = True Then
    'realiza el evento de optclick
    optGiroSinImpuestos_Click
 Else
    ' cambia el valor del optGiroSinImpuestos
    optGiroSinImpuestos.Value = True
 End If

' Deshabilita la 1raParte del formulario
  DeshabilitarHabilitarFormulario False

' Oculta controles
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 lblNumCheque.Visible = False: txtNumCheque.Visible = False

 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False

' Limpia las colecciones
   Set mcolEgreDet = Nothing
   Set mcolEgreImpt = Nothing

'Colecciones para la carga del combo de Proyectos
  Set mcolCodProy = Nothing
  Set mcolCodDesProy = Nothing
 
  Set gcolImpSel = Nothing
  gbImpuestos = False

' Inicializa la variable codigo de cuenta
  msCtaCte = Empty
  msOrden = Empty
  msCaBaAnt = Empty
  gsOrden = Empty
   
' Muestra los resumen
  LimpiaCtrlsResumen
  txtSaldoCB = "0.00"
  
' Maneja estado de los botones del formulario
  HabilitaDeshabilitaBotones "Modificar"

End Sub

Private Sub DeshabilitarHabilitarFormulario(bBoleano As Boolean)
txtProy.Enabled = bBoleano: cboProy.Enabled = bBoleano
txtProg.Enabled = bBoleano: cboProg.Enabled = bBoleano
txtLinea.Enabled = bBoleano: cboLinea.Enabled = bBoleano
txtActiv.Enabled = bBoleano: cboActiv.Enabled = bBoleano
txtTipDoc.Enabled = bBoleano: cboTipDoc.Enabled = bBoleano
txtRUCDNI.Enabled = bBoleano: txtNombrProv.Enabled = bBoleano
txtDocEgreso.Enabled = bBoleano
txtObserv.Enabled = bBoleano
txtTotalDoc.Enabled = bBoleano
txtBanco.Enabled = bBoleano: cboBanco.Enabled = bBoleano
cboCtaCte.Enabled = bBoleano: txtNumCheque.Enabled = bBoleano
fraPagar.Enabled = bBoleano: fraTipoGiro.Enabled = bBoleano
cboProdServ.Enabled = bBoleano
txtCant.Enabled = bBoleano
txtValorVenta.Enabled = bBoleano: txtValorCompra.Enabled = bBoleano
grdDetalle.Enabled = bBoleano
txtRinde.Enabled = bBoleano: txtDescRinde.Enabled = bBoleano: cmdBuscaRinde.Enabled = bBoleano
mskFecDoc.Enabled = bBoleano
End Sub

Private Sub CargarTipoDocumentos()
'---------------------------------------------------------------
'Propósito : Carga las colecciónes de tipo de documentos
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim curRetencionPaga As New clsBD2

'Se carga el combo de Tipo de Documento , y SI generan retenciones
sSQL = ""
sSQL = "SELECT IdTipoDoc, DescTipoDoc,RelacTributo,RelacPaga " _
     & "FROM TIPO_DOCUM WHERE RelacProc='SA' AND IdTipoDoc<>'40' ORDER BY DescTipodoc"
CD_CargarColsCbo cboTipDoc, sSQL, mcolCodTipDoc, mcolCodDesTipDoc

sSQL = ""
sSQL = "SELECT IdTipoDoc, DescTipoDoc,RelacTributo,RelacPaga " _
     & "FROM TIPO_DOCUM WHERE RelacProc='SA' ORDER BY DescTipodoc"
curRetencionPaga.SQL = sSQL

'carga la colección de CodigoDoc y Si genera retencion y si paga Servicios o Productos
If curRetencionPaga.Abrir = HAY_ERROR Then End
Do While Not curRetencionPaga.EOF
   mcolCodRetencionPaga.Add Item:=curRetencionPaga.campo(2) & "¯" & curRetencionPaga.campo(3), Key:=curRetencionPaga.campo(0)
   curRetencionPaga.MoverSiguiente
Loop

'cierra el cursor que carga datos en las colecciones de documento
curRetencionPaga.Cerrar

End Sub

Private Sub LimpiarFormulario()
'---------------------------------------------------------------
'Propósito : Limpia el formulario de egreso
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
' Limpia la 1ra parte formulario
  LimpiarPrimeraParteFormulario
' Limpia la 2da parte formulario
  LimpiarSegundaParteFormulario
' Reset resumen
  LimpiaCtrlsResumen

End Sub

Private Sub CargarColProducto()
'---------------------------------------------------------------
'Propósito : Carga la colección de Productos con su medida
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curMedidaProd As New clsBD2

sSQL = "SELECT idprod, DescProd, Medida, CodCont, Tipo " & _
            "FROM PRODUCTOS " & _
            "ORDER BY DescProd "
            
'sSQL = "SELECT idprod, DescProd, Medida, CodCont FROM PRODUCTOS ORDER BY DescProd"

'Carga la colección de descripcion y medida de los productos
curMedidaProd.SQL = sSQL
If curMedidaProd.Abrir = HAY_ERROR Then
  End
End If
Do While Not curMedidaProd.EOF
    ' Se carga la colección de descripciones + unidades de los productos con la 1º y 2º
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolDesMedidaContProd.Add Key:=curMedidaProd.campo(0), _
                              Item:=curMedidaProd.campo(2) & "¯" & curMedidaProd.campo(3)
    
    'colección de producto y su descripción
    mcolidprod.Add curMedidaProd.campo(0)
    mcolCodDesProd.Add curMedidaProd.campo(0) & "¯" & curMedidaProd.campo(1) & "¯" & curMedidaProd.campo(4)
    
    ' Se avanza a la siguiente fila del cursor
    curMedidaProd.MoverSiguiente
Loop

'Cierra el cursor de medida de productos
curMedidaProd.Cerrar

End Sub

Private Sub CargarColServicio()
'--------------------------------------------------------------
'Propósito : Carga la colección de servcio con su medida
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Se ejecuta desde el form_load
Dim sSQL As String
Dim curMedidaServ As New clsBD2

sSQL = "SELECT IdServ, DescServ, Medida, CodCont, Tipo " & _
            "FROM SERVICIOS " & _
            "ORDER BY DescServ "
            
'sSQL = "SELECT IdServ, DescServ, Medida,CodCont FROM SERVICIOS ORDER BY DescServ"

'Carga la colección de descripcion y medida de los servcios
curMedidaServ.SQL = sSQL
If curMedidaServ.Abrir = HAY_ERROR Then
  End
End If
Do While Not curMedidaServ.EOF
    ' Se carga la colección de descripciones + unidades de los servicios con la 1º y 2º
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolDesMedidaContServ.Add Key:=curMedidaServ.campo(0), _
                                  Item:=curMedidaServ.campo(2) & "¯" & curMedidaServ.campo(3)
 
    'colección de producto y su descripción
    mcolCodServ.Add curMedidaServ.campo(0)
    mcolCodDesServ.Add curMedidaServ.campo(0) & "¯" & curMedidaServ.campo(1) & "¯" & curMedidaServ.campo(4)
    
    ' Se avanza a la siguiente fila del cursor
    curMedidaServ.MoverSiguiente
Loop
'Cierra el cursor de medida de servicios
curMedidaServ.Cerrar
End Sub

Private Sub cboActiv_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboActiv) = True Then SendKeys "{down}"

End Sub

Private Sub cboActiv_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboActiv.ListIndex) = False And cboActiv.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub


Private Sub cboActiv_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboActiv_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboActiv, vbWhite) = True Then
' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboActiv.Text, txtActiv, mcolCodActiv, mcolCodDesActiv
Else '  Vaciar Controles enlazados al combo
    txtActiv.Text = Empty
End If

'Cambia el alto del combo
cboActiv.Height = CBONORMAL

End Sub
Private Sub cboLinea_Change()
' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboLinea) = True Then SendKeys "{down}"

End Sub

Private Sub cboLinea_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboLinea.ListIndex) = False And cboLinea.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboLinea_KeyDown(KeyCode As Integer, Shift As Integer)
' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboLinea_LostFocus()
' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboLinea, vbWhite) = True Then
' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboLinea.Text, txtLinea, mcolCodLinea, mcolCodDesLinea
Else '  Vaciar Controles enlazados al combo
    txtLinea.Text = Empty
End If

'Cambia el alto del combo
cboLinea.Height = CBONORMAL

End Sub
Private Sub cboProg_Change()
' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProg) = True Then SendKeys "{down}"

End Sub

Private Sub cboProg_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProg.ListIndex) = False And cboProg.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub


Private Sub cboProg_KeyDown(KeyCode As Integer, Shift As Integer)
' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboProg_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProg, vbWhite) = True Then
' Se actualiza código (TextBox) correspondiente a descripción introducida
   CD_ActCod cboProg.Text, txtProg, mcolCodProg, mcolCodDesProg
Else '  Vaciar Controles enlazados al combo
    txtProg.Text = Empty
End If

'Cambia el alto del combo
cboProg.Height = CBONORMAL

End Sub
Private Sub cboProy_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProy) = True Then SendKeys "{down}"

End Sub

Private Sub cboProy_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProy.ListIndex) = False And cboProy.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub


Private Sub cboProy_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboProy_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProy, vbWhite) = True Then
    
    ' Se actualiza código (TextBox) correspondiente a descripción introducida
    CD_ActCod cboProy.Text, txtProy, mcolCodProy, mcolCodDesProy
    
Else '  Vaciar Controles enlazados al combo
    txtProy.Text = Empty
End If

'Cambia el alto del combo
cboProy.Height = CBONORMAL

End Sub

Private Sub CambiaroptCajaBancos()
'-------------------------------------------------------------------
'Propósito : Establece los controles de la primera parte del formulario _
             cuando se cambia de optCaja a optBancos bis
'Recibe : Nada
'Entrega : Nada
'-------------------------------------------------------------------
If gsTipoOperacionEgreso = "Nuevo" Then
   If optCaja.Value = True Then
    'Calcula el siguiente orden de Caja y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("CA")
   ElseIf optBanco.Value = True Then
    'Calcula el siguiente orden de Banco y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("BA")
   ElseIf optRendir.Value = True Then
    'Calcula el siguiente orden de Caja y lo muestra en el txtCodEgreso
    txtCodEgreso.Text = Var22("CA")
   End If
End If
  
' De acuerdo a la elección realizada maneja los controles de banco
ManejaControlesBanco

' Cargar Saldo
CargarSaldo

End Sub

Private Sub ManejaControlesBanco()
'-------------------------------------------------------------------
'Propósito  : Establece los controles de banco cuando se cambia de optCaja a optBancos bis
'Recibe     : Nada
'Entrega    : Nada
'-------------------------------------------------------------------
'If optCaja.Value = True Then
' 'Limpia y oculta los controles de Banco
' txtBanco.Text = Empty
' lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False
' lblCtaCte.Visible = False: cboCtaCte.Visible = False
' cmdPBanco.Visible = False: cmdPCtaCte.Visible = False
' txtNumCheque.Text = Empty: lblNumCheque.Visible = False: txtNumCheque.Visible = False
' msCtaCte = Empty
'
'Else
' 'Muestra los controles de banco
' lblBanco.Visible = True: txtBanco.Visible = True: cboBanco.Visible = True
' lblCtaCte.Visible = True: cboCtaCte.Visible = True
' cmdPBanco.Visible = True: cmdPCtaCte.Visible = True
' lblNumCheque.Visible = True: txtNumCheque.Visible = True
'
'End If

If optCaja.Value = True Then
 'Limpia y oculta los controles de Banco y Rendir
 txtBanco.Text = Empty
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 txtNumCheque.Text = Empty: lblNumCheque.Visible = False: txtNumCheque.Visible = False
 msCtaCte = Empty
 txtRinde.Text = Empty
 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False
ElseIf optBanco.Value = True Then
 'Muestra los controles de banco y oculta el resto
 lblBanco.Visible = True: txtBanco.Visible = True: cboBanco.Visible = True: cmdPBanco.Visible = True
 lblCtaCte.Visible = True: cboCtaCte.Visible = True: cmdPCtaCte.Visible = True
 lblNumCheque.Visible = True: txtNumCheque.Visible = True
 txtRinde.Text = Empty
 lblRinde.Visible = False: txtRinde.Visible = False: txtDescRinde.Visible = False: cmdBuscaRinde.Visible = False

ElseIf optRendir.Value = True Then
 'Muestra los controles a rendir y oculta el resto
 lblRinde.Visible = True: txtRinde.Visible = True: txtDescRinde.Visible = True: cmdBuscaRinde.Visible = True
 txtBanco.Text = Empty
 lblBanco.Visible = False: txtBanco.Visible = False: cboBanco.Visible = False: cmdPBanco.Visible = False
 lblCtaCte.Visible = False: cboCtaCte.Visible = False: cmdPCtaCte.Visible = False
 txtNumCheque.Text = Empty: lblNumCheque.Visible = False: txtNumCheque.Visible = False
 msCtaCte = Empty
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Cierra las colecciones usadas
'Colecciones para la carga del combo de Proyectos
Set mcolCodProy = Nothing
Set mcolCodDesProy = Nothing

' Colecciones para la carga del combo de Programas
Set mcolCodProg = Nothing
Set mcolCodDesProg = Nothing

'Colecciones para la carga del combo de Lineas
Set mcolCodLinea = Nothing
Set mcolCodDesLinea = Nothing

'Colecciones para la carga del combo de Actividades
Set mcolCodActiv = Nothing
Set mcolCodDesActiv = Nothing

'Colecciones para la carga del combo de Bancos
Set mcolCodBanco = Nothing
Set mcolCodDesBanco = Nothing

'Colecciones para la carga del combo de Ctas Ctes
Set mcolCodCtaCte = Nothing
Set mcolCodDesCtaCte = Nothing

'Colecciones para la carga del combo de Productos
Set mcolidprod = Nothing
Set mcolCodDesProd = Nothing
Set mcolDesMedidaContProd = Nothing

'Colecciones para la carga del combo de Servicios
Set mcolCodServ = Nothing
Set mcolCodDesServ = Nothing
Set mcolDesMedidaContServ = Nothing

'Colecciones para la carga del combo de tipo de documento
Set mcolCodTipDoc = Nothing
Set mcolCodDesTipDoc = Nothing
Set mcolCodRetencionPaga = Nothing

' Colección para la carga de detalle, impuestos
Set mcolEgreDet = Nothing
Set mcolEgreImpt = Nothing

'Destruye la coleccion global
Set gcolTabla = Nothing

'Colecciones para la carga del combo de Proyectos
Set mcolCodCatGasto = Nothing
Set mcolCodDesCatGasto = Nothing

' Verifica si esta habilitado controles de egreso caja bancos
If gsTipoOperacionEgreso = "Modificar" And mbEgresoCargado = True Then
     mcurEgreso.Cerrar ' Cierra el cursor del ingreso
End If

End Sub

Private Sub grdDetalle_EnterCell()

If ipos <> grdDetalle.Row Then
    '  Verifica si es la última fila
    If grdDetalle.Row > 0 And grdDetalle.Row < grdDetalle.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdDetalle, ipos
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdDetalle.Row
End If

End Sub

Private Sub grdDetalle_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Llama al proceso que cambia la verificación de un producto
    grdDetalle_DblClick
 End If
 
End Sub

Private Sub grdDetalle_Click()

DesmarcarFilaGRID grdDetalle
DesmarcarGrid grdDetalle
' Verifica si se canceló las operaciones en el grid
If mbCancelagrid = False Then
    If grdDetalle.Row > 0 And grdDetalle.Row < grdDetalle.Rows Then
        ' Marca la fila seleccionada
        MarcarFilaGRID grdDetalle, vbWhite, vbDarkBlue
        cmdEliminar.Enabled = True
    End If
End If

End Sub

Private Sub grdDetalle_DblClick()

' Verifica si se canceló las operaciones en el grid
If mbCancelagrid = False Then
  If grdDetalle.Row > 0 Then
  
      ' Verifica si esta seleccionado
    If grdDetalle.CellBackColor <> vbDarkBlue Then
       MarcarFilaGRID grdDetalle, vbWhite, vbDarkBlue
       Exit Sub
    End If
    
    'Verifica que el cboProdServicio este vacio
    If cboProdServ.Text <> Empty Then
        'Termina la ejecucion
        Exit Sub
    End If

    ' carga la fila selecionada
    CargarEditarFila
    
    ' elimina la fila seleccionada del grid
    If grdDetalle.Rows > 2 Then
      ' elimina la fila seleccionada del grid
      grdDetalle.RemoveItem grdDetalle.Row
    Else
      ' estable vacío el grddetalle
      grdDetalle.Rows = 1
      cmdEliminar.Enabled = False
    End If
    
    ' coloca el focus a cbo producto
    cboProdServ.SetFocus
    
    ' Actualiza el ipos
    ipos = 0
    
  End If
End If

DesmarcarFilaGRID grdDetalle
DesmarcarGrid grdDetalle
' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub CargarEditarFila()
'---------------------------------------------------
' Propósito: Carga la fila seleccionada para la edición
' Recibe: Nada
' Entrega: Nada
'---------------------------------------------------

'"Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "CodCont")
    'Recupera en el msidprod el codigo del producto seleccionado
    msIdProdServ = grdDetalle.TextMatrix(grdDetalle.RowSel, 4)
    
    If optProducto.Value = True Then
        'Recupera en el combo el producto del msidprodserv
        'CD_ActVarCbo cboProdServ, msIdProdServ, mcolCodDesProd
        ActualizarVariableCbo cboProdServ, grdDetalle.TextMatrix(grdDetalle.RowSel, 0), mcolCodDesProd
    Else
        'Recupera en el combo el servicio del msidprodserv
        'CD_ActVarCbo cboProdServ, msIdProdServ, mcolCodDesServ
        ActualizarVariableCbo cboProdServ, grdDetalle.TextMatrix(grdDetalle.RowSel, 0), mcolCodDesServ
    End If
    txtMedida.Text = grdDetalle.TextMatrix(grdDetalle.RowSel, 5)
    msCodCont = grdDetalle.TextMatrix(grdDetalle.RowSel, 6)
    

   ' Pone los Montos en sus respectivos controles
    txtCant.Text = grdDetalle.TextMatrix(grdDetalle.RowSel, 1)
   ' Verifica el tipo de giro que tiene el documento
    If optGiroConImpuestos.Value = True Then
        ' pone el monto en valor venta
        txtValorVenta = grdDetalle.TextMatrix(grdDetalle.RowSel, 3)
    Else
      If TipoEgreso = "EMPR" Then
        If txtTipDoc.Text <> "02" Then
          ' pone el valor en valor VENTA y raliza los cálculos
          txtValorVenta = grdDetalle.TextMatrix(grdDetalle.RowSel, 3)
        Else
          ' pone el valor en valor compra y raliza los cálculos
          txtValorCompra = grdDetalle.TextMatrix(grdDetalle.RowSel, 3)
        End If
      Else
        ' pone el valor en valor compra y raliza los cálculos
        txtValorCompra = grdDetalle.TextMatrix(grdDetalle.RowSel, 3)
      End If
    End If
    
    DetalleSubido = True
End Sub

Private Sub grdDetalle_GotFocus()

' Verifica si se puede realizar operaciones en el grid
If fbOperarDetalle = False Then
    ' no se puede operar, cancela las operaciones en el grid
    mbCancelagrid = True
    Exit Sub
Else
    mbCancelagrid = False
End If

End Sub


Private Sub Image1_Click()
'Carga la Var48
Var48
End Sub

Private Sub mskFecDoc_Change()

' Verifica  que NO tenga campos en blanco
If ValidarFechaOblig(mskFecDoc) = True Then
' Los campos coloca a color blanco
   mskFecDoc.BackColor = vbWhite
Else
' Marca los campos obligatorios
   mskFecDoc.BackColor = Obligatorio
End If

' Habilita el botón aceptar
HabilitarBotonAceptar
End Sub

Private Sub mskFecDoc_KeyPress(KeyAscii As Integer)
  ' Si se presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
      SendKeys vbTab
  End If
End Sub

Private Sub optBanco_Click()

' Realiza el cambio de opción a Bancos
 CambiaroptCajaBancos
    
'  Habilita el botón aceptar
 HabilitarBotonAceptar
    
End Sub

Private Sub optBanco_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
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

Private Sub optCaja_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optGiroConImpuestos_Click()

' Cambia la opción de giro con impuestos a giro con impuestos
CambiarOptTipoGiro

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub optGiroConImpuestos_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optGiroSinImpuestos_Click()

' Cambia la opción de giro con impuestos a giro sin impuestos
CambiarOptTipoGiro

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub CambiarOptTipoGiro()

' Verifica el opt elegido
If optGiroConImpuestos.Value = True Then
    ' la opción girado con impuestos, oculta los controles que calculan _
    el monto de compra
    lblValorCompra.Visible = False: txtValorCompra.Visible = False
    lblPrecioUniCompra.Visible = False: txtPrecioUniCompra.Visible = False
    lblValorVenta.Caption = "Costo Total": lblPrecioUniVenta.Caption = "Precio Unit.Total"
Else
    ' la opción girado sin impuestos, Muestraa los controles que calculan _
    el monto de compra
    lblValorCompra.Visible = True: txtValorCompra.Visible = True
    lblPrecioUniCompra.Visible = True: txtPrecioUniCompra.Visible = True
    lblValorVenta.Caption = "Costo Neto": lblPrecioUniVenta.Caption = "Precio Unit.Neto"
    lblValorCompra.Caption = "Costo Total": lblPrecioUniCompra.Caption = "Precio Unit.Total"
End If

End Sub

Private Sub optGiroSinImpuestos_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optProducto_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
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

Private Sub optServicio_Click()

' Realiza el cambio de optProdServ
CambioOptServProd

End Sub

Private Sub optProducto_Click()

' Realiza el cambio de optProdServ
CambioOptServProd

End Sub

Private Sub CambioOptServProd()

' Verifica si la operación no fué cancelada en el otro optServicio
If mbCancelaOptClick = False Then
    
    ' Verifica si el grid detalle tiene elementos
    If grdDetalle.Rows > 1 Then
       ' Pregunta si desea borrar los registros introducidos en el grid
       If MsgBox("Se borrarán los datos introducidos en el Detalle, ¿Desea Proseguir?", _
       vbQuestion + vbYesNo, "SGCcaijo-Egreso con afectación") = vbNo Then
            ' cambia el valor de opt y cancela la operación
            mbCancelaOptClick = True
            mbCancelaCambioTipDoc = True
            If optProducto.Value = True Then
                optServicio.Value = True
            Else
                optProducto.Value = True
            End If
            
            Exit Sub
       End If
    End If
    
    'Limpia el cbo de servicios y los campos de detalle
    LimpiarSegundaParteFormulario
        
    'Carga el combo
    If optProducto.Value = True Then
        'Cambia la etiqueta del cboprocServ y a la columna 0 del grdDetalle
        lblProdServ = "Producto"
        grdDetalle.TextMatrix(0, 0) = "Producto"
        'Carga el cboProducto de acuerdo a la relación
        'CargarCboCols cboProdServ, mcolCodDesProd
        CargarCboProdServ cboProdServ, mcolCodDesProd
    Else
        'Cambiamos la etiqueta del cboprocServ
        lblProdServ = "Servicio"
        grdDetalle.TextMatrix(0, 0) = "Servicio"
        'carga los servicios en el cbo
        'CargarCboCols cboProdServ, mcolCodDesServ
        CargarCboProdServ cboProdServ, mcolCodDesServ
    End If
Else
    ' actualiza la variable que cancela el cambio de Prod a Serv
     mbCancelaOptClick = False
End If

End Sub

Private Sub LimpiarSegundaParteFormulario()
' -------------------------------------------------------------------
' Propósito: limpia los controles de la segunda parte del formulario
' -------------------------------------------------------------------

' Limpia los controles de ingreso al detalle
  LimpiaControlesIngreDet

' Limpia combo productos y servicios
  cboProdServ.Clear

' Limpia el grdDetalle
  grdDetalle.Rows = 1
  
End Sub

Private Sub LimpiarPrimeraParteFormulario()

' Limpia los controles generales del formulario
txtCodEgreso = Empty
txtProy = Empty
txtTipDoc = Empty
txtRUCDNI = Empty
txtNombrProv = Empty
txtDocEgreso = Empty
txtTotalDoc = Empty
txtObserv = Empty
' Limpia controles banco
txtBanco = Empty
txtNumCheque = Empty
' Limpia Rendir
txtRinde = Empty
mskFecDoc = "__/__/____"

End Sub


Private Sub LimpiaControlesIngreDet()
' -------------------------------------------------------------------
' Propósito: Limpia los controles que permiten el ingreso al grdDetalle del Doc
' -------------------------------------------------------------------

' Limpia el combo ProdServ
cboProdServ.ListIndex = -1
cboProdServ.BackColor = Obligatorio

' Limpia los controles txt
txtValorVenta = Empty
txtValorCompra = Empty
txtCant.Text = Empty
txtMedida = Empty

End Sub

Private Sub optServicio_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtActiv_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboActiv, txtActiv, mcolCodDesActiv

' Verifica SI el campo esta vacio
If txtActiv.Text <> "" And cboActiv.Text <> "" Then
    ' Los campos coloca a color blanco
    txtActiv.BackColor = vbWhite
    If TipoEgreso = "EMPR" Then
      RecuperarCtaContableActiv
    End If
Else
  'Marca los campos obligatorios
  txtActiv.BackColor = Obligatorio

End If

'Habilitar btotón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtActiv_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
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


  'Habilita el botón aceptar
  HabilitarBotonAceptar

End Sub

Public Sub ActualizarListcboCtaCte()
'------------------------------------------------------
Dim sSQL As String
Dim curCtaCte As New clsBD2

'Inicializa el cboCtaCte
cboCtaCte.Clear
cboCtaCte.BackColor = Obligatorio

If txtBanco.BackColor <> Obligatorio And Len(txtBanco.Text) = txtBanco.MaxLength Then
  ' Carga la Sentencia para obtener las CtsCts en dólares que pertenecen a el txtBanco
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

Private Sub txtBanco_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  Else
    ' Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If

End Sub

Private Sub txtCant_Change()

If txtCant.Text <> "" Then
   ' SI es Igual al vacio lo pone como obligatorio
   txtCant.BackColor = vbWhite
Else
   ' SI NO lo pone a obligatorio
   txtCant.BackColor = Obligatorio
End If

'Calcula el precio unitario de venta y compra
    CalcularPrecioUniVV
    CalcularPrecioUniVC

'Habilitar añadir
    HabilitaBotonAñadir

End Sub

Private Sub CalculaValores(sControlOrigen As String)

   Select Case sControlOrigen
   Case "ValorVenta"
    ' activa la variable que indica que se está calculando los Precios
     mbCalculando = True
     CalcularValorCompra
     mbCalculando = False
   Case "ValorCompra"
    ' activa la variable que indica que se está calculando los Precios
     mbCalculando = True
     CalcularValorVenta
     mbCalculando = False
   End Select


End Sub

Private Sub CalcularValorCompra()
'--------------------------------------------------------------
' Propósito: Calcula el valor de compra en base al valor de venta _
             introducido
' Recibe: nada
' Entrega: nada
'--------------------------------------------------------------
Dim sRelacImpuestos As String

' Asigna a la variable la relación del documento con el impuesto
sRelacImpuestos = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1)


' Verifica el tipo de documento introducido y su relación con impuestos
Select Case sRelacImpuestos
Case "Registra", "Paga"
    ' Muestra el valor de la compra cuando los impuestos se registran, Pagan
    txtValorCompra = Format(Val(Var37(txtValorVenta)) * (1 + mdblSumImpt), "###,###,##0.00")
Case "Retiene"
    ' Muestra el valor de la compra cuando los impuestos se retienen
    txtValorCompra = Format(Val(Var37(txtValorVenta)) / (1 - mdblSumImpt), "###,###,##0.00")
End Select


End Sub

Private Sub CalcularValorVenta()
'--------------------------------------------------------------
' Propósito: Calcula el valor de venta en base al valor de compra _
             introducido
' Recibe: nada
' Entrega: nada
'--------------------------------------------------------------
Dim sRelacImpuestos As String

' Asigna a la variable la relación del documento con el impuesto
sRelacImpuestos = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1)

' Verifica el tipo de documento introducido y su relación con impuestos
Select Case sRelacImpuestos
Case "Registra", "Paga"
    ' Muestra el valor de la compra cuando los impuestos se registran, Pagan
    txtValorVenta = Format(Val(Var37(txtValorCompra)) / (1 + mdblSumImpt), "###,###,##0.00")
Case "Retiene"
    ' Muestra el valor de la compra cuando los impuestos se retienen
    txtValorVenta = Format(Val(Var37(txtValorCompra)) * (1 - mdblSumImpt), "###,###,##0.00")
End Select

End Sub

Private Sub txtCant_GotFocus()

' Verificar si se a introducido un tipo de documento
If fbOperarDetalle = False Then
    Exit Sub
End If

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Dato Ingresado
Var34 txtCant, 11, KeyAscii

End Sub

Private Sub txtCant_LostFocus()

' pone el maximo número de caracteres
txtCant.MaxLength = 11

' Verifica si se introdujjo un dato valido
If Val(Var37(txtCant)) = 0 Then
    txtCant = Empty
Else
    ' formato a cantidad
    txtCant = Format(Val(Var37(txtCant.Text)), "####0.00")
End If

End Sub

Private Sub TxtCategoriaGasto_Change()
  'SI procede, se actualiza descripción correspondiente a código introducido
  CD_ActDesc cboCategoriaGasto, TxtCategoriaGasto, mcolCodDesCatGasto
  
  ' Verifica SI el campo esta vacio
  If TxtCategoriaGasto.Text <> "" And cboCategoriaGasto.Text <> "" Then
      ' Los campos coloca a color blanco
      TxtCategoriaGasto.BackColor = vbWhite
      
  Else
    'Marca los campos obligatorios
    TxtCategoriaGasto.BackColor = Obligatorio
  
  End If
  
  'Habilitar btotón aceptar
  HabilitarBotonAceptar
End Sub

Private Sub TxtCategoriaGasto_KeyPress(KeyAscii As Integer)
  ' Si se presiona enter pasa al siguiente control
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
          
          ' Verifica si el egreso existe y es con afectación
          If fbCargaEgreso = True Then
             ' Sale y deshabilita el control
             SendKeys vbTab
            
             ' Habilita el formulario
             DeshabilitarHabilitarFormulario True
              
             ' deshabilita el txtcod egreso y el botón buscar, _
               habilita anular
             txtCodEgreso.Enabled = False
             cmdBuscarEgreso.Enabled = False
             cmdAnular.Enabled = True
          End If ' fin de cargar egreso
          
        Else ' vuelve a mayúsulas el txtcodegreso
            txtCodEgreso = UCase(txtCodEgreso)
        End If ' fin verificar mayúsculas
        
      End If ' fin de verofocr el tamalo del texto
      ' VADICK MODIFICACION TENIENDO LA FECHA DE TRABAJO DEL DOCUMENTO RECUPERADO
      ' VERIFICAMOS SI SE HA CERRRADO EL AÑO CONTABLE PARA DETERMINAR QUE PUEDE MODIFICARSE Y QUE NO
      ' HABILITANDO O DESHABILITANDO LOS CONTROLES CORRESPONDIENTES
      If Conta52(Right(mskFecTrab.Text, 4)) = True Then
        ' No se puede realizar la operación
        DeshabilitarControlesPorAnioCerrado False
        'MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
        vbCritical + vbOKOnly, "SGCCAIJO-Caja-Bancos"
        ' Sale
        'Exit Sub
      End If
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

Private Function fbCargaEgreso() As Boolean
' ----------------------------------------------------------
' Propósito: Verifica si existe el código del egreso y carga _
             los datos de el egreso
' Recibe : Nada
' Entrega : Nada
' ----------------------------------------------------------
Dim sSQL As String
If Left(txtCodEgreso, 2) = "BA" Then ' para bancos
    ' Establece el origen
     msCaBaAnt = "BA"
    ' carga la sentencia  que verifica si existe el código de egreso y es con afectación
    sSQL = "SELECT E.Orden, E.IdProy, E.IdProg, E.IdLinea, E.IdActiv, E.NumDoc, " _
     & "E.IdTipoDoc, E.FecMov, E.MontoAfectado, E.MontoCB, P.Numero, " _
     & "P.DescProveedor, P.IdProveedor, E.Anulado, E.Observ, " _
     & "E.GiradoConImpuestos, C.idBanco, C.idCta, E.NumCheque, E.CODCATGASTO, E.FecDoc " _
     & "FROM EGRESOS E, PROVEEDORES P, TIPO_CUENTASBANC C " _
     & "WHERE E.Orden='" & txtCodEgreso & "' and E.Anulado='NO' and " _
     & "E.IdProy<>Null and E.IdProveedor=P.IdProveedor and E.idCta=C.idCta "
ElseIf Left(txtCodEgreso, 2) = "CA" Then ' para caja
    ' carga la sentencia  que verifica si existe el código de egreso y es con afectación
    sSQL = "SELECT E.Orden, E.IdProy, E.IdProg, E.IdLinea, E.IdActiv, E.NumDoc, " _
     & "E.IdTipoDoc, E.FecMov, E.MontoAfectado, E.MontoCB, P.Numero, " _
     & "P.DescProveedor, P.IdProveedor, E.Anulado, E.Observ, " _
     & "E.GiradoConImpuestos, E.Origen, E.CODCATGASTO, E.FecDoc " _
     & "FROM EGRESOS E, PROVEEDORES P " _
     & "WHERE E.Orden='" & txtCodEgreso & "' and E.Anulado='NO' and " _
     & "E.IdProy<>Null and E.IdProveedor=P.IdProveedor "
Else ' codigo incorrecto
    'Mensaje de código incorrecto
    MsgBox "El código de Egreso que se digitó no es correcto", vbExclamation, "Caja-Banco- Egresos"
    ' cierra el cursor y se va
    fbCargaEgreso = False
    Exit Function
End If

' ejecuta la sentencia
mcurEgreso.SQL = sSQL
If mcurEgreso.Abrir = HAY_ERROR Then End
' Cursor abierto
mbEgresoCargado = True

' verifica si existe el egreso con afectación
If mcurEgreso.EOF Then
    'Mensaje de registro de egreso a Caja o Bancos NO existe
    MsgBox "El código de Egreso que se digitó no está registrado como " & _
      "egreso con afectación financiera o está anulado", vbExclamation, "Caja-Banco- Egresos"
    mcurEgreso.Cerrar
    ' cierra el cursor y se va
    fbCargaEgreso = False
Else
    ' Carga el cursor que contiene el detalle del egreso
    sSQL = "SELECT G.Orden, G.Concepto, G.CodConcepto, G.Cantidad, G.Monto,(G.Monto/G.Cantidad) " _
         & "FROM GASTOS G " _
         & "WHERE G.Orden='" & txtCodEgreso & "'"
    ' Ejecuta la sentencia
    mcurDetalleEgreso.SQL = sSQL
    If mcurDetalleEgreso.Abrir = HAY_ERROR Then End
    
    ' Carga el cursor que contiene los impuestos del egreso
    ' "idImp", "ValorImpuesto", "Monto", codcont
    'sSQL = "SELECT M.Orden, M.IdImp, M.ValorImp ,M.Monto, I.CodContable " _
           & "FROM MOV_IMPUESTOS M, TIPO_IMPUESTOS I " _
           & "WHERE M.Orden='" & txtCodEgreso & "' and M.IdImp=I.IdImp"
     sSQL = "SELECT M.Orden, M.IdImp, M.ValorImp ,M.Monto, M.CodContable, M.DescImp " _
           & "FROM MOV_IMPUESTOS M " _
           & "WHERE M.Orden='" & txtCodEgreso & "'"
     ' Eejecuta la sentencia
     mcurImpuestos.SQL = sSQL
     If mcurImpuestos.Abrir = HAY_ERROR Then End
     
    ' Carga los datos generales
    CargaControlesGenerales
    ' Carga colección de impuestos
    CargaImpuestos
    ' Carga los datos del detalle
    CargaControlesDetalle
    ' Devuelve el resultado de la función
    fbCargaEgreso = True
End If

End Function

Private Sub CargaControlesGenerales()
'--------------------------------------------------------------
'Propósito: Carga los controles del formulario referidos  a los _
            datos generales de el egreso con afectación
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------
' Carga los controles editables y de opción
Dim sSQL As String
    ' Actualiza Variables
      msOrden = mcurEgreso.campo(0)
      gsIdProv = mcurEgreso.campo(12)
      gdblMontoTotal = mcurEgreso.campo(8)
    ' Carga opt caja-bancos
    If Left(msOrden, 2) = "CA" Then
      ' Verifica si la operación es de caja o de cuentas a Rendir
        If mcurEgreso.campo(16) = "C" Then
            ' Establece el origen
            msCaBaAnt = "CA"
            ' Habilita la opción de Caja
            optCaja.Value = True
        ElseIf mcurEgreso.campo(16) = "R" Then
            ' Establece el origen
            msCaBaAnt = "ER"
            ' Habilita la opción de Rendir
            optRendir.Value = True
            ' Averigua la cuenta a rendir
            CargarCuentaRendir
        End If
        If Not IsNull(mcurEgreso.campo(17)) Then
          CodigoCategoriaGasto = mcurEgreso.campo(17)
        Else
          CodigoCategoriaGasto = ""
        End If
      mskFecDoc = FechaDMA(mcurEgreso.campo(18))
    Else
        ' Egreso de banco y carga sus controles
        optBanco.Value = True
        txtBanco = mcurEgreso.campo(16)
        msCtaCte = mcurEgreso.campo(17) ' el código de la Ctacte
        CD_ActVarCbo cboCtaCte, msCtaCte, mcolCodDesCtaCte
        txtNumCheque = mcurEgreso.campo(18)
        If Not IsNull(mcurEgreso.campo(19)) Then
          CodigoCategoriaGasto = mcurEgreso.campo(19)
        Else
          CodigoCategoriaGasto = ""
        End If
        CargarSaldo
        mskFecDoc = FechaDMA(mcurEgreso.campo(20))
    End If
    
    ' Carga los datos en sus controles
    mskFecTrab = FechaDMA(mcurEgreso.campo(7))
    
    ' Carga las colecciones necesarias para el manejo del formulario
    ' Carga Proyectos en el modificación
    sSQL = "SELECT idproy,IdProy + '   ' + descproy  " & _
           " FROM Proyectos WHERE idproy IN " & _
           "(SELECT idproy FROM Presupuesto_proy ) " & _
           "ORDER BY idproy"
    CD_CargarColsCbo cboProy, sSQL, mcolCodProy, mcolCodDesProy

    txtProy = mcurEgreso.campo(1)
    txtProg = mcurEgreso.campo(2)
    txtLinea = mcurEgreso.campo(3)
    txtActiv = mcurEgreso.campo(4)
    TxtCategoriaGasto = CodigoCategoriaGasto
    txtDocEgreso = mcurEgreso.campo(5)
    txtTipDoc = mcurEgreso.campo(6)
    txtRUCDNI = mcurEgreso.campo(10)
    txtNombrProv = mcurEgreso.campo(11)
    txtTotalDoc = Format(mcurEgreso.campo(8), "###,###,##0.00")
    If Val(mcurEgreso.campo(8)) <> Val(mcurEgreso.campo(9)) Then
      txtMontoCB = Format(mcurEgreso.campo(8), "###,###,##0.00")
    Else
      txtMontoCB = Format(mcurEgreso.campo(9), "###,###,##0.00")
    End If
    txtObserv = mcurEgreso.campo(14)
    
    ' Carga OptTipoGiro
    If mcurEgreso.campo(15) = "SI" Then
        optGiroConImpuestos.Value = True
    Else
        optGiroSinImpuestos.Value = True
    End If
    
    ' Maneja los estados de las opciones
    Manejaopciones

End Sub

Private Sub CargarCuentaRendir()
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

Private Sub Manejaopciones()
If msCaBaAnt = "CA" Or msCaBaAnt = "ER" Then
    ' Habilita optcaja y optrendir
    optCaja.Enabled = True: optBanco.Enabled = False: optRendir.Enabled = True
ElseIf msCaBaAnt = "BA" Then
    ' No habilita las opciones
    optCaja.Enabled = False: optBanco.Enabled = False: optRendir.Enabled = False
End If
End Sub

Private Sub CargaControlesDetalle()
Dim i As Integer
Dim CodigoProductoServicio As String
Dim DescripcionProductoServicio As String

If (mcurDetalleEgreso.EOF) Then ' verifica que no sea vacio
    MsgBox "Error Egreso sin detalle en BD, Consulte al administrador", , "SGCcaijo-Egreso con afectación"
    Exit Sub
Else
    ' Pone el valor a optPaga Prod o Serv
    If mcurDetalleEgreso.campo(1) = "P" Then
        optProducto.Value = True
    ElseIf mcurDetalleEgreso.campo(1) = "S" Then
        optServicio.Value = True
    End If

    ' carga el grid del detalle
    Do While Not mcurDetalleEgreso.EOF
   'Producto,Cantidad,Precio Unitario,Total,idproducto,Medida,CodCont
        If mcurDetalleEgreso.campo(1) = "P" Then
          CodigoProductoServicio = mcurDetalleEgreso.campo(2)
          DescripcionProductoServicio = ""
          For i = 1 To mcolCodDesProd.Count
            If Var30(mcolCodDesProd(i), 1) = CodigoProductoServicio Then ' Elemento encontrado
              DescripcionProductoServicio = Var30(mcolCodDesProd(i), 2) ' Actualiza código
              Exit For
            End If
          Next
            
          grdDetalle.AddItem DescripcionProductoServicio _
            & vbTab & Format(mcurDetalleEgreso.campo(3), "###0.00") & vbTab & Format(mcurDetalleEgreso.campo(5), "###,###,##0.00") _
            & vbTab & Format(mcurDetalleEgreso.campo(4), "###,###,##0.00") & vbTab & mcurDetalleEgreso.campo(2) _
            & vbTab & Var30(mcolDesMedidaContProd.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Var30(mcolDesMedidaContProd.Item(mcurDetalleEgreso.campo(2)), 2)
            
            'grdDetalle.AddItem Var30(mcolCodDesProd.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Format(mcurDetalleEgreso.campo(3), "###0.00") & vbTab & Format(mcurDetalleEgreso.campo(5), "###,###,##0.00") _
            & vbTab & Format(mcurDetalleEgreso.campo(4), "###,###,##0.00") & vbTab & mcurDetalleEgreso.campo(2) _
            & vbTab & Var30(mcolDesMedidaContProd.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Var30(mcolDesMedidaContProd.Item(mcurDetalleEgreso.campo(2)), 2)
        ElseIf mcurDetalleEgreso.campo(1) = "S" Then
          CodigoProductoServicio = mcurDetalleEgreso.campo(2)
          DescripcionProductoServicio = ""
          For i = 1 To mcolCodDesServ.Count
            If Var30(mcolCodDesServ(i), 1) = CodigoProductoServicio Then ' Elemento encontrado
              DescripcionProductoServicio = Var30(mcolCodDesServ(i), 2) ' Actualiza código
              Exit For
            End If
          Next
          
          grdDetalle.AddItem DescripcionProductoServicio _
            & vbTab & Format(mcurDetalleEgreso.campo(3), "###0.00") & vbTab & Format(mcurDetalleEgreso.campo(5), "###,###,##0.00") _
            & vbTab & Format(mcurDetalleEgreso.campo(4), "###,###,##0.00") & vbTab & mcurDetalleEgreso.campo(2) _
            & vbTab & Var30(mcolDesMedidaContServ.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Var30(mcolDesMedidaContServ.Item(mcurDetalleEgreso.campo(2)), 2)
          
          'grdDetalle.AddItem Var30(mcolCodDesServ.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Format(mcurDetalleEgreso.campo(3), "###0.00") & vbTab & Format(mcurDetalleEgreso.campo(5), "###,###,##0.00") _
            & vbTab & Format(mcurDetalleEgreso.campo(4), "###,###,##0.00") & vbTab & mcurDetalleEgreso.campo(2) _
            & vbTab & Var30(mcolDesMedidaContServ.Item(mcurDetalleEgreso.campo(2)), 1) _
            & vbTab & Var30(mcolDesMedidaContServ.Item(mcurDetalleEgreso.campo(2)), 2)
        End If
        
        ' carga la colección detalle codconcepto, cantidad, monto
        mcolEgreDet.Add Item:=mcurDetalleEgreso.campo(2) & "¯" _
                            & Format(mcurDetalleEgreso.campo(3), "####0.00") & "¯" _
                            & Format(mcurDetalleEgreso.campo(4), "###,###,##0.00"), _
                        Key:=mcurDetalleEgreso.campo(2)
        
        mcurDetalleEgreso.MoverSiguiente
    Loop
    'Mueve al inicio del cursor
    mcurDetalleEgreso.Cerrar
End If

End Sub

Private Sub CargaImpuestos()

    Do While Not mcurImpuestos.EOF
        ' Carga la colección de los impuestos seleccionados
        gcolImpSel.Add Key:=mcurImpuestos.campo(1), _
                       Item:=mcurImpuestos.campo(1) & "¯" & _
                              Format(mcurImpuestos.campo(2), "#0.00") & "¯" & _
                              Format(mcurImpuestos.campo(3), "#0.00") & "¯" & _
                              Trim(mcurImpuestos.campo(4)) & "¯" & Trim(mcurImpuestos.campo(5))
        ' Carga la coleccion de impuestos
        mcolEgreImpt.Add Key:=mcurImpuestos.campo(1), _
                       Item:=mcurImpuestos.campo(1) & "¯" & _
                              Format(mcurImpuestos.campo(2), "#0.00") & "¯" & _
                              Format(mcurImpuestos.campo(3), "#0.00") & "¯" & _
                              Trim(mcurImpuestos.campo(4)) & "¯" & Trim(mcurImpuestos.campo(5))
        ' mueve al siguiente impuesto
        mcurImpuestos.MoverSiguiente
    Loop
    ' Actualiza variables
    mcurImpuestos.Cerrar
    
    ' Actualiza la variable que indica que se seleccionó impuestos
    gbImpuestos = True
    
    ' Averigua el Monto total de impuestos relacionados a el documento
    mdblMontoTotalImpt = SumaMontosImpuestos

    ' Averigua la suma de los porcentajes aplicados a el documento
    mdblSumImpt = SumaImpuestosAplicados

    ' Muestra el monto relacionado a los impuestos
    txtMontoImpuesto.Text = Format(mdblMontoTotalImpt, "###,###,##0.00")
End Sub

Private Sub txtCodEgreso_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtDocEgreso_Change()

' Verifica  que NO tenga campos en blanco
If txtDocEgreso.Text <> "" And InStr(txtDocEgreso, "'") = 0 Then
' Los campos coloca a color blanco
   txtDocEgreso.BackColor = vbWhite
Else
' Marca los campos obligatorios
   txtDocEgreso.BackColor = Obligatorio
End If

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtDocEgreso_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtLinea_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboLinea, txtLinea, mcolCodDesLinea
'Actualiza el combo de Actividades
ActualizarcboActividades
    ' Verifica SI el campo esta vacio
If txtLinea.Text <> "" And cboLinea.Text <> "" Then
   ' Los campos coloca a color blanco
   txtLinea.BackColor = vbWhite
Else
'Los campos coloca a color amarillo
   txtLinea.BackColor = Obligatorio
End If

' habilitar botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub ActualizarcboActividades()
Dim sSQL As String
' Limpia  controles relacionados con Actividad
    cboActiv.Clear
    txtActiv.Text = Empty

' Verifica SI el campo esta vacio
If txtLinea.Text <> "" And cboLinea.Text <> "" Then
    ' Carga el combo con Descripciones de los programa para el proyecto elegidos
    ' Inicializa las colecciones
    Set mcolCodActiv = Nothing
    Set mcolCodDesActiv = Nothing

    sSQL = "SELECT DISTINCT PresProy.IdActiv, PresProy.IdActiv + '   ' + Activ.DescActiv  " & _
           " FROM ACTIVIDADES Activ, PRESUPUESTO_PROY PresProy" & _
           " WHERE Activ.IdActiv=PresProy.IdActiv and PresProy.Idproy=" & "'" & txtProy.Text & "'" & _
           " AND PresProy.Idprog=" & "'" & txtProg.Text & "'" & _
           " AND PresProy.IdLinea=" & "'" & txtLinea.Text & "'  ORDER BY PresProy.IdActiv + '   ' + Activ.DescActiv"

    CD_CargarColsCbo cboActiv, sSQL, mcolCodActiv, mcolCodDesActiv
End If

End Sub

Private Sub txtLinea_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtNombrProv_GotFocus()
    ' Manda el tab al siguiente elemento
    SendKeys vbTab
End Sub

Private Sub txtNumCheque_Change()

' Verifica SI el campo esta vacio
If txtNumCheque.Text <> "" And InStr(txtNumCheque, "'") = 0 Then
' El campos coloca a color blanco
   txtNumCheque.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtNumCheque.BackColor = Obligatorio
End If

'Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    'Convierte a mayusculas
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

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtProg_Change()

'SI procede, se actualiza descripción correspondiente a código introducido
   CD_ActDesc cboProg, txtProg, mcolCodDesProg
'Actualizar el combo de lineas
ActualizarcboLineas
' Verifica SI el campo esta vacio
If txtProg.Text <> "" And cboProg.Text <> "" Then
  ' Los campos coloca a color blanco
   txtProg.BackColor = vbWhite
Else
'Los campos coloca a color amarillo
   txtProg.BackColor = Obligatorio
End If

' Habilitar boton aceptar
HabilitarBotonAceptar

End Sub

Private Sub ActualizarcboLineas()

Dim sSQL As String

'Limpia  controles relacionados con Linea,Actividad
cboLinea.Clear
txtLinea.Text = Empty
cboActiv.Clear
txtActiv.Text = Empty

'Verifica SI es un dato valido
If txtProg.Text <> "" And cboProg.Text <> "" Then
   ' Carga el combo con Descripciones de las Lineas para el proyecto,programa elegidos
   ' Se inicializan las colecciones
    Set mcolCodDesLinea = Nothing
    Set mcolCodLinea = Nothing

    sSQL = "SELECT DISTINCT PresProy.IdLinea, PresProy.IdLinea + '   ' + Lins.DescLinea  " & _
           " FROM LINEAS Lins, PRESUPUESTO_PROY PresProy" & _
           " WHERE Lins.IdLinea=PresProy.IdLinea and PresProy.Idproy=" & "'" & txtProy.Text & "'" & _
           " And PresProy.Idprog=" & "'" & txtProg.Text & "' ORDER BY PresProy.IdLinea + '   ' + Lins.DescLinea "

    CD_CargarColsCbo cboLinea, sSQL, mcolCodLinea, mcolCodDesLinea

End If

End Sub

Private Sub txtProg_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtProy_Change()
Dim sSQL As String

' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboProy, txtProy, mcolCodDesProy

' Actualiza Combos de programas
ActualizarcboProgramas

' Actualiza Combos de Categoria de Gasto
ActualizarcboCategoriaGasto

sSQL = ""
sSQL = "SELECT Tipo " & _
       "FROM Proyectos WHERE idproy = '" & txtProy & "' "

' ejecuta la sentencia
mcurProyectos.SQL = sSQL
If mcurProyectos.Abrir = HAY_ERROR Then End

If Not mcurProyectos.EOF Then
  If mcurProyectos.campo(0) = "PROY" Then
    TipoEgreso = "PROY"
  ElseIf mcurProyectos.campo(0) = "EMPR" Then
    TipoEgreso = "EMPR"
  End If
End If

mcurProyectos.Cerrar

 ' Verifica SI el campo esta vacio
If txtProy.Text <> "" And cboProy.Text <> "" Then
   ' Los campos coloca a color blanco
   txtProy.BackColor = vbWhite
   'Actualiza Financiera y periodo del proyecto
    ActualizarFinanancieraPeriodo
Else
   'Los campos coloca a color amarillo
   txtProy.BackColor = Obligatorio
   txtFinan.Text = Empty
   txtFinan.Visible = False
   lblFinan.Visible = False
End If

' Habilita el botón aceptar
HabilitarBotonAceptar

' Limpia combo productos y servicios
cboProdServ.Clear

'Carga el combo
If optProducto.Value = True Then
    'Cambia la etiqueta del cboprocServ y a la columna 0 del grdDetalle
    lblProdServ = "Producto"
    grdDetalle.TextMatrix(0, 0) = "Producto"
    'Carga el cboProducto de acuerdo a la relación
    'CargarCboCols cboProdServ, mcolCodDesProd
    CargarCboProdServ cboProdServ, mcolCodDesProd
Else
    'Cambiamos la etiqueta del cboprocServ
    lblProdServ = "Servicio"
    grdDetalle.TextMatrix(0, 0) = "Servicio"
    'carga los servicios en el cbo
    'CargarCboCols cboProdServ, mcolCodDesServ
    CargarCboProdServ cboProdServ, mcolCodDesServ
End If
End Sub

Private Sub ActualizarcboProgramas()
Dim sSQL As String
'--------------------------------------------------------------
'Propósito: Actualiza los cboProg y txtProg cuando se cambia el proyecto
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------
' nota :    Llamado desde el evento change de txtproy y Modificar

' Limpia Programa, Linea, Actividad
    cboProg.Clear
    txtProg.Text = Empty
    cboLinea.Clear
    txtLinea.Text = Empty
    cboActiv.Clear
    txtActiv.Text = Empty

' Verifica SI es un dato valido
If txtProy.Text <> "" And cboProy.Text <> "" Then
    ' Carga el combo con Descripciones de los programa para el proyecto elegidos
    ' Se inicializan las colecciones
    Set mcolCodProg = Nothing
    Set mcolCodDesProg = Nothing

    ' Carga los Programas relacionados con los proyectos
    sSQL = ""
    sSQL = "SELECT DISTINCT PresProy.IdProg, PresProy.IdProg + '   ' + Prog.DescProg " & _
           " FROM PROGRAMAS Prog, PRESUPUESTO_PROY PresProy" & _
           " WHERE Prog.IdProg=PresProy.IdProg " & _
           " AND PresProy.Idproy=" & "'" & txtProy.Text & "' ORDER BY PresProy.IdProg + '   ' + Prog.DescProg "

    CD_CargarColsCbo cboProg, sSQL, mcolCodProg, mcolCodDesProg
End If

End Sub

Public Sub ActualizarFinanancieraPeriodo()
'----------------------------------------------------------------------------
'PROPÓSITO: Actualizar los controles referentes a financiera, despues de ingresar un proyecto
'Recibe:    nada
'Devuelve:  nada
'----------------------------------------------------------------------------
' nota: llamado desde textbox, y combobox de proyecto al ingresar un proyecto

Dim sSQL As String
Dim curFinanPerioProy As New clsBD2

'Recupera financiera del proyecto seleccionado
sSQL = ""
    sSQL = "SELECT F.DescFinan, P.PerioProy, P.FecInicio " & _
           "FROM PROYECTOS P, Tipo_Financieras F " & _
           "WHERE P.IdProy=" & "'" & txtProy.Text & "'" & _
           " And P.IdFinan=F.IdFinan"
       
curFinanPerioProy.SQL = sSQL

' ejecuta la consulta y asignamos al txt de proyecto
If curFinanPerioProy.Abrir = HAY_ERROR Then
  Unload Me
  End
End If
'Carga las variables del modulo
txtFinan.Text = curFinanPerioProy.campo(0)
'miPerioProy = curFinanPerioProy.campo(1)
'msFecInicio = curFinanPerioProy.campo(2)
txtFinan.Visible = True
lblFinan.Visible = True

curFinanPerioProy.Cerrar 'Cierra el cursor

End Sub

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
  'NO se encuentra la CtaCte
  msCtaCte = ""
End If

'Carga el saldo de la Cta corriente
CargarSaldo

'Cambia el alto del combo
cboCtaCte.Height = CBONORMAL

' Habilitar el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtProy_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
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

Private Sub txtRinde_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub txtRUCDNI_Change()
On Error GoTo mnjError

'Verifica  que NO tenga campos en blanco
If txtRUCDNI.Text <> Empty And (Len(txtRUCDNI.Text) = 8 _
   Or Len(txtRUCDNI.Text) = 11) Then
   If gbProvNuevoConAfecta = True Then
        'Carga la coleccion con los datos
        CargaColDatos
        'Recupera los datos si se encuentra en la colección
        txtNombrProv.Text = Var30(gcolTabla.Item(txtRUCDNI.Text), 3)
        gsIdProv = Var30(gcolTabla.Item(txtRUCDNI.Text), 2)
        ' Los campos coloca a color blanco
        txtRUCDNI.BackColor = vbWhite
   Else
        'Recupera los datos si se encuentra en la colección
        txtNombrProv.Text = Var30(gcolTabla.Item(txtRUCDNI.Text), 3)
        gsIdProv = Var30(gcolTabla.Item(txtRUCDNI.Text), 2)
        ' Los campos coloca a color blanco
        txtRUCDNI.BackColor = vbWhite
    End If
    '-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        If Len(txtRUCDNI.Text) = 11 Then
            'Muestra el mensaje
            If MsgBox("El código ingresado no existe en la BD" & _
                ", Desea verificar en la ventana de Busqueda? ", vbYesNo + vbInformation, "SGCcaijo-Egreso con Afectación") = vbYes Then
                txtRUCDNI.Text = Empty
                ' Carga los títulos del grid selección
                 giNroColMNSel = 8
                 aTitulosColGrid = Array("Nro.RUC/DNI", " IdProv.", "Descripción", "RUC/DNI", "Dirección", "Teléfono", "Fax", "Representante")
                 aTamañosColumnas = Array(1500, 1000, 4500, 800, 3500, 2000, 2000, 3000)
                'Muestra el formulario de busqueda de proveedores
                frmMNSelecProvCaja.Show vbModal, Me
            Else
                'Termina la ejecución del procediminto
                Exit Sub
            End If
        End If
    End If
Else
    'Limpia el txtDescripción
    txtNombrProv.Text = Empty
    'Marca los campos obligatorios
    txtRUCDNI.BackColor = Obligatorio
End If

'Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub CargaColDatos()
'----------------------------------------------------------------------------
'Propósito: Carga una colección con los datos necesarios para el mantenimiento
'Recibe:  sBoton string que indica el nombre del botón presionado
'Devuelve: Nada
'----------------------------------------------------------------------------
Dim sSQL As String

' Sentencia SQL con cuyos datos se carga el grid
sSQL = "SELECT Numero, IdProveedor ,DescProveedor, RUC_DNI, Direc_Proveedor, " _
     & "Tel_Proveedor , Fax_Proveedor, Repre_Proveedor " _
     & " FROM PROVEEDORES ORDER BY IdProveedor"

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 8, gcolTabla

End Sub

Private Sub txtRUCDNI_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtTipDoc_Change()

' Si procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboTipDoc, txtTipDoc, mcolCodDesTipDoc

' Inicializa la variable que indica si canceló  la operación _
  cambio del tipo del documento
mbCancelaCambioTipDoc = False

' Verifica si se puede habilitar el botón para elegir impuestos
HabilitarBotonImpuestos
 
 ' Verifica SI el campo esta vacio
If txtTipDoc.Text <> "" And cboTipDoc.Text <> "" Then
    ' Los campos coloca a color blanco
    txtTipDoc.BackColor = vbWhite
   
    ' Habilita y carga el cboProdServ
    HabilitarOptsRelacPaga
    
    ' Verifica si se canceló la operación de cambio de tipdoc en OptPaga
    If mbCancelaCambioTipDoc = True Then
        ' pone el anterior valor de el documento
        gbKdown = True
        txtTipDoc = msAntTipDoc
        Exit Sub
    End If
    
    ' Verifica los datos en relación con el impuesto
    VerificarRelacTipDocImpuestos
    
    ' Actualiza el valor de tipo documento
    msAntTipDoc = txtTipDoc
Else
   'Marca los campos obligatorios
   txtTipDoc.BackColor = Obligatorio
   
   ' Deshabilita los opts
    optServicio.Enabled = False: optProducto.Enabled = False
End If

' Habilita botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub VerificarRelacTipDocImpuestos()

' Verifica si ya se aplicó impuestos a el documento
If gbImpuestos = True Then
    ' verifica si se cambió la relación con impuestos de la variable _
      que almacenaba la anterior.
    If gsRelacTributo <> Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1) _
    Then
        ' Verifica que los procesos de calcular montos de impts sean diferrentes
        If gsRelacTributo = "Retiene" Or Var30(mcolCodRetencionPaga.Item(txtTipDoc), 1) = "Retiene" _
        Then
            ' los calculos realizados  para los impuestos, no son correctos ahora _
              por que ha cambiado la relación que tiene el documento con los impuestos
            If mcolCodRetencionPaga.Count > 0 Then ' Verfica si lacoleccion de impuestos no es vacía
                
                ' Mensaje : debe calcular nuevamente los impuestos
                MsgBox " Debe calcular de nuevo los impuestos y añadir el detalle del Doc. " _
                , , "SGCcaijo Caja-Bancos, Egreso con afectación"
                ' Vacia la colección de impuestos
                Set gcolImpSel = Nothing
                ' Actualiza la variable que indica si se aplicó impuestos
                gbImpuestos = False
                
                 ' borra los controles del ingreso a detalle en la segunda parte _
                 del formulario, tambien reset al resumen
                 LimpiaControlesIngreDet
                 LimpiaCtrlsResumen
                
                ' Verifica si ha introducido el totaldoc
                If txtTotalDoc <> Empty Then
                    ' pone el focus a impuestos y muestra el formulario
                    cmdSelImpuestos.SetFocus
                    cmdSelImpuestos_Click
                Else ' pone el focus txttotaldoc
                    txtTotalDoc.Enabled = True
                End If ' fin de verificar si se introdujo el total del documento
                
            End If  ' fin de verificar la coleccion de impuestos
        Else ' Los Procesos no son diferentes pero si los montos que salen de Caja-Bancos
            
             ' Muestra el monto que saldrá de caja o de bancos cuando se cambie los impuestos
              txtMontoCB.Text = Format(CalculaMontoEgresoCB, "###,##0.00")
        
        End If ' fin de verifica procesos relacionados a impuestos sean diferentes
    
    Else ' no se cambio la relación de tipdoc con impuestos
             ' Muestra el monto que saldrá de caja o de bancos cuando se cambie los impuestos
              txtMontoCB.Text = Format(CalculaMontoEgresoCB, "###,##0.00")
    End If ' fin de verificar si se cambio la relación de tipdoc con impuestos
    
End If ' fin de verificar si se aplicó impuestos

End Sub


Private Sub VerificarTotalDocImpuestos()
'-----------------------------------------------------------------------
'Propósito: Verifica las estructuras de datos usados ,  que estas esten _
            ok con el total del documento.
'Recibe: Nada
'Entrega: Nada
'-----------------------------------------------------------------------

' Verifica si ya se aplicó impuestos a el documento
If gbImpuestos = True Then
    ' verifica si se cambió el monto total del documento
    If gdblMontoTotal <> Val(Var37(txtTotalDoc)) Then
      ' los calculos realizados  para los impuestos, no son correctos ahora _
         por que ha cambiado el total del documento.
            If mcolCodRetencionPaga.Count > 0 Then ' Verfica si la colección de impuestos no es vacía
                
                ' Mensaje : debe calcular nuevamente los impuestos
                MsgBox " Debe calcular de nuevo los impuestos, cambió el total del documento " _
                , , "SGCcaijo Caja-Bancos, Egreso con afectación"
                ' Vacia la colección de impuestos
                Set gcolImpSel = Nothing
                ' Actualiza la variable que indica si se aplicó impuestos
                gbImpuestos = False
                
                ' borra los controles del ingreso a detalle en la segunda parte _
                 del formulario, tambien reset al resumen
                 LimpiaControlesIngreDet
                 LimpiaCtrlsResumen
                
                ' Verifica si ha introducido el tipo de documento
                If txtTipDoc <> Empty And cboTipDoc <> Empty Then
                    ' pone el focus a impuestos y muestra el formulario
                    cmdSelImpuestos.SetFocus
                    cmdSelImpuestos_Click
                Else ' pone el focus txttipodocumento
                    txtTipDoc.SetFocus
                End If ' fin de verificar si se introdujo el tipo de documento
                
                ' Verifica si se cambio el alto del control cboProdServ, estado normal
                If cboProdServ.Height = CBOALTO Then CambiarAltocbo cboProdServ
                 
            End If  ' fin de verificar la coleccion de impuestos
    End If ' fin de verificar si se cambio el total del documento
End If ' fin de verificar si se aplicó impuestos

End Sub

Private Sub LimpiaCtrlsResumen()
'----------------------------------------------------------
' Proposito: Limpia los controles resumen del formulario
'----------------------------------------------------------

txtMontoCB = "0.00"
txtMontoImpuesto = "0.00"

End Sub

Private Sub HabilitarOptsRelacPaga()
Dim RelacPaga As String
' Asigna la variable que indica la relación del Doc con lo que paga
RelacPaga = Var30(mcolCodRetencionPaga.Item(txtTipDoc), 2)

'Habilita o desabilita los opts de acuerdo al tipo de documento
'Verifica si es servicio el campo relacion del tipo de documento
If RelacPaga = "S" Then
   optServicio.Value = True
   optServicio.Enabled = True
   optProducto.Enabled = False
   
'Verifica si es producto
ElseIf RelacPaga = "P" Then
   optProducto.Value = True
   optProducto.Enabled = True
   optServicio.Enabled = False
'Si es servicio o producto el campo relacion del tipo de documento
Else
   'habilita las dos opciones para que el usuario elija que pagar
   optProducto.Enabled = True
   optServicio.Enabled = True
End If

End Sub

Private Sub txtTipDoc_KeyPress(KeyAscii As Integer)

' Si se presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtTotalDoc_Change()

'Verifica SI el campo esta vacio
If txtTotalDoc.Text <> "" And Val(txtTotalDoc.Text) <> 0 Then
  'El campos coloca a color blanco
   txtTotalDoc.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtTotalDoc.BackColor = Obligatorio
End If

' Verifica si se puede habilitar el botón para elegir impuestos
HabilitarBotonImpuestos

' Habilitar botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtTotalDoc_GotFocus()

' Pone el alto del cbo a normal
If cboProdServ.Height = CBOALTO Then cboProdServ.Height = CBONORMAL
txtTotalDoc.MaxLength = 12
'Elimina las comas
txtTotalDoc.Text = Var37(txtTotalDoc.Text)
End Sub

Private Sub txtTotalDoc_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Monto ingresado en soles, luego se ubica en la componente inmediata
Var33 txtTotalDoc, KeyAscii


End Sub

Private Sub txtTotalDoc_LostFocus()

' Maximo número de digitos para el monto
txtTotalDoc.MaxLength = 14

If txtTotalDoc.Text <> "" Then
   'Da formato de moneda
   txtTotalDoc.Text = Format(Val(Var37(txtTotalDoc.Text)), "###,###,###,##0.00")
       
   ' Verifica si se cambio los impuestos
   VerificarTotalDocImpuestos
   
Else
   'coloca el control obligatorio
   txtTotalDoc.BackColor = Obligatorio
End If

End Sub

Private Sub HabilitarBotonImpuestos()
' ---------------------------------------------------------------------
' Propósito: habilitar el botón SelImpuestos cuando se ha elegido un tipo
'            de documento y se ha puesto un monto total al documento
' Recibe: Nada
' Entrega : Nada
' ---------------------------------------------------------------------

' Verifica si se ha introducido al control txtTotalDoc y txtTipoDoc

If txtTipDoc <> Empty And cboTipDoc <> Empty Then
    If txtTotalDoc.Text <> Empty Then
        ' se ha introducido los datos en ambos controles, habilita el botón
        cmdSelImpuestos.Enabled = True
    Else ' no se se ha introducido los datos en ambos controles, deshabilita el botón
        cmdSelImpuestos.Enabled = False
    End If
Else
 ' no se se ha introducido los datos en ambos controles, deshabilita el botón
        cmdSelImpuestos.Enabled = False
End If

End Sub

Private Sub txtValorCompra_GotFocus()

' Verificar si se a introducido un tipo de documento
If fbOperarDetalle = False Then
   Exit Sub
End If

' Verifica si se ha introducido el valor en txtcantidad
If fbVerificarCantidad = False Then SendKeys vbTab

txtValorCompra.MaxLength = 12
'Elimina las comas
txtValorCompra.Text = Var37(txtValorCompra.Text)
End Sub

Private Sub txtValorVenta_Change()

' SI NO se ha introducido valor, se marca campo obligatorio
If txtValorVenta.Text <> "" And Val(txtValorVenta.Text) <> 0 Then
   ' coloca el color de correcto
   txtValorVenta.BackColor = vbWhite
   
Else
   ' coloca el color obligatorio al control
   txtValorVenta.BackColor = Obligatorio

End If

' se esta cambiando el valor de venta. Calcula PrecioUniVV y si
 CalcularPrecioUniVV


' Verifica si el cambio del monto de venta se originó en el control o es _
  hecho por el calculo al cambiar el monto de compra
If mbCalculando = False And txtValorVenta <> Empty Then
    ' Calcula los valores de Compra y precio unitario de compra si se giró _
      sin impuestos incluidos en el detalle
    CalculaValores "ValorVenta"
ElseIf txtValorVenta = Empty Then
    ' el control es vacío, actualiza la condición de compra
    mbCalculando = True
    txtValorCompra = Empty
    mbCalculando = False
End If

' HabilitaAñadir
    HabilitaBotonAñadir
    
End Sub

Private Sub HabilitaBotonAñadir()
' -------------------------------------------------------
' Propósito: Verifica si se puede habilitar el botón añadir
' Recibe: Nada
' Entrega: Nada
' -------------------------------------------------------
' verifica si estan completos los controles

If cboProdServ.BackColor <> vbWhite Or txtCant.BackColor <> vbWhite _
   Or txtValorVenta.BackColor <> vbWhite Or txtValorCompra.BackColor <> vbWhite Then
    ' deshabilita el botón añadir
    cmdAñadir.Enabled = False
    cmdEliminar.Enabled = False
Else
    ' habilita e botón añadir
    cmdAñadir.Enabled = True
End If

End Sub


Private Sub CalcularPrecioUniVV()

If Val(Var37(txtCant)) > 0 And Val(Var37(txtValorVenta)) > 0 Then
 ' se tiene una cantidad aceptable, calcula el precio unitario
 txtPrecioUniVenta = Format(Var37(txtValorVenta) / Var37(txtCant), "###,###,##0.00")
Else
 ' la cantidad es cero, no aceptable o el txtcantidad esta vacía, entonces _
   muestra cero
  txtPrecioUniVenta = "0.00"
End If

End Sub

Private Sub CalcularPrecioUniVC()
'----------------------------------------------------------------------
' Propósito: Calcula el precio unitario de compra de acuerdo al monto _
            de compra y la cantidad
' Recibe : nada
' Entrega : nada
'----------------------------------------------------------------------

' Verifica si la cantidad tiene un valor > 0

If Val(Var37(txtCant)) > 0 And Val(Var37(txtValorCompra)) > 0 Then
 ' se tiene una cantidad aceptable, calcula el precio unitario
 txtPrecioUniCompra = Format(Var37(txtValorCompra) / Var37(txtCant), "###,###,##0.00")
Else
 ' la cantidad es cero, no aceptable o el txtcantidad esta vacía, entonces _
   muestra cero
  txtPrecioUniCompra = "0.00"
End If

End Sub


Private Sub txtValorVenta_GotFocus()

' Verificar si se a introducido un tipo de documento
If fbOperarDetalle = False Then
    Exit Sub
End If

' Verifica si se ha introducido el valor en txtcantidad
If fbVerificarCantidad = False Then SendKeys vbTab

txtValorVenta.MaxLength = 12
'Elimina las comas
txtValorVenta.Text = Var37(txtValorVenta.Text)
End Sub

Private Function fbVerificarCantidad() As Boolean
fbVerificarCantidad = False
If txtCant.Text <> Empty Then
    ' se ha introducido la cantidad
    fbVerificarCantidad = True
End If

End Function

Private Sub txtValorVenta_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Monto ingresado en soles, luego se ubica en la componente inmediata
Var33 txtValorVenta, KeyAscii

End Sub

Private Sub txtValorVenta_LostFocus()

'Maxima longitud
txtValorVenta.MaxLength = 14
If txtValorVenta.Text <> "" Then
   'Da formato de moneda
   txtValorVenta.Text = Format(Val(Var37(txtValorVenta.Text)), "###,###,###,##0.00")
Else
   txtValorVenta.BackColor = Obligatorio
End If


End Sub

Private Sub txtValorCompra_Change()

' SI NO se ha introducido valor, se marca campo obligatorio
If txtValorCompra.Text <> "" And Val(txtValorCompra.Text) <> 0 Then
   ' coloca el color de correcto
   txtValorCompra.BackColor = vbWhite
   
   ' se esta cambiando el valor de compra. Calcula PrecioUniVC y si
   CalcularPrecioUniVC

Else
    ' coloca el color obligatorio
   txtValorCompra.BackColor = Obligatorio
End If

' Verifica si el cambio del monto de compra se originó en el control o es _
  hecho por el calculo al cambiar el monto de venta
If mbCalculando = False And txtValorCompra <> Empty Then
    ' Calcula los valores de Venta y precio unitario de venta si se giró _
      sin impuestos incluidos en el detalle
    CalculaValores "ValorCompra"
ElseIf txtValorCompra = Empty Then
    ' el control es vacío, actualiza la condición de compra
    mbCalculando = True
    txtValorVenta = Empty
    mbCalculando = False
End If

End Sub

Private Sub txtValorCompra_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Monto ingresado en soles, luego se ubica en la componente inmediata
Var33 txtValorCompra, KeyAscii

End Sub

Private Sub txtValorCompra_LostFocus()

'Maxima longitud
txtValorCompra.MaxLength = 14
If txtValorCompra.Text <> "" Then
   'Da formato de moneda
   txtValorCompra.Text = Format(Val(Var37(txtValorCompra.Text)), "###,###,###,##0.00")
Else
   txtValorCompra.BackColor = Obligatorio
End If


End Sub

Private Sub DeshabilitarControlesPorAnioCerrado(bBoleano As Boolean)
  txtProy.Enabled = True: cboProy.Enabled = True
  txtProg.Enabled = True: cboProg.Enabled = True
  txtLinea.Enabled = True: cboLinea.Enabled = True
  txtActiv.Enabled = True: cboActiv.Enabled = True
  txtTipDoc.Enabled = bBoleano: cboTipDoc.Enabled = bBoleano
  txtRUCDNI.Enabled = bBoleano: txtNombrProv.Enabled = bBoleano
  txtDocEgreso.Enabled = bBoleano
  txtObserv.Enabled = bBoleano
  txtTotalDoc.Enabled = bBoleano
  txtBanco.Enabled = bBoleano: cboBanco.Enabled = bBoleano
  cboCtaCte.Enabled = bBoleano: txtNumCheque.Enabled = bBoleano
  fraPagar.Enabled = bBoleano: fraTipoGiro.Enabled = bBoleano
  cboProdServ.Enabled = bBoleano
  txtCant.Enabled = bBoleano
  txtValorVenta.Enabled = bBoleano: txtValorCompra.Enabled = bBoleano
  grdDetalle.Enabled = bBoleano
  txtRinde.Enabled = bBoleano: txtDescRinde.Enabled = bBoleano: cmdBuscaRinde.Enabled = bBoleano
  cmdSelImpuestos.Enabled = bBoleano: cmdBuscar.Enabled = bBoleano
  mskFecDoc.Enabled = bBoleano
End Sub

Private Sub ActualizarcboCategoriaGasto()
Dim sSQL As String
Dim mcurCatGasto As New clsBD2
'--------------------------------------------------------------
'Propósito: Actualiza los cboProg y txtProg cuando se cambia el proyecto
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------
' nota :    Llamado desde el evento change de txtproy y Modificar

    cboCategoriaGasto.Clear
    TxtCategoriaGasto.Text = Empty

  ' Verifica SI es un dato valido
  If txtProy.Text <> "" And cboProy.Text <> "" Then
    ' Carga el combo con Descripciones de los programa para el proyecto elegidos
    ' Se inicializan las colecciones
    Set mcolCodCatGasto = Nothing
    Set mcolCodDesCatGasto = Nothing
    
    sSQL = ""
    sSQL = "SELECT CG.CODCATGASTO, CG.CODCATGASTO + '   ' + CG.DESCRIPCIONGASTO " & _
           " FROM PROYECTOS P, CATEG_GASTO CG " & _
           " WHERE P.IDFINAN = CG.IDFINANCIERA " & _
           " AND P.Idproy = " & "'" & txtProy.Text & "' " & _
           " ORDER BY CG.CODCATGASTO + '   ' + CG.DESCRIPCIONGASTO "
    
    mcurCatGasto.SQL = sSQL
    ' Se abre el cursor y se tratan errores
    If mcurCatGasto.Abrir = HAY_ERROR Then
      End
    End If
    
    'No existe ningún producto que empiece con el codigo sCodigo
    If Not mcurCatGasto.EOF Then
      ' Carga los Programas relacionados con los proyectos
      Label10.Visible = True
      TxtCategoriaGasto.Visible = True
      cboCategoriaGasto.Visible = True
      cmdPCategoriaGasto.Visible = True
      
      TxtCategoriaGasto.BackColor = Obligatorio
      
      CD_CargarColsCbo cboCategoriaGasto, sSQL, mcolCodCatGasto, mcolCodDesCatGasto
    Else
      Label10.Visible = False
      TxtCategoriaGasto.Visible = False
      cboCategoriaGasto.Visible = False
      cmdPCategoriaGasto.Visible = False
      
      TxtCategoriaGasto.BackColor = vbWhite
    End If
  End If
End Sub

Private Sub RecuperarCtaContableActiv()
Dim sSQL As String
Dim curCtaContable As New clsBD2

CtaContableActividadEmpresa = ""
' Verifica SI el campo esta vacio
If txtActiv.Text <> "" And cboActiv.Text <> "" Then
    ' Carga el combo con Descripciones de los programa para el proyecto elegidos
    sSQL = "SELECT Distinct PRESUPUESTO_PROY.CtaContable " & _
           " FROM PRESUPUESTO_PROY " & _
           " WHERE PRESUPUESTO_PROY.Idproy=" & "'" & txtProy.Text & "'" & _
           " AND PRESUPUESTO_PROY.Idprog=" & "'" & txtProg.Text & "'" & _
           " AND PRESUPUESTO_PROY.IdLinea=" & "'" & txtLinea.Text & "'" & _
           " AND PRESUPUESTO_PROY.IdActiv=" & "'" & txtActiv.Text & "'"
    ' Ejecuta la sentencia
  curCtaContable.SQL = sSQL
  If curCtaContable.Abrir = HAY_ERROR Then End
  
  ' verifica si tiene algún producto verificado
  If Not curCtaContable.EOF Then
    CtaContableActividadEmpresa = curCtaContable.campo(0)
  End If
  
  curCtaContable.Cerrar
End If
End Sub

Public Sub CargarCboProdServ(cboRec As ComboBox, colRec As Collection)
'----------------------------------------------------------------------------
'Propósito: Carga combo apartir de colecciones
'Recibe:   cboRec (Combo donde se carga), colRec (Coleccion de donde se carga el combo)
'Devuelve: Nada
'----------------------------------------------------------------------------

Dim i As Integer
'Carga Combo apartir de Colecciones
For i = 1 To colRec.Count
  If (TipoEgreso = "PROY") And (Var30(colRec(i), 3) = "PROY") Then
    cboRec.AddItem Var30(colRec(i), 2)
  ElseIf (TipoEgreso = "EMPR") And (Var30(colRec(i), 3) = "EMPR") Then
    cboRec.AddItem Var30(colRec(i), 2)
  End If
Next i

End Sub

Public Sub ActualizarInfoProdServ(sTextoCbo As String, txtRec As String, colCod As Collection, colCodDesc As Collection)
  Dim i As Integer ' Contador de bucle For
  
  ' Se busca la descripción en la colección de códigos+descripciones
  For i = 1 To colCodDesc.Count
    If (Var30(colCodDesc(i), 2) = sTextoCbo) And (Var30(colCodDesc(i), 3) = TipoEgreso) Then  ' Elemento encontrado
      txtRec = colCod(i) ' Actualiza código
      bExisteCod = True
      Exit For
    End If
  Next
End Sub

Public Sub ActualizarVariableCbo(cboRec As ComboBox, txtRec As String, colCodDesc As Collection, Optional sNomCtlProbl As String)
  Dim i As Integer ' Contador de bucle For

  On Error GoTo ErrClaveCol
      ' Tratamiento del error producido al intentar acceder mediante
      ' clave a una colección
  
  ' SI es una activación NO deseada, NO hay nada que hacer
  If sNomCtlProbl = cboRec.Name Then  ' Ejecución NO deseada
      Exit Sub
  End If
  
  ' SI el código es "-", se da por existente
  '     Nota: para ciertos campos, se guarda en la BD un carácter "-" en
  '     vez de " " porque se han observado comportamientos extraños
  '     (_muy_ extraños) en los filtros de Crystal Reports.
  If txtRec = "-" Then
      Exit Sub
  End If
  
  bExisteCod = False
    
  ' SI busca el código en la lista de valores del ComboBox
  For i = 0 To cboRec.ListCount - 1
    'If cboRec.List(i) = colCodDesc(txtRec) Then ' Descripción buscada
    If cboRec.List(i) = txtRec Then ' Descripción buscada
      cboRec.ListIndex = i ' Se actualiza el combo
      cboRec.BackColor = vbWhite
      bExisteCod = True
      Exit For
    End If
  Next i
  
PostErrClaveCol:
  
  ' Se pasa por aquí _después_ de tratar un error producido al
  ' intentar acceder a una colección por clave
  
  If Not bExisteCod Then
      MsgBox "No existe el código '" & Trim(txtRec) & "'", vbExclamation, "Aviso"
  End If
  
  Exit Sub
  
  '-------------------------------------------------------------------
ErrClaveCol:
  
  If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
      bExisteCod = False
      Resume PostErrClaveCol ' La ejecución sigue por aquí
  End If
End Sub
