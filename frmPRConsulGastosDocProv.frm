VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPProyecto 
      Height          =   255
      Left            =   6345
      Picture         =   "frmPRConsulGastosDocProv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   285
      Width           =   220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta"
      Height          =   960
      Left            =   120
      TabIndex        =   14
      Top             =   1140
      Width           =   11415
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   2640
         TabIndex        =   16
         Top             =   120
         Width           =   4935
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   330
            Left            =   1245
            TabIndex        =   17
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   330
            Left            =   3525
            TabIndex        =   18
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha &Fin:"
            Height          =   195
            Left            =   2640
            TabIndex        =   20
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha &Inicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   285
            Width           =   915
         End
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin MSMask.MaskEdBox mskFecConsulta 
         Height          =   315
         Left            =   9960
         TabIndex        =   21
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
      Begin VB.Label Label3 
         Caption         =   "&Fecha consulta:"
         Height          =   255
         Left            =   8280
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   400
      Left            =   8640
      TabIndex        =   1
      Top             =   8160
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   10320
      TabIndex        =   0
      Top             =   8160
      Width           =   1000
   End
   Begin Crystal.CrystalReport rptInforme 
      Left            =   10200
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   10398
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      MergeCells      =   4
   End
   Begin VB.ComboBox cboProyecto 
      Height          =   315
      Left            =   1680
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   240
      Width           =   4920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proyecto"
      Height          =   1160
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11415
      Begin VB.TextBox txtProyecto 
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtNumPeriodo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10320
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtFinan 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   5355
      End
      Begin MSMask.MaskEdBox mskFecInicioProy 
         Height          =   315
         Left            =   9960
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "&Fecha inicio proyecto:"
         Height          =   255
         Left            =   8280
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "&Proyecto:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFinan 
         Caption         =   "Financiera:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número de periodos:"
         Height          =   195
         Left            =   8280
         TabIndex        =   10
         Top             =   720
         Width           =   1470
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
