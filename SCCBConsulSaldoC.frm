VERSION 5.00
Begin VB.Form frmCBConsulSaldoC 
   Caption         =   "Caja y Bancos- Consulta de Saldo de Caja"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   HelpContextID   =   69
   Icon            =   "SCCBConsulSaldoC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEgresos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1725
      Width           =   1770
   End
   Begin VB.TextBox txtIngresos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1245
      Width           =   1770
   End
   Begin VB.TextBox txtSaldo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   765
      Width           =   1785
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   3330
      TabIndex        =   3
      Top             =   2640
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totales de Caja (S/.)"
      Height          =   2295
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4155
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Egresos:"
         Height          =   195
         Left            =   930
         TabIndex        =   7
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ingresos:"
         Height          =   195
         Left            =   885
         TabIndex        =   6
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo de Caja:"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   600
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmCBConsulSaldoC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim sSQL As String
Dim curCuentas As New clsBD2
Dim curIngreso As New clsBD2
Dim curEgreso As New clsBD2
Dim iSaldo As String
Dim iCol As Integer
Dim curEmpresas As New clsBD2
Dim EmpresasExistentes As String
Dim InstrucEmpresas As String
Dim TotalEgresoProyectos As Double
Dim TotalEgresoEmpresasSinRH As Double
Dim TotalEgresoEmpresasSoloRHCB As Double
Dim TotalEgresos As Double

'Hallar el Ingreso
sSQL = "SELECT sum(monto)From Ingresos " & _
        "WHERE  LEFT(Orden,2)= 'CA' AND Anulado='NO'"
          
  curIngreso.SQL = sSQL
  If curIngreso.Abrir = HAY_ERROR Then
    End
  End If
      
'Se muestran los Ingresos
If IsNull(curIngreso.campo(0)) Then
  txtIngresos.Text = "0"
Else
  txtIngresos.Text = Format(curIngreso.campo(0), "###,###,##0.00")
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
txtEgresos.Text = Format(TotalEgresos, "###,###,##0.00")
iSaldo = Val(Var37(txtIngresos.Text)) - Val(Var37(txtEgresos.Text))

'Se muestran los valores
If txtIngresos.Text = "0" And txtEgresos.Text = "0" Then
  txtSaldo.Text = "0.00"
Else
  txtSaldo.Text = Format(iSaldo, "###,###,##0.00")
End If

'Se cierran los cursores
curIngreso.Cerrar


End Sub




