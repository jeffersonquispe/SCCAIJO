VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBConsulSaldoB 
   Caption         =   "Caja y Bancos - Consulta de Saldos Bancarios"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   HelpContextID   =   69
   Icon            =   "SCCBConsulSaldoB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   7800
      TabIndex        =   1
      Top             =   4080
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdSaldos 
      Height          =   3945
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   6959
      _Version        =   393216
      Rows            =   1
      Cols            =   5
   End
End
Attribute VB_Name = "frmCBConsulSaldoB"
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

'Se ajusta el ancho y se cargan los títulos de las columnas del grid
aTitulosColGrid = Array("Nº Cuenta", "Banco", "Ingreso", "Egreso", "Saldo")
aTamañosColumnas = Array(1300, 2800, 1400, 1400, 1400)

grdSaldos.Row = 0
For iCol = 0 To grdSaldos.Cols - 1
  grdSaldos.Col = iCol
  grdSaldos.ColWidth(iCol) = aTamañosColumnas(iCol)
  grdSaldos.Text = aTitulosColGrid(iCol)
Next

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

'Se seleccionan las cuentas en dólares
sSQL = "SELECT  c.desccta, b.descbanco,c.idcta FROM Tipo_Bancos B, TIPO_CUENTASBANC C" & _
       " WHERE b.Idbanco = c.Idbanco AND c.idmoneda = 'SOL' " & _
       " ORDER BY  b.descbanco,c.desccta"
curCuentas.SQL = sSQL
If curCuentas.Abrir = HAY_ERROR Then
  End
End If

Do While Not curCuentas.EOF
       
  'Hallar el Ingreso
  sSQL = "SELECT sum(monto)From Ingresos " & _
        "WHERE idcta='" & curCuentas.campo(2) & "' and Anulado='NO'"
          
  curIngreso.SQL = sSQL
  If curIngreso.Abrir = HAY_ERROR Then
    End
  End If
      
  'Hallar el Egreso
  '*-*-*-*-*
  '*-*-*-*-*  TOTAL DE EGRESOS PARA PROYECTOS CON AFECTACION Y SIN AFECTACION
  '*-*-*-*-*
  sSQL = "SELECT SUM(MontoCB) FROM Egresos" & _
        " WHERE idcta='" & curCuentas.campo(2) & _
        "' AND Anulado='NO' AND " & InstrucEmpresas
          
  curEgreso.SQL = sSQL
  If curEgreso.Abrir = HAY_ERROR Then
    End
  End If
        
  'Se calcula el saldo
  If IsNull(curEgreso.campo(0)) Then
    TotalEgresoProyectos = "0"
  Else
    TotalEgresoProyectos = curEgreso.campo(0)
  End If
  
  'Se cierra el cursor
  curEgreso.Cerrar
  
  'Hallar el Egreso
  '*-*-*-*-*
  '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SIN RH
  '*-*-*-*-*
  sSQL = "SELECT SUM(MontoAfectado) FROM Egresos, PROYECTOS " & _
        " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.idcta='" & curCuentas.campo(2) & _
        "' AND Anulado='NO' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc<>'02') "
          
  curEgreso.SQL = sSQL
  If curEgreso.Abrir = HAY_ERROR Then
    End
  End If
        
  'Se calcula el saldo
  If IsNull(curEgreso.campo(0)) Then
    TotalEgresoEmpresasSinRH = "0"
  Else
    TotalEgresoEmpresasSinRH = curEgreso.campo(0)
  End If
  
  'Se cierra el cursor
  curEgreso.Cerrar
  
  'Hallar el Egreso
  '*-*-*-*-*
  '*-*-*-*-*  TOTAL EGRESOS PARA EMPRESAS SOLO RH
  '*-*-*-*-*
  sSQL = "SELECT SUM(MontoCB) FROM Egresos, PROYECTOS " & _
        " WHERE (EGRESOS.IdProy = PROYECTOS.IdProy) And EGRESOS.idcta='" & curCuentas.campo(2) & _
        "' AND Anulado='NO' And (PROYECTOS.Tipo = 'EMPR') and (IdTipoDoc='02') "
          
  curEgreso.SQL = sSQL
  If curEgreso.Abrir = HAY_ERROR Then
    End
  End If
        
  'Se calcula el saldo
  If IsNull(curEgreso.campo(0)) Then
    TotalEgresoEmpresasSoloRHCB = "0"
  Else
    TotalEgresoEmpresasSoloRHCB = curEgreso.campo(0)
  End If
  
  'Se cierra el cursor
  curEgreso.Cerrar
  
  TotalEgresos = TotalEgresoProyectos + TotalEgresoEmpresasSinRH + TotalEgresoEmpresasSoloRHCB
    
  iSaldo = curIngreso.campo(0) - TotalEgresos
  
  grdSaldos.AddItem curCuentas.campo(0) & vbTab & curCuentas.campo(1) & vbTab & _
                  Format(curIngreso.campo(0), "###,###,##0.00") & vbTab & _
                  Format(TotalEgresos, "###,###,##0.00") & vbTab & _
                  Format(iSaldo, "###,###,##0.00")
                  
  curCuentas.MoverSiguiente
Loop

curCuentas.Cerrar
curIngreso.Cerrar

End Sub


