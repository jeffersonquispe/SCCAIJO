VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPRMEConsulSaldo 
   Caption         =   "Ctas. Moneda Extranjera- Consulta de Saldos"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   6840
      TabIndex        =   1
      Top             =   2280
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdSaldos 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      Cols            =   5
   End
End
Attribute VB_Name = "frmPRMEConsulSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim sSQL As String
Dim curCuentas As New clsBDConsulta
Dim curIngreso As New clsBDConsulta
Dim curEgreso As New clsBDConsulta
Dim iSaldo As String

'Se ajusta el ancho y se cargan los títulos de las columnas del grid
aTitulosColGrid = Array("Nº Cuenta", "Banco", "Ingreso", "Egreso", "Saldo")
aTamañosColumnas = Array(1100, 2800, 1200, 1200, 1200)

grdSaldos.Row = 0
For iCol = 0 To grdSaldos.Cols - 1
  grdSaldos.Col = iCol
  grdSaldos.ColWidth(iCol) = aTamañosColumnas(iCol)
  grdSaldos.Text = aTitulosColGrid(iCol)
Next

'Se seleccionan las cuentas en dólares
sSQL = "SELECT  c.desccta, b.descbanco,c.idcta FROM Tipo_Bancos B, CuentasBanc C" & _
       " WHERE b.Idbanco = c.Idbanco AND c.idmoneda = 'USD' " & _
       " ORDER BY  b.descbanco,c.desccta"
curCuentas.SQL = sSQL
If curCuentas.Abrir = HAY_ERROR Then
  End
End If

Do While Not curCuentas.EOF
       
  'Hallar el Ingreso
  sSQL = "SELECT sum(montodol)From ingreso_ctas_extr " & _
        "WHERE idcta='" & curCuentas.campo(2) & "' and Anulado='No'"
          
  curIngreso.SQL = sSQL
  If curIngreso.Abrir = HAY_ERROR Then
    End
  End If
      
  'Hallar el Egreso
  sSQL = "SELECT SUM(Montodol)" & _
        " FROM Egreso_Ctas_Extr AS E, Det_Egreso_Ctas_Extr AS D" & _
        " WHERE d.idcta='" & curCuentas.campo(2) & _
        "' AND E.idegreso = D.idegreso and E.Anulado='No'"
          
  curEgreso.SQL = sSQL
  If curEgreso.Abrir = HAY_ERROR Then
    End
  End If
        
  'Se calcula el saldo
  If IsNull(curEgreso.campo(0)) Then
    iSaldo = Format(curIngreso.campo(0), "###,###,##0.00")
  
  Else
    '(Ingreso - Egreso)
    iSaldo = Format(curIngreso.campo(0) - curEgreso.campo(0), "###,###,##0.00")
      
  End If
  
  grdSaldos.AddItem curCuentas.campo(0) & vbTab & curCuentas.campo(1) & vbTab & _
                  Format(curIngreso.campo(0), "###,###,##0.00") & vbTab & _
                  Format(curEgreso.campo(0), "###,###,##0.00") & vbTab & _
                  iSaldo
                  
  curCuentas.MoverSiguiente
Loop

curCuentas.Cerrar

End Sub
