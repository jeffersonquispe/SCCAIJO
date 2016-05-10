VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBEGPago_Planillas 
   Caption         =   "Pago de planillas"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   Icon            =   "SCCBEGPagoPlanillas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame frPlanilla 
      Caption         =   "Seleccione la planilla a pagar:"
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid grdPlanilla 
         Height          =   2895
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         HighLight       =   0
         FillStyle       =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monto a pagar:"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   3855
      Begin VB.TextBox txtMonto 
         Height          =   375
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMonto 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Pagar:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   3015
   End
End
Attribute VB_Name = "frmCBEGPago_Planillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variable que indica si el pago es en efectivo o bancario
Dim msTipoPago As String * 8
' Maneja el grid
Dim ipos As Long

Private Sub cmdAceptar_Click()
Dim bTodoOk As Boolean
Dim i As Integer
Dim iTamano As Integer

'Se comprueba que se haya marcasdo algúna fila
If grdPlanilla.Row < 1 Then
  MsgBox "Debe marcar algúna planilla por pagar", vbInformation + vbOKOnly, "SGCcaijo-Pago de planillas"
  Exit Sub
End If

' Verifica que se haya selecionado la primera planilla
If grdPlanilla.Row <> 1 Then
   ' Mensaje
   MsgBox "Primero pagar las planillas anteriores, Seleccione la 1ra planilla mostrada", , _
          "SGCcaijo-Pago de Planillas"
   ' Sale del proceso
   Exit Sub
End If

' Verifica si el monto por pagar es Mayor al saldo que queda por pagar
If Val(Var37(txtMonto)) > _
   Val(Var37(grdPlanilla.TextMatrix(grdPlanilla.Row, 2))) Then
   ' Mensaje
   MsgBox "El monto de pago definido es Mayor al saldo que queda" & Chr(13) _
        & "por pagar de la planilla seleccionada, Ingrese nuevamente", , _
          "SGCcaijo-Pago de Planillas"
   ' Coloca el focus a el txtmonto
   txtMonto.SetFocus
   ' Sale del proceso
   Exit Sub
End If

' Carga el detalle de la planilla
  CargaPlanillaCTB bTodoOk

' Verifica si el proceso CargaPlanilla esta Ok
  If bTodoOk = False Then
    ' Limpia la colección EgresoSA detalle
    Set gcolDetMovCB = Nothing
    ' Sale del proceso
    Exit Sub
  End If
' Pasa los datos del pago de planilla a egreso sin afectación
 iTamano = Len(grdPlanilla.TextMatrix(grdPlanilla.Row, 0))
 frmCBEGSinAfecta.txtDesc.Text = grdPlanilla.TextMatrix(grdPlanilla.Row, 1)
 frmCBEGSinAfecta.txtAfecta.MaxLength = iTamano
 frmCBEGSinAfecta.txtAfecta = grdPlanilla.TextMatrix(grdPlanilla.Row, 0)
 frmCBEGSinAfecta.txtMonto = txtMonto
' Cierra el formulario
  Unload Me

End Sub

Private Sub CargaPlanillaCTB(bTodoOk As Boolean)
'---------------------------------------------------------------
'Propósito: Carga las contracuentas y montos calculados para planilla
'Recibe : Nada
'Entrega : Nada
'----------------------------------------------------------------
Dim sSQL As String
Dim curPlanilla As New clsBD2
Dim dblMonto, dblSaldo As Double

' Asume que el proceso esta Ok
bTodoOk = True

' Carga la consulta que averigua las contracuentas de planilla
sSQL = "SELECT PL.CodPlanilla, PC.CodContable, PC.Monto " _
      & "FROM PLN_PLANILLAS PL, PLN_CTB_TOTALES PC " _
      & "WHERE PL.Codplanilla=PC.CodPlanilla and PL.PagadoCB='NO' and " _
      & "PL.CodPlanilla='" & grdPlanilla.TextMatrix(grdPlanilla.Row, 0) & "' " _
      & "ORDER BY PL.CodPlanilla, PC.CodContable"

' Ejecuta la sentencia
curPlanilla.SQL = sSQL
If curPlanilla.Abrir = HAY_ERROR Then End

' Verifica si no es vacío
If curPlanilla.EOF Then
    ' Error debe calcular la planilla de nuevo
  MsgBox "No se contabilizó correctamente esta planilla," & Chr(13) & _
         "Se deberá calcular nuevamente la planilla", "SGCcaijo-Pago Planillas"
Else ' Reparte el monto definido entre los saldos por pagar de las _
       contracuentas de Planillas
    dblMonto = Val(Var37(txtMonto.Text))
    ' Recorre las contracuentas y los saldos que faltan por pagar
    Do While (Not curPlanilla.EOF) And (Val(dblMonto) > 0)
        ' Averigua el saldo de la contracuenta
        dblSaldo = fAveriguarSaldoContracuenta(curPlanilla.campo(0), _
                                               curPlanilla.campo(1), _
                                               curPlanilla.campo(2))
       ' Verica si el saldo por pagar
        If Val(dblSaldo) < 0 Then
            MsgBox "Error: El sistema informa que se pagó mas que el saldo por pagar" & Chr(13) _
                 & "definido para una contracuenta de la planilla, Consulte al administrador", , _
                   "SGCcaijo-Verificar Pago de planillas"
            ' Error
            bTodoOk = False
            ' Cierra la consulta
            curPlanilla.Cerrar
            ' Sale de la función
            Exit Sub
            
        ElseIf Val(dblSaldo) > 0 Then
            ' Define el detalle del pago para las contracuentas de planillas
             If Val(dblMonto) >= Val(dblSaldo) Then ' El monto definido es Mayor al monto por pagar
             ' Añade al detalle el monto que falta pagar de la contracuenta _
               y reduce el monto por repartir.Código,ctactb,monto
               gcolDetMovCB.Add Item:=curPlanilla.campo(0) & "¯" _
                              & curPlanilla.campo(1) & "¯" _
                              & Format(dblSaldo, "##0.00"), _
                           Key:=curPlanilla.campo(0) & "¯" _
                              & curPlanilla.campo(1)
               dblMonto = Val(dblMonto) - Val(dblSaldo)
               dblMonto = Round(dblMonto, 2)
             Else ' El monto definido es Mayor al monto por pagar
             ' Añade al detalle el monto que falta pagar de la contracuenta _
               y reduce el monto por repartir. Código,ctactb,monto
               gcolDetMovCB.Add Item:=curPlanilla.campo(0) & "¯" _
                              & curPlanilla.campo(1) & "¯" _
                              & Format(dblMonto, "##0.00"), _
                           Key:=curPlanilla.campo(0) & "¯" _
                              & curPlanilla.campo(1)
               dblMonto = 0
             End If
        End If
       ' Si saldo es cero entonces pasa a la siguiente contracuenta
        
       ' Mueve al siguiente monto por pagar de las contracuentas
       curPlanilla.MoverSiguiente
    Loop

End If

' Cierra la componente
curPlanilla.Cerrar

' Verifica si esta Ok el detalle del pago de planillas
If Val(dblMonto) < 0 Then
    MsgBox "Error: No se pudo repartir el monto definido entre los " & Chr(13) _
         & "saldos de las contracuentas de la planilla, Consulte al administrador", , _
           "SGCcaijo-Verificar Pago de planillas"
    ' Error
    bTodoOk = False
ElseIf Val(dblMonto) > 0 Then
    MsgBox "Error: No se pudo repartir el monto definido es Mayor a los " & Chr(13) _
         & "saldos de las contracuentas de la planilla, Consulte al administrador", , _
           "SGCcaijo-Verificar Pago de planillas"
    ' Error
    bTodoOk = False
End If

End Sub

Private Function fAveriguarSaldoContracuenta(CodPlanilla As String, _
                            CtaCtb As String, Monto As Double) As Double
' ------------------------------------------------------------------
' Propósito: Averigua el saldo que falta por pagar para cubrir los montos _
             comprometidos a las contracuentas de planillas.
' Recibe : Código de la planilla, la contracuenta de planilla, _
           el Monto comprometido para la contracuenta.
' Entrega : Saldo por pagar de la contracuenta.
' ------------------------------------------------------------------
Dim curMontoPagado As New clsBD2
Dim sSQL As String

' Averigua los pagos realizados a las contracuentas del planillas
sSQL = "SELECT SUM(PP.Monto) " _
    & "FROM PAGO_PLANILLAS PP, EGRESOS E " _
    & "WHERE PP.CodPlanilla='" & CodPlanilla & "' and PP.CodContable='" & CtaCtb & "' and " _
    & "PP.Orden=E.Orden and E.Anulado='NO'"
' Ejecuta la  sentencia que calcula los pagos ralizados
curMontoPagado.SQL = sSQL
If curMontoPagado.Abrir = HAY_ERROR Then End

' Verifica si la consulta es vacía
If curMontoPagado.EOF Then
    ' No se ha pagado nada del total por pagar de la contracuenta
    ' Saldo=Monto comprometido
    fAveriguarSaldoContracuenta = Monto
Else
    ' Verifica si es nulo el monto pagado
    If IsNull(curMontoPagado.campo(0)) Then
        ' No se ha pagado nada del total por pagar de la contracuenta
        ' Saldo=Monto comprometido
        fAveriguarSaldoContracuenta = Monto
    Else ' La planilla ya tiene montos pagados
        fAveriguarSaldoContracuenta = Monto - curMontoPagado.campo(0)
    End If
End If

End Function

Private Sub cmdSalir_Click()

' Sale del formulario
Unload Me

End Sub

Private Sub Form_Load()

' Limpia la colección EgresoSA detalle
Set gcolDetMovCB = Nothing

' Pone el título al grid
aTitulosColGrid = Array("CodPlanilla", "Planilla", "Monto")
aTamañosColumnas = Array(0, 3800, 1500)

CargarGridTitulos grdPlanilla, aTitulosColGrid, aTamañosColumnas

' Define si el Pago de planillas es por caja(Efectivo) o banco(bancario)
AveriguaTipoPago

' Averigua y carga las planillas por pagar en efectivo o por bancos
CargaPlanillasporPagar

' Maneja los controles de acuerdo al tipo de pago
Manejacontroles

' Inicializa el grid
ipos = 0
gbCambioCelda = False

End Sub

Private Sub Manejacontroles()
' ------------------------------------------------------
' Propósito: Maneja los controles del formulario de acuerdo al tipo _
             de pago elegido(efectivo o bancario)
' Recibe : Nada
' Entrega : Nada
' ------------------------------------------------------
' Pone el color obligatorio a txtMontoPago
txtMonto.BackColor = Obligatorio
' Maneja los atributos de los controles
If msTipoPago = "Efectivo" Then
    ' Coloca el caption del grid y al monto a pagar
    frPlanilla.Caption = "Seleccione la planilla a pagar en efectivo:"
    lblMonto.Caption = "Monto en efectivo"
    ' Deshabilita el monto a pagar
    txtMonto.Enabled = False
Else
    ' Coloca el caption del grid y al monto a pagar
    frPlanilla.Caption = "Seleccione la planilla a pagar por banco:"
    lblMonto.Caption = "Monto por banco"
    ' Habilita el monto a pagar
    txtMonto.Enabled = True
End If

' Deshabilita el botón aceptar
cmdAceptar.Enabled = False

End Sub

Private Sub CargaPlanillasporPagar()
' ------------------------------------------------------
' Propósito: Carga la información de las planillas y los montos _
             por pagar, de acuerdo al tipo _
             de pago elegido (efectivo o bancario)
' Recibe : Nada
' Entrega : Nada
' ------------------------------------------------------
Dim sSQL As String
Dim curTotalFormaPago As New clsBD2
Dim curTotalPagado As New clsBD2
Dim sCajaBanco As String * 2
Dim sTablaPLNTipoPago As String

' De acuerdo al tipo de pago(Efectivo-Bancario), realiza la consulta
If msTipoPago = "Efectivo" Then ' Averigua los pagos en caja
    sCajaBanco = "CA"
    sTablaPLNTipoPago = "PLN_PAGOEFECTIVO TP "
Else ' Averigua los pagos en Banco
    sCajaBanco = "BA"
    sTablaPLNTipoPago = "PLN_PAGOBANCOS TP "
End If

' Carga las planillas que todavía no se han pagado totalmente y _
  sus montos a pagar en efectivo o banco de acuerdo al tipo de pago
sSQL = "SELECT P.CodPlanilla, P.DescPlanilla, sum(TP.Monto) " _
    & "FROM PLN_PLANILLAS P," & sTablaPLNTipoPago _
    & "WHERE P.PagadoCB='NO' and P.CodPlanilla=TP.CodPlanilla and " _
    & "P.CodPlanilla IN (SELECT DISTINCT CP.CodPlanilla FROM CTB_ASIENTOSPLANILLA CP) " _
    & "GROUP BY P.CodPlanilla, P.DescPlanilla " _
    & "ORDER BY P.CodPlanilla"

' Ejecuta la sentencia
curTotalFormaPago.SQL = sSQL
If curTotalFormaPago.Abrir = HAY_ERROR Then End
If curTotalFormaPago.EOF Then
    ' No existen planillas pendientes por pagar en efectivo
    curTotalFormaPago.Cerrar
    Exit Sub
End If

' Carga el grid
Do While Not curTotalFormaPago.EOF

    ' Carga el total de los montos pagados(efectivo o bancario) de las _
      planillas que faltan cancelar totalmente.
    sSQL = "SELECT SUM(PP.Monto) " _
        & "FROM PAGO_PLANILLAS PP, EGRESOS E " _
        & "WHERE PP.CodPlanilla='" & curTotalFormaPago.campo(0) & "' and " _
        & "PP.Orden=E.Orden and E.Anulado='NO' and LEFT(E.Orden,2)='" & sCajaBanco & "'"
    
    ' Ejecuta la sentencia
    curTotalPagado.SQL = sSQL
    If curTotalPagado.Abrir = HAY_ERROR Then End
    
    If curTotalPagado.EOF Then
        ' No se ha pagado nada del total por pagar(Efectivo o bancario).
        ' Añade el monto total por pagar(Efectivo o bancario)
        grdPlanilla.AddItem curTotalFormaPago.campo(0) & _
            vbTab & curTotalFormaPago.campo(1) & _
            vbTab & Format(curTotalFormaPago.campo(2), "###,###,##0.00")
    Else
        ' Verifica si es nulo el monto pagado
        If IsNull(curTotalPagado.campo(0)) Then
            ' No se pago nada, añade el total por pagar(Efectivo o bancario)
            grdPlanilla.AddItem curTotalFormaPago.campo(0) & _
                vbTab & curTotalFormaPago.campo(1) & _
                vbTab & Format(curTotalFormaPago.campo(2), "###,###,##0.00")
        Else ' La planilla ya tiene montos pagados
            ' Calcula y añade el saldo que falta pagar en (Efectivo o bancario)
            If Val(curTotalFormaPago.campo(2)) <= Val(curTotalPagado.campo(0)) Then
                If Val(curTotalFormaPago.campo(2)) < Val(curTotalPagado.campo(0)) Then
                     ' Error se pago mas de lo comprometido
                     MsgBox "Error : Se ha pagado de más para: " & curTotalFormaPago.campo(1) & Chr(13) _
                            & "Verifique los pagos, anule alguno y vuelva a pagar la planilla", , _
                            "SGCCaijo-Verificar saldos por pagar de planillas"
                End If
            Else ' Añade el saldo que falta por pagar
                grdPlanilla.AddItem curTotalFormaPago.campo(0) & _
                    vbTab & curTotalFormaPago.campo(1) & _
                    vbTab & Format(curTotalFormaPago.campo(2) - curTotalPagado.campo(0), "###,###,##0.00")
            End If
        End If
    End If
    
    ' Cierra la consulta TotalPagado
    curTotalPagado.Cerrar
    
    ' Mueve a la siguiente planilla pendiente
    curTotalFormaPago.MoverSiguiente
Loop

' Cierra la componente
curTotalFormaPago.Cerrar

End Sub

Private Sub AveriguaTipoPago()
' ------------------------------------------------------
' Propósito: Averigua el tipo de pago(Efectivo o bancario) a realizar _
             en caja-bancos por el movimineto Pago de planillas
' Recibe : Nada
' Entrega : Nada
' ------------------------------------------------------
' Verifica si la opción Caja fué elegida en el egreso sin afectación
If frmCBEGSinAfecta.optCaja.Value = True Then
    ' Si se paga por caja entonces el pago es en efectivo
    msTipoPago = "Efectivo"
Else
    ' Si se paga por banco entonces el pago es bancario
    msTipoPago = "Bancario"
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If gcolDetMovCB.Count = 0 Then
  ' Pone vacia el concepto del formulario egreso
  frmCBEGSinAfecta.txtCodMov = Empty
End If

End Sub

Private Sub grdPlanilla_Click()

If grdPlanilla.Row > 0 And grdPlanilla.Row < grdPlanilla.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdPlanilla, vbWhite, vbDarkBlue
    ' Pasa el monto de la planilla a el txtPorPagar
    txtMonto = grdPlanilla.TextMatrix(grdPlanilla.Row, 2)
    ' Habilita aceptar
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub HabilitarBotonAceptar()
' -------------------------------------------------------------
' Propósito: Habilita el botón aceptar si las condiciones son aceptables
' Recibe: Nada
' Entrega : Nada
' -------------------------------------------------------------
'Se comprueba que se haya marcado algúna fila
If ipos < 1 Or txtMonto.BackColor <> vbWhite Then
  ' No se seleccionó algúna planilla o no se ingresó el monto a pagar
  cmdAceptar.Enabled = False
  Exit Sub
End If
' Todo Ok, Habilita el botón aceptar
cmdAceptar.Enabled = True

End Sub

Private Sub grdPlanilla_DblClick()

'Hace llamado al evento click del aceptar
cmdAceptar_Click

End Sub

Private Sub grdPlanilla_EnterCell()

If ipos <> grdPlanilla.Row Then
    '  Verifica si es la última fila
    If grdPlanilla.Row > 0 And grdPlanilla.Row < grdPlanilla.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdPlanilla, ipos
            ' Muestra el monto por pagar
            txtMonto = grdPlanilla.TextMatrix(grdPlanilla.Row, 2)
            gbCambioCelda = False
            cmdAceptar.Enabled = True
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdPlanilla.Row
End If

End Sub

Private Sub grdPlanilla_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    SendKeys vbTab
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

