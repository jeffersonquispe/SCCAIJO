VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBVoucher 
   Caption         =   "Voucher de caja - bancos"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consulta del "
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   285
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   255
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "&Generar Informe"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdConsulta 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   0
      BackColorSel    =   -2147483636
      FillStyle       =   1
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   -3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -1920
      Width           =   2535
   End
End
Attribute VB_Name = "frmCBVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Cursores de ingreso
Dim mcurIngresos As New clsBDConsulta
Dim mcurDevPrestIngr As New clsBDConsulta
Dim mcurPerIngr As New clsBDConsulta
Dim mcurTercIngr As New clsBDConsulta
Dim mcurTrasladosIngr As New clsBDConsulta

'Cursores de egreso
Dim mcurEgresos As New clsBDConsulta
Dim mcurAdelantosEgr As New clsBDConsulta
Dim mcurPerEgr As New clsBDConsulta
Dim mcurProvEgr As New clsBDConsulta
Dim mcurTercEgr As New clsBDConsulta
Dim mcurPagoPLEgr As New clsBDConsulta
Dim mcurPrestamosEgr As New clsBDConsulta
Dim mcurTrasladosEgr As New clsBDConsulta

'Colección utilizadas
Dim mcolCtaBanco As New Collection
Dim mcolGastosProdServ As New Collection

Private Sub cmdSalir_Click()

'Descarga el formulario
Unload Me

End Sub

Private Sub Form_Load()

' Carga los títulos del grid
'Id, Fecha, Orden, NroDoc, Prov, CtaCte,Banco, ChequeImporte, CodPresu,Glosa
aTitulosColGrid = Array("Id", "Fecha", "Orden", "Nro Documento", "Proveedor", _
                        "Cuentan Cte", "Banco", "Cheque", "Importe", "Cod Prosupuesto", "Glosa")
aTamañosColumnas = Array(0, 1000, 1100, 1300, 3500, 1300, 2500, 3500, 3500, 1000, 2500)
CargarGridTitulos grdConsulta, aTitulosColGrid, aTamañosColumnas

End Sub

Private Sub CargarAsientos(strOrden As String)
'------------------------------------------
'Propósito  : Carga los asientos relacionados al orden
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim sSQL As String
Dim curCBAsiento As New clsBDConsulta

'Carga el titulo del asiento contable del grdConsulta
'Id, Fecha, Orden, NroDoc, Prov,  CtaCte,Banco, Cheque, Importe,CodPresu,Glosa
grdConsulta.AddItem "" & vbTab & vbTab _
                    & "Cta. Debe" & vbTab _
                    & "Cta. Haber" & vbTab _
                    & "Concepto" & vbTab _
                    & "Monto Debe" & vbTab _
                    & "Monto Haber" & vbTab & vbTab & vbTab _
                    & vbTab & vbTab
                    
'CodContable, DescCuenta,DebeHaber,Monto
sSQL = "SELECT AD.CodContable, PC.DescCuenta, AD.DebeHaber, AD.Monto " _
       & "FROM CTB_ASIENTOSCAJABANCOS C, CTB_ASIENTOS A, CTB_ASIENTOS_DET AD, PLAN_CONTABLE PC " _
       & "WHERE C.Orden='" & strOrden & "' And A.NumAsiento=C.NumAsiento And " _
       & "A.Anulado='NO' And A.NumAsiento=AD.NumAsiento And AD.CodContable=PC.CodContable " _
       & "ORDER BY AD.DebeHaber"
                   
'Copia la sentencia sSQL
curCBAsiento.SQL = sSQL

' Abre el cursor SI hay  error sale indicando la causa del error
If curCBAsiento.Abrir = HAY_ERROR Then
    End
End If

'Verifica la existencia del asiento para el orden
If curCBAsiento.EOF Then
    'Mensaje de registro de Ingreso a Caja o Bancos NO existe
    MsgBox "Error no hay Asientos para este orden", _
      vbExclamation + vbOKOnly, "Caja-Bancos- Voucher"
    curCBAsiento.Cerrar
    Exit Sub

Else
    
    'Hacer mientras no sea fin de
    Do While Not curCBAsiento.EOF
    
        'verificas si es al debe
        If curCBAsiento.campo(2) = "D" Then
                                            
            grdConsulta.AddItem "A" & vbTab & vbTab _
                    & curCBAsiento.campo(0) & vbTab _
                    & vbTab _
                    & curCBAsiento.campo(1) & vbTab _
                    & Format(curCBAsiento.campo(3), "###,###,##0.00") & vbTab _
                    & vbTab & vbTab & vbTab _
                    & vbTab & vbTab
                    

        ElseIf curCBAsiento.campo(2) = "H" Then
        
            grdConsulta.AddItem "A" & vbTab & vbTab _
                    & vbTab _
                    & curCBAsiento.campo(0) & vbTab _
                    & curCBAsiento.campo(1) & vbTab _
                    & vbTab _
                    & Format(curCBAsiento.campo(3), "###,###,##0.00") & vbTab _
                    & vbTab & vbTab _
                    & vbTab & vbTab
                    
        End If
        
        'Mueve al siguiente registro
        curCBAsiento.MoverSiguiente
    
    Loop
           
End If

'Cierra el cursor
curCBAsiento.Cerrar

End Sub

Private Sub CargarAsientoTraslados(strOrden As String, strTipoTraslado As String)
'------------------------------------------
'Propósito  : Carga los asientos relacionados al orden
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim sSQL As String
Dim curCBAsiento As New clsBDConsulta

'Carga el titulo del asiento contable del grdConsulta
'Id, Fecha, Orden, NroDoc, Prov,  CtaCte,Banco, Cheque, Importe,CodPresu,Glosa
grdConsulta.AddItem "" & vbTab & vbTab _
                    & "Cta. Debe" & vbTab _
                    & "Cta. Haber" & vbTab _
                    & "Concepto" & vbTab _
                    & "Monto Debe" & vbTab _
                    & "Monto Haber" & vbTab & vbTab & vbTab _
                    & vbTab & vbTab
         
'Verifica si es asiento de traslados de ingreso o egreso
If strTipoTraslado = "Ingreso" Then
    'CodContable, DescCuenta,DebeHaber,Monto
    sSQL = "SELECT AD.CodContable, PC.DescCuenta, AD.DebeHaber, AD.Monto " _
           & "FROM CTB_TRASLADOCAJABANCOS C, CTB_ASIENTOS A, CTB_ASIENTOS_DET AD, PLAN_CONTABLE PC " _
           & "WHERE C.OrdenIngreso='" & strOrden & "' And A.NumAsiento=C.NumAsiento And " _
           & "A.Anulado='NO' And A.NumAsiento=AD.NumAsiento And AD.CodContable=PC.CodContable " _
           & "ORDER BY AD.DebeHaber"
ElseIf strTipoTraslado = "Egreso" Then
    'CodContable, DescCuenta,DebeHaber,Monto
    sSQL = "SELECT AD.CodContable, PC.DescCuenta, AD.DebeHaber, AD.Monto " _
           & "FROM CTB_TRASLADOCAJABANCOS C, CTB_ASIENTOS A, CTB_ASIENTOS_DET AD, PLAN_CONTABLE PC " _
           & "WHERE C.OrdenEgreso='" & strOrden & "' And A.NumAsiento=C.NumAsiento And " _
           & "A.Anulado='NO' And A.NumAsiento=AD.NumAsiento And AD.CodContable=PC.CodContable " _
           & "ORDER BY AD.DebeHaber"
End If

'Copia la sentencia sSQL
curCBAsiento.SQL = sSQL

' Abre el cursor SI hay  error sale indicando la causa del error
If curCBAsiento.Abrir = HAY_ERROR Then
    End
End If

'Verifica la existencia del asiento para el orden
If curCBAsiento.EOF Then
    'Mensaje de registro de Ingreso a Caja o Bancos NO existe
    MsgBox "Error no hay Asientos para este orden", _
      vbExclamation + vbOKOnly, "Caja-Bancos- Voucher"
    curCBAsiento.Cerrar
    Exit Sub

Else
    
    'Hacer mientras no sea fin de
    Do While Not curCBAsiento.EOF
    
        'verificas si es al debe
        If curCBAsiento.campo(2) = "D" Then
                                            
            grdConsulta.AddItem "A" & vbTab & vbTab _
                    & curCBAsiento.campo(0) & vbTab _
                    & vbTab _
                    & curCBAsiento.campo(1) & vbTab _
                    & Format(curCBAsiento.campo(3), "###,###,##0.00") & vbTab _
                    & vbTab & vbTab & vbTab _
                    & vbTab & vbTab
                    

        ElseIf curCBAsiento.campo(2) = "H" Then
        
            grdConsulta.AddItem "A" & vbTab & vbTab _
                    & vbTab _
                    & curCBAsiento.campo(0) & vbTab _
                    & curCBAsiento.campo(1) & vbTab _
                    & vbTab _
                    & Format(curCBAsiento.campo(3), "###,###,##0.00") & vbTab _
                    & vbTab & vbTab _
                    & vbTab & vbTab
                    
        End If
        
        'Mueve al siguiente registro
        curCBAsiento.MoverSiguiente
    
    Loop
           
End If

'Cierra el cursor
curCBAsiento.Cerrar

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Destruye las colecciones
Set mcolCtaBanco = Nothing
Set mcolGastosProdServ = Nothing

End Sub

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  
  ' Carga consulta
  CargaCBVoucher
    
Else
  mskFechaFin.BackColor = Obligatorio
  grdConsulta.Rows = 1
  
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub mskFechaIni_Change()
' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaIni) Then
  mskFechaIni.BackColor = vbWhite
  
  ' Carga las existencias de almacén
  CargaCBVoucher
  
Else
  'Coloca a obligatorio la FechaIni
  mskFechaIni.BackColor = Obligatorio
  grdConsulta.Rows = 1
  'Habilita el cmdInforme
  cmdInforme.Enabled = False
End If

End Sub

Private Sub CargaCBVoucher()
'-----------------------------------------------------------
'Propósito  : Carga los voucher entre las fechas de consulta
'Recibe     : Nada
'Devuelve   : Nada
'-----------------------------------------------------------

'Limpia el grid
grdConsulta.Rows = 1

' Verifica los datos introducidos para la consulta
If fbOkDatosIntroducidos = False Then
  ' Sale de el proceso y limpia el grid
  grdConsulta.Rows = 1
  'Deshabilita el cmdInforme
  cmdInforme.Enabled = False
  Exit Sub
End If

'Carga la colección de ctas y banco
CargaColCtaBanco
CargaColGastos

'Carga los cursores de ingreso
CurMovPerIngr
CurMovTercIngr
CurTrasladosIngr
CurDevPrestIngr

'Cargar cursores de egreso
CurMovPerEgr
CurMovTercEgr
CurMovProvEgr
CurTrasladosEgr
CurPagoPLEgr
CurAdelantosEgr
CurPrestamosEgr

'Carga los ingresos a CajaBancos
CargaIngresos

'Carga los egresos de CajaBancos
CargaEgresos

'Carga al grid con los datos
CargaDatosGrid
    
End Sub

Private Sub CargaColCtaBanco()
'----------------------------------------------
'Propósito  : Carga la colección de ctas y bancos
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------
Dim sSQL As String
Dim curCtaBanco As New clsBDConsulta

'Sentencia SQL con cuyos datos se carga el grid
sSQL = ""
sSQL = "SELECT TC.IdCta, TC.DescCta, TB.DescBanco " _
     & "FROM TIPO_CUENTASBANC TC, TIPO_BANCOS TB " _
       & "WHERE TC.IdBanco=TB.IdBanco And TC.IdMoneda='SOL' " _
       & "ORDER BY TC.IdCta"
     
'Carga los datos de proveedores
curCtaBanco.SQL = sSQL

'Verifica si hay error
If curCtaBanco.Abrir = HAY_ERROR Then
   'Termina la ejecución
   End
End If

'Hacer mientras no sea fin del registro del cursor
Do While Not curCtaBanco.EOF

    'Agrega los datos a la colección IdCta, DescCta, DescBanco
    mcolCtaBanco.Add Item:=curCtaBanco.campo(0) & "¯" & _
                           curCtaBanco.campo(1) & "¯" & _
                           curCtaBanco.campo(2), _
                      Key:=curCtaBanco.campo(0)
                           
    'Mueve al siguiente registro
    curCtaBanco.MoverSiguiente
Loop

'Cierra el cursor
curCtaBanco.Cerrar

End Sub

Private Sub CargaColGastos()
'----------------------------------------------
'Propósito  : Carga la colección de productos
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------
Dim sSQL As String
Dim curGastos As New clsBDConsulta

'Sentencia SQL con cuyos datos se carga el grid
sSQL = ""
sSQL = "SELECT DISTINCT G.Orden, G.Concepto " _
       & "FROM GASTOS G, EGRESOS E " _
       & "WHERE E.Orden=G.Orden And E.Anulado='NO' And " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' " _
       & "ORDER BY G.Orden"
     
'Carga los datos de proveedores
curGastos.SQL = sSQL

'Verifica si hay error
If curGastos.Abrir = HAY_ERROR Then
   'Termina la ejecución
   End
End If

'Hacer mientras no sea fin del registro del cursor
Do While Not curGastos.EOF

    'Agrega los datos a la colección IdCta, DescCta, DescBanco
    mcolGastosProdServ.Add Item:=curGastos.campo(0) & "¯" & _
                                 curGastos.campo(1), _
                            Key:=curGastos.campo(0)
                           
    'Mueve al siguiente registro
    curGastos.MoverSiguiente
Loop

'Cierra el cursor
curGastos.Cerrar

End Sub


Private Sub CurDevPrestIngr()
'----------------------------------------------
'Propósito  : Carga los cursores con devolución relacionados con ingreso
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------
Dim sSQL As String

'Sentencia SQL con cuyos datos se carga el grid
sSQL = ""
sSQL = "SELECT DISTINCT DP.Orden, ( P.Apellidos & ', ' & P.Nombre) ,I.FecMov " _
     & "FROM DEVOLUCION_PRESTAMOSCB  DP, INGRESOS I, PLN_PERSONAL P " _
       & "WHERE DP.Orden=I.Orden And I.Anulado='NO' And  " _
       & "I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "DP.IdPersona=P.IdPersona " _
       & "ORDER BY I.FecMov, DP.Orden"
     
'Carga los datos de proveedores
mcurDevPrestIngr.SQL = sSQL

'Verifica si hay error
If mcurDevPrestIngr.Abrir = HAY_ERROR Then
   'Termina la ejecución
   End
End If

End Sub

Private Sub CurTrasladosIngr()
'----------------------------------------------
'Propósito  : Carga los cursores de traslados relacionados con ingreso
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------
Dim sSQL As String

'Sentencia SQL con cuyos datos se carga el grid
sSQL = ""
sSQL = "SELECT CT.OrdenIngreso,(P.Apellidos & ', ' & P.Nombre), I.FecMov " _
     & "FROM CTB_TRASLADOCAJABANCOS  CT, INGRESOS I, PLN_PERSONAL P " _
       & "WHERE CT.OrdenIngreso=I.Orden And I.Anulado='NO' And  " _
       & "I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "CT.IdPersona=P.IdPersona " _
       & "ORDER BY I.FecMov, CT.OrdenIngreso"
     
'Carga los datos de proveedores
mcurTrasladosIngr.SQL = sSQL

'Verifica si hay error
If mcurTrasladosIngr.Abrir = HAY_ERROR Then
   'Termina la ejecución
   End
End If

End Sub


Private Sub CurMovTercIngr()
'----------------------------------------------
'Propósito  : Carga los cursores de terceros relacionados con ingreso
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------

Dim sSQL As String

'Sentencia para cargar la colección
sSQL = ""
sSQL = "SELECT I.Orden,TT.DescTerc, I.FecMov " _
       & "FROM MOV_TERCEROS MT, INGRESOS I , TIPO_TERCEROS TT " _
       & "WHERE MT.Orden=I.Orden And I.Anulado='NO' And  " _
       & "I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "MT.IdTercero=TT.IdTerc " _
       & "ORDER BY I.FecMov, I.Orden"

'Carga los datos de terceros
mcurTercIngr.SQL = sSQL

'Verifica si hay error
If mcurTercIngr.Abrir = HAY_ERROR Then
   'Termina la ejecución
   End
End If

End Sub

Private Sub CurMovPerIngr()
' ----------------------------------------------
' Propósito : Carga la colección del personal de ingreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT MP.Orden,( p.Apellidos & ', ' & P.Nombre) , I.FecMov " _
       & "FROM MOV_PERSONAL MP, INGRESOS I, PLN_PERSONAL P " _
       & "WHERE MP.Orden=I.Orden And I.Anulado='NO' And  " _
       & "I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "MP.IdPersona=P.IdPersona " _
       & "ORDER BY I.FecMov, MP.Orden "

'Copia la sentencia SQL
mcurPerIngr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurPerIngr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub

Private Sub CurMovTercEgr()
' ----------------------------------------------
' Propósito : Carga la colección de terceros de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden,TT.DescTerc, E.FecMov " _
       & "FROM MOV_TERCEROS MT, EGRESOS E, TIPO_TERCEROS TT " _
       & "WHERE MT.Orden=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "MT.IdTercero=TT.IdTerc " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurTercEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurTercEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub


Private Sub CurTrasladosEgr()
' ----------------------------------------------
' Propósito : Carga la colección de traslados de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden,(P.Apellidos & ', ' & P.Nombre), E.FecMov " _
       & "FROM CTB_TRASLADOCAJABANCOS CT, EGRESOS E, PLN_PERSONAL P " _
       & "WHERE CT.OrdenEgreso=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "CT.IdPersona=P.IdPersona " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurTrasladosEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurTrasladosEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub

Private Sub CurAdelantosEgr()
' ----------------------------------------------
' Propósito : Carga el cursor con los adelantos de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden, ( P.Apellidos & ', ' & P.Nombre) ,E.FecMov " _
       & "FROM ADELANTOS A, EGRESOS E, PLN_PERSONAL P " _
       & "WHERE A.Orden=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "A.IdPersona=P.IdPersona " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurAdelantosEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurAdelantosEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub

Private Sub CurPrestamosEgr()
' ----------------------------------------------
' Propósito : Carga el cursor con los prestamos de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden,(P.Apellidos & ', ' & P.Nombre), E.FecMov " _
       & "FROM PAGO_PRESTAMOS PP, EGRESOS E, PLN_PERSONAL P " _
       & "WHERE PP.Orden=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "PP.IdPersona=P.IdPersona " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurPrestamosEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurPrestamosEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub

Private Sub CurPagoPLEgr()
' ----------------------------------------------
' Propósito : Carga el cursor de los pagos de planilla de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden,P.DescPlanilla, E.FecMov " _
       & "FROM PAGO_PLANILLAS PP, EGRESOS E, PLN_PLANILLAS P " _
       & "WHERE PP.Orden=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And  " _
       & "PP.CodPlanilla=P.CodPlanilla " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurPagoPLEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurPagoPLEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub


Private Sub CurMovProvEgr()
' ----------------------------------------------
' Propósito : Carga la colección de proveedores de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT E.Orden,P.DescProveedor,E.FecMov " _
       & "FROM PROVEEDORES P, EGRESOS E " _
       & "WHERE E.IdProveedor=P.IdProveedor And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' " _
       & "ORDER BY E.FecMov, E.Orden "

'Copia la sentencia SQL
mcurProvEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurProvEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub


Private Sub CurMovPerEgr()
' ----------------------------------------------
' Propósito : Carga la colección del personal de egreso
' Recibe    : Nada
' Entrega   : Nada
' ----------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT MP.Orden,( p.Apellidos & ', ' & P.Nombre), E.FecMov " _
       & "FROM MOV_PERSONAL MP, EGRESOS E, PLN_PERSONAL P " _
       & "WHERE MP.Orden=E.Orden And E.Anulado='NO' And  " _
       & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' And " _
       & "MP.IdPersona=P.IdPersona " _
       & "ORDER BY E.FecMov, MP.Orden "

'Copia la sentencia SQL
mcurPerEgr.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurPerEgr.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub


Private Function fbOkDatosIntroducidos() As Boolean
' ----------------------------------------------------
' Propósito: Verifica si esta bien los datos para ejecutar _
            la consulta
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------
If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
' Verifica que la fecha de inicio sea menor a la fecha final
    If CompararFechaIniFin(mskFechaIni, mskFechaFin) = True Then
        fbOkDatosIntroducidos = False
        Exit Function
    End If
End If
' Verifica si los datos obligatorios se ha llenado
If mskFechaIni.BackColor <> vbWhite Or _
   mskFechaFin.BackColor <> vbWhite Then
   fbOkDatosIntroducidos = False
   Exit Function
End If

' Devuelve la función ok
fbOkDatosIntroducidos = True

End Function

Private Sub CargaIngresos()
'------------------------------------------
'Propósito  : Carga los ingresos a caja bancos
'             entre las fechas de consulta
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = ""
sSQL = "SELECT I.Orden,I.IdCta,I.FecMov, I.NumDoc, TC.Abreviatura, " & _
       "TM.DescConCB, I.Monto " & _
       "FROM INGRESOS I, TIPO_DOCUM TC, TIPO_MOVCB TM " & _
       "WHERE I.Anulado='NO' And I.IdTipoDoc=TC.IdTipoDoc And I.CodMov=TM.IdConCB And " & _
       "I.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' " & _
       "ORDER BY I.FecMov, I.Orden"

'Copia la sentencia SQL
mcurIngresos.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurIngresos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If
  
End Sub

Private Sub CargaEgresos()
'--------------------------------------------------
'Propósito  : Carga los egresos de caja bancos
'             entre las fechas de consulta
'Recibe     : Nada
'Devuelve   : Nada
'--------------------------------------------------
Dim sSQL As String

'Sentencia SQL
sSQL = "SELECT E.Orden,E.IdProy, E.IdProg,E.IdLinea,E.IdActiv,E.IdCta, " _
        & "E.FecMov, E.NumDoc, TD.Abreviatura, E.MontoAfectado, E.MontoCB, " _
        & "E.NumCheque, TM.DescConCB " _
        & "FROM EGRESOS E, TIPO_DOCUM TD, TIPO_MOVCB TM " _
        & "WHERE E.Anulado='NO' And TD.IdTipoDoc=E.IdTipoDoc And E.CodMov=TM.IdConCB And " _
        & "E.FecMov BETWEEN '" & FechaAMD(mskFechaIni) & "' And '" & FechaAMD(mskFechaFin) & "' " _
        & "ORDER BY E.FecMov, E.Orden"
        
'Copia la sentencia SQL
mcurEgresos.SQL = sSQL

'Verifica si hay error en la sentencia
If mcurEgresos.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

End Sub

Private Sub CargaDatosGrid()
'------------------------------------------
'Propósito  : Carga el grid con los datos
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim blnRecorreCursores As Boolean
Dim blnCargaIng, blnCargaEgre  As Boolean

'Agrega los datos al grid
'Verifica si es que no sea fin de registro
blnRecorreCursores = True

'Hacer mientras el cursor tenga datos
Do While blnRecorreCursores = True

   ' Verifica si se ha terminado de recorrer todos los cursores
   If mcurIngresos.EOF And mcurEgresos.EOF Then
   
       ' Sale de recorrer cursor
       blnRecorreCursores = False
       
   Else
        'Verifica que ninguno de los cursores sea el final del registro
        If Not mcurIngresos.EOF And Not mcurEgresos.EOF Then
        
            'Verifica si el ingreso es antes de la salida
           If mcurIngresos.campo(4) <= mcurEgresos.campo(8) Then
                blnCargaIng = True
                blnCargaEgre = False
                
            'Verifica si la salida es antes del ingreso
            ElseIf mcurIngresos.campo(4) > mcurEgresos.campo(8) Then
                blnCargaIng = False
                blnCargaEgre = True
                
            End If
        
        'El mcurIngresos no es fin del registro
        ElseIf Not mcurIngresos.EOF Then
            blnCargaIng = True
            blnCargaEgre = False
            
        'El mcurSalidas no es fin del registro
        ElseIf Not mcurEgresos.EOF Then
            blnCargaEgre = True
            blnCargaIng = False
        End If
             
        ' añade una fila al grid
        If blnCargaIng Then
        
            ' Ingresos a caja bancos
            IngresosCajaBancos
                   
            ' Mueve al siguiente concepto de ingreso
            mcurIngresos.MoverSiguiente
        End If
        
        ' Agrega fila al grid
        If blnCargaEgre Then
            
           ' Egresos de caja bancos
           EgresosCajaBancos
            
           ' Mueve al siguiente concepto de ingreso
            mcurEgresos.MoverSiguiente
            
        End If
        
    End If
    
Loop ' Fin de hacer mientras sea fin de cursor

' Cierra los cursores
mcurIngresos.Cerrar
mcurEgresos.Cerrar

End Sub

Private Sub IngresosCajaBancos()
'------------------------------------------
'Propósito  : Procesa los ingresos  a caja bancos y determina
'             el proveedor para el ingreso
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
'Verifica que no sea fin del registro
Dim strDescCta As String
Dim strDescBanco As String

If Not mcurPerIngr.EOF Then
    'Verifica que los ordenes y la fechas coincidan
    If mcurIngresos.campo(0) = mcurPerIngr.campo(0) And _
       mcurIngresos.campo(2) = mcurPerIngr.campo(2) Then
       
        'Verifica si el ingreso es de Caja
        If Trim(mcurIngresos.campo(1)) = "" Then
        
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurPerIngr.campo(1) & vbTab _
                                & vbTab & vbTab & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
            

        Else
            'Verifica si el orden es de Bancos
            VerificarBanco strDescCta, strDescBanco, mcurIngresos.campo(1)
            
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurPerIngr.campo(1) & vbTab _
                                & strDescCta & vbTab & strDescBanco & vbTab _
                                & mcurIngresos.campo(5) & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
                                
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        End If
        
        'Carga los Asientos contables para el orden
        CargarAsientos mcurIngresos.campo(0)
            
        'Mueve al siguiente registro
        mcurPerIngr.MoverSiguiente
    End If
End If

'Verifica que no sea fin del registro
If Not mcurDevPrestIngr.EOF Then
    'Compara los datos del mcurIngresos y mcurDevPrestIngr
    If mcurIngresos.campo(0) = mcurDevPrestIngr.campo(0) And _
       mcurIngresos.campo(2) = mcurDevPrestIngr.campo(2) Then
       
       'Verifica si el ingreso es de Caja
        If Trim(mcurIngresos.campo(1)) = "" Then
        
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurDevPrestIngr.campo(1) & vbTab _
                                & vbTab & vbTab & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

            

        Else
            'Verifica si el orden es de Bancos
            VerificarBanco strDescCta, strDescBanco, mcurIngresos.campo(1)
            
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurDevPrestIngr.campo(1) & vbTab _
                                & strDescCta & vbTab & strDescBanco & vbTab _
                                & mcurIngresos.campo(5) & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
                            
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        End If
        
        'Carga los Asientos contables para el orden
        CargarAsientos mcurIngresos.campo(0)
        
       'Mueve al siguiente registro
       mcurDevPrestIngr.MoverSiguiente
    End If
End If

'Verifica que no sea fin del registro
If Not mcurTercIngr.EOF Then
    'Compara los datos del mcurIngresos y mcurTercIngr
    If mcurIngresos.campo(0) = mcurTercIngr.campo(0) And _
        mcurIngresos.campo(2) = mcurTercIngr.campo(2) Then
        
        'Verifica si el ingreso es de Caja
        If Trim(mcurIngresos.campo(1)) = "" Then
        
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurTercIngr.campo(1) & vbTab _
                                & vbTab & vbTab & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

            

        Else
            'Verifica si el orden es de Bancos
            VerificarBanco strDescCta, strDescBanco, mcurIngresos.campo(1)
            
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurTercIngr.campo(1) & vbTab _
                                & strDescCta & vbTab & strDescBanco & vbTab _
                                & mcurIngresos.campo(5) & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        End If
        
        'Carga los Asientos contables para el orden
        CargarAsientos mcurIngresos.campo(0)
        
        'Mueve al siguiente registro
        mcurDevPrestIngr.MoverSiguiente
    End If
    
End If

'Verifica que no sea fin del registro
If Not mcurTrasladosIngr.EOF Then
    'Compara los datos del mcurIngresos y mcurTrasladosIngr
    If mcurIngresos.campo(0) = mcurTrasladosIngr.campo(0) And _
        mcurIngresos.campo(2) = mcurTrasladosIngr.campo(2) Then

        'Verifica si el ingreso es de Caja
        If Trim(mcurIngresos.campo(1)) = "" Then
        
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurTrasladosIngr.campo(1) & vbTab _
                                & vbTab & vbTab & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        Else
            'Verifica si el orden es de Bancos
            VerificarBanco strDescCta, strDescBanco, mcurIngresos.campo(1)
            
            'Agrega datos al grdConsulta
            grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurIngresos.campo(2)) & vbTab _
                                & mcurIngresos.campo(0) & vbTab _
                                & mcurIngresos.campo(4) & "/" & mcurIngresos.campo(3) & vbTab _
                                & mcurTrasladosIngr.campo(1) & vbTab _
                                & strDescCta & vbTab & strDescBanco & vbTab _
                                & mcurIngresos.campo(5) & vbTab _
                                & MonedaLetra(Val(mcurIngresos.campo(6))) & vbTab _
                                & vbTab & mcurIngresos.campo(5) & vbTab
            ' Coloca color al grid
            grdConsulta.Row = grdConsulta.Rows - 1
            MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        End If
        
        'Carga los Asientos contables para el orden
        CargarAsientoTraslados mcurIngresos.campo(0), "Ingreso"
    
        'Mueve al siguiente registro
        mcurTrasladosIngr.MoverSiguiente
    End If
    
End If

End Sub

Private Sub VerificarBanco(strDescCta As String, strDescBanco As String, strCta As String)
'------------------------------------------
'Propósito  : Verifica si el movimiento se hizo con
'             Banco
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim strCtaBanc As String
On Error GoTo ErrClaveCol

'Accede al registro de la colección
strCtaBanc = mcolCtaBanco.Item(strCta)

'Copia los datos de la variable srtCtaBanc a las variables
strDescCta = DescomponerItem(strCtaBanc, 2)
strDescBanco = DescomponerItem(strCtaBanc, 3)

'Si hay error
ErrClaveCol:

' Error al acceder a elemento de varMiObjeto
If Err.Number = 5 Then
    'Devuelve vacio
   strDescCta = ""
   strDescBanco = ""
End If

End Sub

Private Sub EgresosCajaBancos()
'------------------------------------------
'Propósito  : Procesa los Egresos  de caja bancos y determina
'             el proveedor para el ingreso
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim strDescCta As String
Dim strDescBanco As String

'Verifica si es con afectación el egreso
If Trim(mcurEgresos.campo(1)) <> "" Then
    'Afecta al proveedor
    If Not mcurProvEgr.EOF Then
        'Verifica que los ordenes y la fechas coincidan
        If mcurEgresos.campo(0) = mcurProvEgr.campo(0) And _
           mcurEgresos.campo(6) = mcurProvEgr.campo(2) Then
                   
                'Verifica si el ingreso es de Caja
                If Trim(mcurEgresos.campo(5)) = "" Then
                
                    'Agrega datos al grdConsulta
                    grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                        & mcurEgresos.campo(0) & vbTab _
                                        & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                        & mcurProvEgr.campo(1) & vbTab _
                                        & vbTab & vbTab & vbTab _
                                        & MonedaLetra(Val(mcurEgresos.campo(9))) & vbTab _
                                        & mcurEgresos.campo(1) & mcurEgresos.campo(2) & _
                                          mcurEgresos.campo(3) & mcurEgresos.campo(4) _
                                        & vbTab & DeterminarGlosaConAfecta
                    
                    ' Coloca color al grid
                    grdConsulta.Row = grdConsulta.Rows - 1
                    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

                Else
                    'Verifica si el orden es de Bancos
                    VerificarBanco strDescCta, strDescBanco, mcurEgresos.campo(5)
                    
                    'Agrega datos al grdConsulta
                    grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                        & mcurEgresos.campo(0) & vbTab _
                                        & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                        & mcurProvEgr.campo(1) & vbTab _
                                        & strDescCta & vbTab & strDescBanco & vbTab & mcurEgresos.campo(11) & vbTab _
                                        & MonedaLetra(Val(mcurEgresos.campo(9))) & vbTab _
                                        & mcurEgresos.campo(1) & mcurEgresos.campo(2) & _
                                          mcurEgresos.campo(3) & mcurEgresos.campo(4) _
                                        & vbTab & DeterminarGlosaConAfecta
                                        
                    ' Coloca color al grid
                    grdConsulta.Row = grdConsulta.Rows - 1
                    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
                                        
                End If
                
                'Carga los Asientos contables para el orden
                CargarAsientos mcurEgresos.campo(0)
                    
                'Mueve al siguiente registro
                mcurEgresos.MoverSiguiente
                
            End If
        End If
Else 'El egreso es sin afectación
    If Not mcurPerEgr.EOF Then
        'Verifica que los ordenes y la fechas coincidan
        If mcurEgresos.campo(0) = mcurPerEgr.campo(0) And _
           mcurEgresos.campo(6) = mcurPerEgr.campo(2) Then
                   
                'Verifica si el ingreso es de Caja
                If Trim(mcurEgresos.campo(5)) = "" Then
                
                    'Agrega datos al grdConsulta
                    grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                        & mcurEgresos.campo(0) & vbTab _
                                        & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                        & mcurPerEgr.campo(1) & vbTab _
                                        & vbTab & vbTab & vbTab _
                                        & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                        & vbTab & mcurEgresos.campo(12)
                                        
                    ' Coloca color al grid
                    grdConsulta.Row = grdConsulta.Rows - 1
                    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

        
                Else
                    'Verifica si el orden es de Bancos
                    VerificarBanco strDescCta, strDescBanco, mcurEgresos.campo(5)
                    
                    'Agrega datos al grdConsulta
                    grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                        & mcurEgresos.campo(0) & vbTab _
                                        & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                        & mcurPerEgr.campo(1) & vbTab _
                                        & strDescCta & vbTab & strDescBanco & vbTab & mcurEgresos.campo(11) & vbTab _
                                        & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                        & vbTab & mcurEgresos.campo(12)
                    ' Coloca color al grid
                    grdConsulta.Row = grdConsulta.Rows - 1
                    MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray
                  
                End If
                
                'Carga los Asientos contables para el orden
                CargarAsientos mcurEgresos.campo(0)
                    
                'Mueve al siguiente registro
                mcurPerEgr.MoverSiguiente
            End If
        End If
        
        If Not mcurTercEgr.EOF Then
            'Verifica que los ordenes y la fechas coincidan
            If mcurEgresos.campo(0) = mcurTercEgr.campo(0) And _
               mcurEgresos.campo(6) = mcurTercEgr.campo(2) Then
                       
                    'Verifica si el ingreso es de Caja
                    If Trim(mcurEgresos.campo(5)) = "" Then
                    
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurTercEgr.campo(1) & vbTab _
                                            & vbTab & vbTab & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

            
                    Else
                        'Verifica si el orden es de Bancos
                        VerificarBanco strDescCta, strDescBanco, mcurEgresos.campo(5)
                        
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurTercEgr.campo(1) & vbTab _
                                            & strDescCta & vbTab & strDescBanco & vbTab & mcurEgresos.campo(11) & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

                                            
                    End If
                    
                    'Carga los Asientos contables para el orden
                    CargarAsientos mcurEgresos.campo(0)
                        
                    'Mueve al siguiente registro
                    mcurTercEgr.MoverSiguiente
            End If
        End If
        
        ' Mientra no sea fin del cursor mcurTraslados
        If Not mcurTrasladosEgr.EOF Then
            'Verifica que los ordenes y la fechas coincidan
            If mcurEgresos.campo(0) = mcurTrasladosEgr.campo(0) And _
               mcurEgresos.campo(6) = mcurTrasladosEgr.campo(2) Then
                       
                    'Verifica si el ingreso es de Caja
                    If Trim(mcurEgresos.campo(5)) = "" Then
                    
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurTrasladosEgr.campo(1) & vbTab _
                                            & vbTab & vbTab & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                                            
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

            
                    Else
                        'Verifica si el orden es de Bancos
                        VerificarBanco strDescCta, strDescBanco, mcurEgresos.campo(5)
                        
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurTrasladosEgr.campo(1) & vbTab _
                                            & strDescCta & vbTab & strDescBanco & vbTab & mcurEgresos.campo(11) & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                                            
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

                                            
                    End If
                    
                    'Carga los Asientos contables para el orden
                    CargarAsientoTraslados mcurEgresos.campo(0), "Egreso"
                        
                    'Mueve al siguiente registro
                    mcurTrasladosEgr.MoverSiguiente
            End If
        End If
        
        'Verifica que no sea fin del cursor curPagoPL
        If Not mcurPagoPLEgr.EOF Then
            'Verifica que los ordenes y la fechas coincidan
            If mcurEgresos.campo(0) = mcurPagoPLEgr.campo(0) And _
               mcurEgresos.campo(6) = mcurPagoPLEgr.campo(2) Then
                       
                    'Verifica si el ingreso es de Caja
                    If Trim(mcurEgresos.campo(5)) = "" Then
                    
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurPagoPLEgr.campo(1) & vbTab _
                                            & vbTab & vbTab & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

            
                    Else
                        'Verifica si el orden es de Bancos
                        VerificarBanco strDescCta, strDescBanco, mcurEgresos.campo(5)
                        
                        'Agrega datos al grdConsulta
                        grdConsulta.AddItem "M" & vbTab & FechaDMA(mcurEgresos.campo(6)) & vbTab _
                                            & mcurEgresos.campo(0) & vbTab _
                                            & mcurEgresos.campo(8) & "/" & mcurEgresos.campo(7) & vbTab _
                                            & mcurPagoPLEgr.campo(1) & vbTab _
                                            & strDescCta & vbTab & strDescBanco & vbTab & mcurEgresos.campo(11) & vbTab _
                                            & MonedaLetra(Val(mcurEgresos.campo(10))) & vbTab _
                                            & vbTab & mcurEgresos.campo(12)
                        ' Coloca color al grid
                        grdConsulta.Row = grdConsulta.Rows - 1
                        MarcarFilaGRID grdConsulta, vbWhite, vbDarkGray

                                            
                    End If
                    
                    'Carga los Asientos contables para el orden
                    CargarAsientos mcurEgresos.campo(0)
                        
                    'Mueve al siguiente registro
                    mcurTercEgr.MoverSiguiente
            End If
        End If
End If

End Sub

Function DeterminarGlosaConAfecta() As String
'------------------------------------------
'Propósito  : Determina la glosa para el egreso
'             con afectación, Compras o servicios
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------
Dim strProdServ As String

'Accede al registro de la colección
strProdServ = DescomponerItem(mcolGastosProdServ.Item(mcurEgresos.campo(0)), 2)

'Verifica si es producto
If strProdServ = "P" Then
    'Producto
    DeterminarGlosaConAfecta = "POR LAS COMPRAS REALIZADAS"
ElseIf strProdServ = "S" Then
    'Servicio
    DeterminarGlosaConAfecta = "POR LOS SERVICIOS PRESTADOS"
End If

End Function

