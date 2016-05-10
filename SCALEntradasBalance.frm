VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmALEntrBalance 
   Caption         =   "SGCcajo - Almacén Balance"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   HelpContextID   =   86
   Icon            =   "SCALEntradasBalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "A&nular"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   5880
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   1150
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   10155
      Begin VB.CommandButton cmdBuscar 
         Height          =   300
         Left            =   3050
         Picture         =   "SCALEntradasBalance.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   250
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1440
      End
      Begin VB.TextBox txtNumDoc 
         Height          =   315
         Left            =   1540
         MaxLength       =   50
         TabIndex        =   2
         Top             =   705
         Width           =   2235
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   8340
         TabIndex        =   3
         TabStop         =   0   'False
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   7440
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Có&digo :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label15 
         Caption         =   "Número Doc:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   690
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4670
      Left            =   120
      TabIndex        =   17
      Top             =   1150
      Width           =   10155
      Begin VB.CommandButton cmdPProdServ 
         Height          =   255
         Left            =   5720
         Picture         =   "SCALEntradasBalance.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   230
         Width           =   220
      End
      Begin VB.ComboBox cboProdServ 
         Height          =   315
         ItemData        =   "SCALEntradasBalance.frx":0CA4
         Left            =   1035
         List            =   "SCALEntradasBalance.frx":0CA6
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   195
         Width           =   4935
      End
      Begin VB.TextBox txtMedida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txtCant 
         Height          =   315
         Left            =   1035
         MaxLength       =   8
         TabIndex        =   7
         Top             =   585
         Width           =   855
      End
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPrecioUni 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   8955
         TabIndex        =   16
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "A&ñadir"
         Height          =   375
         Left            =   7800
         TabIndex        =   10
         Top             =   600
         Width           =   1000
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetalle 
         Height          =   3500
         Left            =   960
         TabIndex        =   15
         Top             =   1080
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         HighLight       =   0
         FillStyle       =   1
      End
      Begin VB.Label lblProdServ 
         Caption         =   "&Producto:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   195
         Width           =   735
      End
      Begin VB.Label lblCant 
         Caption         =   "&Cantidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   585
         Width           =   855
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "&Costo Total:"
         Height          =   195
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblPrecioUniCompra 
         AutoSize        =   -1  'True
         Caption         =   "Precio u&nit. :"
         Height          =   195
         Left            =   5280
         TabIndex        =   18
         Top             =   600
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmALEntrBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Colecciones para la modificación de el formulario
Private mcolIngreALDet As New Collection

' Colecciones para el manejo del formulario
Private mcolidprod As New Collection
Private mcolCodDesProd As New Collection
Private mcolDesMedidaContProd As New Collection

' Variable para manejo formulario
Private msIdprod As String
Private msCodSuministro As String
Private msCodVariacion As String
Private msCodigo As String
Private mbALIngresoCargado As Boolean

' Cursor para el manejo de formulario
Private mcurIngreAL As New clsBD2
Private mcurIngreALDet As New clsBD2

' Variable para el manejo del grid
Dim ipos As Long

Private Sub cboProdServ_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cmdAceptar_Click()

If gsTipoOperacionAlmacen = "Nuevo" Then
 ' Pregunta aceptación de los datos
   If MsgBox("¿Está conforme con los datos?", _
      vbQuestion + vbYesNo, "Caja-Bancos, Egreso con afectación") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario
    ' Se guardan los datos del egreso
     GuardarAlmacen
   Else: Exit Sub ' sale
   End If
Else
 ' Controla si algún producto ha sido verificado en almacén
 If fbOkProductosAlmacen("Modificar") = False Then
    ' algún dato incorrecto
    Exit Sub
 End If

 ' Mensaje de conformidad de los datos
   If MsgBox("¿Está conforme con las modificaciones realizadas ? ", _
             vbQuestion + vbYesNo, "SGCcaijo- Modificación de ingreso a almacén balance") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario
    ' Se Modifican los datos del egreso
     GuardarModificacionesAlmacen
   Else: Exit Sub ' sale
   End If
End If

'Actualiza la transaccion
Var8 -1, Empty

' Mensaje Ok
MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso con Afectación"

' Limpia la pantalla para una nueva operación, Prepara el formulario
LimpiarFormulario
   
If gsTipoOperacionAlmacen = "Nuevo" Then
 ' Nuevo egreso
   NuevoIngresoAL
   
Else
 ' cierra el control Ingreso a almacén
   If mbALIngresoCargado Then
    mcurIngreAL.Cerrar
    mbALIngresoCargado = False
    mskFecTrab = "__/__/____"
   End If

 ' Se modifican los datos del ingreso a almacén
   ModificarIngresoAL
End If


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

End Sub


Private Sub LimpiarSegundaParteFormulario()
' -------------------------------------------------------------------
' Propósito: limpia los controles de la segunda parte del formulario
' -------------------------------------------------------------------

' Limpia los controles de ingreso al detalle
  LimpiaControlesIngreDet


' Limpia el grdDetalle
  grdDetalle.Rows = 1
  
End Sub

Private Sub LimpiaControlesIngreDet()
' -------------------------------------------------------------------
' Propósito: Limpia los controles que permiten el ingreso al grdDetalle del Doc
' -------------------------------------------------------------------

' Limpia el combo ProdServ
cboProdServ.ListIndex = -1
cboProdServ.BackColor = Obligatorio

' Limpia los controles txt
txtValor = Empty
txtCant.Text = Empty
txtMedida = Empty

End Sub

Private Sub LimpiarPrimeraParteFormulario()
' -------------------------------------------------------------------
' Propósito: limpia los controles de la segunda parte del formulario
' -------------------------------------------------------------------

' Limpia los controles generales del formulario
txtCodigo = Empty
txtNumDoc = Empty

End Sub

Private Sub GuardarAlmacen()
'---------------------------------------------------------
' Proposito :Guarda los datos de ingreso de productos a almacén balance
' Recibe: Nada
' Entrega : Nada
'---------------------------------------------------------
msCodigo = CalcularCodigoAL

'Guarda el registro general del EgresoCA en Egreso
GrabarALBalanceGeneral
               
'Guarda el detalle de moviento en Gastos
GrabarDetAlGastos
               
'Guardar Ingreso a almacén
IngresarProdsBalance

End Sub

Private Sub GuardarModificacionesAlmacen()
' ----------------------------------------------
' Propósito : Guarda las Modificaciones del ingreso almacén balance
' Recibe : Nada
' Entrega : Nada
' ----------------------------------------------
'Guarda el registro general del EgresoCA en Egreso
ModificarAlmacenGeneral
               
'Guarda el detalle de moviento en Gastos
ModificarDetalleEnGastos
               
End Sub

Private Function fsAsignarNumIngreso() As String
'---------------------------------------------------------
' Propósito: obtiene el último numero de ingreso a almacén _
             y retorna el siguiente
' Recibe: Nada
' Entrega: Nada
'---------------------------------------------------------
Dim sCodigo As String
Dim sSQL As String
Dim curNumeroIngreso As New clsBD2
Dim iNumSec As Long

' Concatenamos el codigo AñoMes
  sCodigo = Right(mskFecTrab, 4)
  
' Carga la sentencia
  sSQL = "SELECT Max(NroIngreso)  FROM ALMACEN_INGRESOS WHERE  NroIngreso LIKE '" & sCodigo & "*'"

' Ejecuta la sentencia
  curNumeroIngreso.SQL = sSQL
' Averigua el último orden de ingreso
  If curNumeroIngreso.Abrir = HAY_ERROR Then
     End
  End If
  
' Separa los cuatro últimos caracteres del maximo numero de ingreso
 If IsNull(curNumeroIngreso.campo(0)) Then
  fsAsignarNumIngreso = (sCodigo & "0000001")
 Else
  iNumSec = Val(Right(curNumeroIngreso.campo(0), 5))
  fsAsignarNumIngreso = sCodigo & Format(CStr(iNumSec) + 1, "000000#")
 End If

' Cierra el cursor
  curNumeroIngreso.Cerrar

End Function

Private Sub ModificarDetalleEnGastos()
'------------------------------------------------------------
' Propósito:  Modifica los registros detalle en las tablas _
              gastos
' Recibe : Nada
' Entrega: Nada
'------------------------------------------------------------
Dim sSQL, sRegDetalle, sProdServ As String
Dim modDetIngresoBalance As New clsBD3
Dim i As Integer
Dim sPrecioUnit, sResto As String
On Error GoTo ErrClaveCol

'recorre el grid
For i = 1 To grdDetalle.Rows - 1
    'carga el reg detalle para compararlo con el registro cargado en la colección
    '"idproducto", "Cantidad", "Monto",
    sRegDetalle = grdDetalle.TextMatrix(i, 4) & "¯" & grdDetalle.TextMatrix(i, 1) _
        & "¯" & grdDetalle.TextMatrix(i, 3)
    'verifica SI el registro se encuentra en la colección y SI se modificó
    ' Si no se encuentra se inserta en ErrClaveCol: error 5
    If mcolIngreALDet.Item(grdDetalle.TextMatrix(i, 4)) <> sRegDetalle Then
       
       ' se modifico el registro, entonces se actualiza la BD
        sSQL = "UPDATE GASTOS SET " & _
        "Cantidad=" & grdDetalle.TextMatrix(i, 1) & "," & _
        "Monto=" & Var37(grdDetalle.TextMatrix(i, 3)) & _
        " WHERE Orden='" & msCodigo & "' and " _
             & "CodConcepto='" & grdDetalle.TextMatrix(i, 4) & "'"
        
        modDetIngresoBalance.SQL = sSQL
        'ejecuta la sentencia que modifica el registro en gastos
        If modDetIngresoBalance.Ejecutar = HAY_ERROR Then End
        modDetIngresoBalance.Cerrar
        ' Calcula el Precio unitario y el resto
        sPrecioUnit = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) / Val(Var37(grdDetalle.TextMatrix(i, 1))), "#0.00")
        sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.00")
        If Val(sResto) < 0 Then
           sPrecioUnit = Format(Val(sPrecioUnit) - 0.01, "#0.00")
           sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.00")
        End If
        
       ' Se actualiza los datos de almacén ingresos
        sSQL = "UPDATE ALMACEN_INGRESOS SET CantidadDisponible=" & Var37(grdDetalle.TextMatrix(i, 1)) _
            & ", PrecioUnit=" & sPrecioUnit & ", Resto=" & sResto _
            & " WHERE Orden='" & txtCodigo & "' and IdProd='" & grdDetalle.TextMatrix(i, 4) & "'"
        
        modDetIngresoBalance.SQL = sSQL
        'ejecuta la sentencia que modifica el registro en gastos
        If modDetIngresoBalance.Ejecutar = HAY_ERROR Then End
        modDetIngresoBalance.Cerrar

        ' Modifica la contabilidad
         ' Carga la colección asiento IdProd, Tratamiento, CodSuministro, CodVariación, Monto
            gcolAsientoDet.Add _
                Key:=grdDetalle.TextMatrix(i, 4), _
                Item:=grdDetalle.TextMatrix(i, 4) & "¯" _
                & "IA" & "¯" & grdDetalle.TextMatrix(i, 6) & "¯" _
                & grdDetalle.TextMatrix(i, 7) & "¯" & Var37(grdDetalle.TextMatrix(i, 3))
            
            'Colección que guarda los datos generales para los asientos contable
            'Orden, Fecha, Producto, Glosa, Proceso
            gcolAsiento.Add _
                Key:=txtCodigo, _
                Item:=txtCodigo & "¯" & FechaAMD(mskFecTrab.Text) & "¯" _
                & grdDetalle.TextMatrix(i, 4) & "¯INGRESO A ALMACEN¯" & "Ingreso"
            
            'Asiento de ingreso de almacen
            gsTipoAlmacen = "Ingreso"
            
            'Realiza el asiento automatico de ingreso a almacén
            Conta47
        
            'eliminar el elmento modificado de la colección
            mcolIngreALDet.Remove (grdDetalle.TextMatrix(i, 4))
    
    Else 'registro no se modificó
       'Solo se elimina de la colección para seguir con los demas registros
       mcolIngreALDet.Remove (grdDetalle.TextMatrix(i, 4))
    End If

PostErrClaveCol:
  
Next i

'eliminar los que quedan en la colección
 ElimnarRegsDetEliminados
'-------------------------------------------------------------------
ErrClaveCol:

    If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
        'el registro NO existe en el egreso con afectacion original
        'carga la sentencia que inserta el registro detalle en la base de datos
        sSQL = "INSERT INTO GASTOS VALUES('" & txtCodigo & "','P','" _
            & grdDetalle.TextMatrix(i, 4) & "'," & grdDetalle.TextMatrix(i, 1) & "," _
            & Var37(grdDetalle.TextMatrix(i, 3)) & ")"
        modDetIngresoBalance.SQL = sSQL
        'ejecuta la sentencia que añade registro  a Gastos
        If modDetIngresoBalance.Ejecutar = HAY_ERROR Then End
        modDetIngresoBalance.Cerrar
        
        ' Ingresa el producto a Ingresos almacén
                
                
    ' "Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "VarExistencias"
    ' Calcula PrecioUnit, Resto
    sPrecioUnit = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) / Val(Var37(grdDetalle.TextMatrix(i, 1))), "#0.00")
    sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.00")
    ' Verifica si el resto es negativo
    If Val(sResto) < 0 Then
      sPrecioUnit = Format(Val(sPrecioUnit) - 0.01, "#0.00")
      sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.00")
    End If
   ' Se ingresa el producto a ALMACEN_INGRESOS, carga la sentencia
   ' Orden, Prod,NumIngreso,PrecioUnit,Resto,CantDisponible,Fecha
     sSQL = "INSERT INTO ALMACEN_INGRESOS VALUES('" _
            & txtCodigo & "','" & grdDetalle.TextMatrix(i, 4) & "','" _
            & fsAsignarNumIngreso & "'," & sPrecioUnit & "," _
            & sResto & "," & grdDetalle.TextMatrix(i, 1) & ",'" _
            & FechaAMD(mskFecTrab) & "')"
            
            ' Ejecuta la sentencia
            modDetIngresoBalance.SQL = sSQL
            If modDetIngresoBalance.Ejecutar = HAY_ERROR Then End
            
            'Cierra la instancia de modificacion de la base de datos
            modDetIngresoBalance.Cerrar
            
            ' Carga la colección asiento IdProd, Tratamiento, CodSuministro, CodVariación, Monto
            ' "Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "VarExistencias"
            gcolAsientoDet.Add _
                Key:=grdDetalle.TextMatrix(i, 4), _
                Item:=grdDetalle.TextMatrix(i, 4) & "¯" _
                & "IA" & "¯" & grdDetalle.TextMatrix(i, 6) & "¯" _
                & grdDetalle.TextMatrix(i, 7) & "¯" & Var37(grdDetalle.TextMatrix(i, 3))
            
            'Colección que guarda los datos generales para los asientos contable
            'Orden, Fecha, Producto, Glosa, Proceso
            gcolAsiento.Add _
                Key:=txtCodigo, _
                Item:=txtCodigo & "¯" & FechaAMD(mskFecTrab.Text) & "¯" _
                & grdDetalle.TextMatrix(i, 4) & "¯INGRESO A ALMACEN¯" & "Ingreso"
            
            'Asiento de ingreso de almacen
            gsTipoAlmacen = "Ingreso"
            
            'Realiza el asiento automatico de ingreso a almacén
            Conta47
                
                
        Resume PostErrClaveCol ' La ejecución sigue por aquí
    End If
    

End Sub

Private Sub ElimnarRegsDetEliminados()
'----------------------------------------------------------------
'Propósito : Elimina los registros que han sido eliminados del _
            detalle en la modificación del ingreso almacén balance
'Recibe : Nada
'Devuelve : Nada
'----------------------------------------------------------------
Dim sSQL As String
Dim modRegDetEgreso As New clsBD3
Dim MiObjeto As Variant ' Variables de información.
    
For Each MiObjeto In mcolIngreALDet  ' Recorre los elementos que quedan en la colección
    'Elimina los registros de gastos
    sSQL = "DELETE * FROM GASTOS " _
         & "WHERE Orden ='" & txtCodigo & "'" _
         & " and CodConcepto='" & Var30(MiObjeto, 1) & "'"
    modRegDetEgreso.SQL = sSQL
    ' Ejecuta la sentencia que elimina los registros eliminados del egreso
    If modRegDetEgreso.Ejecutar = HAY_ERROR Then End
    modRegDetEgreso.Cerrar
    
    ' Se borra de almacén ingresos
    sSQL = "DELETE * FROM ALMACEN_INGRESOS WHERE  " _
      & "Orden='" & txtCodigo & "' and IdProd='" & Var30(MiObjeto, 1) & "'"
    modRegDetEgreso.SQL = sSQL
    ' Ejecuta la sentencia que elimina los registros eliminados del egreso
    If modRegDetEgreso.Ejecutar = HAY_ERROR Then End
    modRegDetEgreso.Cerrar
    
    ' Orden , IdProducto
    gcolAsiento.Add Key:=txtCodigo, _
              Item:=txtCodigo & "¯" & Var30(MiObjeto, 1)
    ' Elimina los asientos de los productos eliminados
    Conta44
    
Next MiObjeto

End Sub

Public Sub GrabarDetAlGastos()
'--------------------------------------------------------------------
'Propósito  :Grabar datos del detalle en la tabla 'Gastos'
'Recibe     :Nada
'Devuelve   :Nada
'--------------------------------------------------------------------
Dim i As Integer
Dim sSQL, sProdServ  As String
Dim modGastos As New clsBD3

' Recorre el grid y lo almacena en la BD
For i = 1 To grdDetalle.Rows - 1
    ' Producto,Cantidad,Precio Unitario,Total,CodConcepto,Medida,Suminstro,Variación)
    ' Inserta el detalle en Gastos Orden , Concepto, CodConcepto, Cantidad, Monto
    sSQL = "INSERT INTO GASTOS VALUES('" & msCodigo & "','P','" _
    & grdDetalle.TextMatrix(i, 4) & "'," & Var37(grdDetalle.TextMatrix(i, 1)) & "," _
    & Var37(grdDetalle.TextMatrix(i, 3)) & ")"
    
    ' Ejecuta la sentencia
    modGastos.SQL = sSQL
    If modGastos.Ejecutar = HAY_ERROR Then
      End
    End If
        
    'Se cierra la query
    modGastos.Cerrar
    
Next i

End Sub

Private Sub ModificarAlmacenGeneral()
' -----------------------------------------------
'Propósito: Modificar el ingreso de balance general en la bd
'Recibe : Nada
'Enatrega : Nada
' -----------------------------------------------
Dim sSQL As String
Dim modAlmacen As New clsBD3

' Carga la sentencia que modifica el ingreso
sSQL = "UPDATE ALMACEN_BALANCE SET " & _
   "NumDoc='" & Trim(txtNumDoc) & "' " & _
   "WHERE IdBalance='" & txtCodigo & "'"

' Ejecuta la sentencia que modifica el almacén
modAlmacen.SQL = sSQL
If modAlmacen.Ejecutar = HAY_ERROR Then End

'Cierra la componente
modAlmacen.Cerrar

End Sub


Private Sub GrabarALBalanceGeneral()
'-----------------------------------------------------------------
' Propósito : Guarda los datos generales del Ingreso Balance
' Recibe : Nada
' Devuelve : Nada
'-----------------------------------------------------------------
' Nota: llamado de el click de botón AceptarModificar
Dim modALBalance As New clsBD3
Dim sSQL As String

'Carga la sentencia que inserta un registro
sSQL = "INSERT INTO ALMACEN_BALANCE VALUES('" & msCodigo & "','" _
        & Trim(txtNumDoc.Text) & "','" & FechaAMD(mskFecTrab.Text) & "','NO')"
  
' Ejecuta la sentencia que inserta un registro a Egresos
modALBalance.SQL = sSQL
If modALBalance.Ejecutar = HAY_ERROR Then
 ' cierra error
  End
End If

'Se cierra la query
modALBalance.Cerrar

End Sub

Private Sub IngresarProdsBalance()
'------------------------------------------------------
'Propósito: Ingresa a almacén los productos verificados
'Recibe: Nada
'Devuelve: Nada
'------------------------------------------------------
Dim sSQL As String
Dim modAlmacen As New clsBD3
Dim i As Integer
Dim sPrecioUnit As String
Dim sResto As String
' Recorre el grd
For i = 1 To grdDetalle.Rows - 1
    ' "Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "VarExistencias"
    ' Calcula PrecioUnit, Resto
    sPrecioUnit = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) / Val(Var37(grdDetalle.TextMatrix(i, 1))), "#0.00")
    sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.0000")
    ' Verifica si el resto es negativo
    If Val(sResto) < 0 Then
      sPrecioUnit = Format(Val(sPrecioUnit) - 0.01, "#0.00")
      sResto = Format(Val(Var37(grdDetalle.TextMatrix(i, 3))) - (Val(sPrecioUnit) * Val(Var37(grdDetalle.TextMatrix(i, 1)))), "#0.0000")
    End If
   ' Se ingresa el producto a ALMACEN_INGRESOS, carga la sentencia
   ' Orden, Prod,NumIngreso,PrecioUnit,Resto,CantDisponible,Fecha
     sSQL = "INSERT INTO ALMACEN_INGRESOS VALUES('" _
            & msCodigo & "','" & grdDetalle.TextMatrix(i, 4) & "','" _
            & fsAsignarNumIngreso & "'," & sPrecioUnit & "," _
            & sResto & "," & grdDetalle.TextMatrix(i, 1) & ",'" _
            & FechaAMD(mskFecTrab) & "')"
            
            ' Ejecuta la sentencia
            modAlmacen.SQL = sSQL
            If modAlmacen.Ejecutar = HAY_ERROR Then End
            
            'Cierra la instancia de modificacion de la base de datos
            modAlmacen.Cerrar
            
            ' Carga la colección asiento IdProd, Tratamiento, CodSuministro, CodVariación, Monto
            ' "Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "VarExistencias"
            gcolAsientoDet.Add _
                Key:=grdDetalle.TextMatrix(i, 4), _
                Item:=grdDetalle.TextMatrix(i, 4) & "¯" _
                & "IA" & "¯" & grdDetalle.TextMatrix(i, 6) & "¯" _
                & grdDetalle.TextMatrix(i, 7) & "¯" & Var37(grdDetalle.TextMatrix(i, 3))
            
            'Colección que guarda los datos generales para los asientos contable
            'Orden, Fecha, Producto, Glosa, Proceso
            gcolAsiento.Add _
                Key:=msCodigo, _
                Item:=msCodigo & "¯" & FechaAMD(mskFecTrab.Text) & "¯" _
                & grdDetalle.TextMatrix(i, 4) & "¯INGRESO A ALMACEN¯" & "Ingreso"
            
            'Asiento de ingreso de almacen
            gsTipoAlmacen = "Ingreso"
            
            'Realiza el asiento automatico de ingreso a almacén
            Conta47
    
Next i ' Siguiente fila del grid detalle

End Sub

Private Function fbOkProductosAlmacen(sOperacion As String) As Boolean
'---------------------------------------------------------
' Proposito : Verifica en almacén si los productos ingresados _
              ya fueron sacados de almacén
' Recibe: Nada
' Entrega : Nada
'---------------------------------------------------------
Dim curProdsSalidos As New clsBD2
Dim sSQL As String
Dim sProd As Variant
Dim colProdSalidos As New Collection

' Averigua los productos ingresados y que se les han dado salida
sSQL = "SELECT AI.IdProd FROM ALMACEN_INGRESOS AI, GASTOS G " _
     & "WHERE G.Orden='" & msCodigo & "' and G.Orden=AI.Orden " _
     & "and G.CodConcepto=AI.IdProd and G.Cantidad>AI.CantidadDisponible"

' Ejecuta la sentencia
curProdsSalidos.SQL = sSQL
If curProdsSalidos.Abrir = HAY_ERROR Then End

' verifica si tiene algún producto que se le ha dado salida en almacén
If curProdsSalidos.EOF Then
    ' No hay productos salidos de este ingreso
Else
    ' Existen productos ingresados que se les ha dado salida en almacén
    ' Carga la colección de los productos que se les ha dado salida
    Do While Not curProdsSalidos.EOF
        '
        colProdSalidos.Add Item:=curProdsSalidos.campo(0), _
                              Key:=curProdsSalidos.campo(0)
        ' Mueve al siguiente registro
        curProdsSalidos.MoverSiguiente
    Loop
    curProdsSalidos.Cerrar
    
    ' Si hay productos salidos en almacén , no se puede anular
    If colProdSalidos.Count > 0 And sOperacion = "Anular" Then
        MsgBox "No se puede Anular por que algunos productos " & Chr(13) _
             & "han salido de almacén. Consulte al administrador", , "SGCcaijo-Ingreso a almacén por balance"
        ' Devuelve el resultado de la verificación
        fbOkProductosAlmacen = False
        Set colProdSalidos = Nothing
        Exit Function
    End If
    
End If

' Verifica si se cambio los datos de los productos verificados _
  en almacén
  If sOperacion = "Modificar" Then
    For Each sProd In colProdSalidos
      If mcolIngreALDet.Item(sProd) <> fsEstaenDetalle(sProd) Then
          MsgBox "No se puede guardar los cambios por que el producto : " _
                  & Chr(13) & mcolCodDesProd(sProd) _
                  & Chr(13) & "Se ha dado salida en almacén. Consulte al administrador", , "SGCcaijo-Ingreso por balance almacén "
          ' Devuelve el resultado de la verificación
          fbOkProductosAlmacen = False
          Set colProdSalidos = Nothing
          Exit Function
      End If
    Next sProd
  End If

' Vacía la colección de productos ingresados a almacén
Set colProdSalidos = Nothing

' Inicializa la función asumiendo que no se ha ingresado productos
fbOkProductosAlmacen = True

End Function


Private Sub cmdAnular_Click()
Dim modAnularALIngre As New clsBD3
Dim sSQL As String

'Verificar en almacén, si se puede anular el ingreso a almacén por balance
If fbOkProductosAlmacen("Anular") = False Then Exit Sub

'Preguntar si desea Anular el registro de Ingreso a Banco
'Mensaje de conformidad de los datos
If MsgBox("¿Seguro que desea anular el ingreso a almacén " & msCodigo & "?", _
              vbQuestion + vbYesNo, "Balance de almacén") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario
    
    'Cambiar el campo Anulado de Ingresos a "SI"
     sSQL = "UPDATE ALMACEN_BALANCE SET " & _
        "Anulado='SI'" & _
        "WHERE IdBalance='" & msCodigo & "'"
    
    'SI al ejecutar hay error se sale de la aplicación
     modAnularALIngre.SQL = sSQL
     If modAnularALIngre.Ejecutar = HAY_ERROR Then
      End
     End If
    'Se cierra la query
    modAnularALIngre.Cerrar
     
    'Elimina los registros del detalle que han ingresado a almacén
    ElimnarRegsDetEliminados
    
    'Actualiza la transaccion
    Var8 -1, Empty
    
    ' Mensaje Ok
    MsgBox "Operación efectuada correctamente", , "SGCCaijo-Egreso con Afectación"
    
    ' Limpia la pantalla para una nueva operación, Prepara el formulario
    LimpiarFormulario
       
    If gsTipoOperacionEgreso = "Nuevo" Then
     ' Nuevo ingreso a almacén
       NuevoIngresoAL
    Else
     ' Cierra el control Ingreso a almacén
       If mbALIngresoCargado Then
        mcurIngreAL.Cerrar
        mbALIngresoCargado = False
        mskFecTrab = "__/__/____"
       End If
    
     ' Se modifican los datos del ingreso a almacén
       ModificarIngresoAL
    End If
End If
End Sub

Private Sub cmdAñadir_Click()
'"Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "Variación"

If fsEstaenDetalle(msIdprod) = Empty Then
    '  Añade un elemento al grid
    grdDetalle.AddItem cboProdServ.Text & vbTab & txtCant & vbTab & _
                       txtPrecioUni & vbTab & txtValor & vbTab & _
                       msIdprod & vbTab & txtMedida & vbTab & msCodSuministro & _
                       vbTab & msCodVariacion
    
    ' Vaciar los controles de ingreso del detalle
    LimpiaControlesIngreDet
    
    ' Da el control al cboProdServ
    cboProdServ.SetFocus
    
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

Private Sub cmdBuscar_Click()

' Muestra el formulario para la seleción de el ingreso a modificar
frmALSelEntradaBalance.Show vbModal, Me

End Sub

Private Sub cmdCancelar_Click()

' Limpia el formulario y pone en blanco variables
LimpiarFormulario

' Verifica el tipo operación
If gsTipoOperacionAlmacen = "Nuevo" Then
  ' Prepara el formulario
  NuevoIngresoAL
Else
  ' cierra el control egreso
   If mbALIngresoCargado Then
    mcurIngreAL.Cerrar
    mbALIngresoCargado = False
    mskFecTrab = "__/__/____"
   End If
  ' Prepara el formulario
  ModificarIngresoAL
End If

End Sub

Private Sub cmdEliminar_Click()

' Elimina la fila selccionada del Grid
If grdDetalle.CellBackColor = vbDarkBlue And grdDetalle.Row > 0 Then
      ' elimina la fila seleccionada del grid
    If grdDetalle.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdDetalle.RemoveItem grdDetalle.Row
    Else
            ' estable vacío el grddetalle
            grdDetalle.Rows = 1
    End If
End If

' Actualiza la posición del grid
ipos = 0

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub cmdPProdServ_Click()

If cboProdServ.Enabled Then
    ' alto
     cboProdServ.Height = CBOALTO
    ' focus a cbo
    cboProdServ.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()
 
 'cierra el formulario
 Unload Me
 
End Sub

Private Sub grdDetalle_Click()

If grdDetalle.Row > 0 And grdDetalle.Row < grdDetalle.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdDetalle, vbWhite, vbDarkBlue
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

Private Sub grdDetalle_DblClick()

' Verifica si se canceló las operaciones en el grid
  If grdDetalle.Row > 0 Then
    
    ' Verifica si esta seleccionado
    If grdDetalle.CellBackColor <> vbDarkBlue Then
       MarcarFilaGRID grdDetalle, vbWhite, vbDarkBlue
       Exit Sub
    End If
    
    ' Carga la fila selecionada
    CargarEditarFila
    
    ' Elimina la fila seleccionada del grid
    If grdDetalle.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdDetalle.RemoveItem grdDetalle.Row
    Else
            ' estable vacío el grddetalle
            grdDetalle.Rows = 1
    End If
    ' coloca el focus a cbo producto
    cboProdServ.SetFocus
    
    ' Actualiza el ipos
    ipos = 0
  End If

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
    msIdprod = grdDetalle.TextMatrix(grdDetalle.RowSel, 4)
 
   'Recupera en el combo el producto del msidprodserv
    CD_ActVarCbo cboProdServ, msIdprod, mcolCodDesProd
 
    txtMedida.Text = grdDetalle.TextMatrix(grdDetalle.RowSel, 5)
    msCodSuministro = grdDetalle.TextMatrix(grdDetalle.RowSel, 6)
    msCodVariacion = grdDetalle.TextMatrix(grdDetalle.RowSel, 7)

   ' Pone los Montos en sus respectivos controles
    txtCant.Text = grdDetalle.TextMatrix(grdDetalle.RowSel, 1)
   
   ' Pone el monto en valor venta
    txtValor = grdDetalle.TextMatrix(grdDetalle.RowSel, 3)
    
End Sub

Private Sub Form_Load()

'Carga la colección de producto
CargarColProducto

'Coloca el titulo al Grid
aTitulosColGrid = Array("Producto", "Cantidad", "Precio Unitario", "Total", "idproducto", "Medida", "Suministros", "VarExistencias")
aTamañosColumnas = Array(3900, 1400, 1400, 1400, 0, 0, 0, 0)
CargarGridTitulos grdDetalle, aTitulosColGrid, aTamañosColumnas

'Establece campos obligatorios de la primera parte del formulario
EstableceCamposObligatorios1raParte

'Establece campos obligatorios de la segunda parte del formulario
EstableceCamposObligatorios2daParte

' Inicializa el grid
ipos = 0
gbCambioCelda = False

' Dependiendo de la operación a realizar prepara el formulario
If gsTipoOperacionAlmacen = "Nuevo" Then
    ' Deshabilita el txtCodigo
    txtCodigo.Enabled = False
    
    ' Deshabilita el botón elegir
    cmdBuscar.Enabled = False
    
    ' Coloca la fecha del sistema
    mskFecTrab.Text = gsFecTrabajo
    
    ' Prepara el formulario para un nuevo egreso
    NuevoIngresoAL
Else
    'Prepara el formulario para modificar el egreso
    ModificarIngresoAL
End If

End Sub

Private Sub HabilitarBotonAceptar()
' ----------------------------------------------------------------
' Propósito: habilita el botón aceptar modificar si es que los datos _
             introducidos estan correctos
' Recibe: Nada
' Entrega: Nada
' ----------------------------------------------------------------
' Deshabilita el boton
cmdAceptar.Enabled = False
 
' Verifica si se a introducido los datos obligatorios generales
If grdDetalle.Rows <= 1 Or txtNumDoc.BackColor <> vbWhite Then
   ' Algún obligatorio falta ser introducido
   Exit Sub
End If

' Verifica si se cambio algún dato
If gsTipoOperacionAlmacen = "Modificar" Then
    ' Verifica si se cambio los datos generales
   If fbCambioDetalle = False And mcurIngreAL.campo(2) = Trim(txtNumDoc) Then
        ' No se cambio ningún dato
          Exit Sub
   End If
End If

' Habilita botón aceptar
cmdAceptar.Enabled = True

End Sub

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
If grdDetalle.Rows - 1 <> mcolIngreALDet.Count Then
    ' Se ha cambiado los impuestos
    fbCambioDetalle = True
    Exit Function
Else
    ' Verifica si se cambió el detalle, recorre el cursor
  If mcolIngreALDet.Count = 0 Then
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
        If sReg <> mcolIngreALDet.Item(grdDetalle.TextMatrix(i, 4)) Then
            ' Se ha cambiado el detalle
            fbCambioDetalle = True
        End If
    Next i

  End If
End If
Exit Function
'-----------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' error , elemento no encontrado
        ' Se ha cambiado el detalle
         fbCambioDetalle = True
         Exit Function
    End If
End Function

Private Function fbVerificarCantidad() As Boolean
' ----------------------------------------------------------------
' Propósito: Verifica si se ha introducido la cantidad de el producto o serv
' Recibe: nada
' Entrega : nada
' ----------------------------------------------------------------
' verifica si se ha ingresado la cantidad
fbVerificarCantidad = False
If txtCant.Text <> Empty Then
    ' se ha introducido la cantidad
    fbVerificarCantidad = True
End If

End Function

Private Sub HabilitaBotonAñadir()
' -------------------------------------------------------
' Propósito: Verifica si se puede habilitar el botón añadir
' Recibe: Nada
' Entrega: Nada
' -------------------------------------------------------
' verifica si estan completos los controles

If cboProdServ.BackColor <> vbWhite Or txtCant.BackColor <> vbWhite _
   Or txtValor.BackColor <> vbWhite Then
    ' deshabilita el botón añadir
    cmdAñadir.Enabled = False
Else
    ' habilita e botón añadir
    cmdAñadir.Enabled = True
End If

End Sub

Private Sub MostrarMedidaCont()
'-----------------------------------------------------------------------------
'Propósito : Muestra la medida del producto
'-----------------------------------------------------------------------------

'Muestra la medida del producto seleccionado en el combo
 txtMedida.Text = Var30(mcolDesMedidaContProd.Item(msIdprod), 1)
 msCodSuministro = Var30(mcolDesMedidaContProd.Item(msIdprod), 2)
 msCodVariacion = Var30(mcolDesMedidaContProd.Item(msIdprod), 3)

End Sub

Private Sub NuevoIngresoAL()
'---------------------------------------------------------------
'Propósito : Prepara el formulario para un ingreso a almacén
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
Dim sSQL As String

' Calcula ingreso a almacén
    txtCodigo = CalcularCodigoAL

' deshabilita los botones del formulario
    HabilitaDeshabilitaBotones ("Nuevo")
    
End Sub

Private Sub ModificarIngresoAL()
'-----------------------------------------------------------------
'Propósito : Prepara el formulario para modificar ingreso a almacén
'Recibe : Nada
'Devuelve :Nada
'-----------------------------------------------------------------

' Habilita el txtCodigo
  txtCodigo.Enabled = True
  txtCodigo.BackColor = Obligatorio

' Deshabilita la 1raParte del formulario
  DeshabilitarHabilitarFormulario False

' Limpia las colecciones
   Set mcolIngreALDet = Nothing

' Maneja estado de los botones del formulario
  HabilitaDeshabilitaBotones "Modificar"

End Sub

Private Sub EstableceCamposObligatorios1raParte()
' -------------------------------------------------------------------
'Propósito: Establece que campos son obligatorios en la primera parte _
            del formulario
' -------------------------------------------------------------------
' Nada
txtNumDoc.BackColor = Obligatorio

End Sub

Private Function fsEstaenDetalle(sProdServ As Variant) As String
'------------------------------------------------------
'Propósito: Verificar la existencia de un Producto en el grdDetalle
'Recibe:    Nada
'Devuelve:  string que indica la existencia de la cboProdServ en el grd detalle
'------------------------------------------------------
'Nota:      llamado desde el evento click de cmdAñadir
Dim j As Integer

'Inicializamos a funcion asumiendo que Producto no esta en el grddetalle 1
fsEstaenDetalle = Empty

' Recorremos el grid detalle de Producto verificando la existencia de txtProd
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


Private Sub EstableceCamposObligatorios2daParte()
' -------------------------------------------------------------------
'Propósito: Establece que campos son obligatorios en la segunda parte _
            del formulario
' -------------------------------------------------------------------
cboProdServ.BackColor = Obligatorio
txtCant.BackColor = Obligatorio
txtValor.BackColor = Obligatorio

End Sub

Public Function CalcularCodigoAL() As String
'------------------------------------------------------
'Propósito: Determina el último registro e incrementa en 1 el campo código
'Recibe:    Nada
'Devuelve: Código
'------------------------------------------------------
Dim sCodigo As String
Dim sNroSecuencial As String
Dim sSQL As String
Dim curCodigoAlmacen As New clsBD2
Dim iNumSec As Integer

'Concatenamos el codigo ALAñoMes
sCodigo = "AL" & Right(gsFecTrabajo, 2) & Mid(gsFecTrabajo, 4, 2)

'Se carga un string con el último registro del campo código
sSQL = "SELECT Max(IdBalance) FROM ALMACEN_BALANCE WHERE IdBalance LIKE '" & sCodigo & "*'"

' Ejecuta la sentencia
curCodigoAlmacen.SQL = sSQL
If curCodigoAlmacen.Abrir = HAY_ERROR Then
  End
End If

'Separa los cuatro últimos caracteres del maximo código
If IsNull(curCodigoAlmacen.campo(0)) Then ' No hay registros
  CalcularCodigoAL = (sCodigo & "0001")
Else ' Tiene registros
 iNumSec = Val(Right(curCodigoAlmacen.campo(0), 4))
 CalcularCodigoAL = sCodigo & Format(CStr(iNumSec) + 1, "000#")
End If
'Cierra el cursor
curCodigoAlmacen.Cerrar

End Function

Private Sub HabilitaDeshabilitaBotones(sProceso As String)
'-----------------------------------------------------------------
' Proposito: Coloca la condición de los botones en el proceso
' Recibe: Nada
' Entrega: Nada
'-----------------------------------------------------------------
Select Case sProceso

' depende del proceso habilita y deshabilita botones
Case "Nuevo"
    cmdAñadir.Enabled = False
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    
Case "Modificar"
    cmdBuscar.Enabled = True
    cmdAñadir.Enabled = False
    cmdAceptar.Enabled = False
    cmdAnular.Enabled = False
    
End Select

End Sub

Private Sub DeshabilitarHabilitarFormulario(bBoleano As Boolean)
'---------------------------------------------------------------
'Propósito : Deshabilita controles editables del formulario
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
txtNumDoc.Enabled = bBoleano
cboProdServ.Enabled = bBoleano
txtCant.Enabled = bBoleano
txtValor.Enabled = bBoleano
grdDetalle.Enabled = bBoleano

End Sub

Private Sub CargarColProducto()
'---------------------------------------------------------------
'Propósito : Carga la colección de Productos con su medida y _
             códigos de suministro y variación de existencias
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
Dim sSQL As String
Dim curMedidaProd As New clsBD2
' Carga la sentencia
sSQL = "SELECT idprod, DescProd, Medida, CodSuministro,CodVariacion " _
     & "FROM PRODUCTOS " _
     & "WHERE ActivoFijo='NO' ORDER BY DescProd"

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
                              Item:=curMedidaProd.campo(2) _
                              & "¯" & curMedidaProd.campo(3) _
                              & "¯" & curMedidaProd.campo(4)
    
    'colección de producto y su descripción
    mcolidprod.Add curMedidaProd.campo(0)
    mcolCodDesProd.Add curMedidaProd.campo(1), curMedidaProd.campo(0)

    ' Se avanza a la siguiente fila del cursor
    curMedidaProd.MoverSiguiente
Loop

' Cierra el cursor de medida de productos
curMedidaProd.Cerrar

' Carga el cboProducto de acuerdo a la relación
CargarCboCols cboProdServ, mcolCodDesProd

End Sub

Private Sub cboProdServ_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProdServ) = True Then SendKeys "{down}"

End Sub

Private Sub cboProdServ_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProdServ.ListIndex) = False And cboProdServ.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboProdServ_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProdServ, Obligatorio) = True Then

    ' Se actualiza código de la variable correspondiente a descripción introducida
    CD_ActCboVar cboProdServ.Text, msIdprod, mcolidprod, mcolCodDesProd

    ' Actualiza la medida correspondiente al producto seleccionada y su codcont
   MostrarMedidaCont
Else
   'no se eligió un product o serv
   msIdprod = Empty
End If

'Cambia el alto del combo
cboProdServ.Height = CBONORMAL

'habilitar el boton añadir
HabilitaBotonAñadir

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'  Descarga las coleciones

' Colecciones para la modificación de el formulario
Set mcolIngreALDet = Nothing

' Colecciones para el manejo del formulario
Set mcolidprod = Nothing
Set mcolCodDesProd = Nothing
Set mcolDesMedidaContProd = Nothing

End Sub

Private Sub CalcularPrecioUni()
'----------------------------------------------------------------------
' Propósito: Calcula el precio unitario de compra de acuerdo al monto _
            de compra y la cantidad
' Recibe : nada
' Entrega : nada
'----------------------------------------------------------------------
' Verifica si la cantidad tiene un valor > 0

If Val(Var37(txtCant)) > 0 And Val(Var37(txtValor)) > 0 Then
 ' se tiene una cantidad aceptable, calcula el precio unitario
 txtPrecioUni = Format(Var37(txtValor) / Var37(txtCant), "###,###,##0.00")
Else
 ' la cantidad es cero, no aceptable o el txtcantidad esta vacía, entonces _
   muestra cero
  txtPrecioUni = "0.00"
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
 CalcularPrecioUni
 
'Habilitar añadir
 HabilitaBotonAñadir

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Dato Ingresado
Var34 txtCant, 8, KeyAscii

End Sub

Private Sub txtCant_LostFocus()

' pone el maximo número de caracteres
txtCant.MaxLength = 8

' Verifica si se introdujjo un dato valido
If Val(Var37(txtCant)) = 0 Then
    txtCant = Empty
Else
    ' formato a cantidad
    txtCant = Format(Val(Var37(txtCant.Text)), "#0.00")
End If

End Sub

Private Sub txtCodigo_Change()

' Verifica el proceso que se realiza en el formulario
If gsTipoOperacionAlmacen = "Modificar" Then
    
    ' Verifica si se ha introducido el tamaño de el código
      If Len(txtCodigo) = txtCodigo.MaxLength Then
        ' Verifica mayúsculas
        If UCase(txtCodigo) = txtCodigo Then
          
          ' Verifica si el ingreso existe existe
          If fbCargaIngresoALBL = True Then
             ' Sale y deshabilita el control
             SendKeys vbTab
            
             ' Habilita el formulario
             DeshabilitarHabilitarFormulario True
              
             ' deshabilita el txtcodigo y el botón buscar, _
               habilita anular
             txtCodigo.Enabled = False
             cmdBuscar.Enabled = False
             cmdEliminar.Enabled = True
             cmdAnular.Enabled = True
          End If ' fin de cargar Ingreso por balance
          
        Else ' vuelve a mayúsulas el txtcodigo
            txtCodigo = UCase(txtCodigo)
        End If ' fin verificar mayúsculas
        
      End If ' fin de verificar el tamaño del texto
 End If
 
 ' Maneja el color del control txtcodigo
 If txtCodigo = Empty Then
    ' coloca el color obligatorio al control
    txtCodigo.BackColor = Obligatorio
 Else
    ' coloca el color de edición
    txtCodigo.BackColor = vbWhite
 End If
 
End Sub

Private Function fbCargaIngresoALBL() As Boolean
' ----------------------------------------------------------
' Propósito: Verifica si existe el código del ingreso almacén y carga _
             los datos de el ingreso a almace´n por balance
' Recibe : Nada
' Entrega : Nada
' ----------------------------------------------------------
Dim sSQL As String

' Carga la sentencia  que verifica si existe el código de ingreso a almacén
 sSQL = "SELECT IB.IdBalance, IB.Fecha, IB.NumDoc " _
     & "FROM ALMACEN_BALANCE IB " _
     & "WHERE IB.IdBalance='" & txtCodigo & "' and IB.Anulado='NO'"

' Ejecuta la sentencia
mcurIngreAL.SQL = sSQL
If mcurIngreAL.Abrir = HAY_ERROR Then End
' Cursor abierto
mbALIngresoCargado = True

' Verifica si existe el código de ingreso a almacén
If mcurIngreAL.EOF Then
    'Mensaje de registro no existe
    MsgBox "El código del ingreso a almacén que se digitó no está registrado como " & _
      "ingreso a almacén por balance o está anulado", vbExclamation, "SCCaijo-Ingreso a Almacén por Balance"
    mcurIngreAL.Cerrar
    ' cierra el cursor y se va
    fbCargaIngresoALBL = False
Else
    ' Carga el cursor que contiene el detalle del ingreso a almacén
    sSQL = "SELECT G.Orden, G.Concepto, G.CodConcepto, G.Cantidad, G.Monto,(G.Monto/G.Cantidad) " _
         & "FROM GASTOS G " _
         & "WHERE G.Orden='" & txtCodigo & "'"
    ' Ejecuta la sentencia
    mcurIngreALDet.SQL = sSQL
    If mcurIngreALDet.Abrir = HAY_ERROR Then End
    
    ' Carga los datos generales
    CargaControlesGenerales
    ' Carga los datos del detalle
    CargaControlesDetalle
    ' Devuelve el resultado de la función
    fbCargaIngresoALBL = True
End If

End Function

Private Sub CargaControlesGenerales()
'--------------------------------------------------------------
'Propósito: Carga los controles del formulario referidos  a los _
            datos generales de el ingreso a almacén por balance
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------
' Carga los controles editables y de opción
Dim sSQL As String
'IB.IdBalance, IB.Fecha, IB.Observacion,
' Actualiza Variables
msCodigo = mcurIngreAL.campo(0)
    
' Carga los datos en sus controles
mskFecTrab = FechaDMA(mcurIngreAL.campo(1))
txtNumDoc = mcurIngreAL.campo(2)
    
End Sub

Private Sub CargaControlesDetalle()
'--------------------------------------------------------------
'Propósito: Carga los controles del formulario referidos  a los _
            datos del detalle del Ingreso a almacén por balance
'Recibe:    Nada
'Devuelve:  Nada
'--------------------------------------------------------------
' Carga los datos relacionados al detalle en el grid
If (mcurIngreALDet.EOF) Then ' verifica que no sea vacio
    MsgBox "Error Ingreso a almacén sin detalle en BD, Consulte al administrador", , "SGCcaijo-Ingreso a almacén por balance"
    Exit Sub
Else
    ' Carga el grid del detalle
    Do While Not mcurIngreALDet.EOF
   'Producto,Cantidad,Precio Unitario,Total,idproducto,Medida,CodSuministro, CodVariacion

   grdDetalle.AddItem Var30(mcolCodDesProd.Item(mcurIngreALDet.campo(2)), 1) _
    & vbTab & Format(mcurIngreALDet.campo(3), "#0.00") & vbTab & Format(mcurIngreALDet.campo(5), "###,###,##0.00") _
    & vbTab & Format(mcurIngreALDet.campo(4), "###,###,##0.00") & vbTab & mcurIngreALDet.campo(2) _
    & vbTab & Var30(mcolDesMedidaContProd.Item(mcurIngreALDet.campo(2)), 1) _
    & vbTab & Var30(mcolDesMedidaContProd.Item(mcurIngreALDet.campo(2)), 2) _
    & vbTab & Var30(mcolDesMedidaContProd.Item(mcurIngreALDet.campo(2)), 3)
        
   ' Carga la colección detalle codconcepto, cantidad, monto
        mcolIngreALDet.Add Item:=mcurIngreALDet.campo(2) & "¯" _
                            & Format(mcurIngreALDet.campo(3), "#0.00") & "¯" _
                            & Format(mcurIngreALDet.campo(4), "###,###,##0.00"), _
                        Key:=mcurIngreALDet.campo(2)
        
        mcurIngreALDet.MoverSiguiente
    Loop
    
    'Cierra el cursor de detalle de ingreso a almacén
    mcurIngreALDet.Cerrar
End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtNumDoc_Change()

' Verifica  que NO tenga campos en blanco
If txtNumDoc.Text <> "" And InStr(txtNumDoc, "'") = 0 Then
' Los campos coloca a color blanco
   txtNumDoc.BackColor = vbWhite
Else
' Marca los campos obligatorios
   txtNumDoc.BackColor = Obligatorio
End If

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    'Convierte a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtValor_Change()
' Si no se ha introducido valor, se marca campo obligatorio
If txtValor.Text <> "" And Val(txtValor.Text) <> 0 Then
   ' coloca el color de correcto
   txtValor.BackColor = vbWhite
   
   ' se esta cambiando el valor de compra. Calcula PrecioUniVC y si
   CalcularPrecioUni

Else
    ' coloca el color obligatorio
   txtValor.BackColor = Obligatorio
End If

' HabilitaAñadir
  HabilitaBotonAñadir

End Sub

Private Sub txtValor_GotFocus()

' Verifica si se ha introducido el valor en txtcantidad
If fbVerificarCantidad = False Then SendKeys vbTab
'Elimina las comas
txtValor.MaxLength = 12
txtValor.Text = Var37(txtValor.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Monto ingresado en soles, luego se ubica en la componente inmediata
Var33 txtValor, KeyAscii

End Sub

Private Sub txtValor_LostFocus()

'Maxima longitud
txtValor.MaxLength = 14
If txtValor.Text <> "" Then
   'Da formato de moneda
   txtValor.Text = Format(Val(Var37(txtValor.Text)), "###,###,###,##0.00")
Else
   txtValor.BackColor = Obligatorio
End If

End Sub
