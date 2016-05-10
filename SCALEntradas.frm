VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmALEntrVerif 
   Caption         =   "Almacén- Verificación de Entradas"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   HelpContextID   =   83
   Icon            =   "SCALEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptInformes 
      Left            =   840
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   6735
      Width           =   1125
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   6735
      Width           =   1125
   End
   Begin VB.Frame fVerificarProds 
      Caption         =   "Productos a verificar:"
      Height          =   4005
      Left            =   120
      TabIndex        =   10
      Top             =   2670
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid grdVerificar 
         Height          =   3700
         Left            =   120
         TabIndex        =   7
         Top             =   200
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   6535
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         HighLight       =   0
         FillStyle       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2660
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtProveedor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtTipDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3975
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtProy 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtNumDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtOrden 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskFecEgre 
         Height          =   315
         Left            =   6015
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecTrab 
         Height          =   315
         Left            =   6000
         TabIndex        =   1
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
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblProveedor 
         Caption         =   "Proveedor:"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Caja:"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Left            =   3180
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblProy 
         Caption         =   "Proyecto:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblNumDoc 
         Caption         =   "Num Doc:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblOrden 
         Caption         =   "Orden:"
         Height          =   255
         Left            =   495
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmALEntrVerif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Colección para guardar los productos y su verificación en almacén
Dim mcolProdVerificar As New Collection

' Variable para el manejo de el grid selección
Dim ipos As Long

Private Sub IngresarProdsVerificados()

Dim sSQL As String
Dim modAlmacen As New clsBD3
Dim i As Integer
Dim sPrecioUnit As String
Dim sResto As String
' Recorre el grd
For i = 1 To grdVerificar.Rows - 1
   ' Rebisa el campo Verificado
   If grdVerificar.TextMatrix(i, 4) = "SI" Then
        
        ' Se ha verificado en el grd, se modifica en BD ALMACEN_VERIFICACION
        sSQL = "UPDATE ALMACEN_VERIFICACION SET " _
        & "Verificado = 'SI' " _
        & "WHERE Orden='" & txtOrden & "' and " _
        & "IdProd='" & grdVerificar.TextMatrix(i, 0) & "'"
        
        ' Ejecuta la sentencia
        modAlmacen.SQL = sSQL
        If modAlmacen.Ejecutar = HAY_ERROR Then End
        
        'Cierra la instancia de modificacion de la base de datos
        modAlmacen.Cerrar
        
        
        ' Ingresa los producto como mercaderias o activos fijos
        If grdVerificar.TextMatrix(i, 8) = "SI" Then
            ' Ingresa como activo fijo
            ' Se ingresa el producto a ACTIVOFIJO_INGRESOS, carga la sentencia
            ' Orden, Prod,NumIngreso,Monto,Fecha
            sSQL = "INSERT INTO ACTIVOFIJO_INGRESOS VALUES('" _
            & txtOrden & "','" & grdVerificar.TextMatrix(i, 0) & "','" _
            & fsAsignarNumIngreso(i) & "'," & Var37(grdVerificar.TextMatrix(i, 5)) & ",'" _
            & FechaAMD(mskFecEgre) & "')"
            
            ' Ejecuta la sentencia
            modAlmacen.SQL = sSQL
            If modAlmacen.Ejecutar = HAY_ERROR Then End
            
            'Cierra la instancia de modificacion de la base de datos
            modAlmacen.Cerrar
            
                        
        Else ' Ingresa como mercaderias
  
            '"Id.Prod", "Producto", "Unidad", "Cantidad", "Verificado", "Monto"
            ' Calcula PrecioUnit, Resto
            sPrecioUnit = Format(Val(Var37(grdVerificar.TextMatrix(i, 5))) / Val(Var37(grdVerificar.TextMatrix(i, 3))), "#0.00")
            sResto = Format(Val(Var37(grdVerificar.TextMatrix(i, 5))), "#0.00")
            ' Verifica si el resto es negativo
            If Val(sResto) < Val(Val(sPrecioUnit) * Val(Var37(grdVerificar.TextMatrix(i, 3)))) Then
                sPrecioUnit = Format(Val(sPrecioUnit) - 0.01, "#0.00")
            End If
        
            ' Se ingresa el producto a ALMACEN_INGRESOS, carga la sentencia
            ' Orden, Prod,NumIngreso,PrecioUnit,Resto,CantDisponible,Fecha
            sSQL = "INSERT INTO ALMACEN_INGRESOS VALUES('" _
            & txtOrden & "','" & grdVerificar.TextMatrix(i, 0) & "','" _
            & fsAsignarNumIngreso(i) & "'," & sPrecioUnit & "," _
            & sResto & "," & grdVerificar.TextMatrix(i, 3) & ",'" _
            & FechaAMD(mskFecEgre) & "')"
            
            ' Ejecuta la sentencia
            modAlmacen.SQL = sSQL
            If modAlmacen.Ejecutar = HAY_ERROR Then End
            
            'Cierra la instancia de modificacion de la base de datos
            modAlmacen.Cerrar
            
            'Carga la colección asiento IdProd, Tratamiento, CodSuministro, CodVariación, Monto
            'Id.Prod,Producto,Unidad,Cantidad,Verificado,Monto,CodSuministro,CodVarExistencia
            gcolAsientoDet.Add _
                Key:=grdVerificar.TextMatrix(i, 0), _
                Item:=grdVerificar.TextMatrix(i, 0) & "¯" _
                & "IA" & "¯" & grdVerificar.TextMatrix(i, 6) & "¯" _
                & grdVerificar.TextMatrix(i, 7) & "¯" & Var37(grdVerificar.TextMatrix(i, 5))
            
            'Colección que guarda los datos generales para los asientos contable
            'Orden, Fecha, Producto, Glosa
            gcolAsiento.Add _
                Key:=txtOrden, _
                Item:=txtOrden & "¯" & FechaAMD(mskFecEgre.Text) & "¯" _
                & grdVerificar.TextMatrix(i, 0) & "¯INGRESO A ALMACEN¯Ingreso¯IA"
            
            'Asiento de ingreso de almacen
            gsTipoAlmacen = "Ingreso"
            
            'Realiza el asiento automatico de ingreso a almacén
            Conta47
               
        End If
   End If
Next i

End Sub

Private Sub LlenarTablaRPTALINGRESO()

Dim i As Integer
Dim sSQL As String
Dim sPrecioUnit As String
Dim curFechasVerif As New clsBD2
Dim modAlIngreso As New clsBD3
Dim colFechasVerif As New Collection

' Guarda los datos del documento
sSQL = "INSERT INTO RPTALINGRESOVERIFDOC VALUES('" & txtOrden & "','" _
                                          & FechaAMD(mskFecEgre) & "','" _
                                          & txtTipDoc & "','" & txtNumDoc & "','" _
                                          & txtProveedor & "')"
' Ejecuta la sentencia
Var14 sSQL, False

' Averigua las fechas de las verificaciones de las mercaderías
sSQL = "SELECT AI.IdProd, AI.Fecha " _
    & "FROM ALMACEN_INGRESOS AI " _
    & "WHERE AI.Orden='" & txtOrden & "'"
' Ejecuta la sentencia
curFechasVerif.SQL = sSQL
If curFechasVerif.Abrir = HAY_ERROR Then End

Do While Not curFechasVerif.EOF

    ' Guarda en una colección
    colFechasVerif.Add curFechasVerif.campo(1), curFechasVerif.campo(0)
    
    ' Mueve al siguiente elemento
    curFechasVerif.MoverSiguiente
    
Loop

' Cierra el cursor
curFechasVerif.Cerrar

' Averigua las fechas de las verificaciones de los Activos fijos
sSQL = "SELECT AI.IdProd, AI.Fecha " _
    & "FROM ACTIVOFIJO_INGRESOS AI " _
    & "WHERE AI.Orden='" & txtOrden & "'"
' Ejecuta la sentencia
curFechasVerif.SQL = sSQL
If curFechasVerif.Abrir = HAY_ERROR Then End

Do While Not curFechasVerif.EOF
    
    ' Guarda en una colección
    colFechasVerif.Add curFechasVerif.campo(1), curFechasVerif.campo(0)
    
    ' Mueve al siguiente elemento
    curFechasVerif.MoverSiguiente
Loop

' Cierra el cursor
curFechasVerif.Cerrar

' Guarda los datos de las verificaciones de los productos
For i = 1 To grdVerificar.Rows - 1
    ' Carga los productos verificados
    If grdVerificar.TextMatrix(i, 4) = "SI" Then
    
        ' "Id.Prod", "Producto", "Unidad", "Cantidad", "Verificado", "Monto", "Suministros", "VarExistencias", "ActivoFijo"
        sPrecioUnit = Format(Val(Var37(grdVerificar.TextMatrix(i, 5))) / Val(Var37(grdVerificar.TextMatrix(i, 3))), "#0.00")
         sSQL = "INSERT INTO RPTALINGRESOVERIFDET VALUES " _
         & "('" & txtOrden & "','" _
         & colFechasVerif.Item(grdVerificar.TextMatrix(i, 0)) & "'," _
         & Var37(grdVerificar.TextMatrix(i, 3)) & ",'" _
         & grdVerificar.TextMatrix(i, 2) & "','" _
         & grdVerificar.TextMatrix(i, 0) & "','" _
         & grdVerificar.TextMatrix(i, 1) & "'," _
         & sPrecioUnit & "," _
         & Var37(grdVerificar.TextMatrix(i, 5)) & ")"
        
        'Copia la sentencia sSQL
        modAlIngreso.SQL = sSQL
        
        'Verifica si hay error
        If modAlIngreso.Ejecutar = HAY_ERROR Then
          End
        End If
        
        'Se cierra la query
        modAlIngreso.Cerrar
    End If
Next i

' Limpia la colección de fechas
Set colFechasVerif = Nothing

End Sub

Private Sub ImprimirIngreso()

Dim sSQL As String
Dim rptIngresoAlmacen As New clsBD4

'Llena la tabla con datos
  LlenarTablaRPTALINGRESO

' Formulario
  Set rptIngresoAlmacen.frmRef = Me

' Se asigna el valor del control Crystal del Formulario a la CLASE.
  rptIngresoAlmacen.AsignarRpt

' Clausula WHERE de las relaciones del rpt.
  rptIngresoAlmacen.FiltroSelectionFormula = ""
  
' Nombre del fichero
  rptIngresoAlmacen.NombreRPT = "rptALIngresoVerif.rpt"

' Presentación preliminar del Informe
  rptIngresoAlmacen.ImpresionDirecta

'Sentencia SQL
 sSQL = "DELETE * FROM RPTALINGRESOVERIFDOC"
'Borra la tabla
 Var21 sSQL

'Sentencia SQL
 sSQL = "DELETE * FROM RPTALINGRESOVERIFDET"
'Borra la tabla
 Var21 sSQL


End Sub

Private Function fsAsignarNumIngreso(j As Integer)
Dim sCodigo As String
Dim sSQL As String
Dim curNumeroIngreso As New clsBD2
Dim iNumSec As Long

' Concatenamos el codigo AñoMes
  sCodigo = Right(mskFecTrab, 4)
  
' Verifica si se calcula el número de ingreso de mercaderias o de activos fijos
  If grdVerificar.TextMatrix(j, 8) = "NO" Then ' Activo fijo, carga sentencia
        sSQL = "SELECT Max(NroIngreso)  FROM ALMACEN_INGRESOS WHERE NroIngreso  LIKE '" & sCodigo & "*'"
  Else ' Mercadería, Carga la sentencia
        sSQL = "SELECT Max(NroActivo)  FROM ACTIVOFIJO_INGRESOS WHERE  NroActivo LIKE '" & sCodigo & "*'"
  End If

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

Private Sub HabilitarBotonAceptar()

   If fbOkGrid = True Then ' Verifica si se verificó o desverificó
    'Habilita el boton aceptar
    cmdAceptar.Enabled = True
   Else
    'Habilita el boton aceptar
    cmdAceptar.Enabled = False
   End If
     
End Sub

Private Function fbOkGrid() As Boolean
Dim sSINO As String
Dim i As Integer

' inicializa la función asumiendo que no se verificó o desverificó nada
fbOkGrid = False
' Averigua el proceso en el formulario verificación
If gsTipoOperacionAlmacen = "Nuevo" Then
    sSINO = "SI"
Else
    sSINO = "NO"
End If
' recorre el grdProdverificado para averiguar si hay alguno que falta
For i = 1 To grdVerificar.Rows - 1
    If grdVerificar.TextMatrix(i, 4) = sSINO Then
        'Hay alguno sin verificar
        fbOkGrid = True
        Exit Function
    End If
Next i

End Function


Private Sub cmdAceptar_Click()

'Verifica si el año esta cerrado
If Conta52(Right(mskFecEgre.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Verificación de almacen"
    Exit Sub
End If

'La fecha a verificar es anterior al egreso a caja
If FechaAMD(mskFecTrab.Text) < FechaAMD(mskFecEgre.Text) Then
    'Mensaje
    MsgBox "La fecha de verificación debe ser posterior a la compra!.", vbOKOnly + vbInformation, "SGCcaijo - Verificación de ingreso a almacen"
    'Termina la ejecución del procedimiento
    Exit Sub
End If

' Guarda los datos del formulario en la BD
If gsTipoOperacionAlmacen = "Nuevo" Then
    ' Verifica si se puede verificar los productos seleccionados
    If fbPuedeIngresar = False Then Exit Sub
    ' Verifica si se han verificado todos los Prods
    
    If fbEstanTodosVerificados = False Then
        ' Mensaje de si desea continuar
        If MsgBox("No se han verificado todos los productos del documento " & txtNumDoc & Chr(13) & _
         "Algunos de los productos no estarán disponibles en almacén." & _
         "¿Desea continuar con los verificados?", vbInformation + vbYesNo, _
         "Almacén- Verificación de Ingresos") = vbNo Then Exit Sub
    Else ' Todos los productos fueron verificados
        If MsgBox("¿Está conforme con las verificaciones ?", vbYesNo + vbQuestion, "Almacén- Verificación de Ingresos") = vbNo Then Exit Sub
    End If
        'Actualiza la transaccion
         Var8 1, gsFormulario
         
        ' Si esta Ok Verifica los Prods en Almacén
         IngresarProdsVerificados
    
ElseIf gsTipoOperacionAlmacen = "Modificar" Then
    
    ' Verifica si se puede desverificar de almacén Productos ingresados
    If fbPuedeDesverificar = False Then Exit Sub
        
    ' Mensaje de confirmación
    If MsgBox("¿Está conforme con las modificaciones ?", vbYesNo + vbQuestion, "Almacén- Verificación de Ingresos") = vbNo Then Exit Sub
    
    'Actualiza la transaccion
     Var8 1, gsFormulario
    
    'Desverifica Productos antes verificados e ingresados a almacén
     DesverificarProds

End If


'Actualiza la transaccion
Var8 -1, Empty

' Msg de operación realizada
MsgBox "Operación realizada correctamente", , "SGCcaijo- Verificación de Productos"

' Cierra el formulario
Unload Me

End Sub
 
Private Sub DesverificarProds()

Dim sSQL As String
Dim modAlmacen As New clsBD3
Dim i As Integer
' Recorre el grd
For i = 1 To grdVerificar.Rows - 1
   
   ' Rebisa los productos desverficados
   If grdVerificar.TextMatrix(i, 4) = "NO" Then
   
    ' Desverifica el producto
       If grdVerificar.TextMatrix(i, 8) = "NO" Then ' Elimina el ingreso de la mercadería

         ' Orden, Prod,NumIngreso,PrecioUnit,Resto,CantDisponible,Fecha
         sSQL = "DELETE * FROM ALMACEN_INGRESOS WHERE  " _
         & "Orden='" & txtOrden & "' and IdProd='" & grdVerificar.TextMatrix(i, 0) & "'"
        
         ' Ejecuta la sentencia
         modAlmacen.SQL = sSQL
         If modAlmacen.Ejecutar = HAY_ERROR Then End
         
         'Cierra la instancia de modificacion de la base de datos
         modAlmacen.Cerrar
    
                  ' Carga la colección asiento para eliminar el ingreso  contable
         ' Orden , IdProducto
         gcolAsiento.Add Key:=txtOrden, _
                         Item:=txtOrden & "¯" & grdVerificar.TextMatrix(i, 0)
         ' Actualiza contabilidad
         Conta44
         
       Else ' Elimina los activos fijos

         ' Orden, Prod,NumActivo,Monto,Fecha
         sSQL = "DELETE * FROM ACTIVOFIJO_INGRESOS WHERE  " _
         & "Orden='" & txtOrden & "' and IdProd='" & grdVerificar.TextMatrix(i, 0) & "'"
        
         ' Ejecuta la sentencia
         modAlmacen.SQL = sSQL
         If modAlmacen.Ejecutar = HAY_ERROR Then End
         
         'Cierra la instancia de modificacion de la base de datos
         modAlmacen.Cerrar
        
       End If ' Fin de verificar si es mercadería o activo fijo
        
       ' Se desverifica en el grd, se modifica en BD ALMACEN_VERIFICACION
       sSQL = "UPDATE ALMACEN_VERIFICACION SET " _
        & "Verificado = 'NO' " _
        & "WHERE Orden='" & txtOrden & "' and " _
        & "IdProd='" & grdVerificar.TextMatrix(i, 0) & "'"
        
        ' Ejecuta la sentencia
       modAlmacen.SQL = sSQL
       If modAlmacen.Ejecutar = HAY_ERROR Then End
        
        'Cierra la instancia de modificacion de la base de datos
       modAlmacen.Cerrar

    End If ' fin de eliminar los desverificados
Next i

End Sub

Private Function fbPuedeDesverificar() As Boolean

Dim i As Integer
Dim sSQL As String
Dim curProdAnterior As New clsBD2

'Inicializa la función asumiendo que no hay Productos anteriores que verificar
fbPuedeDesverificar = True

' Recorre el grdProdverificado para averiguar si hay alguno que falta
For i = 1 To grdVerificar.Rows - 1
    If grdVerificar.TextMatrix(i, 4) = "NO" And grdVerificar.TextMatrix(i, 8) = "NO" Then
        ' Hay alguno para desverificar, Comprueba si el total del ingreso _
          del producto es Igual que el total disponible del mismo
        ' Carga la sentencia
        sSQL = "SELECT CantidadDisponible FROM ALMACEN_INGRESOS " _
           & "WHERE Orden='" & txtOrden & "' and IdProd='" _
           & grdVerificar.TextMatrix(i, 0) & "'"
           
        
        ' ejecuta la sentencia
        curProdAnterior.SQL = sSQL
        If curProdAnterior.Abrir = HAY_ERROR Then End
        ' Verifica si existen Productos, Ingresados a almacén
        If curProdAnterior.EOF Then
            MsgBox "Existen en BD Productos verificados pero no están disponibles, Consultar con el Administrador" _
                    , , "SGCcaijo- Verificación de Productos "
                fbPuedeDesverificar = False
                curProdAnterior.Cerrar
                Exit Function
                    
        Else ' existen productos ingresados
        '"Id.Prod", "Producto", "Unidad", "Cantidad", "Verificado", "Monto"
            If Val(Var37(grdVerificar.TextMatrix(i, 3))) <> Val(curProdAnterior.campo(0)) _
            Then
                MsgBox "No se puede desverificar se ha dado salida a algunos" _
                      & " de los productos: " & Chr(13) & grdVerificar.TextMatrix(i, 1) & Chr(13) _
                      & ", Consulte al Administrador", , _
                      "SGCcaijo- Verificación de Productos"
                fbPuedeDesverificar = False
                curProdAnterior.Cerrar
                Exit Function
             End If
        End If
        ' cierra la consulta
        curProdAnterior.Cerrar
    End If

Next i

End Function

Private Function fbPuedeIngresar() As Boolean

'  Dim i As Integer
'  Dim sSQL As String
'  Dim curProdAnterior As New clsBD2
'
'  'Inicializa la función asumiendo que no hay Productos anteriores que verificar
'  fbPuedeIngresar = True
'
'  ' recorre el grdProdverificado para averiguar si hay alguno que falta
'  For i = 1 To grdVerificar.Rows - 1
'    If grdVerificar.TextMatrix(i, 4) = "SI" Then
'        ' Hay alguno para verificar, Comprueba si existen Productos _
'          del mismo tipo sin verificar anteriores al Orden
'        ' Carga la sentencia
'        If Mid(txtOrden, 1, 2) = "CA" Then
'
'            sSQL = "SELECT AV.Orden,AV.IdProd FROM ALMACEN_VERIFICACION AV,EGRESOS E " _
'           & "WHERE AV.Orden=E.Orden and IdProd='" & grdVerificar.TextMatrix(i, 0) & "' and " _
'           & "Verificado='NO' and (E.FecMov<'" & FechaAMD(mskFecEgre) & "' Or " _
'           & "AV.Orden<'" & txtOrden & "') " _
'           & "ORDER BY E.FecMov, AV.Orden"
'
'        Else
'        ' Ejecuta la sentencia
'          curProdAnterior.SQL = sSQL
'        End If
'        If curProdAnterior.Abrir = HAY_ERROR Then End
'        ' Verifica si existen Productos, si los hay sale de la función
'        If curProdAnterior.EOF Then
'        Else ' existen productos anteriore que modificar
'            MsgBox "No se pueden verificar los Productos por que falta" & Chr(13) _
'                  & "verificar: " & grdVerificar.TextMatrix(i, 1) & Chr(13) _
'                  & "en un movimiento anterior con el orden: " & curProdAnterior.campo(0), , _
'                  "SGCcaijo- Verificación de Productos"
'            fbPuedeIngresar = False
'            curProdAnterior.Cerrar
'            Exit Function
'        End If
'        ' cierra la consulta
'        curProdAnterior.Cerrar
'    End If
'  Next i
Dim i As Integer
Dim sSQL As String
Dim curProdAnterior As New clsBD2

'Inicializa la función asumiendo que no hay Productos anteriores que verificar
fbPuedeIngresar = True

' recorre el grdProdverificado para averiguar si hay alguno que falta
For i = 1 To grdVerificar.Rows - 1
    If grdVerificar.TextMatrix(i, 4) = "SI" Then
        ' Hay alguno para verificar, Comprueba si existen Productos _
          del mismo tipo sin verificar anteriores al Orden
        ' Carga la sentencia
      sSQL = "SELECT AV.Orden,AV.IdProd FROM ALMACEN_VERIFICACION AV,EGRESOS E " _
           & "WHERE AV.Orden=E.Orden and IdProd='" & grdVerificar.TextMatrix(i, 0) & "' and " _
           & "Verificado='NO' and (E.FecMov<'" & FechaAMD(mskFecEgre) & "' Or " _
           & "AV.Orden<'" & txtOrden & "') " _
           & "ORDER BY E.FecMov, AV.Orden"
        ' Ejecuta la sentencia
        curProdAnterior.SQL = sSQL
        If curProdAnterior.Abrir = HAY_ERROR Then End
        ' Verifica si existen Productos, si los hay sale de la función
        If curProdAnterior.EOF Then
        Else ' existen productos anteriore que modificar
            MsgBox "No se pueden verificar los Productos por que falta" & Chr(13) _
                  & "verificar: " & grdVerificar.TextMatrix(i, 1) & Chr(13) _
                  & "en un movimiento anterior con el orden: " & curProdAnterior.campo(0), , _
                  "SGCcaijo- Verificación de Productos"
            fbPuedeIngresar = False
            curProdAnterior.Cerrar
            Exit Function
        End If
        ' cierra la consulta
        curProdAnterior.Cerrar
    End If
Next i
End Function

Private Function fbEstanTodosVerificados() As Boolean

Dim i As Integer

'Inicializa la función asumiendo que todos los productos estan verificados
fbEstanTodosVerificados = True

' recorre el grdProdverificado para averiguar si hay alguno que falta
For i = 1 To grdVerificar.Rows - 1
    If grdVerificar.TextMatrix(i, 4) <> "SI" Then
        'Hay alguno sin verificar
        fbEstanTodosVerificados = False
        Exit Function
    End If
Next i

End Function

Private Sub cmdSalir_Click()
' cierra el formulario
    Unload Me
End Sub

Private Sub cmdTratamiento_Click()

End Sub

Private Sub Form_Load()

' Pone los datos generales
With frmALSelEntrada.grdIngreso
    txtOrden = .TextMatrix(.Row, 0)
    txtProy = .TextMatrix(.Row, 1)
    txtNumDoc = .TextMatrix(.Row, 2)
    txtTipDoc = .TextMatrix(.Row, 3)
    txtProveedor = .TextMatrix(.Row, 4)
    mskFecEgre = .TextMatrix(.Row, 5)
    mskFecTrab = gsFecTrabajo
End With

' Coloca el titulo al Grid
aTitulosColGrid = Array("Id.Prod", "Producto", "Unidad", "Cantidad", "Verificado", "Monto", "Suministros", "VarExistencias", "ActivoFijo")
aTamañosColumnas = Array(0, 4700, 900, 800, 800, 0, 0, 0, 0)
CargarGridTitulos grdVerificar, aTitulosColGrid, aTamañosColumnas

' Inicializa el grid
ipos = 0
gbCambioCelda = False

End Sub


Private Sub CargarProdVerificar(sSINO As String)

Dim sSQL As String
Dim curProdsVerificar As New clsBD2

' Carga la sentencia que consulta a la BD acerca del los productos a verificar
sSQL = "SELECT A.IdProd,P.DescProd, P.Medida, G.Cantidad, A.Verificado, G.Monto, P.CodSuministro, P.CodVariacion, P.ActivoFijo " & _
     "FROM ALMACEN_VERIFICACION A,GASTOS G, PRODUCTOS P WHERE " & _
     "A.Orden='" & txtOrden & "' and A.Verificado='" & sSINO & _
     "' and A.Orden=G.Orden and A.IdProd=G.Codconcepto and " & _
     "G.Concepto='P' and A.IdProd=P.IdProd " & _
     "ORDER BY P.DescProd "

' Ejecuta la consulta
curProdsVerificar.SQL = sSQL
' Abre el cursor si hay error sale indicando la causa del error
If curProdsVerificar.Abrir = HAY_ERROR Then End

If curProdsVerificar.EOF Then
  'No existen Productos para este Orden, error en la BD
  MsgBox "Error Integridad BD, No se pudieron cargar los Prod para este Documento. " & Chr(13) & _
         "Consulte al Administrador ", vbInformation + vbOKOnly, "S.G.Ccaijo-Salida de Almacen"
  ' cierra el formulario
  Unload Me
Else
  'Verifica la existencia del registro de Egreso
    Do While Not curProdsVerificar.EOF
        If curProdsVerificar.campo(4) = sSINO Then
        'Añade el nuevo registro al grid
        grdVerificar.AddItem (curProdsVerificar.campo(0) & vbTab & curProdsVerificar.campo(1) _
        & vbTab & curProdsVerificar.campo(2) & vbTab & curProdsVerificar.campo(3) & vbTab _
        & curProdsVerificar.campo(4) & vbTab & curProdsVerificar.campo(5) & vbTab _
        & curProdsVerificar.campo(6) & vbTab & curProdsVerificar.campo(7)) & vbTab _
        & curProdsVerificar.campo(8)
        
      End If
        'Mueve al siguiente registro del cursor
        curProdsVerificar.MoverSiguiente
    Loop

End If
'Cierra la instancia de consulta
curProdsVerificar.Cerrar

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' colección para guardar los productos y su verificación en almacén
Set mcolProdVerificar = Nothing

End Sub

Private Sub grdVerificar_Click()

If grdVerificar.Row > 0 And grdVerificar.Row < grdVerificar.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdVerificar, vbWhite, vbDarkBlue
End If

End Sub

Private Sub grdVerificar_EnterCell()

If ipos <> grdVerificar.Row Then
    '  Verifica si es la última fila
    If grdVerificar.Row > 0 And grdVerificar.Row < grdVerificar.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdVerificar, ipos
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdVerificar.Row
End If

End Sub

Private Sub grdVerificar_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Llama al proceso que cambia la verificación de un producto
    grdVerificar_DblClick
 End If
 
End Sub


Private Sub grdVerificar_DblClick()

' si la fila esta selecionada(color azul), se cambia la columna
' verificado de Vacio a el caracter constante
 grdVerificar.Col = 4
 
If grdVerificar.CellBackColor = vbDarkBlue Then
    
    If grdVerificar.Text = "NO" Then
      ' si el producto no ha sido verificado, lo vuelve verificado
        grdVerificar.Text = "SI"
        
    ElseIf grdVerificar.Text = "SI" Then
    ' el producto es verificado, lo vuelve no verificado
        grdVerificar.Text = "NO"
    End If

End If

' habilita el botón aceptar
HabilitarBotonAceptar

End Sub

Private Sub txtOrden_Change()

' Verifica el tipo de operación a realizar
If gsTipoOperacionAlmacen = "Nuevo" Then
    'Cargar los registros de los productos ha verificar en almacén, en el grd
    CargarProdVerificar "NO"
    
ElseIf gsTipoOperacionAlmacen = "Modificar" Then
    ' Carga los Productos verificados a desverificar
    CargarProdVerificar "SI"
End If

' Habilita el botón aceptar
HabilitarBotonAceptar

End Sub
