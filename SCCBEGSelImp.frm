VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBEGSelImp 
   Caption         =   "Selección de Impuestos"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Vuelve a la pantalla anerior"
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   400
      Left            =   5160
      TabIndex        =   5
      Top             =   2280
      Width           =   1000
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   400
      Left            =   6240
      TabIndex        =   4
      ToolTipText     =   "Vuelve a la pantalla anerior"
      Top             =   2280
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdImp 
      Height          =   1815
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      Cols            =   3
   End
   Begin VB.CommandButton cmdAceptarModificar 
      Caption         =   "&Aceptar"
      Height          =   400
      Left            =   5160
      TabIndex        =   0
      Top             =   4680
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdRetenciones 
      Height          =   1935
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      Cols            =   5
   End
   Begin VB.Label Label2 
      Caption         =   "Retenciones Aplicadas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6840
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Marque el Impuesto a Aplicar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmCBEGSelImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Define colecciones para la carga de las retenciones
Dim mcolRetenciones As New Collection

'Variable que representa la operacion a realizar por los botones
Dim msOperacion As String

Private Sub cmdAceptarModificar_Click()
Dim i As Integer
Dim sSQL As String
Dim modEgresoCajaBanco As New clsBDModificacion

'verifica el tipo de operacion realizada en el formulario
Select Case msOperacion

Case "Nuevo"
    ' Mensaje de conformidad
    If MsgBox("¿Esta conforme con las retenciones a Aplicar?", vbQuestion + vbYesNo, _
              "Retención de Impuestos") = vbYes Then
        
        'Graba los impuestos en BD
        GrabarImpuestosenBD
            
        'Se descarga el formulario
        Unload Me
   End If
Case "Modificar"
     'Mensaje de conformidad de los datos
    If MsgBox("¿Esta conforme con las modificaciones?", vbQuestion + vbYesNo, _
            "Modificación de Retenciones por Impuestos ") = vbYes Then
       
        'se modifica los registros en mov impuestos
        ModificarRetenidos
        
        'Se descarga el formulario
        Unload Me
    End If
End Select

End Sub

Private Sub ModificarRetenidos()
'--------------------------------------------------------
'Propósito  :Modifica los registros en mov_impuestos
'--------------------------------------------------------
' Nota Llamado desde el Click Aceptar

Dim sSQL, sRegRetenido As String
Dim modRegRetencionImp As New clsBDModificacion
Dim i As Integer

On Error GoTo ErrClaveCol


'recorre el grid
For i = 1 To grdRetenciones.Rows - 1
    'carga el reg detalle para compararlo con el registro cargado en la colección
    '"idImp" "ValorImpuesto", "MontoRetenido",
    sRegRetenido = grdRetenciones.TextMatrix(i, 1) & " " & grdRetenciones.TextMatrix(i, 3) _
        & " " & grdRetenciones.TextMatrix(i, 4)
    'verifica si el registro se encuentra en la coleccion y si se modificó
    ' Si no se encuentra se inserta en ErrClaveCol: error 5
    If mcolRetenciones.Item(grdRetenciones.TextMatrix(i, 1)) <> sRegRetenido Then
       
       ' se modifico el registro, entonces se actualiza la BD
        sSQL = "UPDATE MOV_IMPUESTOS SET " & _
        "Monto=" & grdRetenciones.TextMatrix(i, 4) & "," & _
        "ValorImp=" & EliminarComas(grdRetenciones.TextMatrix(i, 3)) & _
        " WHERE Orden='" & gsOrden & "' and " _
             & "IdImp='" & grdRetenciones.TextMatrix(i, 1) & "'"
        
        modRegRetencionImp.SQL = sSQL
        'ejecuta la sentencia que modifica el registro en gastos
        If modRegRetencionImp.Ejecutar = HAY_ERROR Then End
        modRegRetencionImp.Cerrar
        'eliminar el elmento modificado de la coleccion
        mcolRetenciones.Remove (grdRetenciones.TextMatrix(i, 1))
    Else 'registro no se modifico
      'Solo se elimina de la coleccion para seguir con los demas registros
        mcolRetenciones.Remove (grdRetenciones.TextMatrix(i, 1))
    End If
PostErrClaveCol:
    
Next i

'eliminar los que quedan en la coleccion
 EliminarRetencionesEliminadas
'-------------------------------------------------------------------
ErrClaveCol:

    If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
        'el registro no existe en el egreso con afectacion original
        'carga la sentencia que inserta la retencion en Mov_Impuestos en base de datos
        'Orden,IdImpuesto,MontoRetenido,ValorImpuesto
        sSQL = "INSERT INTO MOV_IMPUESTOS VALUES('" & gsOrden & "','" _
            & grdRetenciones.TextMatrix(i, 1) & "'," & EliminarComas(grdRetenciones.TextMatrix(i, 4)) & "," _
            & EliminarComas(grdRetenciones.TextMatrix(i, 3)) & ")"
        modRegRetencionImp.SQL = sSQL
        'ejecuta la sentencia que añade registro  a Gastos
        If modRegRetencionImp.Ejecutar = HAY_ERROR Then End
        modRegRetencionImp.Cerrar
        Resume PostErrClaveCol ' La ejecución sigue por aquí
    End If
    
End Sub

Private Sub EliminarRetencionesEliminadas()
'--------------------------------------------------------
'Propósito  :Elimina las retenciones que fueron eliminadas del grd
'            Retenciones
'--------------------------------------------------------
Dim sSQL As String
Dim modRegRetenciones As New clsBDModificacion
Dim MiObjeto ' Variables de información.
    
For Each MiObjeto In mcolRetenciones  ' Recorre los elementos que quedanen la coleccion
    'Elimina los registros de gastos
    sSQL = "DELETE * FROM MOV_IMPUESTOS " _
         & "WHERE Orden ='" & gsOrden & "'" _
         & " and IdImp='" & Trim(Left(MiObjeto, 3)) & "'"
    modRegRetenciones.SQL = sSQL
    'ejecuta la sentencia que elimina los registros eliminados del egreso
    If modRegRetenciones.Ejecutar = HAY_ERROR Then End
    modRegRetenciones.Cerrar
    
Next MiObjeto

End Sub

Private Sub GrabarImpuestosenBD()
'--------------------------------------------------------
'Propósito  :Almacena en  la tabla Mov_Impuestos las retenciones
'            seleccionados en el grid de Impuestos calculando
'            las % correspondientes
'--------------------------------------------------------
Dim sSQL As String
Dim modMovImpuestos As New clsBDModificacion
Dim i As Integer

' Recorremos el Grid retenciones guarda en BD las Filas
For i = 1 To grdRetenciones.Rows - 1 'Se recorren las filas
    
              
        sSQL = "INSERT INTO MOV_IMPUESTOS VALUES('" & grdRetenciones.TextMatrix(i, 0) & "','" _
                & grdRetenciones.TextMatrix(i, 1) & "'," & EliminarComas(grdRetenciones.TextMatrix(i, 4)) & "," _
                & EliminarComas(grdRetenciones.TextMatrix(i, 3)) & ")"
        
        'Si al ejecutar hay error se sale de la aplicación
        modMovImpuestos.SQL = sSQL
        If modMovImpuestos.Ejecutar = HAY_ERROR Then
          End
        End If
        
        'Se cierra la query
        modMovImpuestos.Cerrar
Next i

End Sub

Private Sub cmdAñadir_Click()

'verifica si el impuesto elegido en el grd impuesto esta ya en grdretenciones
If fbEstaImpuestoRetenido = True Then
 MsgBox "El impuesto elegido ya esta retenido, " & Chr(13) _
         & "Elimine primero si desea usarlo", , "Retenciones de Impuestos"
 Exit Sub
End If

' Añade una fila al grd retenciones con los alculos necesarios
grdRetenciones.AddItem (gsOrden & vbTab & grdImp.TextMatrix(grdImp.RowSel, 0) _
                & vbTab & grdImp.TextMatrix(grdImp.RowSel, 1) _
                & vbTab & grdImp.TextMatrix(grdImp.RowSel, 2) _
                & vbTab & fsCalcularMontoRetenido)
            
'habilita el boton aceptar
HabilitaBotonAceptarModificar

End Sub

Private Sub HabilitaBotonAceptarModificar()
'------------------------------------------------------------------------
'Proposito :Habilita el botón AceptarModificar de acuerdo al tipo d operación
'           a realizar en el formulario
'Recibe: Nada
'Entrega : Nada
'------------------------------------------------------------------------

' Deshabilita botón AceptarModificar
    cmdAceptarModificar.Enabled = False

Select Case gsTipoOperacionEgreso

Case "Nuevo"
    'verifica si se ha efectuado retenciones
    If grdRetenciones.Rows > 1 Then cmdAceptarModificar.Enabled = True
Case "Modificar"
    'Verifica si se ha realizado modificaciones en las retenciones
        'Habilita botón AceptarModificar
        cmdAceptarModificar.Enabled = True
End Select

End Sub

Private Function fbEstaImpuestoRetenido() As Boolean
'------------------------------------------------------------------------
'Proposito : verificar si en grd Retenidos, ya se encuentra un impuesto_
'            que se eligio de grdImpuestos
'Recibe: nada
'Entrega : booleano que indica si un impuesto ya esta en el grd de retenidos
'------------------------------------------------------------------------
Dim i As Integer
'asume que no se encuentra retenido
fbEstaImpuestoRetenido = False
'recorre el grd retenciones para averiguar si el impuesto elegido ya fué _
 retenido
 If grdRetenciones.Rows = 1 Then 'grid vacio
    Exit Function 'sale de la funcion
 Else
    For i = 1 To grdRetenciones.Rows - 1
     'Compara el codigo del impuesto para averiguar si existe
        If grdImp.TextMatrix(grdImp.RowSel, 0) = grdRetenciones.TextMatrix(i, 1) Then
            fbEstaImpuestoRetenido = True
            Exit Function
        End If
    Next
End If
End Function

Private Function fsCalcularMontoRetenido() As String
'------------------------------------------------------------------------
'Proposito : Calcula el monto retenido de acuerdo al valor del impuesto _
'             y el valor del Egreso CA
'Recibe: nada
'Entrega : string que representa al monto calculado de la retención _
'          con formato numerico
'------------------------------------------------------------------------
Dim dMonto As Double
'inicializa la función
fsCalcularMontoRetenido = "0.00"
dMonto = 0
'calcula el monto retenido
dMonto = CDbl(Val(EliminarComas(frmCBEGConAfecta.txtTotalDoc.Text))) _
         * CDbl(EliminarComas(grdImp.TextMatrix(grdImp.RowSel, 2))) / 100
'da formato string al monto
fsCalcularMontoRetenido = Format(dMonto, "###,###,##0.00")

End Function

Private Sub cmdCancelar_Click()

'Cierra el formulario
Unload Me

End Sub

Private Sub cmdEliminar_Click()

'Elimina las filas marcadas en grd de Impuestos retenidos
EliminarFilaMarcadaGRID grdRetenciones

'Habilitar botón aceptarmodificar
HabilitaBotonAceptarModificar

End Sub

Private Sub cmdSalir_Click()
Select Case msOperacion

Case "Nuevo"
'Si la operación en el formulario es: nuevas retenciones
    Unload Me
Case "Modificar"
'Si la operación es modificación de retenciones
    Unload Me
End Select

End Sub

Private Sub Form_Load()
Dim sSQL As String

'Se construye la sentencia (si la fecha de baja es <> '' no se muestra)
sSQL = "SELECT idimp, descimp , valorimp FROM TIPO_IMPUESTOS"

' Se carga un array con los títulos de las columnas y otro con los tamaños para
' Pasárselos a la función que carga el grid que muestra los impuestos a elegir
aTitulosColGrid = Array("Cód.", "Descripción", "Valor")
aTamañosColumnas = Array(600, 3500, 500)

CargarGrid grdImp, sSQL, aTitulosColGrid, aTamañosColumnas

If grdImp.Rows = 1 Then
    MsgBox "No existen Impuestos", _
          vbInformation + vbOKOnly, "S.G.Ccaijo"

End If

' Se carga un array con los títulos de las columnas y otro con los tamaños para
'pasárselos a la función que carga el grid que muestra los impuestos a elegir
aTitulosColGrid = Array("Orden", "IdImpuesto", "Descripción", "% Impuesto", "Monto Retenido")
aTamañosColumnas = Array(1200, 0, 3500, 950, 1500)

CargarGridTitulos grdRetenciones, aTitulosColGrid, aTamañosColumnas

' Deshabilita botón Aceptar
cmdAceptarModificar.Enabled = False
' Deshabilita botón añadir
cmdAñadir.Enabled = False

'Verifica el tipo de operacion ha realizar en el formulario
If gsTipoOperacionEgreso = "Nuevo" Then
    'El tipo de operación ha realizar en el formulario es retencion _
    por pago de Impuestos
    msOperacion = "Nuevo"
    
Else
    'El tipo de operación ha realizar en el formulario es modificar _
    la retencion de Impuestos
    msOperacion = "Modificar"
    'Carga las retenciones aplicadas al egreso CA, de MovImpuestos
    CargarRetencionesAplicadas

End If

End Sub

Private Sub CargarRetencionesAplicadas()
'------------------------------------------------------------------------
'Proposito : Carga las retenciones aplicadas al egreso para porder modificar
'Recibe: Nada
'Entrega : Nada
'------------------------------------------------------------------------
Dim sSQL As String
Dim curRetenciones As New clsBDConsulta

' Carga la sentencia select que averigua las retenciones que causo el egreso _
 con Afectación
sSQL = "SELECT I.IdImp,I.DescImp,M.ValorImp,M.Monto " _
    & "FROM TIPO_IMPUESTOS I, MOV_IMPUESTOS M " _
    & "WHERE I.IdImp=M.IdImp and Orden='" & gsOrden & "'"
' Ejecuta la consulta
curRetenciones.SQL = sSQL
If curRetenciones.Abrir = HAY_ERROR Then End

' Verifica si el cursor es vacío
If curRetenciones.EOF Then
    ' No tiene retenciones
    curRetenciones.Cerrar
    Exit Sub
Else
    ' Tiene retenciones, carga la colección de retenciones
    ' Carga el grid retenciones
    Do While Not curRetenciones.EOF
        
        'Añade una fila a grd Orden,IdImp,DescImp,ValorImp,Monto
        grdRetenciones.AddItem (gsOrden & vbTab _
        & curRetenciones.campo(0) & vbTab _
        & curRetenciones.campo(1) & vbTab _
        & curRetenciones.campo(2) & vbTab _
        & Format(curRetenciones.campo(3), "###,###,#0.00"))
        
        'Añade un item a la colección IdImp,ValorImp,Monto
        mcolRetenciones.Add _
        Item:=curRetenciones.campo(0) & " " _
            & curRetenciones.campo(2) & " " _
            & Format(curRetenciones.campo(3), "###,###,#0.00"), _
        Key:=curRetenciones.campo(0)
        
        'Mueve al siguiente registro del cursor
        curRetenciones.MoverSiguiente
    Loop
End If
'Cierra el cursor de la consulta
curRetenciones.Cerrar

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Elimina las colecciones usadas en el formulario
  Set mcolRetenciones = Nothing

End Sub

Private Sub grdImp_Click()



'Si se pincha el grid y está vacío no hace nada
If grdImp.Rows = 1 Then
  Exit Sub
End If

'Marca o Desmarca solo una fila en el grd
MarcarUnaFilaGrid grdImp

'Habilitar el boton añadir
HabilitaBotonAñadir


End Sub

Private Sub HabilitaBotonAñadir()
'------------------------------------------------------------------------
'Proposito : Habilita el botón añadir si se ha elegido alguna fila en grdImpuestos
'Recibe: Nada
'Entrega : Nada
'------------------------------------------------------------------------
Dim i As Integer
'Inicializa el botona añadir deshabilitado
 cmdAñadir.Enabled = False
 'recorre el grd Impuesto para saber si alguna fila de grd esta en azul
For i = 1 To grdImp.Rows - 1
 'Verifica si hay alguna fila seleccionada
    If grdImp.CellBackColor = vbDarkBlue Then 'hay una fila seleccionada
          cmdAñadir.Enabled = True
          Exit Sub
    End If
Next

End Sub

Private Sub grdRetenciones_Click()
'Si se pincha el grid y está vacío no hace nada
If grdRetenciones.Rows = 1 Then
  Exit Sub
End If

'Marca o Desmarca solo una fila en el grd
MarcarDesmarcarFilaGRID grdRetenciones

End Sub
