VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmALSalida 
   Caption         =   "Almacén- Salidas de Almacén"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   HelpContextID   =   89
   Icon            =   "SCALSalida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   6720
      Width           =   1000
   End
   Begin VB.CommandButton cmdPProyecto 
      Height          =   255
      Left            =   6870
      Picture         =   "SCALSalida.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   780
      Width           =   220
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   6720
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptarModificar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   6720
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   6720
      Width           =   1000
   End
   Begin VB.ComboBox cboProy 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   750
      Width           =   5200
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   80
      TabIndex        =   19
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtPersonal 
         Height          =   315
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtProy 
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   2
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox txtIdSalida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1875
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4730
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   6600
         Picture         =   "SCALSalida.frx":0BA2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   490
      End
      Begin MSMask.MaskEdBox mskFecha 
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
      Begin VB.Label Label4 
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre:"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Proyecto:"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº Salida:"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   705
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetSalida 
      Height          =   3315
      Left            =   195
      TabIndex        =   17
      Top             =   3195
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   80
      TabIndex        =   24
      Top             =   1680
      Width           =   7335
      Begin VB.CommandButton cmdPProducto 
         Height          =   255
         Left            =   6840
         Picture         =   "SCALSalida.frx":0CA4
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Width           =   220
      End
      Begin VB.ComboBox cboProducto 
         Height          =   315
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Top             =   225
         Width           =   5895
      End
      Begin VB.TextBox txtCantidad 
         Height          =   315
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   10
         Top             =   675
         Width           =   975
      End
      Begin VB.TextBox txtMedida 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2235
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtStock 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   675
         Width           =   1335
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "A&ñadir"
         Height          =   375
         Left            =   5060
         TabIndex        =   13
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label Label5 
         Caption         =   "P&roducto:"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "&Cantidad :"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Stock del producto:"
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   690
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmALSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Colecciones para la carga del combo de Proyectos
Private mcolCodProy As New Collection
Private mcolCodDesProy As New Collection

Private mcolDesMedidaProd  As New Collection

Private mcolidprod As New Collection
Private mcolCodDesProd As New Collection

Private mcolCodPersonal As New Collection
Private mcolCodDesPersonal As New Collection

' Colección para guardar las salidas de los ingresos con saldos disponibles
Private mcolSalidasDisponible As New Collection

'Colección que almacena los productos actuales a modificar
Private mcolSalidaModificar As New Collection

Private mcurRegSalidaAlmacen As New clsBD2
Private mcurDetSalidaAlmacen As New clsBD2
Private mcurMedidaProd As New clsBD2
Private mcurRepartir As New clsBD2

'Cursor que almacena el total de las salidas de una linea de producto posterior
Private mcurSalidasPost As New clsBD2


Private msIdprod As String
Private msIdProdElim As String
Private miDiferencia As Integer
Private msSigno As String

'Variable que identifica el tipo de operacion de los botones
Private msOperacion As String

'Variable que identifica la modificacion de almacen y detalle de almacen
Private mbModificaAlmacen As Boolean
Private mbModificaDetAlmacen As Boolean

Private mcurCtasProd As New clsBD2

'Variable para el manejo del grid
Dim ipos As Long

Dim mcurProyectos As New clsBD2

Private Sub cboProducto_Change()

' verifica si lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProducto) = True Then SendKeys "{down}"

End Sub

Private Sub cboProducto_Click()

' Verifica si el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProducto.ListIndex) = False And cboProducto.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboProducto_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica si es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboProducto_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboProducto, Obligatorio) = True Then

   ' Se actualiza código (TextBox) correspondiente a descripción introducida
   'CD_ActCboVar cboProducto.Text, msIdprod, mcolidprod, mcolCodDesProd
   ActualizarInfoProd cboProducto.Text, msIdprod, mcolidprod, mcolCodDesProd
   
   'Actualiza la medida correspondiente al producto seleccionada
   ActualizaCampoMedida
   
   'Carga el stock del producto en Almacén
   CargarStock
   
Else

   'No se encuentra el producto
   msIdprod = ""
   cboProducto.BackColor = Obligatorio
   txtMedida.Text = Empty
   txtStock.Text = Empty
End If

'Cambia el alto del combo
cboProducto.Height = CBONORMAL

'Habilita boton añadir
HabilitarBotonAñadir

End Sub

Private Sub ActualizaCampoMedida()

' Actualiza la medida del producto
If cboProducto.Text <> "" Then
  cboProducto.BackColor = vbWhite
  
  'Muestra la medida del producto
  MostrarMedida (msIdprod)

Else
  cboProducto.BackColor = Obligatorio

End If

End Sub

Private Sub MostrarMedida(sProducto As String)

'Muestra la medida del producto seleccionado en el combo
txtMedida.Text = mcolDesMedidaProd.Item(Trim(sProducto))

End Sub

Private Sub cboProy_Change()

' verifica si lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboProy) = True Then SendKeys "{down}"

End Sub

Private Sub cboProy_Click()

' Verifica si el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboProy.ListIndex) = False And cboProy.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub

Private Sub cboProy_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica si es enter para salir o flechas para recorrer
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

Private Sub cmdAceptarModificar_Click()
'Verifica si el año esta cerrado
If Conta52(Right(mskFecha.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Salidas de almacén"
    Exit Sub
End If

' Se guarda los nuevos datos
If msOperacion = "Nuevo" Then
    'Verifica que la salida no sea anterior al ingreso
    If SalidaPosterior Then
        'Mensaje de conformidad de los datos
        If MsgBox("¿Esta conforme con los datos ?", _
                      vbQuestion + vbYesNo, "Almacén- Salida de almacén") = vbYes Then
            'Actualiza la transaccion
             Var8 1, gsFormulario
    
            
            GuardarSalidaAlmacen
            
            'Guarda el detalle en Almacen_Sal_Det
            GuardaDetSalidaAlmacen
            
            'Limpia y pregunta si desea hacer una nueva salida
            cmdCancelar_Click
            
            'Actualiza la transaccion
             Var8 -1, Empty
           
            MsgBox "Operación efectuada correctamente", vbInformation, _
               "S.G.CCAIJO - Salidas de Almacén"
               
        End If
    Else
        'Mensaje  de salida
        MsgBox "Fecha de salida anterior al ultimo ingreso", vbInformation + vbOKOnly, "Salidas almacén"
    End If

ElseIf msOperacion = "Modificar" Then 'Se guardan las modificaciones

    If MsgBox("¿Seguro que desea modificar la salida de almacén ?", _
                      vbQuestion + vbYesNo, "Almacén- Salida de almacén") = vbYes Then
        'Actualiza la transaccion
         Var8 1, gsFormulario
        
        If mbModificaAlmacen And mbModificaDetAlmacen Then
        
            ' Se modifica la salida en SALIDA_ALMACEN
            ModificarSalidaAlmacen
            
            ' Se modifica el detalle en Almacen_Sal_Det
            ModificarDetSalidaAlmacen
            
        Else
        
           If mbModificaAlmacen Then
           
               ' Se modifica la salida en SALIDA_ALMACEN
                ModificarSalidaAlmacen
                
           Else
           
               ' Se modifica el detalle en Almacen_Sal_Det
                ModificarDetSalidaAlmacen
                
           End If
           
        End If
    
        'Actualiza la transaccion
         Var8 -1, Empty
        
        MsgBox "Modificación efectuada correctamente", vbInformation, _
           "S.G.CCAIJO - Modificación de Salidas de Almacén"

        'Descarga el formulario
        Unload Me
        
    Else
    
        'Ubica el cursor en el proyecto
        txtProy.SetFocus
        
    End If
 
End If
 
End Sub

Private Function SalidaPosterior() As Boolean
Dim sSQL As String
Dim curUltimoIng As New clsBD2

'Sentencia que determina el ultimo ingreso
sSQL = ""
sSQL = "SELECT MAX(A.Fecha)" & _
       "FROM ALMACEN_INGRESOS A"

'Copia la sentencia SQL
curUltimoIng.SQL = sSQL

'Verifica si hay error
If curUltimoIng.Abrir = HAY_ERROR Then
    'Termina la ejecución
    End
End If

'Compara las fechas
If FechaAMD(mskFecha) < curUltimoIng.campo(0) Then
    'Error
    SalidaPosterior = False
Else
    'La salida es correcta
    SalidaPosterior = True
End If

'Cierra el cursor
curUltimoIng.Cerrar

End Function

Private Sub GuardaDetSalidaAlmacen()

Dim i As Integer
Dim dblMontoSalida As Double
'Recorre el detalle de la salida
For i = 1 To grdDetSalida.Rows - 1

    'Cargar Colección salida línea mercadería
    CargaColeccionSalidasDisponible grdDetSalida.TextMatrix(i, 3), Val(grdDetSalida.TextMatrix(i, 2))

    GuardaSalidasDisponible grdDetSalida.TextMatrix(i, 3), dblMontoSalida, txtIdSalida
    
    DeterminaCtas grdDetSalida.TextMatrix(i, 3)
    
    'Generar Asiento salida por la salida del producto
    GenerarAsientoSalidaLineaProducto txtIdSalida.Text, grdDetSalida.TextMatrix(i, 3), dblMontoSalida, "Registrar"
    
    'Limpia la colección mcolSalidaDisponibles
    Set mcolSalidasDisponible = Nothing
    mcurCtasProd.Cerrar
Next

End Sub

Private Sub GuardaSalidasDisponible(sIdProd As String, dblMontoSalida As Double, sIdSalida As String)
Dim varObjeto As Variant

' Inicia el monto de la salida
dblMontoSalida = 0

For Each varObjeto In mcolSalidasDisponible

    ' Guarda en la BD la salida y actualiza disponibles del ingreso
    GuardaSalidaDisponibleBD Var30(varObjeto, 1), _
                             sIdSalida, _
                             sIdProd, _
                             Var30(varObjeto, 2), _
                             Var30(varObjeto, 3)
    ' Actualiza disponibles ingreso
    DisminuirDisponiblesBD Var30(varObjeto, 1), _
                           sIdProd, _
                           Var30(varObjeto, 2), _
                           Var30(varObjeto, 3), _
                           Var30(varObjeto, 4), _
                           Var30(varObjeto, 5)
    
    'fredi
    ' Acumula al monto total de la salida
    dblMontoSalida = Round(Val(dblMontoSalida) + Val(Var30(varObjeto, 3)), 2)
        
Next varObjeto

End Sub

Private Sub IncrementarDisponiblesBD(Orden As String, IdProd As String, _
                                   Cantidad As Double, Monto As Double, _
                                   CantDisponible As Double, Resto As Double)
Dim modIngreso As New clsBD3
Dim sSQL As String

' Carga la sentencia para guardar la salida de producto
sSQL = "UPDATE Almacen_Ingresos SET " _
        & " Resto=" & Format(Monto + Resto, "##0.00") & " , " _
        & " CantidadDisponible=" & Format(CantDisponible + Cantidad, "##0.00") _
        & " WHERE Orden='" & Orden & "' and IdProd='" & IdProd & "'"
      
' Ejecuta la sentencia
modIngreso.SQL = sSQL
If modIngreso.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modIngreso.Cerrar

End Sub

Private Sub DisminuirDisponiblesBD(Orden As String, IdProd As String, _
                                   Cantidad As Double, Monto As Double, _
                                   CantDisponible As Double, Resto As Double)
Dim modIngreso As New clsBD3
Dim sSQL As String

' Carga la sentencia para guardar la salida de producto
sSQL = "UPDATE Almacen_Ingresos SET " _
        & " Resto=" & Format(Resto - Monto, "##0.00") & " ," _
        & " CantidadDisponible=" & Format(CantDisponible - Cantidad, "##0.00") _
        & " WHERE Orden='" & Orden & "' and IdProd='" & IdProd & "'"
      
' Ejecuta la sentencia
modIngreso.SQL = sSQL
If modIngreso.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modIngreso.Cerrar

End Sub

Private Sub GuardaSalidaDisponibleBD(Orden As String, IdSalida As String, _
                                     IdProd As String, Cantidad As String, _
                                     Monto As String)
Dim modSalidaDet As New clsBD3
Dim sSQL As String

' Carga la sentencia para guardar la salida de producto
sSQL = "INSERT INTO Almacen_Sal_Det " _
        & "VALUES('" & Orden & "','" & IdSalida & "','" _
        & IdProd & "'," & Cantidad & "," & Monto & ")"
          
' Ejecuta la sentencia
modSalidaDet.SQL = sSQL
If modSalidaDet.Ejecutar = HAY_ERROR Then End

' Cierra la componente
modSalidaDet.Cerrar

End Sub

Private Sub CargaColeccionSalidasDisponible(sIdProd As String, ByVal dblCantidad As Double)
Dim sSQL As String
Dim curRegDisponible As New clsBD2
Dim dblMonto As Double

'Sentencia SQL
sSQL = "SELECT A.Orden, A.PrecioUnit, A.Resto, A.CantidadDisponible " & _
        "FROM ALMACEN_INGRESOS A " _
        & "WHERE A.CantidadDisponible > 0 and '" & sIdProd & "'=A.IdProd " _
        & "ORDER BY A.NroIngreso"
        
'Se ejecuta la sentencia SQL
curRegDisponible.SQL = sSQL

'Abre el cursor
If curRegDisponible.Abrir = HAY_ERROR Then End
   
'Reparte sa salida total entre las salidas de los ingresos disponibles
Do While Val(dblCantidad) > 0
    'verifica si la cantidad de salida es Menor Igual a la cantidad disponible del ingreso
    If Val(dblCantidad) >= curRegDisponible.campo(3) Then
        
        'Añade a la colecion de registros de ingreso afectados por la salida
        'Orden, CantidadSalida, MontoTotalSalida,CantidadDisponible,MontoDisponible
        mcolSalidasDisponible.Add Item:=curRegDisponible.campo(0) & "¯" & _
                                         curRegDisponible.campo(3) & "¯" & _
                                         curRegDisponible.campo(2) & "¯" & _
                                         curRegDisponible.campo(3) & "¯" & _
                                         curRegDisponible.campo(2), _
                                         Key:=curRegDisponible.campo(0)
        'Actualiza la cantidad de salida
        dblCantidad = Round(dblCantidad - curRegDisponible.campo(3), 2)
        
    Else 'La cantidad de salida es Menor a disponible
                                       
        'Añade a la colecion de registros de ingreso afectados por la salida
        'Orden, CantidadSalida, MontoTotalSalida,CantidadDisponible,MontoDisponible
        mcolSalidasDisponible.Add Item:=curRegDisponible.campo(0) & "¯" & _
                                         dblCantidad & "¯" & _
                                         Var7(dblCantidad * curRegDisponible.campo(1), 2) & "¯" & _
                                         curRegDisponible.campo(3) & "¯" & _
                                         curRegDisponible.campo(2), _
                                         Key:=curRegDisponible.campo(0)
        'Actualiza el acumulador
        dblCantidad = 0
        
     End If
     
     'Mueve el cursor al siguiente registro
     curRegDisponible.MoverSiguiente
Loop

'Cierra el cursor curRegDisponible
curRegDisponible.Cerrar
   
End Sub

Private Sub GenerarAsientoSalidaLineaProducto(sIdSalida As String, _
                                              sIdProd As String, _
                                              dblMonto As Double, _
                                              sTipoAsiento As String)
     gcolAsientoDet.Add _
     Key:=sIdProd, _
     Item:=sIdProd & "¯" & "EA¯" _
         & mcurCtasProd.campo(0) & "¯" _
         & mcurCtasProd.campo(1) & "¯" _
         & Format(dblMonto, "##0.00")
    
    'Colección que guarda los datos generales para los asientos contable
    
    gcolAsiento.Add _
        Key:=sIdSalida, _
        Item:=sIdSalida & "¯" & FechaAMD(mskFecha.Text) & "¯" _
            & sIdProd & "¯SALIDA DE ALMACEN¯" & sTipoAsiento & "¯SA"

     'Realiza el asiento automatico de ingreso a almacén
    Conta45

End Sub

Private Sub DeterminaCtas(sIdProdRec As String)
Dim sSQL As String

' Carga la sentencia que consulta a la BD acerca del los prods aha verificar
    sSQL = ""
    sSQL = "SELECT  P.CodSuministro, P.CodVariacion " & _
           "FROM PRODUCTOS P WHERE " & _
           "P.IdProd='" & sIdProdRec & "' "
           
           
mcurCtasProd.SQL = sSQL

' Abre el cursor si hay  error sale indicando la causa del error
If mcurCtasProd.Abrir = HAY_ERROR Then
    End
End If

If mcurCtasProd.EOF Then

  'No tiene Cuenta Contable, Auxiliar o Monto para el ingreso a almacén del producto
  MsgBox "Error Integridad BD, no Existen cuenta contable de suministro, variaciòn o Monto para el ingreso a almacén del producto. " & Chr(13) & _
         "Consulte al Administrador ", vbInformation + vbOKOnly, "S.G.Ccaijo-Ingreso a Almacen"
  'Vacia las colecciones
  Set gcolAsiento = Nothing
  Set gcolAsientoDet = Nothing
  
  Exit Sub
End If

End Sub


Private Sub ModificarSalidaAlmacen()

Dim sSQL As String
Dim modSalidaAlmacen As New clsBD3

' Guardar los  datos
sSQL = "UPDATE Almacen_Salidas SET " & _
        "IdSalida='" & txtIdSalida.Text & "'," & _
        "IdProy='" & txtProy.Text & "', " & _
        "IdPersona='" & txtPersonal.Text & "' " & _
        "WHERE IdSalida='" & txtIdSalida.Text & "'"
           
'Si al ejecutar hay error se sale de la aplicación
modSalidaAlmacen.SQL = sSQL
If modSalidaAlmacen.Ejecutar = HAY_ERROR Then
 End
End If

'Se cierra el componente de mod
modSalidaAlmacen.Cerrar

End Sub


Private Sub GuardarSalidaAlmacen()
Dim modAlSalida As New clsBD3
Dim sSQL As String

'Propósito Guarda los datos generales de la salida de Almacén
sSQL = "INSERT INTO Almacen_Salidas VALUES('" & txtIdSalida.Text & "','" & txtProy.Text _
        & "','" & txtPersonal.Text & "','" & FechaAMD(mskFecha.Text) & "','NO')"

'Si al ejecutar hay error se sale de la aplicación
modAlSalida.SQL = sSQL

'Verifica si hay error
If modAlSalida.Ejecutar = HAY_ERROR Then
  End
End If

'Se cierra la query
modAlSalida.Cerrar

End Sub

Private Sub cmdAnular_Click()
'Verifica si el año esta cerrado
If Conta52(Right(mskFecha.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Salidas de almacén"
    Exit Sub
End If

 'Mensaje de conformidad de los datos
 If MsgBox("¿Seguro que desea anular la salida de almacén ?", _
                      vbQuestion + vbYesNo, "Almacén- Salida de almacén") = vbYes Then
    'Actualiza la transaccion
     Var8 1, gsFormulario
    
    'Anular el detalle de la salida
    AnularDetSalida
    
    'Anula el registro general de la salida
    AnularSalida
    
    'Actualiza la transaccion
     Var8 -1, Empty
    
    MsgBox "Modificación efectuada correctamente", vbInformation, _
       "S.G.CCAIJO - Modificación de Salidas de Almacén"

    'Descarga el formulario
    Unload Me
        
Else

    'Ubica el cursor en el proyecto
    txtProy.SetFocus
    
End If
End Sub

Private Sub AnularDetSalida()
Dim Objeto As Variant

For Each Objeto In mcolSalidaModificar
    'Realiza la modificación de la salida del producto
    ModificarSalidaProducto Var30(Objeto, 1), Var30(Objeto, 2), "Anular"
Next Objeto

End Sub


Private Sub AnularSalida()
Dim sSQL As String
Dim modAnularSalida As New clsBD3

'Modifica el campo Anulado a "Si" para la salida seleccionada msSalida
 sSQL = "UPDATE Almacen_Salidas SET " & _
        "Anulado='SI'" & _
        "WHERE IdSalida= '" & txtIdSalida & "'"

'Ejecuta la sentencia
modAnularSalida.SQL = sSQL

'Si al ejecutar hay error se sale de la aplicación
If modAnularSalida.Ejecutar = HAY_ERROR Then
  End
End If

'Se cierra el cursor modAnularAlSalida
modAnularSalida.Cerrar

End Sub

Private Sub cmdAñadir_Click()
Dim dblVariable As Double

If msOperacion = "Modificar" Then
    
   'Verifica si el producto se encuentra en el detalle
   If fbEstaProdenDet = False Then

        'Determina si se ingreso un nuevo producto
        If DeterminarNuevoProducto Then
            
            'Verifica que cantidad sea Menor o Igual a stock
            If Val(Var37(txtCantidad.Text)) > Val(Var37(txtStock.Text)) Then
                MsgBox "La cantidad introducida es Mayor al stock disponible", _
                         vbInformation, "Almacén- Salida de Almacén"
                txtCantidad.SetFocus
                Exit Sub
            End If
            
                          
            'Añade el nuevo registro al grid
            grdDetSalida.AddItem (cboProducto.Text & vbTab & txtMedida & vbTab & txtCantidad.Text _
                            & vbTab & msIdprod)

                 
        Else 'Si DeterminarNuevoProducto es falso
            'IdProd, Cantidad
            'Verifica que cantidad sea Menor o Igual a stock
            If Val(Var37(txtCantidad.Text)) - Var30(mcolSalidaModificar.Item(msIdprod), 2) > Val(Var37(txtStock.Text)) Then
                MsgBox "La cantidad introducida es Mayor al stock disponible", _
                        vbInformation, "Almacén- Salida de Almacén"
                txtCantidad.SetFocus
                Exit Sub
            End If
            
                     
                'Añade el registro al grid sin modificaciones
                grdDetSalida.AddItem (cboProducto.Text & vbTab & txtMedida & vbTab & txtCantidad.Text _
                                & vbTab & msIdprod)
                                
        End If ' Fin del DeterminarNuevoProducto
        
    Else
       ' Envia mensaje
       MsgBox "Este producto ya ha sido ingresado al detalle de Salida de almacén", _
               vbInformation + vbOKOnly, "Almacén- Salida de Almacén"
       
    End If 'Fin del fbEstaProdenDet
'

ElseIf msOperacion = "Nuevo" Then
    
    'Verifica que cantidad sea Menor o Igual a stock
    If Val(Var37(txtCantidad.Text)) > Val(Var37(txtStock.Text)) Then
        MsgBox "La cantidad introducida es Mayor al stock disponible", _
                 vbInformation, "Almacén- Salida de Almacén"
        txtCantidad.SetFocus
        Exit Sub
    End If
    
    'Inicializa variable
'     dblVariable = 0
     
    'Verifica si el producto se encuentra en el detalle
     If fbEstaProdenDet = False Then
'         ' Calcula el monto de salida
         'dblVariable = CalcularMonto(Val(Var37(txtCantidad.Text)))
'
'         If dblVariable = 0 Then
'              MsgBox "Error: El monto de salida es cero, Debe revisar los precios de ingresos y salidas de almacén", _
'                     vbInformation + vbOKOnly, "SGCcaijo-Verificación de precios de Salida"
'
'            'Termina la ejecución
'             Unload Me
'
'         Else
           
             'Añade el nuevo registro al grid
             grdDetSalida.AddItem (cboProducto.Text & vbTab & txtMedida & vbTab & txtCantidad.Text _
                             & vbTab & msIdprod)
'         End If
     
     Else
         ' Envia mensaje
         MsgBox "Esta producto ya ha sido ingresado al detalle de salida de almacén", _
                    vbInformation + vbOKOnly, "ALmacén- Salida de Almacén"
                    
     End If
   
End If

 'Habilita el boton aceptar
HabilitarAceptarModificarSalida
   
'Limpia los campos
LimpiarCamposDetalle

' Limpia la cuenta para  dar opcion a elegir
cboProducto.SetFocus
cmdAñadir.Enabled = False


End Sub

Function DeterminarNuevoProducto() As Boolean
'-------------------------------------------------------------
'Propósito : Determina si el producto que se ingreso es nuevo
'Recibe    : Nada
'Devuelve  : Verdad, Si es nuevo; falso, si este producto se esta modificando su cantidad
'-------------------------------------------------------------
On Error GoTo mnjError

'El producto ingresado no es nuevo
DeterminarNuevoProducto = False
mcolSalidaModificar.Item (msIdprod)
Exit Function
mnjError:
'Verifica si es error, codigo no encontrado en la colección
If Err.Number = 5 Then
    'El producto ingresado no es nuevo
    DeterminarNuevoProducto = True
End If

End Function

Function DeterminarTotalDisponible() As Double
Dim sSQL As String
Dim curTotalDisponible As New clsBD2

'Sentencia SQL para determinar el total de disponible en almacén
sSQL = ""
sSQL = "SELECT SUM(CantidadDisponible) as Total FROM ALMACEN_INGRESOS  " & _
       "WHERE IdProd= '" & msIdprod & "' "

curTotalDisponible.SQL = sSQL

'Abre el cursor
If curTotalDisponible.Abrir = HAY_ERROR Then End

'Verifica si el cursor es vacio
If curTotalDisponible.EOF Then
   MsgBox "No existe ninguna salida para este producto", vbInformation, "Almacén- Salida de almacén"
   DeterminarTotalDisponible = 0
   curTotalDisponible.Cerrar
   Exit Function
Else
    If IsNull(curTotalDisponible.campo(0)) Then
        DeterminarTotalDisponible = 0
    Else
        'Devuelve el total de salidas de almacen para el producto seleccionado
        DeterminarTotalDisponible = curTotalDisponible.campo(0)
   End If
End If

'Cierra el cursor
curTotalDisponible.Cerrar

End Function

Function DeterminarTotalSalidas() As Double

Dim sSQL As String
Dim curTotalDetalle As New clsBD2

'Sentencia SQL para determinar el total de las salidas de almacén
sSQL = ""
sSQL = "SELECT SUM(A.Cantidad) as Total FROM Almacen_Sal_Det A " & _
       "WHERE '" & txtIdSalida.Text & "' <= A.IdSalida and " & _
       "IdProd= '" & msIdprod & "' And " & _
       "A.IdSalida=( SELECT S.IdSalida FROM ALMACEN_SALIDAS S WHERE " & _
       "A.IdSalida=S.IdSalida And S.Anulado='NO')"
       
curTotalDetalle.SQL = sSQL

'Abre el cursor
If curTotalDetalle.Abrir = HAY_ERROR Then End

'Verifica si el cursor es vacio
If curTotalDetalle.EOF Then
   MsgBox "No existe ninguna salida para este producto", vbInformation, "Almacén- Salida de almacén"
   DeterminarTotalSalidas = 0
   curTotalDetalle.Cerrar
   Exit Function
Else
   If IsNull(curTotalDetalle.campo(0)) Then
        DeterminarTotalSalidas = 0
   Else
        'Devuelve el total de salidas de almacen para el producto seleccionado
        DeterminarTotalSalidas = curTotalDetalle.campo(0)
   End If
End If

'Cierra el cursor
curTotalDetalle.Cerrar

End Function

Private Sub CalcularMonto(srtIdProd As String, dblCantidadRec As Double)

Dim sSQL As String
Dim curRegDisponible As New clsBD2
Dim dblMonto As Double

'Inicializa la variable dblMonto
dblMonto = 0
'Sentencia SQL
sSQL = ""
sSQL = "SELECT A.Orden, A.IdProd, A.PrecioUnit, A.Resto, A.CantidadDisponible " & _
        "FROM ALMACEN_INGRESOS A " _
        & "WHERE A.CantidadDisponible > 0 and '" & srtIdProd & "'=A.IdProd " _
        & "ORDER BY A.NroIngreso"
        
'Se ejecuta la sentencia SQL
curRegDisponible.SQL = sSQL

'Abre el cursor
If curRegDisponible.Abrir = HAY_ERROR Then End

'Verifica si el cursor es vacio
If curRegDisponible.EOF Then
  MsgBox "Egreso sin gastos.Error en la base de datos, consulte a su administrador"
  Exit Sub
  
Else

    'Halla el dblMonto de la salida de Almacén
   If dblCantidadRec < curRegDisponible.campo(4) Then
   
            dblMonto = dblCantidadRec * curRegDisponible.campo(2)
            
            'Añade a la colecion de registros afectados por la salida
            'Orden, IdProd, PrecioTotal, Cantidad
'            mcolCodCantidadAlmacen.Add Item:=curRegDisponible.campo(0) & "¯" & _
'                                             curRegDisponible.campo(1) & "¯" & _
'                                             dblMonto & "¯" & _
'                                             dblCantidadRec, _
'                                             Key:=curRegDisponible.campo(0) & curRegDisponible.campo(1)
            'GuardarBD
        
   'Verifica si es Igual a la cantidad del curRegDisponible.campo(4)
    ElseIf dblCantidadRec = curRegDisponible.campo(4) Then
           dblMonto = dblCantidadRec * curRegDisponible.campo(2) + curRegDisponible.campo(3)
           
           'Añade a la colecion de registros afectados por la salida
           'Orden, IdProd, PrecioU, Cantidad
'           mcolCodCantidadAlmacen.Add Item:=curRegDisponible.campo(0) & "¯" & _
'                                            curRegDisponible.campo(1) & "¯" & _
'                                            dblMonto & "¯" & _
'                                            dblCantidadRec, _
'                                            Key:=curRegDisponible.campo(0) & curRegDisponible.campo(1)
        'GuardaBD
   Else
   
        'Si dblCantidadRec afecta a mas de un registro de ingreso a Almacén
        Do While dblCantidadRec >= curRegDisponible.campo(4)
            dblMonto = dblMonto + curRegDisponible.campo(4) * curRegDisponible.campo(2) + curRegDisponible.campo(3)
            
            dblCantidadRec = dblCantidadRec - curRegDisponible.campo(4)
            
            'Añade a la colecion de registros afectados por la salida
            'Orden, IdProd, MontoTotal, Cantidad
'            mcolCodCantidadAlmacen.Add Item:=curRegDisponible.campo(0) & "¯" & _
'                                             curRegDisponible.campo(1) & "¯" & _
'                                             curRegDisponible.campo(4) * curRegDisponible.campo(2) + curRegDisponible.campo(3) & "¯" & _
'                                             curRegDisponible.campo(4), _
'                                             Key:=curRegDisponible.campo(0) & curRegDisponible.campo(1)
'
            'curRegDisponible.campo(4) * curRegDisponible.campo(2) + curRegDisponible.campo(3) =Monto
            'GuardaBD
            'Mueve el cursor al siguiente registro
            curRegDisponible.MoverSiguiente
        Loop
        
        If dblCantidadRec > 0 Then
        
            If dblCantidadRec = curRegDisponible.campo(4) Then
            
                'Calcula el dblMonto total de la salida de Almacén
                dblMonto = dblMonto + dblCantidadRec * curRegDisponible.campo(2) + curRegDisponible.campo(3)
                
                'Añade a la colecion de registros afectados por la salida
                'Orden, IdProd, PrecioU, Cantidad
'                mcolCodCantidadAlmacen.Add Item:=curRegDisponible.campo(0) & "¯" & _
'                                                 curRegDisponible.campo(1) & "¯" & _
'                                                 curRegDisponible.campo(4) * curRegDisponible.campo(2) + curRegDisponible.campo(3) & "¯" & _
'                                                 dblCantidadRec, _
'                                                 Key:=curRegDisponible.campo(0) & curRegDisponible.campo(1)
                'GuardaBD
            Else
            
                'Calcula el dblMonto total de la salida de Almacén
                dblMonto = dblMonto + dblCantidadRec * curRegDisponible.campo(2)
                
                'Añade a la colecion de registros afectados por la salida
                'Orden, IdProd, PrecioU, Cantidad
'                mcolCodCantidadAlmacen.Add Item:=curRegDisponible.campo(0) & "¯" & _
'                                                 curRegDisponible.campo(1) & "¯" & _
'                                                 dblCantidadRec * curRegDisponible.campo(2) & "¯" & _
'                                                 dblCantidadRec, _
'                                                 Key:=curRegDisponible.campo(0) & curRegDisponible.campo(1)
                'GuardarBD
            End If
            
       End If
   End If
   
End If
    
'Se da formato a dblMonto y devuelve el valor de la funcion
'CalcularMonto = Format(dblMonto, "###,###,##0.0000")

'Cierra el frmModificar
curRegDisponible.Cerrar

End Sub

Private Sub LimpiarCamposDetalle()
'-------------------------------------------
'Propósito: Limpia los campos de la pantalla 2 (Detalle)
'-------------------------------------------
cboProducto.ListIndex = -1
cboProducto.BackColor = Obligatorio
txtMedida.Text = Empty
txtCantidad.Text = Empty
txtCantidad.BackColor = Obligatorio
txtStock.Text = Empty

End Sub
Private Function fbEstaProdenDet() As Boolean

Dim j As Integer

'Inicializamos a funcion asumiendo que Procuto no esta en el grdDetSalida
fbEstaProdenDet = False

' recorremos el grid detalle de Producto verificando la existencia de txtProd
For j = 1 To grdDetSalida.Rows - 1
 If grdDetSalida.TextMatrix(j, 0) = cboProducto.Text Then
    fbEstaProdenDet = True
    Exit Function
 End If
Next j

End Function

Private Sub cmdBuscar_Click()

' Carga los títulos del grid selección
  giNroColMNSel = 4
  aTitulosColGrid = Array("IdPersona", "Apellidos y Nombres", "Condición", "Activo")
  aTamañosColumnas = Array(1000, 4500, 1500, 600)
' Muestra el formulario de busqueda
  frmMNSeleccion.Show vbModal, Me

' Verifica si se eligió algun dato a modificar
  If gsCodigoMant <> Empty Then
    txtPersonal.Text = gsCodigoMant
    SendKeys "{tab}"
  Else ' No se eligió nada a modificar
    ' Verifica si txtcodigo es habilitado
    If txtPersonal.Enabled = True Then txtPersonal.SetFocus
  End If
  
End Sub

Private Sub cmdCancelar_Click()
If msOperacion = "Nuevo" Then
    
    ' Limpia la las cajas de texto del formulario
    LimpiarPantalla
    
    If cmdCancelar.Value = False Then
       'Prepara el Formulario para un nuevo ingreso
       NuevaSalida
       
    Else: EstableceCamposObligatorios ' establece los campos obligatorios
    End If
    
Else
        
    'Limpia el formulario
    LimpiarCamposDetalle
    grdDetSalida.Rows = 1
   
   'Carga los controles con datos de la salida de Almacén y Habilita los controles
    CargarControlesSalidaAlmacen
    
    'Vacia la colección original de salida
    Set mcolSalidaModificar = Nothing
    CargaDetSalidaAlmacen
     
    'El cursor se ubica en el grid
    grdDetSalida.SetFocus
    
    ' Inicializa el grid
    ipos = 0
    gbCambioCelda = False
    
    'Desabilita el cmdAceptarModificar y eliminar
    cmdAceptarModificar.Enabled = False

End If

End Sub

Private Sub DeshabilitaHabilitaControles()
'-------------------------------------------------------------
'Propósito: Deshabilita, habilita los controles segun la condicion de estos
'Recibe: Nada
'Devuelve: Nada
'-------------------------------------------------------------
' Nota Llamado desde el evento formLoad y despues ingresar el codigo de salida
txtProy.Enabled = Not txtProy.Enabled
cboProy.Enabled = Not cboProy.Enabled
txtPersonal.Enabled = Not txtPersonal.Enabled
cmdBuscar.Enabled = Not cmdBuscar.Enabled

End Sub

Private Sub DeshabilitaHabilitaControlesDet()

cboProducto.Enabled = Not cboProducto.Enabled
txtCantidad.Enabled = Not txtCantidad.Enabled

End Sub

Private Sub LimpiarPantalla()
'Si la operacion elegida en el menu es Modificar se limpia cod Ingreso
If gsTipoOperacionAlmacen = "Modificar" Then
    If cmdCancelar.Value = False Then txtIdSalida.Text = Empty
    mskFecha.Text = "__/__/____"
End If

txtProy.Text = Empty
txtPersonal.Text = Empty
cboProducto.ListIndex = -1
txtCantidad.Text = Empty
txtStock.Text = Empty
txtMedida.Text = Empty
grdDetSalida.Rows = 1


End Sub


Private Sub cmdEliminar_Click()

Dim i As Integer

' Verifica si el grid esta vacio para habilitar o deshabilitar cmdEliminar
If grdDetSalida.Rows = 1 Or grdDetSalida.CellBackColor <> vbDarkBlue Then
    ' No hace nada
Else    'Actualiza monto total detalle
   
    ' Elimina la fila seleccionada del grid
    If grdDetSalida.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdDetSalida.RemoveItem grdDetSalida.Row
    Else
            ' estable vacío el grddetalle
            grdDetSalida.Rows = 1
    End If
    
    ' Actualiza la posición del grid
    ipos = 0
    
    'Habilita el botón aceptarModificar para la Modificación salida de almacen
    HabilitarAceptarModificarSalida
           
End If

End Sub

Private Sub cmdPProducto_Click()

If cboProducto.Enabled Then
    ' alto
     cboProducto.Height = CBOALTO
    ' focus a cbo
    cboProducto.SetFocus
End If

End Sub

Private Sub cmdPProyecto_Click()

If cboProy.Enabled Then
    ' alto
     cboProy.Height = CBOALTO
    ' focus a cbo
    cboProy.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()

'Cierra el formulario
Unload Me

End Sub

Private Sub Form_Load()
Dim sSQL As String

' Verifica el tipo de operacion a realizar en el formulario
If gsTipoOperacionAlmacen = "Modificar" Then
    
     Me.Caption = "Almacén - Modificación de salidas de almacén"

    'muestra el caption del boton aceptar con modificar
    cmdAceptarModificar.Caption = "Modificar"

    'Prepara campos para modificar algun registro de modificacion de salidas de almacén
    ModificarEgreso
    
Else
    'Carga la fecha de trabajo
    mskFecha.Text = gsFecTrabajo
    'Carga los proyectos activos
    sSQL = ""
    sSQL = "SELECT Idproy, Idproy + '   ' + descproy FROM Proyectos " & _
           "WHERE Idproy IN " & _
           "(SELECT idproy FROM Presupuesto_proy ) And " & _
           "'" & FechaAMD(mskFecha.Text) & "' BETWEEN FecInicio And FecFin " & _
           " ORDER BY Idproy + '   ' + descproy "
    CD_CargarColsCbo cboProy, sSQL, mcolCodProy, mcolCodDesProy
    
    ' Deshabilitamos el botón Aceptar y añadir
    cmdAceptarModificar.Enabled = False
    cmdAñadir.Enabled = False
    
    'Nuevo Salida de Almacén
    Me.Caption = "Almacén- Salidas de almacén"
           
   'Deshabilita txtIdSalida
    txtIdSalida.Enabled = False
       
    'Prepara campos para la salida
    NuevaSalida
    
    
End If
    
    'Carga las colecciones
    CargarColPersonal
    CargarColProducto
        
    'Coloca el titulo al Grid
    aTitulosColGrid = Array("Producto", "Medida", "Cantidad", "Idproducto")
    aTamañosColumnas = Array(3990, 1000, 1000, 0)
    
    CargarGridTitulos grdDetSalida, aTitulosColGrid, aTamañosColumnas

End Sub

Private Sub CargarColPersonal()
Dim sSQL As String
If gsTipoOperacionAlmacen = "Modificar" Then
    'Sentencia SQL
    sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
          & " PP.Condicion, PP.Activo " _
          & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
          & " WHERE P.IdPersona=PP.IdPersona " _
          & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"
Else
    'Sentencia SQL
    sSQL = "SELECT P.IdPersona, ( p.Apellidos & ', ' & P.Nombre), " _
          & " PP.Condicion, PP.Activo " _
          & " FROM Pln_Personal P, PLN_PROFESIONAL PP " _
          & " WHERE P.IdPersona=PP.IdPersona and PP.Activo='SI' " _
          & " ORDER BY ( p.Apellidos & ', ' & P.Nombre)"
End If

' Vacía la colección de datos
Set gcolTabla = Nothing
     
' Carga la colecccion de los registros de Productos
Var18 sSQL, 4, gcolTabla

End Sub

Private Sub ModificarEgreso()
msOperacion = "Modificar"

' Inicializa el grid
ipos = 0
gbCambioCelda = False

'Controles eliminar y añadir desabilita

cmdAñadir.Enabled = False

' habilita, deshabilita botones refentes al botón modificar
DeshabilitarBotones (msOperacion)

End Sub


Private Sub HabilitarBotonAñadir()
'----------------------------------------------------------------------------
'PROPÓSITO: Se habilita "Añadir" si se han rellenado los campos obligatorios
'           sino se desabilita
'----------------------------------------------------------------------------
   
If cboProducto.BackColor <> Obligatorio And cboProducto.ListIndex <> -1 _
    And txtCantidad.BackColor <> Obligatorio And txtCantidad.Text <> "" Then
    'Habilita el boton aceptar
    cmdAñadir.Enabled = True
    cmdCancelar.Enabled = True

Else
   'Deshabilita boton añadir
   cmdAñadir.Enabled = False
   
End If
     
End Sub
Private Sub NuevaSalida()
'--------------------------------------------------------------
'Propósito : Realiza la operacion de Salida de Almacén
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------

'Determina el numero siguiente de la salida de Almacén
txtIdSalida.Text = CalcularNroEgreso

' Inicializa el grid
ipos = 0
gbCambioCelda = False

'Carga la variable de tipo de operacion
msOperacion = "Nuevo"
    
'Establece los Campos Obligatorios si estan vacios
EstableceCamposObligatorios
    
'Deshabilita botones refentes a nuevo
DeshabilitarBotones (msOperacion)

End Sub

Private Sub DeshabilitarBotones(ByVal sBoton As String)
Select Case sBoton
    
Case "Nuevo"
     cmdCancelar.Enabled = True
     cmdAnular.Enabled = False
Case "Modificar"
     cmdCancelar.Enabled = True
     cmdAnular.Enabled = True
Case "Cancelar"
     cmdAceptarModificar.Enabled = False
     cmdCancelar.Enabled = False
End Select

End Sub


Private Sub EstableceCamposObligatorios()
'--------------------------------------------------------------
'Propósito : Establece los campos obligatorios
'Recibe : Nada
'Devuelve :Nada
'---------------------------------------------------------------
'Verifica si el habilitado o desabilitado los controles
If txtProy.Enabled Then
    If txtProy.Text = "" Then txtProy.BackColor = Obligatorio
Else
    txtProy.BackColor = vbWhite
End If
If txtPersonal.Enabled = True Then
    If txtPersonal.Text = "" Then txtPersonal.BackColor = Obligatorio
Else
    txtPersonal.BackColor = vbWhite
End If
If cboProducto.Enabled = True Then
    If cboProducto.Text = "" Then cboProducto.BackColor = Obligatorio
Else
    cboProducto.BackColor = vbWhite
End If
If txtCantidad.Enabled = True Then
    If txtCantidad.Text = "" Then txtCantidad.BackColor = Obligatorio
Else
    txtCantidad.BackColor = vbWhite
End If

End Sub

Private Function CalcularNroEgreso() As String
Dim sCodigo As String
Dim sSQL As String
Dim curNumeroEgreso As New clsBD2
Dim iNumSec As Long

' Concatenamos el codigo AñoMes
  sCodigo = Right(mskFecha, 4)

'Carga la sentencia
  sSQL = "SELECT Max(IdSalida) FROM ALMACEN_SALIDAS WHERE  IdSalida LIKE '" & sCodigo & "*'"
  
' Ejecuta la sentencia
  curNumeroEgreso.SQL = sSQL
' Averigua el último número de egreso
  If curNumeroEgreso.Abrir = HAY_ERROR Then
     End
  End If
  
' Separa los cuatro últimos caracteres del maximo numero de ingreso
 If IsNull(curNumeroEgreso.campo(0)) Then
  CalcularNroEgreso = (sCodigo & "0000001")
 Else
  iNumSec = Val(Right(curNumeroEgreso.campo(0), 6))
  CalcularNroEgreso = sCodigo & Format(CStr(iNumSec) + 1, "000000#")
 End If

' Cierra el cursor
  curNumeroEgreso.Cerrar

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Destruye las colecciones
Set mcolCodProy = Nothing
Set mcolCodDesProy = Nothing
Set mcolDesMedidaProd = Nothing

Set mcolidprod = Nothing
Set mcolCodDesProd = Nothing

Set mcolSalidaModificar = Nothing
Set gcolTabla = Nothing

End Sub

Private Sub grdDetSalida_Click()

If grdDetSalida.Row > 0 And grdDetSalida.Row < grdDetSalida.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdDetSalida, vbWhite, vbDarkBlue
End If

End Sub

Private Sub grdDetSalida_EnterCell()

If ipos <> grdDetSalida.Row Then
    '  Verifica si es la última fila
    If grdDetSalida.Row > 0 And grdDetSalida.Row < grdDetSalida.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdDetSalida, ipos
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdDetSalida.Row
End If

End Sub

Private Sub grdDetSalida_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Carga el producto a editar
    grdDetSalida_DblClick
 End If
 
End Sub

Private Sub grdDetSalida_DblClick()

' Selecciona toda la iFila
If grdDetSalida.Rows > 1 Then
    ' Verifica si esta seleccionado
    If grdDetSalida.CellBackColor <> vbDarkBlue Then
       MarcarFilaGRID grdDetSalida, vbWhite, vbDarkBlue
       Exit Sub
    End If
    
    'Verifica que el cboProducto este vacio
    If cboProducto.Text <> Empty Then
        'Termina la ejecucion
        Exit Sub
    End If

    'Recupera en el msidprod el codigo del producto seleccionado
    msIdprod = grdDetSalida.TextMatrix(grdDetSalida.RowSel, 3)
    
    'Recupera en el combo el producto del msidprod
    CD_ActVarCbo cboProducto, msIdprod, mcolCodDesProd
    
    'Determina si el producto es nuevo
    If DeterminarNuevoProducto Then

        'Recupera la cantidad del producto seleccionado
        txtCantidad.Text = grdDetSalida.TextMatrix(grdDetSalida.RowSel, 2)

    Else
        'Coloca la cantidad anterior para la modificación
        txtCantidad.Text = Var30(mcolSalidaModificar.Item(msIdprod), 2)

    End If
    
    'Recupera la medida del producto seleccionado
    txtMedida.Text = grdDetSalida.TextMatrix(grdDetSalida.RowSel, 1)
    
    
    'Carga el stock del producto
    CargarStock
    
    ' LLama al procedimiento eliminar
    cmdEliminar_Click
    
    ' coloca el focus a cbo producto
    cboProducto.SetFocus
    
    ' Actualiza el ipos
    ipos = 0
   
End If

End Sub

Private Sub txtCantidad_Change()

'Determina txtCantidad si el obligatorio
If txtCantidad.Text <> "" And Val(txtCantidad.Text) <> 0 Then
     txtCantidad.BackColor = vbWhite
Else
   txtCantidad.BackColor = Obligatorio
End If

'Habilita boton añadir
HabilitarBotonAñadir

'Habilitar boton aceptarModificar
 HabilitarAceptarModificarSalida
 
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

'Se tabula con INTRO
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If

'Valida Dato Ingresado
Var34 txtCantidad, 11, KeyAscii

End Sub

Private Sub txtCantidad_LostFocus()

' pone el maximo número de caracteres
txtCantidad.MaxLength = 11

' Verifica si se introdujjo un dato valido
If Val(Var37(txtCantidad)) = 0 Then
    txtCantidad = Empty
Else
    ' formato a cantidad
    txtCantidad = Format(Val(Var37(txtCantidad.Text)), "#0.00")
End If

End Sub


Private Sub txtPersonal_Change()
'Verifica si el tamaño del txt es Igual al tamaño definido
If Len(txtPersonal) = txtPersonal.MaxLength Then
    'Actualiza el txtDesc
    ActualizaDesc
Else
    'Limpia el txtDescAfecta
    txtDesc.Text = Empty
End If

 ' Verifica si el campo esta vacio
If txtPersonal.Text <> Empty And txtDesc.Text <> Empty Then
   ' Los campos coloca a color blanco
   txtPersonal.BackColor = vbWhite
Else
   'Los campos coloca a color amarillo
   txtPersonal.BackColor = Obligatorio
End If

  'Habilitar el boton aceptar de la segunda parte del formulario
  HabilitarAceptarModificarSalida


End Sub

Private Sub ActualizaDesc()
'--------------------------------------------------------------
'PROPÓSITO  : Actualiza la descripcion de la persona
'Recive     : Nada
'Devuelve   : Nada
'--------------------------------------------------------------
On Error GoTo mnjError
'Copia la descripción
txtDesc.Text = Var30(gcolTabla.Item(txtPersonal.Text), 2)

' Maneja el error si no es indice
'-----------------------------------------------------------------
mnjError:
    If Err.Number = 5 Then ' Error al acceder al elemento de la colección
        'Muestra el mensaje
        MsgBox "El código ingresado no existe", , "SGCcaijo-Ingresos"
        'Limpia la descripción
        txtDesc.Text = Empty
    End If
End Sub

Private Sub txtPersonal_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtProy_Change()
Dim sSQL As String

' Si procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboProy, txtProy, mcolCodDesProy

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

cboProducto.Clear

'Carga el cbo de Producto
'CargarCboCols cboProducto, mcolCodDesProd
CargarCboProd cboProducto, mcolCodDesProd
     
 ' Verifica si el campo esta vacio
If txtProy.Text <> "" And cboProy.Text <> "" Then
   ' Los campos coloca a color blanco
   txtProy.BackColor = vbWhite
Else
   'Los campos coloca a color amarillo
   txtProy.BackColor = Obligatorio
End If

'Habilitar el boton aceptar de la segunda parte del formulario
HabilitarAceptarModificarSalida

End Sub

Private Sub HabilitarAceptarModificarSalida()

'Desabilita el botón aceptar
cmdAceptarModificar.Enabled = False

'Inicializa variables que indican si se modificó el egreso y que parte
mbModificaAlmacen = False
mbModificaDetAlmacen = False

If txtProy.BackColor <> Obligatorio And txtPersonal.BackColor <> Obligatorio Then
  If msOperacion = "Modificar" Then 'Verifica si la operación realizada es modificacion
    If txtProy.Text <> mcurRegSalidaAlmacen.campo(0) Or txtPersonal.Text <> mcurRegSalidaAlmacen.campo(1) Then
      'Habilita el botón aceptar
      cmdAceptarModificar.Enabled = True
      ' actualizamos varialble que indica si modifico los datos generales del egreso
      mbModificaAlmacen = True
    End If
    If fbModDetalle = True Then 'Verifica si se modificó el detalle del Almacen
      cmdAceptarModificar.Enabled = True
      mbModificaDetAlmacen = True
        
    Else 'Verifica si grid Detalle no tiene elementos
      If grdDetSalida.Rows = 1 Then
        cmdAceptarModificar.Enabled = False
      End If
    End If
  Else
    If grdDetSalida.Rows > 1 Then
        'Habilita el botón aceptar
        cmdAceptarModificar.Enabled = True
        Exit Sub
    End If
  End If
End If
End Sub

Private Function fbModDetalle() As Boolean
Dim i As Integer
Dim sClave As String
Dim bencontrado As Boolean
On Error GoTo ErrClaveCol
 
'inicializa la funcion
fbModDetalle = False

'Verifica si se eliminó o agregó algun elemento
If (mcolSalidaModificar.Count <> grdDetSalida.Rows - 1) Then

    'hubo una modificacion en el detalle del egreso
    fbModDetalle = True
    Exit Function
End If

'Verifica si se modifico los datos del registro detalle en el grd
i = 1
bencontrado = False
Do While i <= grdDetSalida.Rows - 1 And bencontrado = False
    
    'Concatena IdProd, Cantidad
    sClave = grdDetSalida.TextMatrix(i, 3) & "¯" & grdDetSalida.TextMatrix(i, 2)
    
    'Verifica si se modifico algun dato del grid
    If sClave <> mcolSalidaModificar.Item(grdDetSalida.TextMatrix(i, 3)) Then
       fbModDetalle = True
       bencontrado = True
    End If
    
ErrClaveCol:

    If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
       fbModDetalle = True
       bencontrado = True
    End If
    i = i + 1
Loop

End Function

Private Sub ModificarDetSalidaAlmacen()

Dim i As Integer
Dim strRegSalida As String
Dim Objeto As Variant
On Error GoTo mnjError

'Identifica el tipo de modificación y realiza el proceso
For i = 1 To grdDetSalida.Rows - 1
    'Arma el registro de salida del grdDetalle
    strRegSalida = grdDetSalida.TextMatrix(i, 3) & "¯" & Var37(grdDetSalida.TextMatrix(i, 2))
    'Verifica si el registro de salida es Igual al registro original
    If strRegSalida <> mcolSalidaModificar.Item(grdDetSalida.TextMatrix(i, 3)) Then
        'Realiza la modificación de la salida del producto
        ModificarSalidaProducto grdDetSalida.TextMatrix(i, 3), grdDetSalida.TextMatrix(i, 2), "Modificar"
        'Elimina la salida del producto de la colección original
        mcolSalidaModificar.Remove (grdDetSalida.TextMatrix(i, 3))
    Else
        'Elimina la salida del producto de la colección original
        mcolSalidaModificar.Remove (grdDetSalida.TextMatrix(i, 3))
    End If
ResumePostError:

Next i

'Verifica si existen salidas de productos que han sido eliminadas
If mcolSalidaModificar.Count > 0 Then
    'Recorre los elementos eliminados de la coleccion salidas de almacén
    For Each Objeto In mcolSalidaModificar
        'Realiza la modificación de la salida del producto
        ModificarSalidaProducto Var30(Objeto, 1), Var30(Objeto, 2), "Eliminar"
    Next Objeto
End If
Exit Sub

mnjError:
'Verifica si el error es elemento no encontrado en la colección
If Err.Number = 5 Then
    'Realiza una nueva salida de un producto
    ModificarSalidaProducto grdDetSalida.TextMatrix(i, 3), grdDetSalida.TextMatrix(i, 2), "Registrar"
    'Continua con la siguientes salidas de almacén
    Resume ResumePostError
End If

End Sub

Private Sub ModificarSalidaProducto(sIdProd As String, dblCantidad As Double, sTipoMod As String)
RestaurarIngresos txtIdSalida.Text, sIdProd

'Carga salidas posteriores al número de salida a modificar
CargarSalidasPosteriores txtIdSalida.Text, sIdProd

'Elimina el detalle de las salidas a partir del numero de salida a modificar
EliminarDetSalidas txtIdSalida.Text, sIdProd, sTipoMod

'Averigua el codigo de suministro y variación de almacen
DeterminaCtas sIdProd

'Modifica el registro de salida actual del producto
ModificarSalida txtIdSalida.Text, sIdProd, dblCantidad, sTipoMod

'Actualiza las salidas posteriores
ActualizarSalidasPosteriores sIdProd

End Sub

Private Sub ActualizarSalidasPosteriores(sIdProd As String)
Do While Not mcurSalidasPost.EOF
    ' Realiza el proceso de actualizar datos de las salidas posteriores a la modificada
    ModificarSalida mcurSalidasPost.campo(0), sIdProd, mcurSalidasPost.campo(1), "Modificar"
    ' Mueve al siguiente registro
    mcurSalidasPost.MoverSiguiente
Loop

'Cierra el cursor
mcurSalidasPost.Cerrar

End Sub

Private Sub ModificarSalida(sIdSalida As String, sIdProd As String, dblCantidad As Double, sTipoMod As String)
'--------------------------------------------------------------
'Propósito  : Realiza la salida para registro de salida actual
'Recibe     : sIdSalida, Identificador de la salida
'             sIdProd, Identificador del producto
'             dblCantidad, Cantidad de salida de un producto
'             sTipoMod, Tipo de modificación de registro de salida
'Devuelve   : Nada
'--------------------------------------------------------------
Dim dblMontoSalida As Double

'Inicializa la variable
dblMontoSalida = 0

'Selecciona el tipo de Modificación a realizar
Select Case sTipoMod
Case "Modificar"
    'Actualiza la cantidad modificada de la linea de productos
    'Cargar Colección salida línea mercadería
    CargaColeccionSalidasDisponible sIdProd, dblCantidad

    'Guarda salida de línea de mercaderia de almacén y calcula el monto de la salida
    GuardaSalidasDisponible sIdProd, dblMontoSalida, sIdSalida
    
    'Generar Asiento salida por la salida del producto
    GenerarAsientoSalidaLineaProducto sIdSalida, sIdProd, dblMontoSalida, sTipoMod
    
    'Limpia la colección mcolSalidaDisponibles
    Set mcolSalidasDisponible = Nothing

Case "Eliminar"
    'Generar Asiento salida por la salida del producto
    GenerarAsientoSalidaLineaProducto sIdSalida, sIdProd, 0, sTipoMod
    
Case "Registrar"
    'Guarda las salidas de una linea de productos
    'Cargar Colección salida línea mercadería
    CargaColeccionSalidasDisponible sIdProd, dblCantidad

    'Guarda salida de línea de mercaderia de almacén y calcula el monto de la salida
    GuardaSalidasDisponible sIdProd, dblMontoSalida, sIdSalida
    
    'Generar Asiento salida por la salida del producto
    GenerarAsientoSalidaLineaProducto sIdSalida, sIdProd, dblMontoSalida, sTipoMod
   
   'Limpia la colección mcolSalidaDisponibles
    Set mcolSalidasDisponible = Nothing
Case "Anular"
    
    'Generar Asiento salida por la salida del producto
    GenerarAsientoSalidaLineaProducto sIdSalida, sIdProd, 0, sTipoMod
    
End Select

End Sub

Private Sub EliminarDetSalidas(sIdSalida As String, sIdProd As String, sTipoMod As String)
Dim sSQL As String
Dim modDetAlmacen As New clsBD3
If sTipoMod = "Anular" Then
    'Carga la sentencia que elimina los registros
    sSQL = "DELETE SA.* " _
        & "FROM  ALMACEN_SAL_DET AD " _
        & "WHERE AD.IdSalida>'" & sIdSalida & "' And AD.IdProd= '" & sIdProd & "' and " _
        & " AD.IdSalida in (SELECT SA.IdSalida FROM ALMACEN_SALIDAS SA WHERE SA.IdSalida>='" & sIdSalida & "' And SA.Anulado='NO')"
Else
    'Carga la sentencia que elimina los registros
    sSQL = "DELETE SA.* " _
        & "FROM  ALMACEN_SAL_DET AD " _
        & "WHERE AD.IdSalida>='" & sIdSalida & "' And AD.IdProd= '" & sIdProd & "' and " _
        & " AD.IdSalida in (SELECT SA.IdSalida FROM ALMACEN_SALIDAS SA WHERE SA.IdSalida>='" & sIdSalida & "' And SA.Anulado='NO')"

End If

'ejecuta la sentencia sQL
modDetAlmacen.SQL = sSQL

'Verifica si hay error
If modDetAlmacen.Ejecutar = HAY_ERROR Then End

End Sub

Private Sub CargarSalidasPosteriores(sIdSalida As String, sIdProd As String)
Dim sSQL As String

'Carga las salidas a posteriores del número de salida
'Sentencia SQL
sSQL = "SELECT SA.IdSalida, SUM(AD.Cantidad) " _
     & "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD " _
     & "WHERE SA.IdSalida=AD.IdSalida And AD.IdSalida > '" & sIdSalida & "' " _
     & "And SA.Anulado='NO' And AD.IdProd='" & sIdProd & "' " _
     & "GROUP BY SA.IdSalida " _
     & "ORDER BY SA.IdSalida "

'Copia la sentencia
mcurSalidasPost.SQL = sSQL

'Verifica si hay error
If mcurSalidasPost.Abrir = HAY_ERROR Then End

End Sub

Private Sub RestaurarIngresos(sIdSalida As String, sIdProd As String)
Dim sSQL As String
Dim curDetSalidasPost As New clsBD2

'Carga el detalle de salidas a partir del número de salida
'Sentencia SQL
sSQL = "SELECT AD.Orden, SUM(AD.Cantidad), SUM(AD.Precio), AI.CantidadDisponible, AI.Resto " _
     & "FROM ALMACEN_SALIDAS SA, ALMACEN_SAL_DET AD, ALMACEN_INGRESOS AI " _
     & "WHERE SA.IdSalida=AD.IdSalida And AD.IdSalida >= '" & sIdSalida & "' " _
     & "And SA.Anulado='NO' And AD.IdProd='" & sIdProd & "' And AD.Orden=AI.Orden And " _
     & "AD.IdProd=AI.IdProd " _
     & "GROUP BY AD.Orden,AI.CantidadDisponible, AI.Resto"

'Copia la sentencia
curDetSalidasPost.SQL = sSQL

'Verifica si hay error
If curDetSalidasPost.Abrir = HAY_ERROR Then End

'Devuelve los datos de cada salida a partir del número de salida
Do While Not curDetSalidasPost.EOF

  'Devuelve los datos de cada salida al ingreso correspondiente
  IncrementarDisponiblesBD curDetSalidasPost.campo(0), sIdProd, _
                           curDetSalidasPost.campo(1), curDetSalidasPost.campo(2), _
                           curDetSalidasPost.campo(3), curDetSalidasPost.campo(4)
                        
  'Mueve al siguiente registro del cursor
  curDetSalidasPost.MoverSiguiente
  
Loop

'Cierra el cursor
curDetSalidasPost.Cerrar

End Sub

Private Sub CargarColProducto()
Dim sSQL As String

'Verifica el tipo de operación
If msOperacion = "Nuevo" Then
    'Sentencia SQL para el nuevo producto
    sSQL = "SELECT DISTINCT A.Idprod,P.DescProd,P.Medida, Tipo " _
            & " FROM PRODUCTOS P, ALMACEN_INGRESOS A " _
            & " WHERE P.IdProd=A.IdProd and A.CantidadDisponible >0 " _
            & " ORDER BY DescProd"
Else
    'Sentencia SQL para la modificación del producto
    sSQL = "SELECT DISTINCT A.Idprod,P.DescProd,P.Medida, Tipo " _
            & " FROM PRODUCTOS P, ALMACEN_INGRESOS A " _
            & " WHERE P.IdProd=A.IdProd and A.CantidadDisponible >=0 " _
            & " ORDER BY DescProd"
End If

'Carga la coleccion de descripcion y medida de los productos
mcurMedidaProd.SQL = sSQL

If mcurMedidaProd.Abrir = HAY_ERROR Then
  End
End If

Do While Not mcurMedidaProd.EOF
    
    ' Se carga la colección de descripciones + unidades de los productos con la 1º y 2º
    ' columnas del cursor.  Nótese que la primera columna del cursor
    ' hace de clave de la colección.
    mcolDesMedidaProd.Add mcurMedidaProd.campo(2), mcurMedidaProd.campo(0)
    
    'Coleccion de producto y su descripción
    mcolidprod.Add mcurMedidaProd.campo(0)
    'mcolCodDesProd.Add Item:=mcurMedidaProd.campo(1), Key:=mcurMedidaProd.campo(0)
    mcolCodDesProd.Add mcurMedidaProd.campo(0) & "¯" & mcurMedidaProd.campo(1) & "¯" & mcurMedidaProd.campo(3)

    ' Se avanza a la siguiente fila del cursor
    mcurMedidaProd.MoverSiguiente

Loop

'Cierra el cursor de medida de productos
mcurMedidaProd.Cerrar

End Sub

Private Sub CargarStock()
Dim curStock As New clsBD2
Dim sSQL As String
Dim dblStock As Double

dblStock = 0 'Inicializa variable saldo

'Averigua el ingreso del producto a Almacén
sSQL = "SELECT SUM(CantidadDisponible) as Ingreso FROM ALMACEN_INGRESOS " _
      & "WHERE IdProd='" & msIdprod & "'"

'Copia la sentencia SQL
curStock.SQL = sSQL

'Verifica si hay error
If curStock.Abrir = HAY_ERROR Then End

'Verifica si es nulo
If Not IsNull(curStock.campo(0)) Then dblStock = curStock.campo(0)

'Cierra el cursor
curStock.Cerrar

'Devuelve el stock
txtStock = Format(dblStock, "###,###,##0.00")

End Sub

Public Sub CargarRegSalidaAlmacen()

'Coloca los campos obligatorios
EstableceCamposObligatorios

'Carga salida de Almacén de la BD
 CargaSalidaAlmacen
 
'Carga el detalle de la salida de Almacén
CargaDetSalidaAlmacen

End Sub

Private Sub CargaDetSalidaAlmacen()
Dim curDetSalidaAlmacen As New clsBD2
Dim sSQL As String

' Carga la sentencia que consulta a la BD acerca del registo de salida de Almacén
    sSQL = ""
    sSQL = "SELECT D.IdProd, P.DescProd, P.Medida, SUM(D.Cantidad), SUM(D.Precio)" & _
           "FROM Almacen_Sal_Det D, PRODUCTOS P WHERE " & _
           "'" & txtIdSalida.Text & "'= D.IdSalida and D.IdProd=P.IdProd " & _
           "Group by D.IdProd, P.DescProd, P.Medida "

curDetSalidaAlmacen.SQL = sSQL

' Abre el cursor si hay  error sale indicando la causa del error
If curDetSalidaAlmacen.Abrir = HAY_ERROR Then
    End
End If

If curDetSalidaAlmacen.EOF Then
  
  'No existe cuentas asociadas
  MsgBox "No existen productos para esta salida." & Chr(13) & _
         "Verifique si se han eliminado estos productos anteriormente", _
          vbInformation + vbOKOnly, "S.G.Ccaijo-Salida de Almacén"
Else

    'Verifica la existencia del registro de Egreso
    Do While Not curDetSalidaAlmacen.EOF

        'Añade a la colección las salidas de lineas de producto
        'IdProd, Cantidad
        mcolSalidaModificar.Add Item:=curDetSalidaAlmacen.campo(0) & "¯" & _
                                 curDetSalidaAlmacen.campo(3), _
                                 Key:=curDetSalidaAlmacen.campo(0)

        'Añade el nuevo registro al grid
        grdDetSalida.AddItem (curDetSalidaAlmacen.campo(1) & vbTab & curDetSalidaAlmacen.campo(2) & vbTab & Format(curDetSalidaAlmacen.campo(3), "###,###,##0.00") _
                            & vbTab & curDetSalidaAlmacen.campo(0))

        'Mueve al siguiente registro del cursor
        curDetSalidaAlmacen.MoverSiguiente
    
    Loop

End If

'Cierra el cursor curDetSalidaAlmacen
curDetSalidaAlmacen.Cerrar

End Sub

Private Sub CargaSalidaAlmacen()
Dim sSQL As String

' Determina la salida de Almacén
sSQL = ""
sSQL = "SELECT DISTINCT S.IdProy, S.IdPersona, S.Fecha " _
& "FROM Almacen_Salidas S " _
& "WHERE S.IdSalida=" & "'" & Trim(txtIdSalida.Text) & "'"

mcurRegSalidaAlmacen.SQL = sSQL

' Abre el cursor si hay  error sale indicando la causa del error
If mcurRegSalidaAlmacen.Abrir = HAY_ERROR Then
    End
End If

'Verifica la existencia del registro de salida de Almacén
If mcurRegSalidaAlmacen.EOF Then

    'Mensaje de registro de egreso a Caja o Bancos no existe
    MsgBox "El código de salida de almacén no esta registrado como salida de Almacén", _
            vbInformation + vbOKOnly, "Almacén- Salida de Almacén"

    'La salida de Almacén no existe
    mcurRegSalidaAlmacen.Cerrar

Else
    'Carga los controles con datos de la salida de Almacén y Habilita los controles
    CargarControlesSalidaAlmacen
    
End If

End Sub

Private Sub CargarControlesSalidaAlmacen()
Dim sSQL As String
'Destruye la coleccion
Set mcolCodProy = Nothing
Set mcolCodDesProy = Nothing

'Carga los proyectos activos
sSQL = "SELECT Idproy, Idproy + '   ' + descproy FROM Proyectos " & _
       "WHERE  idproy IN " & _
       "(SELECT idproy FROM Presupuesto_proy ) " & _
       "ORDER BY Idproy + '   ' + descproy"
CD_CargarColsCbo cboProy, sSQL, mcolCodProy, mcolCodDesProy

'Deshabilita IdSalida
txtIdSalida.BackColor = vbWhite
txtIdSalida.Enabled = False

'Rellena los controles de salida de Almacén
txtProy.Text = mcurRegSalidaAlmacen.campo(0)
txtPersonal.Text = mcurRegSalidaAlmacen.campo(1)

'Habilita Botones Cancelar, Anular de Salida de Almacén
cmdCancelar.Enabled = False

End Sub

Private Sub txtProy_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Public Sub CargarCboProd(cboRec As ComboBox, colRec As Collection)
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

Public Sub ActualizarInfoProd(sTextoCbo As String, txtRec As String, _
                     colCod As Collection, colCodDesc As Collection)
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

