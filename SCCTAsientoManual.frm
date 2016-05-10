VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCTAsientoManual 
   Caption         =   "Contabilidad- Asientos Manuales"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   101
   Icon            =   "SCCTAsientoManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6240
      Left            =   40
      TabIndex        =   22
      Top             =   1080
      Width           =   11775
      Begin VB.CommandButton cmdPCtaContable 
         Height          =   255
         Left            =   10560
         Picture         =   "SCCTAsientoManual.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   250
         Width           =   220
      End
      Begin VB.ComboBox cboCtaContable 
         Height          =   315
         Left            =   5280
         Style           =   1  'Simple Combo
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         Width           =   5535
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   600
         TabIndex        =   23
         Top             =   120
         Width           =   2175
         Begin VB.OptionButton optHaber 
            Caption         =   "Haber"
            Height          =   255
            Left            =   1200
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optDebe 
            Caption         =   "Debe"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtCodContable 
         Height          =   315
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   5
         Top             =   225
         Width           =   975
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Top             =   615
         Width           =   1575
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "A&ñadir"
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   10200
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtTotalDebe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtTotalHaber 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10320
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   5760
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grdDebe 
         Height          =   4215
         Left            =   100
         TabIndex        =   13
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         HighLight       =   0
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdHaber 
         Height          =   4215
         Left            =   5930
         TabIndex        =   14
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         HighLight       =   0
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cod.Contable:"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Total Haber:"
         Height          =   255
         Left            =   9360
         TabIndex        =   27
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Total Debe:"
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Haber:"
         Height          =   375
         Left            =   5930
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Debe:"
         Height          =   375
         Left            =   100
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton cmdAceptarModificar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10815
      TabIndex        =   12
      Top             =   7350
      Width           =   900
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   40
      TabIndex        =   18
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtNumAsiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   600
         Width           =   7455
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   8520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   3600
         Picture         =   "SCCTAsientoManual.frx":0BA2
         Stretch         =   -1  'True
         Top             =   170
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "NºAsiento:"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   7800
         TabIndex        =   20
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCTAsientoManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Colecciones para cargar los combos
Dim mcolCodPlanCont As New Collection
Dim mcolDesCodPlanCont As New Collection

' colecciones que cargan los registros de detalle
Dim mcolAsientosDebe As New Collection
Dim mcolAsientosHaber As New Collection

' cursor usado en la modificacion
Dim mcurAsiento As New clsBD2

Dim msNumAsiento  As String
Dim msGlosa As String
Dim msFecha As String

' Codigo del asiento
Dim msCodAsiento As String

'Monto total del debe
Dim mdblSumaMontoDebe As Double

'Monto total del haber
Dim mdblSumaMontoHaber As Double

'booleano  para averiguar SI se cambiolos datos
Dim mbCambioAsiento As Boolean
Dim mbCambioDetalle As Boolean

' Variables para el manejo de los grids
Dim iposH As Long
Dim iposD As Long

Private Sub cboCtaContable_Change()

' verifica SI lo ingresado esta en la lista del combo
If VerificarTextoEnLista(cboCtaContable) = True Then SendKeys "{down}"

End Sub

Private Sub cboCtaContable_Click()

' Verifica SI el evento ha sido activado por el teclado o Mouse
If VerificarClick(cboCtaContable.ListIndex) = False And cboCtaContable.Height = CBOALTO Then SendKeys "{tab}" 'evento activado por mouse

End Sub


Private Sub cboCtaContable_KeyDown(KeyCode As Integer, Shift As Integer)

' Verifica SI es enter para salir o flechas para recorrer
VerificaKeyDowncbo (KeyCode)

End Sub

Private Sub cboCtaContable_LostFocus()

' sale del combo y acualiza datos enlazados
If ValidarDatoCbo(cboCtaContable, vbWhite) = True Then
' Se actualiza código (TextBox) correspondiente a descripción introducida
  CD_ActCod cboCtaContable.Text, txtCodContable, mcolCodPlanCont, mcolDesCodPlanCont
  If InStr(txtCodContable.Text, "104") > 0 Or InStr(txtCodContable.Text, "108") > 0 Then
    If CtaAnulada Then
      MsgBox "La Cuenta Bancaria esta anulada. No se puede continuar.", , "SGCcaijo - Asiento Manual"
      LimpiarCamposDetalle
    End If
  End If
Else
  'NO se encuentra la CtaContable
  txtCodContable = Empty
End If

'Cambia el alto del combo
cboCtaContable.Height = CBONORMAL

End Sub



Private Sub cmdAceptarModificar_Click()
'Verifica si los datos son correctos
If DatosOK = False Then
    'Termana la ejecucion del procedimientos
    Exit Sub
End If
' Verifica el tipo de operacion realizada en ele formulario
If gsTipoOperacionContabilidad = "Nuevo" Then
      ' Mensaje de conformidad de los datos
      If MsgBox("¿Está conforme con los datos?.", _
                  vbQuestion + vbYesNo, "SGCcaijo - Asiento Manual") = vbYes Then
            'Actualiza la transaccion
            Var8 1, gsFormulario
            
            'Guarda los registros generales
            GuardaRegGenerales
            
            'Guarda los registros detalle
            GuardaRegDetalle
     
            'Actualiza la transaccion
            Var8 -1, Empty
           
            'Mensaje de Operación realizada adecuadamente
            MsgBox "Operación realizada correctamente", , "SGCcaijo - Asiento Manual"
        
            'Genera la nueva pantalla para un nuevo siento manual
            NuevoAsiento
       End If
Else
    ' mensaje de confirmación
    If MsgBox("¿Esta de Acuerdo con las modificaciones?", vbInformation + vbYesNo, "Contabilidad: Asiento Manual") = vbYes Then
         'Actualiza la transaccion
        Var8 1, gsFormulario
       
        ' Verifica SI se modifico el asiento
        If mbCambioAsiento = True Then
            ' Guarda las modificaciones del asiento en la BD
            ModificaAsiento
        End If
        
        ' verifica SI se modificó el detalle
        If mbCambioDetalle = True Then
            
            'Elimina los registros de detalle
            Conta2 (txtNumAsiento.Text)
            
            ' Guarda las modificaciones del asiento en la BD
            GuardaRegDetalle
            
        End If
        
        'Actualiza la transaccion
        Var8 -1, Empty
        
        ' manda mensaje de operación realizada satisfactoriamente
         MsgBox "Modificación realizada correctamente", , "Contabilidad, asiento manual"
         
        ' cierra el formulario
         Unload Me
                     
     End If
End If

End Sub

Private Function DatosOK() As Boolean
'Verifica si el año esta cerrado
If Conta52(Right(mskFecha.Text, 4)) = True Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable ha sido cerrado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Asiento Auxiliar"
    'Devuelve el resultado
    DatosOK = False
    Exit Function
End If
'Verifica si existe asiento de apertura
If Conta53(Right(mskFecha.Text, 4)) = False Then
    ' No se puede realizar la operación
    MsgBox "No se puede realizar la operación, el periodo contable no ha sido aperturado.", _
    vbCritical + vbOKOnly, "SGCCAIJO-Asiento Auxiliar"
    'Devuelve el resultado
    DatosOK = False
    Exit Function
End If
'Devuelve el resultado
DatosOK = True

End Function

Private Sub ModificaAsiento()
' ------------------------------------------------------------------
'Proposito: Modifica en la bd los datos generales del asiento
'Recibe: nada
'Entrega: Nada
' ------------------------------------------------------------------
Dim sSQL As String
Dim modAsiento As New clsBD3

    ' se debe modificar el registro de asiento manual
    sSQL = "UPDATE CTB_ASIENTOS " _
         & "SET Glosa='" & txtGlosa & "' " _
         & "WHERE NumAsiento = '" & txtNumAsiento.Text & "'"
     ' ejecuta la modificacion
     modAsiento.SQL = sSQL
     If modAsiento.Ejecutar = HAY_ERROR Then End
    ' cierra componente de modificación
     modAsiento.Cerrar

End Sub


Private Sub GuardaRegDetalle()
' --------------------------------------------------------
' Proposito :Guarda los registros detalle en la tabla CTB_ASIENTOS_DET
' Recibe : Nada
' Entrega: Nada
' --------------------------------------------------------
Dim i As Integer
Dim sSQL As String
Dim dblIdAsiento As Double
Dim modDebeHaber As New clsBD3

'Determina el siguiente codigo del asiento manual
dblIdAsiento = 1

' Recorre el grid y lo almacena en la BD
For i = 1 To grdDebe.Rows - 1
    
    ' Inserta el detalle en Gastos
    sSQL = "INSERT INTO CTB_ASIENTOS_DET VALUES('" & txtNumAsiento.Text & "','" & Format(dblIdAsiento, "0000#") & "','" _
    & grdDebe.TextMatrix(i, 0) & "','D'," _
    & Var37(grdDebe.TextMatrix(i, 2)) & ")"

    modDebeHaber.SQL = sSQL
    If modDebeHaber.Ejecutar = HAY_ERROR Then
      End
    End If
        
    'Se cierra la query
    modDebeHaber.Cerrar
    
    'Incremtenta el 1 en codigo
    dblIdAsiento = dblIdAsiento + 1
Next i
    
' Recorre el grid y lo almacena en la BD
For i = 1 To grdHaber.Rows - 1

    ' Inserta el detalle en Gastos
    sSQL = "INSERT INTO CTB_ASIENTOS_DET VALUES('" & txtNumAsiento.Text & "','" & Format(dblIdAsiento, "0000#") & "','" _
    & grdHaber.TextMatrix(i, 0) & "','H'," _
    & Var37(grdHaber.TextMatrix(i, 2)) & ")"

    modDebeHaber.SQL = sSQL
    If modDebeHaber.Ejecutar = HAY_ERROR Then
      End
    End If
        
    'Se cierra la query
    modDebeHaber.Cerrar
    
   'Incremtenta el 1 en codigo
    dblIdAsiento = dblIdAsiento + 1
        
Next i

End Sub

Private Sub GuardaRegGenerales()
' --------------------------------------------------------
' Proposito :Guarda los registros generales en la tabla CTB_ASIENTOS
' Recibe : Nada
' Entrega: Nada
' --------------------------------------------------------
Dim sSQL As String
Dim modAsiento As New clsBD3

 ' Se registra el nuevo asiento
    sSQL = "INSERT INTO CTB_ASIENTOS VALUES('" & txtNumAsiento.Text & "','" _
            & FechaAMD(mskFecha.Text) & "','" & txtGlosa.Text & "','NO','AM')"
            
    ' Se ejecuta la sentencia que registra el asiento
    modAsiento.SQL = sSQL
    
    ' SI al ejecutar ha y un error sale
    If modAsiento.Ejecutar = HAY_ERROR Then End
    
    ' cierra la componente sql
    modAsiento.Cerrar
    
End Sub

Private Sub NuevoAsiento()
' --------------------------------------------------------
' Proposito :Limpia las cajas de texto para el ingreso de un nuevo _
            asiento contable
' Recibe : Nada
' Entrega: Nada
' --------------------------------------------------------
' Limpia las cajas de texto
txtCodContable.Text = Empty
txtMonto.Text = Empty
txtGlosa.Text = Empty

' Inicializa el grid
iposH = 0
iposD = 0
gbCambioCelda = False

' asigna campos obligatorios
'Determina el siguiente número del asiento manual
txtNumAsiento.Text = Conta4("MA", gsFecTrabajo)

' Define los campos obligatorios
' EstableceCamposObligatorios
EstableceCamposObligatorios

'Inicializa la variable mdblSumaMontoDebe
mdblSumaMontoDebe = 0
mdblSumaMontoHaber = 0
End Sub

Private Function AsignarDebeHaber() As String
'----------------------------------------------------
'Proposito: Según Opt devuelve debe,haber
'----------------------------------------------------
If optDebe.Value = True Then
    ' asiento al debe
    AsignarDebeHaber = "D"
Else ' asiento en el haber
    AsignarDebeHaber = "H"
End If
End Function




Private Sub cmdAñadir_Click()
Dim i As Integer

If optDebe.Value Then
    If VerificarElemento(grdDebe) = False Then
    
        'Añade el nuevo registro al grid
        grdDebe.AddItem (txtCodContable.Text & vbTab & cboCtaContable.Text & vbTab & Format(txtMonto.Text, "###,###,###,##0.00") & vbTab & "D")
        
        ' Actualiza el Campo monto total
        mdblSumaMontoDebe = mdblSumaMontoDebe + Val(Var37(txtMonto.Text))
        
        'Actualiza el monto total al txtSumaMontoDebe
        txtTotalDebe.Text = Format(mdblSumaMontoDebe, "###,###,###,##0.00")
        
         'Limpia los campos
        LimpiarCamposDetalle
        
        ' Habilita el botón aceptar
        If optDebe.Enabled Then optDebe.SetFocus

    Else
        ' Envia mensaje
        MsgBox "La cuenta contable ya ha sido ingresado al detalle del Debe", _
                   vbInformation + vbOKOnly, "SGCcaijo - Asiento Manual"
        ' limpia la cuenta para  dar opcion a elegir
        cboCtaContable.SetFocus
        cmdAñadir.Enabled = False
        Exit Sub
    End If
    
Else
    If VerificarElemento(grdHaber) = False Then
    
        'Añade el nuevo registro al grid
        grdHaber.AddItem (txtCodContable.Text & vbTab & cboCtaContable.Text & vbTab & Format(txtMonto.Text, "###,###,###,##0.00") & vbTab & "H")
        
        ' Actualiza el Campo monto total
        mdblSumaMontoHaber = mdblSumaMontoHaber + Val(Var37(txtMonto.Text))
        'Actualiza el monto total al txtSumaMontoHaber
        txtTotalHaber.Text = Format(mdblSumaMontoHaber, "###,###,###,##0.00")
         'Limpia los campos
        LimpiarCamposDetalle
        
        ' Habilita el botón aceptar
        If optDebe.Enabled Then optDebe.SetFocus
    Else
        ' Envia mensaje
        MsgBox "La cuenta contable ya ha sido ingresado al detalle del Haber", _
                   vbInformation + vbOKOnly, "SGCcaijo - Asiento Manual"
        ' limpia la cuenta para  dar opcion a elegir
        cboCtaContable.SetFocus
        cmdAñadir.Enabled = False
        Exit Sub
    End If
End If

'Habilitar el boton AceptarModificar
HabilitarAceptarModificar

'Habilita el boton Cancelar
cmdCancelar.Enabled = True

End Sub



Private Sub LimpiarCamposDetalle()
'-------------------------------------------
'Propósito: Limpia los campos de la pantalla 2 (Detalle)
'-------------------------------------------
cboCtaContable.ListIndex = -1
txtCodContable.Text = ""
txtCodContable.BackColor = Obligatorio
txtMonto.Text = ""
txtMonto.BackColor = Obligatorio

End Sub

Private Function VerificarElemento(grdDebeHaberRec As MSFlexGrid) As Boolean
'------------------------------------------------------
'Propósito: Verificar la existencia de un Producto en el grdDetalle
'Recibe:    Nada
'Devuelve:  booleano que indica la existencia de la cboProdServ en el grd detalle
'------------------------------------------------------
'Nota:      llamado desde el evento click de cmdAñadir
Dim j As Integer

'Inicializamos a funcion asumiendo que Procuto NO esta en el grddetalle 1
VerificarElemento = False

' recorremos el grid detalle de Producto verificando la existencia de txtProd
For j = 1 To grdDebeHaberRec.Rows - 1
 If grdDebeHaberRec.TextMatrix(j, 0) = txtCodContable.Text Then
    VerificarElemento = True
    Exit Function
 End If
Next j

End Function



Private Sub cmdCancelar_Click()
Dim vObjeto As Variant

If gsTipoOperacionContabilidad = "Nuevo" Then
    'Inicializa La variable que acumula las cantidades del detalle
    mdblSumaMontoDebe = 0
    mdblSumaMontoHaber = 0
    txtTotalDebe.Text = "0.00"
    txtTotalHaber.Text = "0.00"
    
    ' Habilitar Datos generales y botones Aceptar, Cancelar
    cmdAceptarModificar.Enabled = False
    
    'Establece campos obligatorios
    EstableceCamposObligatorios
    
    ' Limpia el Grid
    grdDebe.Rows = 1
    grdHaber.Rows = 1
    
    ' Inicializa el grid
    iposH = 0
    iposD = 0
    gbCambioCelda = False
    
Else 'formulario para la modificacion de registros
    ' restaura los valores predeterminados del formulario en la modificación
    EstableceCamposObligatorios
    txtGlosa = msGlosa
    ' establece las celdas del grd
    grdDebe.Rows = 1
    grdHaber.Rows = 1
    For Each vObjeto In mcolAsientosDebe
       grdDebe.AddItem Var30(vObjeto, 1) & vbTab & _
                         mcolDesCodPlanCont(Var30(vObjeto, 1)) & vbTab & _
                         Var30(vObjeto, 3)
    Next vObjeto
    For Each vObjeto In mcolAsientosHaber
       grdHaber.AddItem Var30(vObjeto, 1) & vbTab & _
                         mcolDesCodPlanCont(Var30(vObjeto, 1)) & vbTab & _
                         Var30(vObjeto, 3)
    Next vObjeto
    ' carga los totales del debe y haber
    CargarTotales
    
    ' Inicializa el grid
    iposH = 0
    iposD = 0
    gbCambioCelda = False
End If
    
End Sub

Private Sub cmdEliminar_Click()
Dim i As Integer

'Verifica SI hay alguna fila seleccionada
If grdDebe.CellBackColor = vbDarkBlue And grdDebe.Row > 0 Then
      
      ' elimina la fila seleccionada del grid
    If grdDebe.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdDebe.RemoveItem grdDebe.Row
    Else
            ' estable vacío el grddetalle
            grdDebe.Rows = 1
    End If
    
    ' Inicializa la suma de las cantidades del detalle en 0
    mdblSumaMontoDebe = 0
        
    ' Verifica SI el grid esta vacio para habilitar o deshabilitar cmdEliminar
    If grdDebe.Rows = 1 Then
        cmdAceptarModificar.Enabled = False
        cmdCancelar.Enabled = False
        
    Else    'Actualiza monto total detalle
        For i = 1 To grdDebe.Rows - 1
             mdblSumaMontoDebe = mdblSumaMontoDebe + Val(Var37(grdDebe.TextMatrix(i, 2)))
        Next
        cmdCancelar.Enabled = True
    End If
    
    'Asigna el monto al txtTotalDebe
    txtTotalDebe.Text = Format(Var37(mdblSumaMontoDebe), "###,###,###,##0.00")
    
    ' Inicializa la variable del grid debe
    iposD = 0
    
ElseIf grdHaber.CellBackColor = vbDarkBlue And grdHaber.Row > 0 Then
               
    ' Elimina la fila selccionada del Grid
    If grdHaber.Rows > 2 Then
            ' elimina la fila seleccionada del grid
            grdHaber.RemoveItem grdHaber.Row
    Else
            ' estable vacío el grddetalle
            grdHaber.Rows = 1
    End If
    
    
    ' Inicializa la suma de las cantidades del detalle en 0
    mdblSumaMontoHaber = 0
    
    ' Verifica SI el grid esta vacio para habilitar o deshabilitar cmdEliminar
    If grdHaber.Rows = 1 Then
        cmdAceptarModificar.Enabled = False
    Else    'Actualiza monto total detalle
        For i = 1 To grdHaber.Rows - 1
             mdblSumaMontoHaber = mdblSumaMontoHaber + Val(Var37(grdHaber.TextMatrix(i, 2)))
        Next
    End If

    txtTotalHaber.Text = Format(Var37(mdblSumaMontoHaber), "###,###,###,##0.00")
    
   ' Inicializa la variable del grid Haber
    iposH = 0
End If


'Habilita el boton aceptar modificar
HabilitarAceptarModificar

End Sub

Private Sub cmdPCtaContable_Click()

If cboCtaContable.Enabled Then
    ' alto
     cboCtaContable.Height = CBOALTO
    ' focus a cbo
    cboCtaContable.SetFocus
End If

End Sub

Private Sub cmdSalir_Click()
' cierra el formulario
Unload Me

End Sub


Private Sub Form_Load()
Dim sSQL As String

'Actualiza fecha de formulario
If gsTipoOperacionContabilidad = "Nuevo" Then

    ' Se carga la fecha de trabajo
    mskFecha = gsFecTrabajo
    'Determina el siguiente número del asiento manual
    txtNumAsiento.Text = Conta4("MA", gsFecTrabajo)
    'Desabilita el control añadir
    cmdAñadir.Enabled = False
    cmdAceptarModificar.Enabled = False
    cmdCancelar.Enabled = False
    'Inicializa las variables con 0
    mdblSumaMontoDebe = 0
    mdblSumaMontoHaber = 0
Else
    With frmCTASelAsientos.grdAsientoM
    
        'Asigna el numero de asiento seleccionado en el formulario de selección
        msNumAsiento = .TextMatrix(.Row, 0)
        msGlosa = .TextMatrix(.Row, 1)
        msFecha = .TextMatrix(.Row, 2)
        mskFecha = msFecha
        txtGlosa = msGlosa
    End With
End If

'Coloca los titulos de grdDebe y grdHaber
aTitulosColGrid = Array("Código", "Cuenta Contable", "Monto")
aTamañosColumnas = Array(600, 3800, 1030)
CargarGridTitulos grdDebe, aTitulosColGrid, aTamañosColumnas

aTitulosColGrid = Array("Código", "Cuenta Contable", "Monto")
aTamañosColumnas = Array(600, 3800, 1030)
CargarGridTitulos grdHaber, aTitulosColGrid, aTamañosColumnas

' Inicializa el grid
iposH = 0
iposD = 0
gbCambioCelda = False

' Carga las colecciones de contabilidad
   sSQL = "SELECT CodContable, CodContable & ' ' & Left(DescCuenta,55) FROM PLAN_CONTABLE " & _
          "WHERE (len(CodContable)=5) " _
          & " ORDER BY CodContable"
    CD_CargarColsCbo cboCtaContable, sSQL, mcolCodPlanCont, mcolDesCodPlanCont
          
If gsTipoOperacionContabilidad = "Nuevo" Then
    ' Define los campos obligatorios
    EstableceCamposObligatorios
Else
    ' Llama al procedimiento que se encarga de preparar el formulario _
   para una nueva modificación
   NuevaModificacion
    
End If

'Determina la forma a alinear
grdDebe.ColAlignment(1) = 1
grdHaber.ColAlignment(1) = 1

End Sub

Private Sub NuevaModificacion()
'---------------------------------------------------------ç
'Proposito: Prepara el formulario para una nueva modificacion
'Recibe: nada
'Entrega: nada
'---------------------------------------------------------
' Prepara el formulario para la modificación
If msNumAsiento <> Empty Then
    ' coloca los campos obligatorios
    EstableceCamposObligatorios
    ' carga los componentes necesarios y los controles del asiento
    CargacurCamposAsiento
    ' deshabilita los botones
    cmdAñadir.Enabled = False
    cmdAceptarModificar.Enabled = False
    cmdCancelar.Enabled = True
Else
   'NO hace nada manda mensaje de error
   MsgBox " Error al cargar formulario, codigo de asiento NO pasado" & Chr(13) _
    & "Cierre el formualrio", , "Contabilidad-Asiento Manual"
End If

End Sub

Private Sub CargacurCamposAsiento()
'---------------------------------------------------------ç
'Proposito: Carga los campos del cursor asiento asi como _
            los controles del formulario
'Recibe: nada
'Entrega: nada
'---------------------------------------------------------
Dim sSQL As String

' Carga la consulta que realiza las modificaciones
sSQL = "SELECT IdAsiento,CodContable,DebeHaber,Monto " _
    & "FROM CTB_ASIENTOS_DET " _
    & "WHERE NumAsiento='" & msNumAsiento & "' " _
    & "ORDER BY IdAsiento, DebeHaber"
    
' Ejecuta la sentencia
mcurAsiento.SQL = sSQL
If mcurAsiento.Abrir = HAY_ERROR Then End
' Verifica SI realizó carrectamente la consulta
If mcurAsiento.EOF Then ' NO se realizó la consulta
    MsgBox "Error en carga de asiento, consulte al admiistrador", , "Contabilidad - Asiento Manual"
    mcurAsiento.Cerrar
    Unload Me
Else
    ' rellena los controles con los valores sacados de la base de datos
   CargarControlesAsiento
    ' cargar los montos del debe y haber
   CargarTotales
   
End If

End Sub

Private Sub CargarControlesAsiento()
'---------------------------------------------------------
'Proposito: Carga los controles del formulario
'Recibe: nada
'Entrega: nada
'---------------------------------------------------------
Dim sDescripcionCta As String
' carga datos generales
txtNumAsiento = msNumAsiento
txtGlosa = msGlosa
' carga los registros de detalle
Do While Not mcurAsiento.EOF
   'recorre el cursor consulta
                                
   'Coloca en grid los sientos detalle de acuerdo SI va al debe o haber
   Select Case mcurAsiento.campo(2)
   Case "D"
        ' carga la coleccion de det asientos Debe IdAsiento,CodCont,DebeHaber,Monto "
        mcolAsientosDebe.Add Key:=mcurAsiento.campo(1), Item:=mcurAsiento.campo(1) & "¯" _
                        & Format(mcurAsiento.campo(3), "###,###,##0.00")

        ' asiento al debe
        sDescripcionCta = mcolDesCodPlanCont(mcurAsiento.campo(1))
        grdDebe.AddItem mcurAsiento.campo(1) & vbTab & _
                        sDescripcionCta & vbTab & _
                        Format(mcurAsiento.campo(3), "###,###,##0.00")
                        
   Case "H"
        ' carga la coleccion de det asientos Haber IdAsiento,CodCont,DebeHaber,Monto "
        mcolAsientosHaber.Add Key:=mcurAsiento.campo(1), Item:=mcurAsiento.campo(1) & "¯" _
                        & Format(mcurAsiento.campo(3), "###,###,##0.00")

        ' asiento al haber
        sDescripcionCta = mcolDesCodPlanCont(mcurAsiento.campo(1))
        grdHaber.AddItem mcurAsiento.campo(1) & vbTab & _
                        sDescripcionCta & vbTab & _
                        Format(mcurAsiento.campo(3), "###,###,##0.00")
   Case Else
        MsgBox "Error en la BD : NO esta correcto el campo debe-haber" _
              & Chr(13) & "Consulte al administrador  ", , "Contabilidad- Asiento Manual"
   
   End Select
   
   ' mueve al siguienete elemento del cursor
   mcurAsiento.MoverSiguiente

Loop
   'Se cierra el cursor
   mcurAsiento.Cerrar

End Sub


Private Sub CargarOptDebeHaber()
'Proposito : Pone el valor de Debe,Haber dependiendo de los _
             controles opt
'             NumAsiento,CodCont,DebeHaber,Monto,IdCod,Origen,Glosa
If mcurAsiento.campo(2) = "D" Then
    optDebe.Value = True
Else
   If mcurAsiento.campo(2) = "H" Then
       optHaber.Value = True
   End If
   
End If
End Sub

Private Sub CargarTotales()
'----------------------------------------------------------------------------
'Propósito  : Carga los totales del debe y el haber
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
Dim i As Integer
' inicializa los totales del debe y el haber
mdblSumaMontoDebe = 0
mdblSumaMontoHaber = 0

For i = 1 To grdDebe.Rows - 1
    ' suma los montos de los asientos detalle para el debe
    mdblSumaMontoDebe = mdblSumaMontoDebe + Val(Var37(grdDebe.TextMatrix(i, 2)))
Next i
For i = 1 To grdHaber.Rows - 1
    ' suma los montos de los asientos detalle para el debe
    mdblSumaMontoHaber = mdblSumaMontoHaber + Val(Var37(grdHaber.TextMatrix(i, 2)))
Next i
' asigna los montos a los controles totales del debe y el haber
txtTotalDebe = Format(mdblSumaMontoDebe, "###,###,##0.00")
txtTotalHaber = Format(mdblSumaMontoHaber, "###,###,##0.00")

End Sub

Private Sub HabilitarAceptarModificar()
'----------------------------------------------------------------------------
'Propósito  : Determina SI se debe habilitar o NO el boton aceptar Modificar luego _
              de realizar distintas operaciones
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------------------------------------------
' Inicializa los datos boleanos para evaluar SI se habilita o NO el botón aceptar modificar
mbCambioAsiento = False
mbCambioDetalle = False
' verifica SI existen campos obligatorios
If txtGlosa.BackColor <> vbWhite Or _
   Val(Var37(txtTotalDebe.Text)) <> Val(Var37(txtTotalHaber.Text)) _
   Or grdDebe.Row < 1 Then
   ' NO estran los datos correctos
   cmdAceptarModificar.Enabled = False
Else
   If gsTipoOperacionContabilidad = "Nuevo" Then
    ' El tipo de operacion realizada en el formulario es nuevo
      cmdAceptarModificar.Enabled = True
   Else
    ' IdAsiento,CodCont,DebeHaber,Monto
    ' Verifica SI se han cambiado los datos
     If txtGlosa <> msGlosa Then
        ' se han cambiado los datos del asiento
        mbCambioAsiento = True
     End If
     ' verifica SI se cambiaron los asientos detalle
     If CambioDetalle = True Then
        mbCambioDetalle = True
     End If
     ' verifica SI hubo cambios
     If mbCambioAsiento Or mbCambioDetalle Then
        ' hubo cambios
        cmdAceptarModificar.Enabled = True
     Else
        ' NO hubo cambios
        cmdAceptarModificar.Enabled = False
     End If
   End If
End If

End Sub

Private Function CambioDetalle() As Boolean
'--------------------------------------------------------------
'Propósito: Verifica SI se modificó el detalle del registro de Egreso
'Recibe:    booleano  que indica SI se modificó el detalle
'Devuelve:  Nada
'--------------------------------------------------------------
' nota :    Llamado desde procedimiento habilitar boton aceptarModificar
Dim i As Integer
Dim sRegAsiento As String
On Error GoTo ErrClaveCol

'inicializa la funcion

CambioDetalle = False

'verifica SI se eliminó o agregó algun elemento en debe
If (mcolAsientosDebe.Count <> grdDebe.Rows - 1) And (grdDebe.Rows > 1) Then
    'hubo una modificacion en el detalle del egreso
    CambioDetalle = True
    Exit Function
Else 'verifica SI el grdDetalle es vacio
    If grdDebe.Rows <= 1 Then Exit Function
End If
'verifica SI se eliminó o agregó algun elemento en haber
If (mcolAsientosHaber.Count <> grdHaber.Rows - 1) And (grdHaber.Rows > 1) Then
    'hubo una modificacion en el detalle del egreso
    CambioDetalle = True
    Exit Function
Else 'verifica SI el grdDetalle es vacio
    If grdHaber.Rows <= 1 Then Exit Function
End If

'verifica SI se modifico los registros detalle en el grd debe
For i = 1 To grdDebe.Rows - 1
   'carga el registrro de la forma IdAsiento,CodCont,DebeHaber,Monto
    sRegAsiento = grdDebe.TextMatrix(i, 0) & "¯" & _
                  grdDebe.TextMatrix(i, 2)
   'verifica SI el registro esta en la colección que almacena los registros _
   detalle originales del egreso
   If sRegAsiento <> mcolAsientosDebe.Item(grdDebe.TextMatrix(i, 0)) Then
        CambioDetalle = True ' se modifico el detalle sale de la funcion
        Exit Function
   End If
Next i
'verifica SI se modifico los registros detalle en el grd haber
For i = 1 To grdHaber.Rows - 1
   'carga el registrro de la forma IdAsiento,CodCont,DebeHaber,Monto
    sRegAsiento = grdHaber.TextMatrix(i, 0) & "¯" & _
                  grdHaber.TextMatrix(i, 2)
   'verifica SI el registro esta en la colección que almacena los registros _
   detalle originales del egreso
   If sRegAsiento <> mcolAsientosHaber.Item(grdHaber.TextMatrix(i, 0)) Then
        CambioDetalle = True ' se modifico el detalle sale de la funcion
        Exit Function
   End If
Next i

'-----------------------------------------------------
ErrClaveCol:

    If Err.Number = 5 Then ' Error al acceder a elemento de colCodDesc
        CambioDetalle = True 'se modifico el detalle del egreso, sale de la funcion
        Exit Function
    End If

End Function



Private Function CalcularSigCodigo() As Double
'----------------------------------------------------------------------------
'Propósito  : Determina el ultimo registro e incrementa en 1 el campo asiento
'             manula
'Recibe     : Nada
'Devuelve   : Número del asiento manual incremtado en 1
'----------------------------------------------------------------------------

Dim sSQL As String
Dim curNroAsientoM As New clsBD2
Dim iNumSec As Integer


'Se carga un string con el ultimo registro del campo NumAsiento
sSQL = ""
sSQL = "SELECT Max(IdAsiento) FROM CTB_ASIENTOS_DET "
curNroAsientoM.SQL = sSQL

' Averigua el ultimo orden de ingreso
If curNroAsientoM.Abrir = HAY_ERROR Then
  Unload Me
End If
  
'Separa los cuatro últimos caracteres del maximo número de asientos manuales
If IsNull(curNroAsientoM.campo(0)) Then ' NO hay registros
    CalcularSigCodigo = 1
Else
    
    CalcularSigCodigo = Val(curNroAsientoM.campo(0)) + 1
    
End If

'Cierra el cursor
curNroAsientoM.Cerrar

End Function


Private Sub EstableceCamposObligatorios()
'--------------------------------------------------------------
'Proposito: Pone los campos obligatorios
'Recibe: Nada
'Entrega : Nada
'--------------------------------------------------------------

' Pone los demas campos obligatorios a color amarillo

txtCodContable.BackColor = Obligatorio
txtMonto.BackColor = Obligatorio
txtGlosa.BackColor = Obligatorio
txtGlosa.Text = Empty
txtMonto.Text = Empty
txtCodContable = Empty
cboCtaContable.ListIndex = -1
grdDebe.Rows = 1
grdHaber.Rows = 1

txtTotalDebe.Text = "0.00"
txtTotalHaber.Text = "0.00"


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Se limpian las colecciones
Set mcolCodPlanCont = Nothing
Set mcolDesCodPlanCont = Nothing

Set mcolAsientosDebe = Nothing
Set mcolAsientosHaber = Nothing


End Sub

Private Sub grdDebe_Click()

If grdDebe.Row > 0 And grdDebe.Row < grdDebe.Rows Then
    ' Verifica si la fila del haber esta seleccionada
    If grdHaber.CellBackColor = vbDarkBlue Then
       DesmarcarFilaGRID grdHaber
    End If
    ' Marca la fila seleccionada
    MarcarFilaGRID grdDebe, vbWhite, vbDarkBlue
End If

End Sub

Private Sub grdDebe_EnterCell()

If iposD <> grdDebe.Row Then
    '  Verifica si es la última fila
    If grdDebe.Row > 0 And grdDebe.Row < grdDebe.Rows Then
         If gbCambioCelda = False Then
            ' Verifica si la fila del haber esta seleccionada
            If grdHaber.CellBackColor = vbDarkBlue Then
               DesmarcarFilaGRID grdHaber
            End If
       
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdDebe, iposD
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    iposD = grdDebe.Row
End If

End Sub

Private Sub grdDebe_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Llama al proceso que edita el registro
    grdDebe_DblClick
 End If
 
End Sub

Private Sub grdDebe_DblClick()

' Selecciona toda la iFila
If grdDebe.Rows > 1 Then
    
    ' Verifica si esta seleccionado
    If grdDebe.CellBackColor <> vbDarkBlue Then
        ' Verifica si la fila del haber esta seleccionada
        If grdHaber.CellBackColor = vbDarkBlue Then
           DesmarcarFilaGRID grdHaber
        End If
        ' Marca la fila del grid
        MarcarFilaGRID grdDebe, vbWhite, vbDarkBlue
        Exit Sub
    End If

    'Recupera en el msIdAlmacen el codigo del producto seleccionado
    txtCodContable.Text = grdDebe.TextMatrix(grdDebe.RowSel, 0)
    txtMonto.Text = grdDebe.TextMatrix(grdDebe.RowSel, 2)
    
    'Selecciona el optHaber
    optDebe.Value = True
   
    'Elimina la fila seleccionados en el grid
    cmdEliminar_Click
    
    'Marca o desmarca la iFila
    If grdDebe.Rows > 1 Then DesmarcarFilaGRID grdDebe
    
    'Coloca el focus a txtMonto
    If txtMonto.Enabled Then txtMonto.SetFocus
    
End If

End Sub

Private Sub grdHaber_Click()

If grdHaber.Row > 0 And grdHaber.Row < grdHaber.Rows Then
    ' Verifica si la fila del Debe esta seleccionada
    If grdDebe.CellBackColor = vbDarkBlue Then
       DesmarcarFilaGRID grdDebe
    End If
    
    ' Marca la fila seleccionada
    MarcarFilaGRID grdHaber, vbWhite, vbDarkBlue
End If

End Sub

Private Sub grdHaber_EnterCell()

If iposH <> grdHaber.Row Then
    '  Verifica si es la última fila
    If grdHaber.Row > 0 And grdHaber.Row < grdHaber.Rows Then
         If gbCambioCelda = False Then
            ' Verifica si la fila del Debe esta seleccionada
            If grdDebe.CellBackColor = vbDarkBlue Then
               DesmarcarFilaGRID grdDebe
            End If
         
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdHaber, iposH
            gbCambioCelda = False
         End If
    End If
    ' Actualiza el valor de la fila
    iposH = grdHaber.Row
End If

End Sub

Private Sub grdHaber_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    ' Llama al proceso que cambia la verificación de un producto
    grdHaber_DblClick
 End If
 
End Sub

Private Sub grdHaber_DblClick()
' Selecciona toda la iFila
If grdHaber.Rows > 1 Then
    ' Verifica si esta seleccionado
    If grdHaber.CellBackColor <> vbDarkBlue Then
        ' Verifica si la fila del Debe esta seleccionada
        If grdDebe.CellBackColor = vbDarkBlue Then
           DesmarcarFilaGRID grdDebe
        End If
        
        ' Marca la fila
        MarcarFilaGRID grdHaber, vbWhite, vbDarkBlue
        Exit Sub
    End If
    
    'Recupera en el msIdAlmacen el codigo del producto seleccionado
    txtCodContable.Text = grdHaber.TextMatrix(grdHaber.RowSel, 0)
    txtMonto.Text = grdHaber.TextMatrix(grdHaber.RowSel, 2)
    
    'Selecciona el optHaber
    optHaber.Value = True
   
    'Elimina la fila seleccionados en el grid
    cmdEliminar_Click
    
    'Marca o desmarca la iFila
    If grdHaber.Rows > 1 Then DesmarcarFilaGRID grdHaber
    
    'Coloca el focus a txtMonto
    If txtMonto.Enabled Then txtMonto.SetFocus
    
End If

End Sub

Private Sub Image1_Click()
'Carga la Var48
Var48
End Sub

Private Sub optDebe_Click()

' Habilita boton añadir
HabilitarAñadir

End Sub

Private Sub optDebe_KeyPress(KeyAscii As Integer)

' Si presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optHaber_Click()

' Habilita boton añadir
HabilitarAñadir

End Sub

Private Sub optHaber_KeyPress(KeyAscii As Integer)

' Si presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtCodContable_Change()

' SI procede, se actualiza descripción correspondiente a código introducido
CD_ActDesc cboCtaContable, txtCodContable, mcolDesCodPlanCont

' Verifica SI el campo esta vacio
If txtCodContable.Text <> "" And cboCtaContable.Text <> "" Then
   ' Los campos coloca a color blanco
   txtCodContable.BackColor = vbWhite
Else
  'Marca los campos obligatorios
   txtCodContable.BackColor = Obligatorio
End If

'habilitar el boton añadir
HabilitarAñadir

End Sub

Private Sub HabilitarAñadir()
'----------------------------------------------------------------------------
'Propósito: Se habilita "Añadir" SI se han rellenado los campos obligatorios
'Recibe:   Nada
'Devuelve: Nada
'----------------------------------------------------------------------------

If txtCodContable.BackColor <> Obligatorio And txtMonto.BackColor <> Obligatorio Then
    'Habilita el botón aceptar
    cmdAñadir.Enabled = True
Else
    cmdAñadir.Enabled = False
End If

End Sub

Private Sub txtCodContable_KeyPress(KeyAscii As Integer)

' Si presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtGlosa_Change()

' Verifica SI el campo esta vacio
If txtGlosa.Text <> Empty And InStr(txtGlosa, "'") = 0 Then
' El campos coloca a color blanco
   txtGlosa.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtGlosa.BackColor = Obligatorio
End If

'habilita el botón aceptar
HabilitarAceptarModificar

End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)

' Si presiona enter pasa al siguiente control
If KeyAscii = 13 Then
    SendKeys vbTab
Else
    'Convierte a mayusculas la descripción
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub txtMonto_Change()

' Verifica SI el campo esta vacio
If txtMonto.Text <> "" Then
' El campos coloca a color blanco
   txtMonto.BackColor = vbWhite
Else
'Marca los campos obligatorios
   txtMonto.BackColor = Obligatorio
End If

'habilitar el boton añadir
HabilitarAñadir

End Sub

Private Sub txtMonto_GotFocus()
'Define el tamaño del txtMonto y elimina comas
txtMonto.MaxLength = 12
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
   
txtMonto.MaxLength = 14
If txtMonto.Text <> "" Then
   'Da formato de moneda
   txtMonto.Text = Format(Val(Var37(txtMonto.Text)), "###,###,###,##0.00")
Else
   txtMonto.BackColor = Obligatorio
End If

End Sub

Function CtaAnulada() As Boolean
  Dim curCtaAnulada As New clsBD2
  
  ' ///// VERIFICAR EL ESTADO DE LA CTA (ANULADA / NO ANULADA)
  curCtaAnulada.SQL = "SELECT ANULADO " _
                  & "FROM TIPO_CUENTASBANC " _
                  & "WHERE CODCONT = '" & txtCodContable & "'"
  
  ' Se abre el cursor y se tratan errores
  If curCtaAnulada.Abrir = HAY_ERROR Then
    End
  End If
  
  If curCtaAnulada.campo(0) = "SI" Then  'ANULADA
    CtaAnulada = True
  Else
    CtaAnulada = False
  End If
  
  curCtaAnulada.Cerrar
End Function

