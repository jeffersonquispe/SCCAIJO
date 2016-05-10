VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCTBAsientosCierre 
   Caption         =   "Contabilidad- Asientos de Cierre Contable Anual"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   HelpContextID   =   105
   Icon            =   "SCCTBAsientosCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Siguiente "
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   120
      TabIndex        =   6
      Top             =   -10
      Width           =   10935
      Begin VB.TextBox txtNumAsiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   650
         Width           =   8655
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "NºAsiento:"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   3720
         Picture         =   "SCCTBAsientosCierre.frx":08CA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label7 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   495
      End
      Begin VB.Label lblPasos 
         AutoSize        =   -1  'True
         Caption         =   "PASO 1"
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   7920
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdAsientos 
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      HighLight       =   0
      FillStyle       =   1
   End
End
Attribute VB_Name = "frmCTBAsientosCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim miPaso As Single
Dim mbTodoOK As Boolean
Dim mdblTotalDebe As Double
Dim mdblTotalHaber As Double

Private Function DatosOK() As Boolean

' Verifica si el total del debe y el total del haber sean iguales
If Val(mdblTotalDebe) <> Val(mdblTotalHaber) Then
    'Mensaje de error
    MsgBox "Los montos totales son diferentes", vbCritical + vbOKOnly, "SGCcaijo-Cierre Contable Anual"
    DatosOK = False
    Exit Function
End If

' Todo Ok
DatosOK = True

End Function

Private Sub cmdAceptar_Click()
'Verifica si los datos estan OK
If DatosOK = False Then
    'Termina la ejecucion del procedimiento
    Exit Sub
End If

If Val(mdblTotalDebe) > 0.0001 Then

    'Actualiza la transaccion
     Var8 1, gsFormulario
    
    Select Case miPaso
    Case 1, 2, 3, 4, 5
        'Guarda el asiento de cierre
         Conta6 txtNumAsiento.Text, FechaAMD("31/12/" & mskAnio.Text), _
                               Empty, Empty, txtGlosa.Text, "C" & miPaso
     Case 6
         ' Guarda el asiento de cierre
         Conta6 txtNumAsiento.Text, FechaAMD("01/01/" & Str(Val(mskAnio.Text) + 1)), _
                               Empty, Empty, txtGlosa.Text, "AA"
         ' Indica que ha finalizado el proceso de cierre
         mbTodoOK = True
    End Select
    
    'Guarda el detalle del asiento
     GuardaDetAsientoCierre
     
    'Actualiza la transaccion
     Var8 -1, Empty
     
    ' MEnsaje de Ok
     MsgBox "El Asiento de Cierre :" & Chr(13) _
     & txtGlosa & "." & Chr(13) _
     & "Fué Realizado Correctamente", vbInformation + vbOKOnly, "SGCcaijo-Cierre Contable Anual"
Else
    ' MEnsaje de Ok
     MsgBox "No se guardo este asiento. No existe movimiento contable" & Chr(13) _
           & "Se continua con el siguiente paso del cierre contable", vbInformation + vbOKOnly, "SGCcaijo-Cierre Contable Anual"
End If

'Verifica si el paso es 6
 If miPaso = 6 Then
    'Termina la ejecucion del procedimiento
    Unload Me
    Exit Sub
 End If
 
'Selecciona el tipo de operación
 GeneraAsientoCierre miPaso + 1, mskAnio

End Sub

Private Sub GuardaDetAsientoCierre()
'------------------------------------------------------------
' Propósito : Guarda el detalle del asientos de cierre de la clase 6 gastos
' Recibe    : Nada
' Entrega   : Nada
'------------------------------------------------------------
Dim i As Integer
Dim dblIdAsiento As Integer

'Determina el siguiente codigo del asiento manual
dblIdAsiento = 1

'Para i=1 hasta total de filas - 2
For i = 1 To grdAsientos.Rows - 2
    
    'Verifica si es del debe o haber
    If grdAsientos.TextMatrix(i, 4) = "D" Then
        'Guarda el detalle del asiento
        Conta10 dblIdAsiento, txtNumAsiento.Text, grdAsientos.TextMatrix(i, 0), _
        "D", Var37(grdAsientos.TextMatrix(i, 2))
    Else
        'Guarda el detalle del asiento
        Conta10 dblIdAsiento, txtNumAsiento.Text, grdAsientos.TextMatrix(i, 0), _
        "H", Var37(grdAsientos.TextMatrix(i, 3))
    End If
    
    'Incrementa el contador
    dblIdAsiento = dblIdAsiento + 1
Next i

End Sub

Private Sub cmdCancelar_Click()
' Descarga el formulario
  Unload Me
End Sub

Private Sub Form_Load()
' Carga los títulos del grid
' Cuenta, Descripción, debe, haber
aTitulosColGrid = Array("Cuenta", "Descripción", "Debe", "Haber", "DH")
aTamañosColumnas = Array(700, 4500, 1200, 1200, 0)
CargarGridTitulos grdAsientos, aTitulosColGrid, aTamañosColumnas

' Inicializa la variable mbTodoOK
mbTodoOK = False

End Sub

Public Sub GeneraAsientoCierre(iPaso As Single, sAnio As String)
'------------------------------------------------------------
' Propósito : Genera asientos de cierre del año
' Recibe    : iPaso, Pasos de cierre; sAnio, Año del cierre
' Entrega   : Nada
'------------------------------------------------------------
' Carga el Paso en la variable de modulo
miPaso = iPaso
'Selecciona el paso de cierre contable anual
Select Case iPaso
Case 1, 2, 3, 4
    'Inicializa controles
    Manejacontroles iPaso, sAnio
    'Asiento de cierre 1,2,3,4
    Conta48 sAnio, iPaso
    'Cargar el grid con los datos de la colección
    CargarGridAsiento "Cierre"
     'Cierra la coleccion
    Set gcolAsientoCierre = Nothing
Case 5
    'Inicializa controles
    Manejacontroles iPaso, sAnio
    'Asiento de cierre 5
    Conta48 sAnio, iPaso
    'Cargar el grid con los datos de la colección
    CargarGridAsiento "Cierre"
    'Cierra la coleccion
Case 6
    'Inicializa controles
    Manejacontroles iPaso, sAnio
    'Cargar el grid con los datos de la colección
    CargarGridAsiento "Apertura"
    'Cierra la coleccion
    Set gcolAsientoCierre = Nothing
End Select

End Sub

Private Sub Manejacontroles(iPaso As Single, sAnio As String)
'Maneja los controles del formulario
Select Case iPaso
Case 1
    'Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asientos de Cierre Contable Anual Paso 1"
    lblPasos.Caption = "Cerrando Gastos- Clase 6"
    txtGlosa.Text = "POR EL TRASLADO DE LOS GASTOS A RESULTADOS DEL EJERCICIO"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Conta4("MA", "31/12/" & sAnio)
Case 2
    'Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asientos de Cierre Contable Anual Paso 2"
    lblPasos.Caption = "Resultados del Ejercicio- Clase 7"
    txtGlosa.Text = "POR EL TRASLADO DE LOS INGRESOS A RESULTADOS DEL EJERCICIO"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Conta4("MA", "31/12/" & sAnio)
        
Case 3
    'Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asientos de Cierre Contable Anual Paso 3"
    lblPasos.Caption = "Mayorización de la Clase 8"
    txtGlosa.Text = "POR EL TRASLADO DE RESULTADOS DEL EJERCICIO A UTILIDADES O PERDIDAS"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Conta4("MA", "31/12/" & sAnio)
Case 4
    ' Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asientos de Cierre Contable Anual Paso 4"
    lblPasos.Caption = "Cerrando Cuentas de Costos- Clases 9 y 7"
    txtGlosa.Text = "POR EL TRASLADO DE CUENTAS DE COSTOS A CARGAS IMPUTABLES"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Conta4("MA", "31/12/" & sAnio)
Case 5
    ' Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asientos de Cierre Contable Anual Paso 5"
    lblPasos.Caption = "Cerrando Cuentas del Activo, Pasivo, Patrimonio - Clases 1,2,3,4,5"
    txtGlosa.Text = "POR EL ACTIVO, PASIVO Y PATRIMONIO AL CIERRE DEL PERIODO"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Conta4("MA", "31/12/" & sAnio)
Case 6
    ' Cambia la etiquetas del formulario
    Me.Caption = "Contabilidad- Asiento de Apertura- Paso 6"
    lblPasos.Caption = "Generando el Asiento de Apertura del Periodo Siguiente"
    txtGlosa.Text = "POR EL ACTIVO, PASIVO Y PATRIMONIO AL INICIO DEL PERIODO"
    ' Genera el numero de asiento auxiliar
    txtNumAsiento.Text = Right(Str(Val(sAnio) + 1), 2) & "01MA0001"
End Select

'Titulo
grdAsientos.Rows = 1

'Actualzia controles
mskAnio.Text = sAnio

End Sub

Private Sub CargarGridAsiento(sTipo As String)
'------------------------------------------------------------
' Propósito : Carga el grid con los datos de la colección gcolAsientoCierre
' Recibe    : sTipo si es un asiento de Cierre o Apertura
' Entrega   : Nada
'------------------------------------------------------------
Dim MiObjeto As Variant

'Inicializa las variables
mdblTotalDebe = 0
mdblTotalHaber = 0

'Limpia el grd
grdAsientos.Rows = 1
'Recorre la colección
For Each MiObjeto In gcolAsientoCierre
    If sTipo = "Cierre" Then ' Carga los elementos del asiento de cierre
        If Var30(MiObjeto, 2) = "D" Then
            'Añade al grid los datos
            grdAsientos.AddItem Var30(MiObjeto, 1) & vbTab & _
            Var30(MiObjeto, 4) & vbTab & _
            Format(Var30(MiObjeto, 3), "###,###,##0.00") & vbTab & vbTab & "D"
            'Acumula los totales debe
            mdblTotalDebe = mdblTotalDebe + Val(Var30(MiObjeto, 3))
        Else
            'Añade al grid los datos
            grdAsientos.AddItem Var30(MiObjeto, 1) & vbTab & _
            Var30(MiObjeto, 4) & vbTab & vbTab & _
             Format(Var30(MiObjeto, 3), "###,###,##0.00") & vbTab & "H"
            'Acumula los totales haber
            mdblTotalHaber = mdblTotalHaber + Val(Var30(MiObjeto, 3))
        End If
    ElseIf sTipo = "Apertura" Then ' Cruza y carga los elementos del asiento de cierre
        If Var30(MiObjeto, 2) = "D" Then
            'Añade al grid los datos
            grdAsientos.AddItem Var30(MiObjeto, 1) & vbTab & _
            Var30(MiObjeto, 4) & vbTab & vbTab & _
             Format(Var30(MiObjeto, 3), "###,###,##0.00") & vbTab & "H"
            'Acumula los totales haber
            mdblTotalHaber = mdblTotalHaber + Val(Var30(MiObjeto, 3))
        Else
            'Añade al grid los datos
            grdAsientos.AddItem Var30(MiObjeto, 1) & vbTab & _
            Var30(MiObjeto, 4) & vbTab & _
            Format(Var30(MiObjeto, 3), "###,###,##0.00") & vbTab & vbTab & "D"
            'Acumula los totales debe
            mdblTotalDebe = mdblTotalDebe + Val(Var30(MiObjeto, 3))
        End If
    End If
Next MiObjeto

'Añade los totales al grid
 grdAsientos.AddItem vbTab & "TOTALES" & vbTab & Format(mdblTotalDebe, "###,###,##0.00") _
 & vbTab & Format(mdblTotalHaber, "###,###,##0.00")

' Coloca color al grid
grdAsientos.Row = grdAsientos.Rows - 1
MarcarFilaGRID grdAsientos, vbBlack, vbGray

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Indica en la variable que esta saliendo
Set gcolAsientoCierre = Nothing
' Verifica que el proceso no ha finalizado correctamente
  If mbTodoOK = False Then
        'Actualiza la transaccion
       Var8 1, gsFormulario
    
       ' Elimina los registros de cierre realizados anteriormente
       Conta51 (mskAnio.Text)
            
       'Actualiza la transaccion
       Var8 -1, Empty
  End If

End Sub

Private Sub Image1_Click()

'Carga la Var48
Var48

End Sub

Private Sub txtGlosa_Change()
' Verifica SI el campo esta vacio
If txtGlosa.Text <> Empty Then
   ' Los campos coloca a color blanco
   txtGlosa.BackColor = vbWhite
Else
   'Marca los campos obligatorios, y limpia el combo
   txtGlosa.BackColor = Obligatorio
End If

'Habilita botón Aceptar
HabilitarBotonAceptar

End Sub

Private Sub HabilitarBotonAceptar()
'----------------------------------------------------------------------------
'PROPÓSITO: *Se habilita "Aceptar del formulario " en Ingreso de un Nuevo registro
'               Si se han rellenado los campos obligatorios
'           *Se habilita "Aceptar" en Modificacion
'               Si se han rellenado los campos, y Si se realizo algun cambio al registro
'----------------------------------------------------------------------------
 
' Verifica si se a introducido los datos obligatorios generales
If txtGlosa.BackColor <> vbWhite Then
   ' Algún obligatorio falta ser introducido
   ' Deshabilita el botón
   cmdAceptar.Enabled = False
   Exit Sub
End If

' Habilita botón aceptar
cmdAceptar.Enabled = True

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
