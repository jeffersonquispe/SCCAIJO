VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCBEGSelImp 
   Caption         =   "Caja y Bancos- Egreso- Selección de Impuestos"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "SCCBEGSelImpuestos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMonto 
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraImpuestos 
      Caption         =   "Seleccione los impuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid grdImp 
         Height          =   3495
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FillStyle       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   100
      Width           =   3975
      Begin VB.OptionButton optSinImpuestos 
         Caption         =   "Sin impuestos"
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optConImpuestos 
         Caption         =   "Con impuestos"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   7350
      Width           =   975
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "A&plicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6420
      TabIndex        =   3
      Top             =   2280
      Width           =   1000
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar  Aplicados"
      Height          =   375
      Left            =   6195
      TabIndex        =   8
      ToolTipText     =   "Vuelve a la pantalla anerior"
      Top             =   2880
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptarModificar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   7350
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdImptAplicados 
      Height          =   1815
      Left            =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      HighLight       =   0
      FillStyle       =   1
   End
   Begin VB.Frame fraImpuestosAplicados 
      Caption         =   "Impuestos aplicados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   225
      TabIndex        =   9
      Top             =   4920
      Width           =   7455
   End
   Begin MSFlexGridLib.MSFlexGrid GrdAFPs 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6840
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FillStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCBEGSelImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Colecciones usadas para operar los impuestos
Private mcolImpuestosAplicados As New Collection
Private mcolImpuestos As New Collection

' Variables de modulo
Private mdblSumaImpuestos As Double
Private mbGridObligatorio As Boolean
Private miFilaAnterior As Integer

Private Sub cmdAceptarModificar_Click()
Dim i As Integer
Dim sSQL As String
Dim modEgresoCajaBanco As New clsBD3

    ' Mensaje de conformidad
    If MsgBox("¿Esta conforme con los impuestos a aplicar?", vbQuestion + vbYesNo, _
              "Egresos- Selección de Impuestos") = vbYes Then
        
        'Carga una colección global con las retenciones elegidas
        CargarColImpSeleccionados
            
        'Se descarga el formulariod
        Unload Me
        
   End If

'Asigna a gbImpuestos a verdadero despues de ejecutar el frmImpuestos
gbImpuestos = True

End Sub


Private Sub CargarColImpSeleccionados()
'-------------------------------------------------------------------------------------------
'Propósito  : Carga en la colección global de retenciones elegidas en grdImptAplicados
'Recibe:      Nada
'Devuelve :    Colección global llena con las retenciones aplicadas al egreso
'-------------------------------------------------------------------------------------------
Dim i As Integer

' Limpia la colección global de los impuestos aplicados
Set gcolImpSel = Nothing

'verifica que haya retenciones en el grd retenciones
If grdImptAplicados.Rows > 1 Then

    'recorre el grd retenciones
    For i = 1 To grdImptAplicados.Rows - 1
    
        ' añade un impuesto a la colección  global que guarda las retenciones
        '"idImp", "ValorImpuesto", "Monto", codcont
        gcolImpSel.Add Item:=grdImptAplicados.TextMatrix(i, 0) & "¯" & grdImptAplicados.TextMatrix(i, 2) _
            & "¯" & Var37(grdImptAplicados.TextMatrix(i, 3)) & "¯" & grdImptAplicados.TextMatrix(i, 4) & "¯" & grdImptAplicados.TextMatrix(i, 1) _
            , Key:=grdImptAplicados.TextMatrix(i, 0)
            
    Next i
    
End If

End Sub

Private Sub CargarColImpuestosAplicado()
'------------------------------------------------------------------------
'Propósito  : Carga la colección con los impuestos aplicados y acumula la suma _
'             de los impuestos
'Recibe     : Nada
'Devuelve   : mcolImpuestosAplicados, Impuestos aplicados y mdblMontoImpuestos
'------------------------------------------------------------------------

Dim i As Integer
Dim j As Integer

i = 1
mdblSumaImpuestos = 0
Do While i <= grdImp.Rows - 1
    grdImp.Row = i
    If grdImp.CellBackColor = vbDarkBlue Then
      '"IdImpuesto", "Descripción", "% Impuesto", "Monto Retenido", "CodCont")
      
      '"Cód.", "Descripción", "Valor", "CodCont")
      If InStr(grdImp.TextMatrix(i, 1), "MIXTA") Then
        j = 1
        Do While j <= GrdAFPs.Rows - 1
          If grdImp.TextMatrix(i, 1) = GrdAFPs.TextMatrix(j, 1) Then
            '*/*/*/*/*/*/*/
            '*/*/*/*/*/*/*/   APORTE AFP
            '*/*/*/*/*/*/*/
            If gdblMontoTotal <= Val(GrdAFPs.TextMatrix(j, 8)) Then '/*/*/*/*/*/ Verificando si es menor de 1125 que es el TOPE para elegir porcentaje de aporte
              mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "1" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) & " APORTE" _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 2)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "1"
              'Acumula los impuestos seleccionados
              mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 2))
            Else
              mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "1" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) & " APORTE" _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 4)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "1"
              'Acumula los impuestos seleccionados
              mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 4))
            End If
            
            '*/*/*/*/*/*/*/
            '*/*/*/*/*/*/*/   COMISION AFP
            '*/*/*/*/*/*/*/
            mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "2" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) & " COMISION" _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 5)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "2"
            
            'Acumula los impuestos seleccionados
            mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 5))
            
            '*/*/*/*/*/*/*/
            '*/*/*/*/*/*/*/   SEGURO AFP
            '*/*/*/*/*/*/*/
            mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "3" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) & " SEGURO" _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 6)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "3"
            
            'Acumula los impuestos seleccionados
            mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 6))
          End If
          j = j + 1
        Loop
      Else ' /*/*/*/*/ Puede ser SNP o los otros impuestos
        If InStr(grdImp.TextMatrix(i, 1), "SNP") Then
          j = 1
          Do While j <= GrdAFPs.Rows - 1
            If grdImp.TextMatrix(i, 1) = GrdAFPs.TextMatrix(j, 1) Then ' /*/*/*/ BUSCA LA SNP EN EL GRID AFP
              If gdblMontoTotal <= Val(GrdAFPs.TextMatrix(j, 8)) Then '/*/*/*/*/*/ Verificando si es menor de 1125 que es el TOPE para elegir porcentaje de aporte
                mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "1" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 2)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "1"
                
                'Acumula los impuestos seleccionados
                mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 2))
              Else
                mcolImpuestosAplicados.Add Item:=GrdAFPs.TextMatrix(j, 0) & "1" _
                            & "¯" & GrdAFPs.TextMatrix(j, 1) _
                            & "¯" & Var37(GrdAFPs.TextMatrix(j, 4)) _
                            & "¯" & GrdAFPs.TextMatrix(j, 3), Key:=GrdAFPs.TextMatrix(j, 0) & "1"
                
                'Acumula los impuestos seleccionados
                mdblSumaImpuestos = mdblSumaImpuestos + Val(GrdAFPs.TextMatrix(j, 4))
              End If
            End If
            j = j + 1
          Loop
        '/*/*/*/* OTROS IMPUESTOS
        Else
          mcolImpuestosAplicados.Add Item:=grdImp.TextMatrix(i, 0) _
                        & "¯" & grdImp.TextMatrix(i, 1) _
                        & "¯" & Var37(grdImp.TextMatrix(i, 2)) _
                        & "¯" & grdImp.TextMatrix(i, 3), Key:=grdImp.TextMatrix(i, 0)
        
          'Acumula los impuestos seleccionados
          mdblSumaImpuestos = mdblSumaImpuestos + Val(grdImp.TextMatrix(i, 2))
        End If
      End If
    End If
    i = i + 1
Loop

End Sub

Private Sub cmdAplicar_Click()
Dim MiObjeto

'Vacia la colección
Set mcolImpuestosAplicados = Nothing

'Verifica si el grdImptAplicados  tiene datos
If grdImptAplicados.Rows > 1 Then

    'Vacia el grid si tiene algún registro
    grdImptAplicados.Rows = 1
    
End If

'Carga la colección y detemina la suma del los impuestos aplicados
CargarColImpuestosAplicado

'Agrega los impuestos seleccionados al grdImptAplicados con sus montos
'Recorre el grid para determinar que fila esta seleccionada
For Each MiObjeto In mcolImpuestosAplicados

    '"IdImpuesto", "Descripción", "% Impuesto", "Monto Retenido", "CodCont")
    '"Cód.", "Descripción", "Valor", "CodCont")
    ' Añade una fila al grd retenciones con los alculos necesarios
    grdImptAplicados.AddItem Var30(MiObjeto, 1) _
                  & vbTab & Var30(MiObjeto, 2) _
                  & vbTab & Var30(MiObjeto, 3) _
                  & vbTab & Format(CalcularImpuestos(Var30(MiObjeto, 3)), "###,###,###,##0.00") _
                  & vbTab & Var30(MiObjeto, 4)
Next

' Habilita el boton Eliminar
cmdEliminar.Enabled = True

'Inicializa los mbGridObligatorio
mbGridObligatorio = False

'Habilita el boton aceptar
HabilitaDeshabilitaBotonAceptar

'Ubica el cursor en la primera celda
UbicacionCursorGrid

End Sub

Private Sub UbicacionCursorGrid()
'--------------------------------------------------------------
'Propósito: Ubica el cursor en la primera celda del grd
'--------------------------------------------------------------

'Ubica la primera celda
grdImptAplicados.Col = 3
grdImptAplicados.Row = 1

'Verifica SI el txt esta vacio o con dato
If txtMonto.Text = "" Or Val(txtMonto.Text) <= 0 Then

    txtMonto.BackColor = Obligatorio
Else
    txtMonto.BackColor = vbWhite

End If

'Enter cel grid
EnterCellGrid

End Sub

Private Sub HabilitaDeshabilitaBotonAceptar()
'------------------------------------------------------------------------
'Propósito  : Habilita o deshabilita el boton aceptar de acuerdo a los datos
'Recibe     : Nada
'Devuelve   : Nada
'------------------------------------------------------------------------

'Deshabilita el boton cmdAceptarModificar
cmdAceptarModificar.Enabled = False

'Verifica si el optSinImpuestos no esta seleccionado
If optSinImpuestos.Value = False Then
    If grdImptAplicados.Rows > 1 Then
        cmdAceptarModificar.Enabled = True
    End If
Else
    cmdAceptarModificar.Enabled = True
End If

End Sub

Private Function CalcularImpuestos(dblImpuestos As Double)
'------------------------------------------------------------------------
'Propósito  : Calcula los impuestos
'Recibe     : dblImpuestos, Impuesto a definir
'Devuelve   : dblImpuestos,
'------------------------------------------------------------------------

'Verifica si se paga, retiene o registra
If gsRelacTributo = "Paga" Or gsRelacTributo = "Registra" Then
    'Se giro sin impuestos
    CalcularImpuestos = ImpuestosPagaRegistra(dblImpuestos)
    
ElseIf gsRelacTributo = "Retiene" Then

    'Se giro sin impuestos
    CalcularImpuestos = ImpuestosRetiene(dblImpuestos)
    
End If

End Function

Private Function ImpuestosRetiene(dblImpuesto As Double)
'------------------------------------------------------------------------
'Propósito  : Determina el impuesto para los documentos girados sin impuestos
'             y el impuesto se paga y registra
'Recibe     : Nada
'Devuelve   : Impuesto a monto a pagar para cada impuesto
'------------------------------------------------------------------------

If dblImpuesto <> 0 Then

    'Calcula el impuesto
    ImpuestosRetiene = gdblMontoTotal * dblImpuesto / 100
    
End If

End Function

Private Function ImpuestosPagaRegistra(dblImpuesto As Double)
'------------------------------------------------------------------------
'Propósito  : Determina el impuesto para los documentos girados sin impuestos
'             y el impuesto se paga y registra
'Recibe     : Nada
'Devuelve   : Impuesto a monto a pagar para cada impuesto
'------------------------------------------------------------------------

If dblImpuesto <> 0 Then

    'Calcula el impuesto
    ImpuestosPagaRegistra = (gdblMontoTotal / (1 + mdblSumaImpuestos / 100)) * dblImpuesto / 100
    
End If

End Function


Private Function MontoImpuestosNoPagaRegistra() As String
'------------------------------------------------------------------------
'Propósito  : Calcula el monto retenido de acuerdo al valor del impuesto _
'             y el valor del Egreso CA
'Recibe     : Nada
'Devuelve   : string que representa al monto calculado de la retención _
'             con formato numerico
'------------------------------------------------------------------------
Dim dMonto As Double
'inicializa la función
MontoImpuestosNoPagaRegistra = "0.00"
dMonto = 0

'Calcula el monto retenido del detalle
dMonto = Val(Var37(gdSumaMontoDetalle)) _
         * Val(Var37(grdImp.TextMatrix(grdImp.RowSel, 2))) / 100


'da formato string al monto
MontoImpuestosNoPagaRegistra = Format(dMonto, "###,###,##0.00")

End Function

Private Sub cmdEliminar_Click()

' Limpia el grid
grdImptAplicados.Rows = 1

' Habilita el boton aceptar
HabilitaDeshabilitaBotonAceptar

' Invisible el txtMonto
txtMonto.Visible = False

' Deshabilita el botón eliminar
cmdEliminar.Enabled = False

End Sub


Private Sub cmdSalir_Click()

' Mensaje de conformidad
If MsgBox("¿Está seguro que desea salir?", vbQuestion + vbYesNo, _
          "Egresos- Retención de Impuestos") = vbYes Then
     
    'Cierra el formulario
    Unload Me
    
End If

End Sub

Private Sub Form_Load()

' Se carga un array con los títulos de las columnas y otro con los tamaños para
aTitulosColGrid = Array("Cód.", "Descripción", "Valor", "CodCont")
aTamañosColumnas = Array(500, 4150, 550, 0)

CargarGridTitulos grdImp, aTitulosColGrid, aTamañosColumnas

' Rellena el grid de impuestos y carga la colección de impuestos
CargaImpuestos
          
' Se carga un array con los títulos de las columnas y otro con los tamaños para
'pasárselos a la función que carga el grid que muestra los impuestos a elegir
aTitulosColGrid = Array("Código", "Descripción", "% Impuesto", "Monto", "CodCont")
aTamañosColumnas = Array(800, 3000, 950, 1350, 0)

'carga los titulos al grid
CargarGridTitulos grdImptAplicados, aTitulosColGrid, aTamañosColumnas

' Deshabilita botón añadir
cmdAplicar.Enabled = False
cmdEliminar.Enabled = False

'Si gbImpuestos es verdadero carga los impuestos antes aplicados
If gbImpuestos = True Then

    'Carga el grdImptAplicados cuando se accede por segunda vez
    'sin modificar tipo de documento y monto
    CargaGrdRetenciones
    If grdImptAplicados.Rows > 1 Then cmdEliminar.Enabled = True
    
End If

End Sub

Private Sub CargaImpuestos()
Dim sSQL As String
Dim curImpt As New clsBD2
Dim curAFP As New clsBD2

'Se construye la sentencia
sSQL = "SELECT IdImp, DescImp, ValorImp, CodContable FROM TIPO_IMPUESTOS ORDER BY descimp"

'Ejecuta la sentencia que carga los impuestos
curImpt.SQL = sSQL
If curImpt.Abrir = HAY_ERROR Then End

If curImpt.EOF Then
    MsgBox "No existen impuestos", _
          vbInformation, vbOKOnly, "Egresos - Selección de Impuestos"
Else ' Carga el grid impuestos
    Do While Not curImpt.EOF
        '"Cód.", "Descripción", "Valor", "CodCont")
        ' Añade una fila al grid impuestos
        grdImp.AddItem curImpt.campo(0) & vbTab & curImpt.campo(1) _
                     & vbTab & Format(curImpt.campo(2), "#0.00") & vbTab & curImpt.campo(3)
        ' Añade un elemento a la colección
        mcolImpuestos.Add Item:=curImpt.campo(0) & "¯" & curImpt.campo(1) _
                     & "¯" & Format(curImpt.campo(2), "#0.00") & "¯" & curImpt.campo(3) _
                     , Key:=curImpt.campo(0)
        ' Mueve al siguiente elemento de la colección
        curImpt.MoverSiguiente
    Loop
End If

'*/*/*/*/*/CARGAMOS LAS AFP

' Se carga un array con los títulos de las columnas y otro con los tamaños para
aTitulosColGrid = Array("Cód.", "Descripción", "AporMenor", "CodCont", "AporActual", "Comis", "Seguro", "TopeAct", "TopeRH")
aTamañosColumnas = Array(300, 800, 600, 600, 600, 400, 450, 600, 550)

CargarGridTitulos GrdAFPs, aTitulosColGrid, aTamañosColumnas

'Se construye la sentencia
sSQL = ""
sSQL = "SELECT IdFP, DescFP, AporteMenor, CodCont, AporteActual, ComisionActual, PrimaActual, TopeActual, TopeControlRH " & _
        "FROM TIPO_FP " & _
        "WHERE DescFP LIKE '*MIXTA' OR DescFP='SNP' " & _
        "ORDER BY DescFP "

'Ejecuta la sentencia que carga los impuestos
curAFP.SQL = sSQL
If curAFP.Abrir = HAY_ERROR Then End

If curAFP.EOF Then
    MsgBox "No existen AFPs", _
          vbInformation, vbOKOnly, "Egresos - Selección de AFPs"
Else ' Carga el grid impuestos
    Do While Not curAFP.EOF
        '"Cód.", "Descripción", "Valor", "CodCont")
        ' Añade una fila al grid impuestos
        grdImp.AddItem curAFP.campo(0) & vbTab & curAFP.campo(1) _
                     & vbTab & Format(curAFP.campo(2) * 100, "#0.00") & vbTab & curAFP.campo(3)
        ' Añade un elemento a la colección
        mcolImpuestos.Add Item:=curAFP.campo(0) & "¯" & curAFP.campo(1) _
                     & "¯" & Format(curAFP.campo(2) * 100, "#0.00") & "¯" & curAFP.campo(3) _
                     , Key:=curAFP.campo(0)
        
        GrdAFPs.AddItem curAFP.campo(0) & vbTab & curAFP.campo(1) _
                     & vbTab & Format(curAFP.campo(2) * 100, "#0.00") & vbTab & curAFP.campo(3) _
                     & vbTab & Format(curAFP.campo(4) * 100, "#0.00") & vbTab & Format(curAFP.campo(5) * 100, "#0.00") _
                     & vbTab & Format(curAFP.campo(6) * 100, "#0.00") & vbTab & Format(curAFP.campo(7), "#0.00") _
                     & vbTab & Format(curAFP.campo(8), "#0.00")
        
        ' Mueve al siguiente elemento de la colección
        curAFP.MoverSiguiente
    Loop
End If

' cierra el cursor de los impuestos
curImpt.Cerrar
curAFP.Cerrar

End Sub


Private Sub CargaGrdRetenciones()
'---------------------------------------------------------
'Propósito  : Carga el grid con las retenciones aplicadas
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------
Dim MiObjeto
Dim ConceptoAFP As String

'Si la coleecion esta vacia
If gcolImpSel.Count = 0 Then
   
   'Coloca a optSinImpuestos a verdadero
   optSinImpuestos.Value = True
   
Else

   'Coloca a optConImpuestos a verdadero
   optConImpuestos.Value = True
   
   'Agrega los impuestos seleccionados al grdImptAplicados con sus montos
   'Recorre el grid para determinar que fila esta seleccionada
   For Each MiObjeto In gcolImpSel
   
      '"IdImpuesto", "Descripción", "% Impuesto", "Monto Retenido", "CodCont")
      '"idImp","Descripción", "ValorImpuesto", CodCont mcolimpuesto
      '"idImp", "ValorImpuesto", "Monto", codcont gcolImpsel
      ' Añade una fila al grd retenciones con los alculos necesarios
      If Left(Var30(MiObjeto, 1), 2) = "00" Then
        grdImptAplicados.AddItem Var30(MiObjeto, 1) _
                     & vbTab & Var30(mcolImpuestos(Var30(MiObjeto, 1)), 2) _
                     & vbTab & Var30(MiObjeto, 2) _
                     & vbTab & Format(Var30(MiObjeto, 3), "###,###,##0.00") _
                     & vbTab & Var30(MiObjeto, 4)
      Else
        If Right(Var30(MiObjeto, 1), 1) = "1" Then
          ConceptoAFP = Left(Var30(MiObjeto, 1), 2)
          ConceptoAFP = Var30(mcolImpuestos(ConceptoAFP), 2)
          ConceptoAFP = ConceptoAFP & " APORTE"
        ElseIf Right(Var30(MiObjeto, 1), 1) = "2" Then
          ConceptoAFP = Left(Var30(MiObjeto, 1), 2)
          ConceptoAFP = Var30(mcolImpuestos(ConceptoAFP), 2)
          ConceptoAFP = ConceptoAFP & " COMISION"
        Else
          ConceptoAFP = Left(Var30(MiObjeto, 1), 2)
          ConceptoAFP = Var30(mcolImpuestos(ConceptoAFP), 2)
          ConceptoAFP = ConceptoAFP & " SEGURO"
        End If
        grdImptAplicados.AddItem Var30(MiObjeto, 1) _
                     & vbTab & ConceptoAFP _
                     & vbTab & Var30(MiObjeto, 2) _
                     & vbTab & Format(Var30(MiObjeto, 3), "###,###,##0.00") _
                     & vbTab & Var30(MiObjeto, 4)
      End If
   Next
   
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Cierra las colecciones locales
' Colecciones usadas para operar los impuestos
Set mcolImpuestosAplicados = Nothing
Set mcolImpuestos = Nothing


End Sub

Private Sub grdImp_Click()
Dim i As Integer
Dim iFila As Integer

iFila = grdImp.Row
i = 1
' Selecciona toda la iFila
If grdImp.Rows > 1 Then

    'Marca o Desmarca solo una fila en el grd
    MarcarDesmarcarFilaGRID grdImp
    
    'Verifica SI hay iFilas marcadas para habilitar el botón Eliminar y  deshabilitar Aceptar2
   cmdAplicar.Enabled = False
 
   Do While i <= grdImp.Rows - 1 'Se recorren las iFilas
        grdImp.Row = i
        If grdImp.CellForeColor = vbWhite Then ' La iFila está marcada
            cmdAplicar.Enabled = True
        End If
        i = i + 1
   Loop
   
   ' se selecciona la iFila marcada
   grdImp.Row = iFila
   
End If


End Sub


Private Sub grdImptAplicados_Click()
Dim i As Integer
Dim iFila As Integer

' Si la fila es cero , no hace nada
If grdImptAplicados.Row < 1 Then Exit Sub

' Oculta el cuadro de Texto txtMonto
txtMonto.Visible = False

If grdImptAplicados.Col < 3 Or grdImptAplicados.Col > 3 Then

    '    ' Coloca a obligatorio la variable mbGridObligatorio
    '    mbGridObligatorio = True
    '
    '    ' Establece las Celdas Obligatoris del Grid
    '    GridCeldasObligatorios
    
    ' Coloca el control en
    grdImptAplicados.Col = 3
    
ElseIf grdImptAplicados.Col = 3 Then
    
    ' Coloca en falso la variable mbGridObligatorio
    mbGridObligatorio = False
    
    ' Ingresa a la celda donde se hizo el click
    EnterCellGrid
    
End If

End Sub

Private Sub GridCeldasObligatorios()
'--------------------------------------------------------------
'Propósito  : Establece los campos obligatorios del grid
'Recibe     : Nada
'Devuelve   : Nada
'---------------------------------------------------------------
Dim i As Integer

If grdImptAplicados.Rows > 1 Then

    'Verifica SI la celda tiene datos o NO
    For i = 1 To grdImptAplicados.Rows - 1
    
        'Inicializa la variable en estado de verificacion
        mbGridObligatorio = True
    
        If grdImptAplicados.TextMatrix(i, 3) = "" Or Val(grdImptAplicados.TextMatrix(i, 3)) <= 0 Then
           grdImptAplicados.Row = i
           grdImptAplicados.Col = 3
           grdImptAplicados.CellBackColor = Obligatorio
        Else
            grdImptAplicados.Row = i
           grdImptAplicados.Col = 3
           grdImptAplicados.CellBackColor = vbWhite
           grdImptAplicados.CellForeColor = vbBlack
        End If
    Next i

End If
    
End Sub

Private Sub grdImptAplicados_EnterCell()

'Verifica que el grdonceptos tenga datos
If grdImptAplicados.Rows > 1 Then

    'Verifica que la columna sea la de montos
    If grdImptAplicados.Col = 3 Then
    
        'mbGridObligatorio pone en falso
        If mbGridObligatorio = False Then
        
            'Adecua el txtMonto a la celda del Grid
            EnterCellGrid
            
        End If
        
    End If
    
End If

End Sub

Private Sub EnterCellGrid()
'--------------------------------------------------------------
'Propósito: Se ejecuta automaticamente cuando una celda
'           del control MSFlexGrid sea seleccionada
'--------------------------------------------------------------
'nota : Es llamado desde el evento entercell del grdImptAplicados

'Oculta el cuadro de Texto txtMonto
txtMonto.Visible = False

' Borra el contenido del cuadro de texto
txtMonto.Text = Empty

'Situa el control txtMonto sobre la celda seleccionada
txtMonto.Top = grdImptAplicados.Top + grdImptAplicados.CellTop
txtMonto.Left = grdImptAplicados.Left + grdImptAplicados.CellLeft

'Ajusta el tamaño del control al tamaño de la celda seleccionada
txtMonto.Width = grdImptAplicados.CellWidth
txtMonto.Height = grdImptAplicados.CellHeight

'Asigna el contenido de la celda seleccionada a la propiedad text del control txtFormatoCelGrid
txtMonto.Text = Var37(grdImptAplicados.Text)

'Visualiza el control txtMonto
txtMonto.Visible = True

'Cursor se ubica en el txtMonto
txtMonto.SetFocus

End Sub


Private Sub grdImptAplicados_LeaveCell()

'Verifica que el grdImptAplicados contenga datos
If grdImptAplicados.Rows > 1 Then
    
    'Verifica que la columna sea >= a la de monto
    If grdImptAplicados.Col = 3 Then
    
        'Evalua si esta verificando los datos del grid
        If mbGridObligatorio = False Then
        
            'Asigna el contenido del txtMonto a la celda
            LeaveCellGrid
            
            'Cambia el color del grdImptAplicados
            If grdImptAplicados.Text = "" Or Val(grdImptAplicados.Text) <= 0 Then
               grdImptAplicados.CellBackColor = Obligatorio
            Else
               grdImptAplicados.CellBackColor = vbWhite
            End If

        End If

    End If
    
End If

End Sub

Private Sub LeaveCellGrid()
'--------------------------------------------------------------
'Propósito: Se ejecuta automaticamente cuando una celda
'           del control MSFlexGrid sea abandonada
'--------------------------------------------------------------

'Asigna el contenido de cuadro de texto txtMonto a la celda
'activa antes de ser abandonada

If grdImptAplicados.Row <> 0 Then
  
  If txtMonto.Text <> "" Then
      
    txtMonto.MaxLength = 14
    
    'Da formato de moneda
    grdImptAplicados.Text = Format(Val(Var37(grdImptAplicados.Text)), "###,###,###,##0.00")
    
  Else
  
    'Vacia el grid
    grdImptAplicados.Text = ""
    
  End If
  
End If

End Sub


Private Sub optConImpuestos_Click()

'Habilita los controles
HabilitarControles

'Habilita el boton aceptar
HabilitaDeshabilitaBotonAceptar

End Sub

Private Sub optConImpuestos_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub optSinImpuestos_Click()

'Deshabilita los grid de los impuestos
DesHabilitarControles

'Vacia el grid de los impuestos aplicados
grdImptAplicados.Rows = 1
txtMonto.Visible = False

'Desmarcar grdImp
DesmarcarGrid grdImp

'Vacia la colección
Set mcolImpuestosAplicados = Nothing

'Habilita el boton aceptar
HabilitaDeshabilitaBotonAceptar

End Sub

Private Sub DesHabilitarControles()
'---------------------------------------
'Propósito  : Deshabilita los controles
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------
fraImpuestos.Enabled = False
fraImpuestosAplicados.Enabled = False
cmdAplicar.Enabled = False
cmdEliminar.Enabled = False

End Sub

Private Sub HabilitarControles()
'---------------------------------------
'Propósito  : Habilita los controles
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------
optConImpuestos.Value = True
fraImpuestos.Enabled = True
fraImpuestosAplicados.Enabled = True

End Sub

Private Sub optSinImpuestos_KeyPress(KeyAscii As Integer)

' Si se presiona enter se pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If
  
End Sub

Private Sub txtMonto_Change()
Dim bVacio As Boolean

If grdImptAplicados.Col = 3 Then

    'Verifica SI el txt esta vacio o con dato
    If txtMonto.Text = "" Or Val(txtMonto.Text) <= 0 Then
    
        txtMonto.BackColor = Obligatorio
    Else
        txtMonto.BackColor = vbWhite
        
    End If
    
    If txtMonto.Text <> "" Then
    
        'Vacia el contenido de txtMonto al grdImptAplicados
        grdImptAplicados.Text = txtMonto.Text
    
    End If
    
End If

End Sub

Function HabilitarAceptar()
'---------------------------------------
'Propósito  : Habilita o deshabilita el boton aceptar
'Recibe     : Nada
'Devuelve   : Nada
'----------------------------------------

Dim i As Integer

cmdAceptarModificar.Enabled = True

'Habilita el boton aceptar cuando el tipo de operacion a realizar es nuevo
'y se haya ingresado algun dato
If txtMonto.BackColor = Obligatorio Then
    cmdAceptarModificar.Enabled = False
    Exit Function
Else
    i = 1
    Do While i <= grdImptAplicados.Rows - 1
        If grdImptAplicados.TextMatrix(i, 3) = "" Or Val(grdImptAplicados.TextMatrix(i, 3)) <= 0 Then
            cmdAceptarModificar.Enabled = False
            Exit Function
        End If
    i = i + 1
    Loop
End If

End Function

Private Sub txtMonto_GotFocus()

If grdImptAplicados.Col >= 3 Then
    'Maximo tamaño del monto
    txtMonto.MaxLength = 12
    txtMonto.Text = Var37(txtMonto.Text)
End If

End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)

If grdImptAplicados.Col >= 3 Then

    'Ubica el cursor en la posición seleccionada
    MovimientoCeldas grdImptAplicados, KeyCode
    
End If

End Sub

Private Sub MovimientoCeldas(grdRec As MSFlexGrid, iTecla As Integer)
'--------------------------------------------------------------
'Propósito  : Mueve el cursor de acuerdo a la ubicación
'Recibe     : El grid donde se ubica el cursor y la tecla presionado
'Devuelve   : Nada
'--------------------------------------------------------------
'Nota:        Llamdo del evento KeyDown del txtMonto

'Selecciona la tecla presionada
Select Case iTecla
'Tecla enter
Case vbKeyReturn

  'Verifica SI es la ultima fila
  If grdRec.RowSel < grdRec.Rows - 1 Then
    grdRec.Row = grdRec.RowSel + 1

  'Verifica SI es la ultima columna
  ElseIf grdRec.ColSel = grdRec.Cols - 1 Then
    grdRec.Col = 2
    grdRec.Row = 1
    Else
        grdRec.Row = 1
        grdRec.Col = grdRec.ColSel + 1
  End If
  
' Tecla cursor arriba
Case vbKeyUp
    'Verifica SI es la primera fila de la datos
    If grdRec.RowSel > 1 Then
        grdRec.Row = grdRec.RowSel - 1
    End If
  
'Tecla cursor abajo
Case vbKeyDown
    'Verifica SI es la ultima fila de la datos
    If grdRec.RowSel < grdRec.Rows - 1 Then
      grdRec.Row = grdRec.RowSel + 1
    End If

End Select
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If grdImptAplicados.Col >= 3 Then
    
    'Valida Monto ingresado en soles, luego se ubica en la celda inmediata
    Var33 txtMonto, KeyAscii
    
 End If
 
End Sub

Private Sub txtMonto_LostFocus()

'Verifica que la columna sea la de monto
If grdImptAplicados.Col = 3 Then

    If txtMonto.Text <> "" Then
    
        'Vacia el contenido de txtMonto al grdImptAplicados
        grdImptAplicados.Text = Format(Val(Var37(txtMonto.Text)), "###,###,###,##0.00")
    
    Else
    
        'Vacia la celda
        grdImptAplicados.Text = ""
        
    End If

End If

End Sub

