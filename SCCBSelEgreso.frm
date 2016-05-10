VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBSelOrden 
   Caption         =   "SCCaijo Selección de Orden"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   Icon            =   "SCCBSelEgreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7095
      TabIndex        =   3
      Top             =   4380
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8295
      TabIndex        =   4
      Top             =   4380
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   40
      TabIndex        =   5
      Top             =   750
      Width           =   9240
      Begin MSFlexGridLib.MSFlexGrid grdOrden 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FillStyle       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   1245
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   330
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha &Inicio:"
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha &Fin:"
         Height          =   330
         Left            =   3090
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCBSelOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable para el manejo del grid selección
Dim ipos As Long

Private Sub mskFechaFin_Change()

' Se valida que la fecha Inicio de la consulta
If ValidarFecha(mskFechaFin) Then
    mskFechaFin.BackColor = vbWhite
    'Carga el grid
    CargarGrid
Else
    mskFechaFin.BackColor = Obligatorio
    grdOrden.Rows = 1
    cmdAceptar.Enabled = False
End If

End Sub

Private Sub CargarGrid()
'Carga con los datos el grid
Dim sSQL As String

If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
    
       'Verifica si la fecha inicio es anterior a la fecha fin
       If CompararFechaIniFin(mskFechaIni, mskFechaFin) = False Then

           'Limpia el grid grdOrden
           grdOrden.Rows = 1
           
           Select Case gsTipoSeleccionOrden
           Case "EgresoCA"
            'Se seleccionan los Egresos con afectación
            sSQL = "SELECT E.Orden, TD.DescTipoDoc, E.NumDoc,E.MontoAfectado, E.FecMov " & _
                  "FROM EGRESOS E, TIPO_DOCUM TD " & _
                  "WHERE E.Anulado='NO' And E.FecMov BETWEEN  '" & _
                   FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                  "And E.IdTipoDoc=TD.IdTipoDoc and E.idProy<>'' And E.Orden Not IN " & _
                 "(SELECT OrdenIngreso FROM CTB_TRASLADOCAJABANCOS) " & _
                 "ORDER BY  E.Orden"
           
           Case "Ingreso"
           'Se seleccionan los Ingresos
            sSQL = "SELECT I.Orden, TD.DescTipoDoc, I.NumDoc, I.Monto, I.FecMov " & _
                 "FROM INGRESOS I, TIPO_DOCUM TD " & _
                 "WHERE I.Anulado='NO' And I.FecMov BETWEEN  '" & _
                  FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                 "And I.IdTipoDoc=TD.IdTipoDoc And I.Orden NOT IN " & _
                 "(SELECT OrdenIngreso FROM CTB_TRASLADOCAJABANCOS) " & _
                 "ORDER BY  I.Orden"
               
            Case "EgresoSA"
            sSQL = "SELECT E.Orden, TD.DescTipoDoc, E.NumDoc,E.MontoCB, E.FecMov " & _
                  "FROM EGRESOS E, TIPO_DOCUM TD " & _
                  "WHERE E.Anulado='NO' And E.FecMov BETWEEN  '" & _
                   FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                  "And E.IdTipoDoc=TD.IdTipoDoc and E.idProy='' And E.Orden Not IN " & _
                 "(SELECT OrdenEgreso FROM CTB_TRASLADOCAJABANCOS) " & _
                 "ORDER BY  E.Orden"
            Case "Ventas"
            'Se seleccionan los Egresos con afectación
            sSQL = "SELECT V.Orden, TD.DescTipoDoc, V.NumDoc, V.MontoTotal, V.FecMov " & _
                  "FROM VENTAS V, VENTAS_TIPO_DOCUM TD " & _
                  "WHERE V.Anulado='NO' And V.FecMov BETWEEN  '" & _
                   FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                  "And V.IdTipoDoc=TD.IdTipoDoc and V.Cancelado = 'NO'" & _
                  "ORDER BY  V.Orden"
       
           End Select
            'Se carga un array con los títulos de las columnas y otro con los tamaños para
            'pasárselos a la función que carga el grid
            aTitulosColGrid = Array("Orden", "Tipo Doc", "Número", "Monto", "Fecha")
            aTamañosColumnas = Array(1200, 3000, 1600, 1600, 1200)
            aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Importe", "fmt_Fecha")
             
            CargarGridConFormatos grdOrden, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
             
            If grdOrden.Rows = 1 Then
               MsgBox "No existen Registros entre estas Fechas", _
                        vbInformation + vbOKOnly, "S.G.Ccaijo"
    
            End If
        Else
            'Limpia el grid
           grdOrden.Rows = 1
           MsgBox "Fecha inicio es posterior a fecha fin", vbInformation + vbOKOnly, " Selección de Orden"
        End If

Else

    'Limpia el grid
    grdOrden.Rows = 1
    
End If

End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub

Private Sub mskFechaIni_Change()

' Se valida que la fecha Inicio de la consulta
If ValidarFecha(mskFechaIni) Then
    mskFechaIni.BackColor = vbWhite
    'Carga el grid
    CargarGrid
Else
  mskFechaIni.BackColor = Obligatorio
  grdOrden.Rows = 1
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub cmdAceptar_Click()
'Se comprueba que se haya mardcado algúna fila
If grdOrden.Row < 1 Then
  MsgBox "Debe marcar algún documento", vbInformation + vbOKOnly, "SGCcaijo-Selección de registros"
  Exit Sub
End If

Select Case gsTipoSeleccionOrden
Case "EgresoCA"
    'Carga el txtIdSalida con datos de la salida a modificar
    gsOrden = grdOrden.TextMatrix(grdOrden.Row, 0)
Case "Ingreso"
    'Carga el txtCodIngreso con datos del ingreso a modificar
    frmCBIngresos.txtCodIngreso.Text = grdOrden.TextMatrix(grdOrden.Row, 0)
Case "EgresoSA"
    'Carga el txtIdSalida con datos de la salida a modificar
    gsOrden = grdOrden.TextMatrix(grdOrden.Row, 0)
Case "Ventas"
    'Carga el txtIdSalida con datos de la salida a modificar
    gsOrden = grdOrden.TextMatrix(grdOrden.Row, 0)
End Select
 

'Sale del formulario
 Unload Me

End Sub

Private Sub cmdSalir_Click()

'Termina la ejecucion del formulario
Unload Me

End Sub

Private Sub Form_Load()

'Coloca a obligatorio
mskFechaIni.BackColor = Obligatorio

'Carga la fecha del sistema
mskFechaFin.Text = gsFecTrabajo
      
'Desabilita los controles
cmdAceptar.Enabled = False

' Inicializa el grid
ipos = 0
gbCambioCelda = False

' Alinea la columna de número de documento
grdOrden.ColAlignment(2) = 1

End Sub
 
Private Sub grdOrden_Click()

If grdOrden.Row > 0 And grdOrden.Row < grdOrden.Rows Then
    ' Marca la fila seleccionada
    MarcarFilaGRID grdOrden, vbWhite, vbDarkBlue
    ' Habilita aceptar
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub grdOrden_DblClick()

'Hace llamado al evento click del aceptar
cmdAceptar_Click

End Sub

Private Sub grdOrden_EnterCell()

If ipos <> grdOrden.Row Then
    '  Verifica si es la última fila
    If grdOrden.Row > 0 And grdOrden.Row < grdOrden.Rows Then
         If gbCambioCelda = False Then
            gbCambioCelda = True
            ' Marca la fila
            MarcarSoloUnaFilaGrid grdOrden, ipos
            gbCambioCelda = False
            cmdAceptar.Enabled = True
         End If
    End If
    ' Actualiza el valor de la fila
    ipos = grdOrden.Row
End If

End Sub

Private Sub grdOrden_KeyPress(KeyAscii As Integer)

' Verifica si se apretó el enter
 If KeyAscii = 13 Then
    SendKeys vbTab
 End If
 
End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)

  ' Si presiona enter pasa al siguiente control
  If KeyAscii = 13 Then
    SendKeys vbTab
  End If

End Sub
