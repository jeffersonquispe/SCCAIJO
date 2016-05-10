VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCBSelOrdenTraslado 
   Caption         =   "SCCaijo Selección de Orden de Traslado"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   HelpContextID   =   56
   Icon            =   "SCCBSelTraslado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optTrasladoDestino 
         Caption         =   "Por Orden de Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optTrasladoOrigen 
         Caption         =   "Por Orden de Origen"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   7095
      TabIndex        =   5
      Top             =   4800
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   8295
      TabIndex        =   6
      Top             =   4800
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   40
      TabIndex        =   7
      Top             =   1110
      Width           =   9240
      Begin MSFlexGridLib.MSFlexGrid grdOrden 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FillStyle       =   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   330
         Left            =   1125
         TabIndex        =   0
         Top             =   315
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
         Left            =   3480
         TabIndex        =   1
         Top             =   315
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha &Fin:"
         Height          =   330
         Left            =   2610
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCBSelOrdenTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable para el manejo del grid
Dim ipos As Long

Private Sub mskFechaFin_Change()

' Se valida que la fecha fin de la consulta
If ValidarFecha(mskFechaFin) Then
  mskFechaFin.BackColor = vbWhite
  'Muestra el grid entre la FechaIni y FechaFin
  MostrarGrid
Else
  mskFechaFin.BackColor = Obligatorio
  grdOrden.Rows = 1
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub MostrarGrid()
Dim sSQL As String

If mskFechaIni.BackColor <> Obligatorio And mskFechaFin.BackColor <> Obligatorio Then
    
        'Verifica si la fecha inicio es anterior a la fecha fin
        If CompararFechaIniFin(mskFechaIni, mskFechaFin) = False Then

            'Limpia el grid grdOrden
            grdOrden.Rows = 1
            
            'Verifica si esta activado el optDestino u optOrigen
            If optTrasladoOrigen.Value Then
            
                'Se seleccionan los traslados
                sSQL = "SELECT E.Orden, TD.DescTipoDoc, E.NumDoc,E.MontoCB, E.FecMov " & _
                      "FROM EGRESOS E, CTB_TRASLADOCAJABANCOS T, TIPO_DOCUM TD " & _
                      "WHERE T.OrdenEgreso=E.Orden And E.Anulado='NO' And E.FecMov BETWEEN  '" & _
                       FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                      "And E.IdTipoDoc=TD.IdTipoDoc and E.idProy='' ORDER BY  E.Orden"
                      
                'Se carga un array con los títulos de las columnas y otro con los tamaños para
                'pasárselos a la función que carga el grid
                aTitulosColGrid = Array("Origen Traslado", "Tipo Doc de Origen", "Nro de Doc Origen ", "Monto", "Fecha")
                aTamañosColumnas = Array(1200, 3000, 1600, 1600, 1200)
                aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Importe", "fmt_Fecha")

            ElseIf optTrasladoDestino.Value Then
                'Se seleccionan los Ingresos
                sSQL = "SELECT I.Orden, TD.DescTipoDoc, I.NumDoc, I.Monto, I.FecMov " & _
                     "FROM CTB_TRASLADOCAJABANCOS T,INGRESOS I, TIPO_DOCUM TD " & _
                     "WHERE T.OrdenIngreso=I.Orden And I.Anulado='NO' And I.FecMov BETWEEN  '" & _
                      FechaAMD(mskFechaIni.Text) & "' And '" & FechaAMD(mskFechaFin.Text) & "' " & _
                     "And I.IdTipoDoc=TD.IdTipoDoc ORDER BY  I.Orden"
                     
               'Se carga un array con los títulos de las columnas y otro con los tamaños para
               'pasárselos a la función que carga el grid
               aTitulosColGrid = Array("Destino Traslado", "Tipo Doc de Destino", "Nro de Doc Destino", "Monto", "Fecha")
               aTamañosColumnas = Array(1200, 3000, 1600, 1600, 1200)
               aFormatos = Array("fmt_Normal", "fmt_Normal", "fmt_Normal", "fmt_Importe", "fmt_Fecha")
    
            End If
           
             'Carga el grid con su respectivo formato
             CargarGridConFormatos grdOrden, sSQL, aTitulosColGrid, aTamañosColumnas, aFormatos
              
             If grdOrden.Rows = 1 Then
                MsgBox "No existen traslados entre estas fechas", _
                         vbInformation + vbOKOnly, "S.G.Ccaijo modificaión"
     
             End If
         Else
             'Limpia el grid
            grdOrden.Rows = 1
            MsgBox "Fecha inicio es posterior a fecha fin", vbInformation + vbOKOnly, " Selección de Orden de traslados"
         End If
Else

    'Limpia el grid
    grdOrden.Rows = 1
    
End If
End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de controlç
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub mskFechaIni_Change()

' Se valida que la fecha Inicio de la consulta
If ValidarFecha(mskFechaIni) Then
  mskFechaIni.BackColor = vbWhite
  
  'Muestra el grid entre las fechas
  MostrarGrid

Else
  grdOrden.Rows = 1
  mskFechaIni.BackColor = Obligatorio
  cmdAceptar.Enabled = False
End If

End Sub

Private Sub cmdAceptar_Click()
'Se comprueba que se haya marcado algúna fila
If grdOrden.Row < 1 Then
  MsgBox "Debe marcar algún documento", vbInformation + vbOKOnly, "SGCcaijo-Selección de registros"
  Exit Sub
End If

'Verifica que opt se selecciono
If optTrasladoOrigen.Value Then

    'Carga el txtIdSalida con ddatos de la salida a modificar
    frmCBTraslado.txtOrdenOrigen.Text = grdOrden.TextMatrix(grdOrden.Row, 0)
ElseIf optTrasladoDestino.Value Then

    'Carga el txtIdSalida con datos de la salida a modificar
    frmCBTraslado.txtOrdenDestino.Text = grdOrden.TextMatrix(grdOrden.Row, 0)
End If

'Muestra el formulario de Traslados
frmCBTraslado.Show vbModal, Me

'Actualiza el grid si se anulo el movimiento
MostrarGrid

'Deshabilita el botón aceptar
cmdAceptar.Enabled = False
  
End Sub


Private Sub cmdSalir_Click()

'Termina la ejecucion del formulario
Unload Me

End Sub

Private Sub Form_Load()

'Coloa a obligatorio
mskFechaIni.BackColor = Obligatorio

'Carga la fecha del sistema
mskFechaFin.Text = gsFecTrabajo
      
'Desabilita los controles
cmdAceptar.Enabled = False

' Inicializa el grid
ipos = 0
gbCambioCelda = False

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

' Si presiona enter cambia de controlç
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optTrasladoDestino_Click()

'Muestra el grid entre las fechas
MostrarGrid

End Sub

Private Sub optTrasladoDestino_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de controlç
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub optTrasladoOrigen_Click()

'Muestra el grid entre las fechas
MostrarGrid

End Sub

Private Sub optTrasladoOrigen_KeyPress(KeyAscii As Integer)

' Si presiona enter cambia de controlç
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub
