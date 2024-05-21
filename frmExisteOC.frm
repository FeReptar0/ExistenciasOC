VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExisteOC 
   Caption         =   "Reporte de existencias de órdenes de compra"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14985
   Icon            =   "frmExisteOC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRepositorio 
      Caption         =   "Repositorio"
      Height          =   495
      Left            =   12600
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Información al..."
      Height          =   1575
      Left            =   240
      TabIndex        =   27
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton optHoy 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optAyer 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Tiempo real:"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Repositorio del"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtDesdeArticulo 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Text            =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtHastaArticulo 
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Text            =   "ZZZZZZ"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   12015
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   11775
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   0
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   20770
         _ExtentY        =   8070
         _StockProps     =   79
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10560
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExporta 
      Caption         =   "Exportar a Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      Height          =   1575
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      Begin VB.ComboBox cmbHastaSucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Text            =   "cmbHastaSucursal"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtHastaNoOC 
         Height          =   285
         Left            =   6720
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbHastaMarca 
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Text            =   "Hasta Marca"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Rango de Fechas OC"
         Height          =   255
         Left            =   8400
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDesdeNoOC 
         Height          =   285
         Left            =   6720
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Número de OC"
         Height          =   255
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbDesdeMarca 
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Text            =   "Desde Marca"
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Marca"
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Artículo"
         Height          =   195
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbDesdeSucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "cmbDesdeSucursal"
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   8400
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   126222337
         CurrentDate     =   41646
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   8400
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   126222337
         CurrentDate     =   41549
      End
      Begin VB.Label lblHasta 
         Caption         =   "Hasta:"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDesde 
         Caption         =   "Desde:"
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
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4080
      TabIndex        =   21
      Top             =   6720
      Width           =   3975
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Revisión:14/05/2024"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   420
         Width           =   2085
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Version 4.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   225
         TabIndex        =   23
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sage 300 Ver. 2023"
         Height          =   195
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmExisteOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'26/08/2021 -- U9tima Actualizacion
'' -- Se agrega la funcion fnStrGetNoPvC, para que consulte del campo de referencia
'' -- Sales Order se toma de la Orden de Compra (SalesO)

'Option Explicit
Dim FacturaRM As String

Dim usuario As String
Dim sUserRepositorio As String
Dim Empresa As String
Dim EmpresaNombre As String
Dim lSignonID As Long
Dim no_producto As String


Private mSession As AccpacSession
Private mSessMgr As AccpacSessionMgr ' this is useful if you need to use the AccpacMeter


Private Sub Check1_Click()
If Check1.Value = 1 Then
    txtDesdeNoOC.Visible = True
    txtHastaNoOC.Visible = True
Else
    txtDesdeNoOC.Visible = False
    txtHastaNoOC.Visible = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    txtDesdeArticulo.Visible = True
    txtHastaArticulo.Visible = True
    
Else
    txtDesdeArticulo.Visible = False
    txtHastaArticulo.Visible = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    cmbDesdeMarca.Visible = True
    cmbHastaMarca.Visible = True
Else
    cmbDesdeMarca.Visible = False
    cmbHastaMarca.Visible = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    DTPicker1.Visible = True
    DTPicker2.Visible = True
Else
    DTPicker1.Visible = False
    DTPicker2.Visible = False
End If
End Sub

Private Function Existencia(itemno As String) As Double
Dim ssql As String
Dim rsOpcional As Recordset
Set rsOpcional = New Recordset

ssql = "select SUM(QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST) EXISTENCIA from ICILOC where ITEMNO='" & itemno & "' GROUP BY ITEMNO"
rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsOpcional.EOF = False And rsOpcional.BOF = False Then
    Existencia = Trim(rsOpcional!Existencia)
Else
    Existencia = 0
End If
rsOpcional.Close
Set rsOpcional = Nothing
End Function

Private Function ValorSALEORDER(PORHSEQ As String) As String
Dim ssql As String
Dim rsOpcional As Recordset
Set rsOpcional = New Recordset

ssql = "select VALUE from POPORHO where PORHSEQ='" & PORHSEQ & "' and OPTFIELD='SALEORDER'"
rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsOpcional.EOF = False And rsOpcional.BOF = False Then
    ValorSALEORDER = Trim(rsOpcional!Value)
Else
    ValorSALEORDER = "0"
End If
rsOpcional.Close
Set rsOpcional = Nothing
End Function

Private Function ValorCosto(itemno As String) As String
On Error Resume Next
Dim ssql As String
Dim rsOpcional As Recordset
Set rsOpcional = New Recordset

ssql = "select VALUE from ICITEMO where ITEMNO='" & itemno & "' and OPTFIELD='COSTFABRI'"
rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsOpcional.EOF = False And rsOpcional.BOF = False Then
    ValorCosto = CDbl(Trim(rsOpcional!Value))
Else
    ValorCosto = "0"
End If
rsOpcional.Close
Set rsOpcional = Nothing
End Function

Private Function ValorPL(itemno As String) As String
On Error Resume Next
Dim ssql As String
Dim rsOpcional As Recordset
Set rsOpcional = New Recordset

ssql = "select VALUE from ICITEMO where ITEMNO='" & itemno & "' and OPTFIELD='PRECIOVTA'"
rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsOpcional.EOF = False And rsOpcional.BOF = False Then
    ValorPL = CDbl(Trim(rsOpcional!Value))
Else
    ValorPL = "0"
End If
rsOpcional.Close
Set rsOpcional = Nothing
End Function

Private Sub Buscar_informacion_de_ayer()
On Error GoTo RutinaError

Dim ssql As String
Dim rsReporte As Recordset
Set rsReporte = New Recordset
Dim PasaWhere As Boolean
Dim strCadena As String

Screen.MousePointer = vbHourglass
frmExisteOC.Enabled = False

Generagrid

ssql = "SELECT [Usuario],[Sucursal],[Articulo],[Codigo2],[Cantidad],[SerLot],[FechaRM],[NoRM],[NoFacturaRM]"
ssql = ssql & " ,[comentariosRM],[NoOC],[NoLineaOC],[SalesOrd],[NoLinea],[CostoOC],[FechaTRF],[NoTRF]"
ssql = ssql & " ,[Costo],[PrecioLista],[Comentarios],[ComentariosTRF],[Proveedor],[NoProducto],[DealID]"
ssql = ssql & " ,[IDEndUser],[EndUserName],[OCCliente],[NUMPV],[IDCliente],[NombreCliente],[IDEjecutivo]"
ssql = ssql & " ,[EjecutivoName],[FechaOC],[FacturaComer],[FacturaPack],[Category]"
ssql = ssql & " From [dbo].[Existencias_OC_Repositorio]"
ssql = ssql & " Where Usuario='REPOSITORI' "
PasaWhere = True
If cmbDesdeSucursal.Text <> "TODAS" Or cmbHastaSucursal.Text <> "TODAS" Then 'SUCURSAL
    If Not PasaWhere Then
        ssql = ssql & "  where  (Sucursal between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
        PasaWhere = True
    Else
        ssql = ssql & "  and  (Sucursal between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
    End If
End If
If Check2.Value = 1 Then 'ARTÍCULO
    If Not PasaWhere Then
        ssql = ssql & "  where  (Articulo between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and  '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "'("
        PasaWhere = True
    Else
        ssql = ssql & "  and  (Articulo between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "')"
    End If
End If
If Check3.Value = 1 Then 'CATEGORÍA
    If Not PasaWhere Then
        ssql = ssql & "  where  (Category  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
        PasaWhere = True
    Else
        ssql = ssql & "  and  (Category  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
    End If
End If
If Check1.Value = 1 Then 'NO OC
    If Not PasaWhere Then
        ssql = ssql & "  where  (NoOC between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
        PasaWhere = True
    Else
        ssql = ssql & "  and  (NoOC  between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
    End If
End If
If Check4.Value = 1 Then 'RANGO FECHA
    If Not PasaWhere Then
        ssql = ssql & "  where  (FechaOC between '" & Year(DTPicker1.Value) & "-" & Month(DTPicker1.Value) & "-" & Day(DTPicker1.Value) & "' and '" & Year(DTPicker2.Value) & "-" & Month(DTPicker2.Value) & "-" & Day(DTPicker2.Value) & "')"
        PasaWhere = True
    Else
        ssql = ssql & "  and  (FechaOC between '" & Year(DTPicker1.Value) & "-" & Month(DTPicker1.Value) & "-" & Day(DTPicker1.Value) & "' and '" & Year(DTPicker2.Value) & "-" & Month(DTPicker2.Value) & "-" & Day(DTPicker2.Value) & "')"
    End If
End If
Debug.Print ssql
rsReporte.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsReporte.EOF = False And rsReporte.BOF = False Then
    rsReporte.MoveFirst
    Do While rsReporte.EOF = False
        strCadena = rsReporte!Sucursal & vbTab & rsReporte!articulo & vbTab & rsReporte!Codigo2 & vbTab & rsReporte!Cantidad & vbTab
        strCadena = strCadena & rsReporte!SerLot & vbTab & rsReporte!FechaRM & vbTab & rsReporte!NoRM & vbTab & rsReporte!NoFacturaRM & vbTab & rsReporte!comentariosRM & vbTab
        strCadena = strCadena & rsReporte!NoOC & vbTab & rsReporte!NoLineaOC & vbTab & rsReporte!SalesOrd & vbTab & rsReporte!NoLinea & vbTab & rsReporte!CostoOc & vbTab
        strCadena = strCadena & rsReporte!FechaTRF & vbTab & rsReporte!NoTRF & vbTab & rsReporte!Costo & vbTab & rsReporte!PrecioLista & vbTab & rsReporte!Comentarios & vbTab
        strCadena = strCadena & rsReporte!ComentariosTRF & vbTab & rsReporte!Proveedor & vbTab & rsReporte!NoProducto & vbTab & rsReporte!DealID & vbTab & rsReporte!IDEndUser & vbTab
        strCadena = strCadena & rsReporte!EndUserName & vbTab & rsReporte!OCCliente & vbTab & rsReporte!NUMPV & vbTab & rsReporte!IDCliente & vbTab & rsReporte!NombreCliente & vbTab
        strCadena = strCadena & rsReporte!IDEjecutivo & vbTab & rsReporte!EjecutivoName & vbTab & rsReporte!FechaOC & vbTab & rsReporte!FacturaComer & vbTab & rsReporte!FacturaPack
        
        SSOleDBGrid1.AddItem strCadena
        rsReporte.MoveNext
    Loop
    MsgBox "Se ha completado la busqueda de información con los parámetros establecidos", vbInformation + vbOKOnly, "Información"
Else
    MsgBox "No existe información disponible", vbCritical + vbOKOnly, "Error"
End If
rsReporte.Close
Set rsReporte = Nothing

Screen.MousePointer = vbNormal
frmExisteOC.Enabled = True

cmdExporta.Enabled = True

Exit Sub

RutinaError:
If Err.Number <> -2147217873 Then 'NO SE MUESTRA EL MENSAJE SI ES REGISTRO DUPLICADO
    MsgBox Err.Number & "  " & Err.Description, vbCritical + vbOK, "Error"
End If
Resume Next
End Sub

Private Sub Buscar_informacion_del_dia()
On Error GoTo RutinaError

Screen.MousePointer = vbHourglass
frmExisteOC.Enabled = False

Dim PasaWhere As Boolean

Dim InfoDisponible As Boolean

Dim ArticulosSinSerieNiLote As String
Dim ssql As String
Dim ssql2 As String
Dim ssql3 As String
Dim ssql4 As String
Dim strCadena As String
Dim Serie As String

Dim rsBusca As Recordset
Set rsBusca = New Recordset

Dim rsSerLot As Recordset
Dim rsSerLotVal As Recordset

Set rsSerLot = New Recordset
Set rsSerLotVal = New Recordset

Dim rsRecibo As Recordset
Set rsRecibo = New Recordset
Dim rsTrans As Recordset
Set rsTrans = New Recordset
Dim rsDemas As Recordset
Set rsDemas = New Recordset

Dim strPORHSEQ As String
Dim ArticuloUsa As Integer
Dim strSerLot As String
Dim Codigo2 As String
Dim FechaRM As String
Dim comentariosRM As String
Dim comenRM As String
Dim SalesO As String
Dim NoLinea As String
Dim CostoOc As String
Dim FechaTRF As String
Dim NoTFR As String
Dim Costo As String
Dim PricoLista As String
Dim Comentario As String
Dim ComentarioTRF As String
Dim NoFacRM As String
Dim NoLineaOC As Long
Dim Cantidad As Double

Dim SumatoriaTotal As Double
Dim strProveedor As String

'---------------------------------
Dim strNoPV As String

Dim sDatosAgregados As String
Dim lDatosAgregados() As String

Dim sDEALID As String
Dim sENDUSER As String
Dim sOCCLIENTE As String
Dim sSALEORDER As String

Dim strEjecutivo As String
Dim strEndUser As String
Dim sIDENDUSER As String
Dim strIDEjecutivo As String
Dim strVendorID As String

''20240206
Dim sFACTCOMER As String
Dim sFACTPACK As String

strNoPV = ""
'---------------------------------

InfoDisponible = False
Generagrid
cmdExporta.Enabled = False
PasaWhere = True

'---------------------------------------------------
'---------------------------------------------------
cnDB.Execute "DELETE FROM Existencias_OC where Usuario='" & usuario & "'"
'---------------------------------------------------
'---------------------------------------------------
'MsgBox "Exixtencias_OC_DePrueba LIMPIA"

SumatoriaTotal = 0

If Check1.Value = 1 Or Check4.Value = 1 Then 'SE FILTRA POR OC O FECHA DE OC
            
            'CASO DE OC OC00036831
            'ssql2 = "select uno.VDCODE,UNO.PORHSEQ, DOS.ITEMNO, TRES.CATEGORY,DOS.LOCATION,"
            'ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=TRES.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
            'ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA,(select QTYONHAND from ICILOC where ICILOC.ITEMNO=DOS.ITEMNO and ICILOC.LOCATION=DOS.LOCATION) as QTYONHAND"
            'ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA,A.QTYONHAND"
            'ssql2 = ssql2 & "  from POPORH1 as UNO  left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ"
            'ssql2 = ssql2 & "  left outer join ICITEM as TRES on TRES.ITEMNO=DOS.ITEMNO"
            'ssql2 = ssql2 & "  left outer join ICILOC as A ON A.ITEMNO=DOS.ITEMNO and A.LOCATION=DOS.LOCATION"
            'ssql2 = ssql2 & "  Where DOS.OQRECEIVED>0 AND A.QTYONHAND>0"
            
            ssql2 = "select uno.VDCODE,UNO.PORHSEQ, DOS.ITEMNO, TRES.CATEGORY,/*C.LOCATION,*/CUATRO.RCPHSEQ,"
            

            ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=TRES.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
            ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA/*,A.QTYONHAND  */,"
            'CALSULAR EL COSTO DE LAS LÍNEAS DE LA OC CUANDO NO TENGAN NI SERIE NI LOTE
            ssql2 = ssql2 & "  case when CUATRO.DISCPCT=100 then 0 "
            ssql2 = ssql2 & "  Else ISNULL(CUATRO.EXTENDED-CUATRO.DISCOUNT,0)/CUATRO.OQRECEIVED  end as VALCOSTOOC  "
            'CALSULAR EL COSTO DE LAS LÍNEAS DE LA OC CUANDO NO TENGAN NI SERIE NI LOTE
            '------------------------------------
            ssql2 = ssql2 & ",  UNO.DATE AS FECHAOC"
            '------------------------------------
            ssql2 = ssql2 & "  from POPORH1 as UNO"
            ssql2 = ssql2 & "  left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ"
            ssql2 = ssql2 & "  /*left outer join PORCPL as C on UNO.PORHSEQ=C.PORHSEQ AND C.PORLSEQ=DOS.PORLSEQ  */"
            ssql2 = ssql2 & "  left outer join ICITEM as TRES on TRES.ITEMNO=DOS.ITEMNO"
            ssql2 = ssql2 & "  /*left outer join ICILOC as A ON A.ITEMNO=DOS.ITEMNO and A.LOCATION=C.LOCATION  */"
            'SE AGREGA LA TABLA DE RECIBO PARA HACER EL CALCULO DEL COSTO DE LOS ITEMS DESDE LA OC
            ssql2 = ssql2 & "  left outer join PORCPL as CUATRO on CUATRO.PORHSEQ=DOS.PORHSEQ AND CUATRO.PORLSEQ=DOS.PORLSEQ"
            'SE AGREGA LA TABLA DE RECIBO PARA HACER EL CALCULO DEL COSTO DE LOS ITEMS DESDE LA OC
            ssql2 = ssql2 & "  Where (DOS.OQRECEIVED>0 and CUATRO.OQRECEIVED>0) /*AND A.QTYONHAND>0  */"
            
            If cmbDesdeSucursal.Text <> "TODAS" Or cmbHastaSucursal.Text <> "TODAS" Then 'SUCURSAL
                If Not PasaWhere Then
                    'ssql2 = ssql2 & "  where   DOS.LOCATION='" & cmbSucursal.Text & "'"
                    ssql2 = ssql2 & "  where  (DOS.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (DOS.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
                End If
            End If
            If Check2.Value = 1 Then 'ARTÍCULO
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (DOS.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and  '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "'("
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (DOS.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "')"
                End If
            End If
            If Check3.Value = 1 Then 'CATEGORÍA
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (TRES.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (TRES.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
                End If
            End If
            If Check1.Value = 1 Then 'NO OC
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (UNO.PONUMBER between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (UNO.PONUMBER  between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
                    '" & txtNoOC.Text & "'"
                End If
            End If
            If Check4.Value = 1 Then 'RANGO FECHA
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (UNO.[DATE] between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (UNO.[DATE] between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "')"
                End If
            End If
            Debug.Print ssql2
            
            rsBusca.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
            If rsBusca.EOF = False And rsBusca.BOF = False Then
               Do While rsBusca.EOF = False
                    ArticuloUsa = ArticuloCon(rsBusca!itemno)
                    
'''            If Trim(rsBusca!itemno) = "CPZP005" Then
''' '               MsgBox "Revisar este Caso"
'''            End If
                    
                    strProveedor = Trim(rsBusca!VDCODE)
                    
''                    If Trim(rsBusca!itemno) = "CM2WA000020" Then
''                        MsgBox "Aquí"
''                    End If
                    
                    Select Case ArticuloUsa
                        Case 1 'serie
                            ssql2 = "select uno.RCPNUMBER,UNO.[DATE] RMFECHA,uno.PORHSEQ, uno.PONUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, tres.SERIALNUMF,"
                            ssql2 = ssql2 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED,"
                            ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2 ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER "
                            '---------------------------
                            ssql2 = ssql2 & ",  DOS.PORLSEQ"
                            '---------------------------
                            ssql2 = ssql2 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ "
                            ssql2 = ssql2 & "  left outer join PORCPLS as TRES on dos.RCPHSEQ=tres.RCPHSEQ and dos.RCPLREV=tres.RCPLREV and dos.RCPLSEQ=tres.RCPLSEQ "
                            ssql2 = ssql2 & "  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            ssql2 = ssql2 & "  where UNO.RCPHSEQ='" & Trim(rsBusca!RCPHSEQ) & "' AND UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND dos.PORLSEQ='" & Trim(rsBusca!PORLSEQ) & "' AND dos.OQRECEIVED>0"
                            Debug.Print ssql2
                            rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                rsRecibo.MoveFirst
                                Do While rsRecibo.EOF = False
                                    FacturaRM = Trim(rsRecibo!INVNUMBER)
                                    comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                    comentariosRM = Replace(comentariosRM, Chr(13), "")
                                    comentariosRM = Replace(comentariosRM, "'", "''")
                                    
                                    ''20240212 -- quita tabulador de la cadena de texto
                                    comentariosRM = Replace(comentariosRM, vbTab, " ")
                                    
                                    
                                    '-----------------------
                                    strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                    '-----------------------
                                    
                                    'Serie = Replace(RTrim(rsRecibo!SERIALNUMF), "'", "''")
                                    'Correccion Porque Mandaba Error
                                    '-----------------------
                                    If IsNull(rsRecibo!SERIALNUMF) Then
                                        Serie = ""
                                    Else
                                        Serie = Replace(RTrim(rsRecibo!SERIALNUMF), "'", "")
                                    End If
                                    '-----------------------

                                    'ssql3 = "select * from ICXSER where SERIALNUMF='" & Serie & "' and [STATUS]=1 AND QTYONHAND>0"
                                    ssql3 = "select * from ICXSER as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.SERIALNUMF='" & Serie & "' AND ITEMNO='" & rsRecibo!itemno & "' and A.[STATUS]=1 and ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0)"  'B.QTYONHAND>0"
                                    Debug.Print ssql3
                                    rsSerLot.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                                        Serie = Replace(RTrim(rsSerLot!SERIALNUMF), "'", "''")
                                        ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                                        ssql4 = ssql4 & "  left outer join ICTREDS as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                                        ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsSerLot!Location) & "' AND dos.ITEMNO='" & Trim(rsRecibo!itemno) & "' and tres.SERIALNUMF='" & Serie & "' AND UNO.DOCTYPE=3"
                                        Debug.Print ssql4
                                        rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                                        If rsTrans.EOF = False And rsTrans.BOF = False Then
                                        
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                        & "'1','" & RTrim(Serie) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                        & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                        & "','" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsRecibo!itemno)) & "'" _
                                                        & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        ' & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                        Else
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                        & "'1','" & RTrim(Serie) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                        & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                        & "','01-01-1900',' ','" & ValorCosto(Trim(rsRecibo!itemno)) _
                                                        & "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & Trim(rsBusca!FechaOC) & "')" '  almacena cadena
                                                        '& "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                    
                                        End If
                                        Debug.Print strCadena
                                        cnDB.Execute strCadena
                                        rsTrans.Close
                                    End If
                                    rsSerLot.Close
                                    rsRecibo.MoveNext
                                Loop
                            Else
                            '---------------------------------------------------
                            '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'0',' ',' ',' ',' '" _
                                            & ",'" & ValorSALEORDER(Trim(rsBusca!PORHSEQ)) & "'," & "'0','" & 0 _
                                            & "',' ',' ','" & ValorCosto(Trim(rsBusca!itemno)) _
                                            & "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ',' ','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & Trim(rsBusca!FechaOC) & "')" '  almacena cadena **&&
                                            '& "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ',' ','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "')" 'almacena cadena
                            '---------------------------------------------------
                            '---------------------------------------------------
                                cnDB.Execute strCadena
                            End If
                            rsRecibo.Close
                        Case 2 'lote
                        
                            'ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            'ssql2 = ssql2 & " UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, tres.LOTNUMF,(Select Value FROM ICITEMO Where "
                            'ssql2 = ssql2 & " " ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
                            
                            '''*** 20230710
                            '''*** ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            ssql2 = "select DOS.LOCATION, uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            
                            ssql2 = ssql2 & " UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, REPLACE(tres.LOTNUMF,CHAR(39),'') AS LOTNUMF, "
                            ssql2 = ssql2 & " (Select Value FROM ICITEMO Where ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
                            
                            ssql2 = ssql2 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED"
                            ssql2 = ssql2 & "  ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER  "
                            
                            '---------------------------
                            ssql2 = ssql2 & ",  DOS.PORLSEQ"
                            '---------------------------

                            ssql2 = ssql2 & " from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ"
                            ssql2 = ssql2 & "  left outer join PORCPLL as TRES on dos.RCPHSEQ=tres.RCPHSEQ and dos.RCPLREV=tres.RCPLREV and dos.RCPLSEQ=tres.RCPLSEQ  "
                            ssql2 = ssql2 & "  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            
                            ssql2 = ssql2 & "  where UNO.RCPHSEQ='" & Trim(rsBusca!RCPHSEQ) & "' AND UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND dos.PORLSEQ='" & Trim(rsBusca!PORLSEQ) & "' and DOS.OQRECEIVED>0"
                            Debug.Print ssql2
                            rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                rsRecibo.MoveFirst
                                Do While rsRecibo.EOF = False
                                    FacturaRM = Trim(rsRecibo!INVNUMBER)
                                    comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                    comentariosRM = Replace(comentariosRM, Chr(13), "")
                                    comentariosRM = Replace(comentariosRM, "'", "''")
                                    
                                    ''20240212 -- quita tabulador de la cadena de texto
                                    comentariosRM = Replace(comentariosRM, vbTab, " ")
                                    
                                    
                                    '-----------------------
                                    strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                    '-----------------------
                                    
                                    'ssql3 = "select * from ICXLOT where LOTNUMF='" & Trim(rsRecibo!LOTNUMF) & "' and QTYAVAIL>0 AND QTYONHAND>0"
                                    
                                    If Check4.Value = 1 Then '--- si BUSCA por RANGO DE FECHAS de OC
                                        '''*** 20230710
                                        '''*** ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                         ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         
                                         Debug.Print ssql3
                                         
                                            rsSerLotVal.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsSerLotVal.EOF = False And rsSerLotVal.BOF = False Then
                                                '''---
                                            Else
                                                ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                                Debug.Print ssql3
                                            End If
                                            rsSerLotVal.Close
                                         
                                    Else
                                        '''*** 20230710
                                        'ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                        
                                        ''*** 20240131 -- Correccion, debido a que cuando hay transferencias cambia el Location y no muestra el renglon
                                         ''ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         
                                         ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         Debug.Print ssql3
                                         
                                            rsSerLotVal.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsSerLotVal.EOF = False And rsSerLotVal.BOF = False Then
                                                '''---
                                            Else
                                                ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                                Debug.Print ssql3
                                            End If
                                            rsSerLotVal.Close
                                    
                                    End If
                                    Debug.Print ssql3
                                    rsSerLot.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                                        rsSerLot.MoveFirst
                                        Do While rsSerLot.EOF = False
                                                                                
                                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                                            ssql4 = ssql4 & "  left outer join ICTREDL as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsSerLot!Location) & "' AND dos.ITEMNO='" & Trim(rsRecibo!itemno) & "' and tres.LOTNUMF='" & RTrim(rsSerLot!LOTNUMF) & "' AND UNO.DOCTYPE=3"
                                            Debug.Print ssql4
                                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                            & Trim(rsSerLot!QTYAVAIL) & ",'" & RTrim(rsSerLot!LOTNUMF) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                            & "','" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsRecibo!itemno)) & "'" _
                                                            & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                            '& ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                            
                                            Else
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                            & Trim(rsSerLot!QTYAVAIL) & ",'" & RTrim(rsSerLot!LOTNUMF) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                            & "','01-01-1900',' ','" & ValorCosto(Trim(rsRecibo!itemno)) _
                                                            & "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                            '& "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                
                                            End If
                                            rsTrans.Close
                                            cnDB.Execute strCadena
                                            rsSerLot.MoveNext
                                        Loop
                                    Else
                                        'strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!ITEMNO) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                        '            & "'0',' ',' ','" & Trim(rsRecibo!PONUMBER) & "',' '" _
                                        '            & ",'" & ValorSALEORDER(Trim(rsBusca!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                        '            & "',' ',' ','" & ValorCosto(Trim(rsBusca!ITEMNO)) _
                                        '            & "','" & ValorPL(Trim(rsBusca!ITEMNO)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "')" 'almacena cadena
                                        'cnDB.Execute strCadena
                                    End If
                                    rsSerLot.Close
                                    rsRecibo.MoveNext
                                Loop
                            End If
                            rsRecibo.Close
                        Case 3
                            'If ArticulosSinSerieNiLote <> "" Then
                            'Cantidad = Existencia(Trim(rsBusca!itemno))
                            Dim rsOpcional As Recordset
                            Set rsOpcional = New Recordset
                            ssql = "select SUM(QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST) EXISTENCIA,LOCATION from ICILOC where ITEMNO='" & Trim(rsBusca!itemno) & "' GROUP BY ITEMNO,LOCATION"
                            Debug.Print ssql
                            rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsOpcional.EOF = False And rsOpcional.BOF = False Then
                                rsOpcional.MoveFirst
                                Do While rsOpcional.EOF = False
                                    If rsOpcional!Existencia > 0 Then
                                    
                                        ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, uno.[DATE] as RMFECHA,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ "
                                        ssql2 = ssql2 & "   left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ  where UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND ITEMNO='" & Trim(rsBusca!itemno) & "'"
                                        
                                        Debug.Print ssql2
                                        rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                                        If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                            FacturaRM = Trim(rsRecibo!INVNUMBER)
                                            comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                            comentariosRM = Replace(comentariosRM, Chr(13), "")
                                            comentariosRM = Replace(comentariosRM, "'", "''")
                                            
                                            ''20240212 -- quita tabulador de la cadena de texto
                                            comentariosRM = Replace(comentariosRM, vbTab, " ")
                                            
                                            
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsOpcional!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                                        & "'" & rsOpcional!Existencia & "',' ','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" & Trim(rsRecibo!RCPNUMBER) & "','" _
                                                        & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        '& ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "')" 'sin datos
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                        
                                            Debug.Print strCadena
                                        Else
                                            FacturaRM = " "
                                            comentariosRM = " "
                                            
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "',' ','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                                        & "'" & rsOpcional!Existencia & "',' ','01-01-1900',' ',' ',' ',' ','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "','" & Left(Trim(rsBusca!NUMPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        '& "'" & rsOpcional!Existencia & "',' ','01-01-1900',' ',' ',' ',' ','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "')" 'sin datos
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                        
                                            Debug.Print strCadena
                                        End If
                                        cnDB.Execute strCadena
                                        rsRecibo.Close
                                    End If
                                    rsOpcional.MoveNext
                                Loop
                                rsOpcional.Close
                            End If
                    End Select
                    rsBusca.MoveNext
                Loop
                InfoDisponible = True
            Else
                'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
                InfoDisponible = False
            End If
            
            rsBusca.Close
    
Else 'NO SE FILTRA POR OC O FECHA DE OC

    ssql = "select uno.LOCATION,uno.ITEMNO,(Select Value FROM ICITEMO Where ITEMNO=UNO.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,(QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST) QTYONHAND  from ICILOC AS UNO left outer join ICITEM as DOS on uno.ITEMNO=dos.ITEMNO where (QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST)>0"
    'ssql = "select uno.LOCATION,uno.ITEMNO,(Select Value FROM ICITEMO Where ITEMNO=UNO.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,UNO.QTYONHAND  from ICILOC AS UNO left outer join ICITEM as DOS on uno.ITEMNO=dos.ITEMNO where (UNO.QTYONHAND+UNO.QTYRENOCST)>0"
    If cmbDesdeSucursal.Text <> "TODAS" Or cmbHastaSucursal.Text <> "TODAS" Then 'SUCURSAL
        If Not PasaWhere Then
            ssql = ssql & "  where  (UNO.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
            PasaWhere = True
        Else
            ssql = ssql & "  and  (UNO.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
        End If
    End If
    If Check2.Value = 1 Then 'ARTÍCULO
        If Not PasaWhere Then
            ssql = ssql & "  where  (UNO.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and  '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "'("
            PasaWhere = True
        Else
            ssql = ssql & "  and  (UNO.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "')"
        End If
    End If
    If Check3.Value = 1 Then 'CATEGORÍA
        If Not PasaWhere Then
            ssql = ssql & "  where  (DOS.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
            PasaWhere = True
        Else
            ssql = ssql & "  and  (DOS.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
        End If
    End If
    Debug.Print ssql
    rsBusca.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBusca.EOF = False And rsBusca.BOF = False Then
        rsBusca.MoveFirst
        Do While rsBusca.EOF = False
            ArticuloUsa = ArticuloCon(rsBusca!itemno)
            
''''            If Trim(rsBusca!itemno) = "ATCM010" Then
''''                                         'MK4LI000018
''''                                         'CB2YI000001
''''                                         'XT4SE000022
''''             MsgBox "Revisar este Caso"
''''            End If

'''If fnGetFirst() <> "CB2YI000001" Then
'''    'MsgBox "Revisar este Caso"
'''End If
                       
            Select Case ArticuloUsa
                Case 1 'series
                    ssql2 = "select SERIALNUMF,1 as CANTIDAD from ICXSER where ITEMNUM='" & Trim(rsBusca!itemno) & "' and [STATUS]=1 and LOCATION='" & Trim(rsBusca!Location) & "'" ' and LIFECONT1=0 REVISIÓN DE MARICELA TERRAZAS, NO PASA FILTRO DE LIFECONT
                    Debug.Print ssql2
                    rsSerLot.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                        Do While rsSerLot.EOF = False
                        
                            Serie = Replace(RTrim(rsSerLot!SERIALNUMF), "'", "''")
                            
                            ssql3 = "select uno.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,"
                            ssql3 = ssql3 & "  isnull(CUATRO.INVNUMBER,'') as INVNUMBER, DOS.UNITWEIGHT,"
                            ssql3 = ssql3 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED"
                            'ssql3 = ssql3 & "  (select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA"
                            '--------------------------------
                            ssql3 = ssql3 & ",  DOS.PORLSEQ"
                            '--------------------------------
                            ssql3 = ssql3 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ"
                            ssql3 = ssql3 & "  left outer join PORCPLS as TRES on DOS.RCPHSEQ=TRES.RCPHSEQ and DOS.RCPLSEQ=TRES.RCPLSEQ  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.SERIALNUMF='" & Serie & "'"
                            
                            ''20231002
                            TipoArticulo = Left(Trim(rsBusca!itemno), 2)
                            
                            ''-- 20240116
                            ''-- If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Then
                            If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Or TipoArticulo = "CM" Then
                                If Mid(Trim(rsBusca!itemno), 3, 1) = 5 Or Mid(Trim(rsBusca!itemno), 3, 1) = 6 Or Mid(Trim(rsBusca!itemno), 3, 1) = 7 Then
                                    ssql3 = ssql3 & "  ORDER BY UNO.DATE"
                                End If
                            End If
                            
                            Debug.Print ssql3
                            
                            rsRecibo.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                            'INSERT INTO Existencias_OC VALUES('Usuario','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                            'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                strProveedor = Trim(rsRecibo!VDCODE)
                                FacturaRM = Trim(rsRecibo!INVNUMBER)
                                comentariosRM = Trim(rsRecibo!DESCRIPTIO) & Trim(rsRecibo!REFERENCE)
                                comentariosRM = Replace(comentariosRM, Chr(13), "")
                                comentariosRM = Replace(comentariosRM, "'", "''")
                                
                                ''20240212 -- quita tabulador de la cadena de texto
                                comentariosRM = Replace(comentariosRM, vbTab, " ")
                                
                                
                                '-----------------------
                                strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                '-----------------------

                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsSerLot!Cantidad) & "','" & Trim(Serie) & "','" & Trim(rsRecibo!Date) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) & "'," 'datos RM
                                            strPORHSEQ = Trim(rsRecibo!PORHSEQ)
                                            NoLineaOC = EncuentraLineaOC(Trim(rsRecibo!PORHSEQ), Trim(rsBusca!itemno), Trim(rsRecibo!UNITWEIGHT))
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                            Else
                                strProveedor = " "
                                FacturaRM = " "
                                comentariosRM = " "
                                strNoPV = ""
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' " _
                                            & "',' '," & "' ',' '," 'sin datos de RM
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                                            strPORHSEQ = "0"
                                            NoLineaOC = 0
                            End If
                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                            ssql4 = ssql4 & "  left outer join ICTREDS as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsBusca!Location) & "' AND dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.SERIALNUMF='" & Serie & "' AND UNO.DOCTYPE=3"
                            Debug.Print ssql4
                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                strCadena = strCadena & "'" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')" 'sin datos de transferencia
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                            Else
                                strCadena = strCadena & "'01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) _
                                & "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')" 'almacena cadena
                                '& "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'sin datos de transferencia
                            End If
                            Debug.Print strCadena
                            cnDB.Execute strCadena
                            rsTrans.Close
                            rsRecibo.Close
                            
                            rsSerLot.MoveNext
                        Loop
                    Else
                    
                    End If
                    rsSerLot.Close
                Case 2 'lotes
                    '---------------------------------
                    strLote = ""
                    '---------------------------------
                    
                    Debug.Print strNoPV

                    ssql2 = "select LOTNUMF,QTYAVAIL from ICXLOT where ITEMNUM='" & Trim(rsBusca!itemno) & "' and QTYAVAIL>0 and LOCATION='" & Trim(rsBusca!Location) & "'"
                    Debug.Print ssql2
                    rsSerLot.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                        Do While rsSerLot.EOF = False
                            '---------------------------------
                            strLote = rsSerLot!LOTNUMF
                            '---------------------------------
                            
                            ' CON EL DESCUENTO DEL 100% EN LA OC EL COSTO DEL ARTICULO SERÁ 0 - MARICELA TERRAZAS 25 01 2016
                            'ssql3 = "select UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,case when DOS.DISCPCT=100 then ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED "
                            'ssql3 = ssql3 & "  Else ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED  end as EXTENDED" ',(select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA
                            
                            '''*** 20230710
                            '''*** ssql3 = "select UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,case when DOS.DISCPCT=100 then 0 "
                            ssql3 = "select DOS.location, UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,case when DOS.DISCPCT=100 then 0 "
                            
                            ssql3 = ssql3 & "  Else ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED  end as EXTENDED" ',(select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA
                            ssql3 = ssql3 & "  ,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER, DOS.UNITWEIGHT "
                            
                            '--------------------------------
                            ssql3 = ssql3 & ",  DOS.PORLSEQ"
                            '--------------------------------

                            ssql3 = ssql3 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ  "
                            ssql3 = ssql3 & "  left outer join PORCPLL as TRES on DOS.RCPHSEQ=TRES.RCPHSEQ and DOS.RCPLSEQ=TRES.RCPLSEQ  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            
                            '''*** 20230710
                            '''*** ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and TRES.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "'"
                            ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and DOS.location='" & Trim(rsBusca!Location) & "' and TRES.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "'"
                            
                            Debug.Print ssql3
                            
                            rsRecibo.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                            
                            'INSERT INTO Existencias_OC VALUES('Usuario','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                            'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                            
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                strProveedor = Trim(rsRecibo!VDCODE)
                                FacturaRM = Trim(rsRecibo!INVNUMBER)
                                comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                comentariosRM = Replace(comentariosRM, Chr(13), "")
                                comentariosRM = Replace(comentariosRM, "'", "''")
                                
                                ''20240212 -- quita tabulador de la cadena de texto
                                comentariosRM = Replace(comentariosRM, vbTab, " ")
                                
                                
                                '-----------------------
                                strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                '-----------------------
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsSerLot!QTYAVAIL) & "','" & Trim(strLote) & "','" & Trim(rsRecibo!Date) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) & "'," 'datos RM
                                            strPORHSEQ = Trim(rsRecibo!PORHSEQ)
                                            NoLineaOC = EncuentraLineaOC(Trim(rsRecibo!PORHSEQ), Trim(rsBusca!itemno), Trim(rsRecibo!UNITWEIGHT))
                                '---------------------------------------------------
                                '---------------------------------------------------

                            Else
                                strProveedor = " "
                                FacturaRM = " "
                                comentariosRM = ""
                                
                                strNoPV = ""
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' " _
                                            & "',' '," & "' ',' '," 'sin datos de RM
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                                            strPORHSEQ = "0"
                                            NoLineaOC = 0
                            End If
                            
                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                            ssql4 = ssql4 & "  left outer join ICTREDL as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsBusca!Location) & "' AND dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "' AND UNO.DOCTYPE=3"
                            
                            Debug.Print ssql4
                            
                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                strCadena = strCadena & "'" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')"
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                            Else
                                strCadena = strCadena & "'01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')"
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'sin datos de transferencia
                            End If
                            Debug.Print strCadena
                            cnDB.Execute strCadena
                            rsTrans.Close
                            rsRecibo.Close
                        
                            rsSerLot.MoveNext
                        Loop
                    Else
                    
                    End If
                    rsSerLot.Close
                Case 3 'nada
                        'INSERT INTO Existencias_OC VALUES('Usuario','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                        'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                        
                        '---------------------------------------------------
                        '---------------------------------------------------
                        strCadena = "INSERT INTO Existencias_OC VALUES ('" & usuario & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                    & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' ',' ',' ','0','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','0','" & strVendorID & "','','','')" 'sin datos
                                    '& "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' ',' ',' ','0','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','0','','')" 'sin datos
                        '---------------------------------------------------
                        '---------------------------------------------------
                                    
                        Debug.Print strCadena
                        cnDB.Execute strCadena
            End Select
            rsBusca.MoveNext
        Loop
        InfoDisponible = True
        cmdExporta.Enabled = True
    Else
        'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
        InfoDisponible = False
    End If
    rsBusca.Close
    
End If

If InfoDisponible Then
    '---------------------------------------------------
    '---------------------------------------------------
    ssql = "select Articulo,SUM(Cantidad) Cantidad,NoOC,NoRM,NoLinea,Sucursal from Existencias_OC where Usuario='" & usuario & "' group by NoOC,NoRM,Sucursal,Articulo,NoLinea"
    '---------------------------------------------------
    '---------------------------------------------------
    Debug.Print ssql
    
    rsBusca.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBusca.EOF = False And rsBusca.BOF = False Then
        Do While rsBusca.EOF = False
        
''''            If Trim(rsBusca!articulo) = "MK4LI000018" Then
''''                 MsgBox "Revisar este Caso"
''''            End If
        
            '---------------------------------------------------
            '---------------------------------------------------
            ssql2 = "select * from Existencias_OC where Usuario='" & usuario & "' and NoOC='" & Trim(rsBusca!NoOC) & "' and NoRM='" & Trim(rsBusca!NoRM) & "' and Sucursal='" & Trim(rsBusca!Sucursal) & "' and Articulo='" & Trim(rsBusca!articulo) & "' and NoLinea='" & Trim(rsBusca!NoLinea) & "'"
            '---------------------------------------------------
            '---------------------------------------------------
            Debug.Print ssql2
            
            rsDemas.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
            If rsDemas.EOF = False And rsDemas.BOF = False Then
                Do While rsDemas.EOF = False
                    If Trim(rsDemas!SerLot) <> "" Then strSerLot = Trim(rsDemas!SerLot) & "," & strSerLot
                    Codigo2 = Trim(rsDemas!Codigo2)
                    FechaRM = Trim(rsDemas!FechaRM)
                    
                    ''**  Sales Order se toma de la Orden de Compra
                    SalesO = Trim(rsDemas!SalesOrd)
                    
                    NoLinea = Trim(rsDemas!NoLinea)
                    CostoOc = rsDemas!CostoOc
                    FechaTRF = Trim(rsDemas!FechaTRF)
                    NoTFR = Trim(rsDemas!NoTRF)
                    strProveedor = Trim(rsDemas!Proveedor)
                    Costo = Format(Trim(rsDemas!Costo), "###,###.#0")
                    PricoLista = Format(Trim(rsDemas!PrecioLista), "###,###.#0")
                    Comentario = Replace(Trim(rsDemas!Comentarios), Chr(9), " ")
                    ComentarioTRF = Replace(Trim(rsDemas!ComentariosTRF), Chr(9), " ")
                    NoFacRM = Trim(rsDemas!NoFactura)
                    NoLineaOC = Trim(rsDemas!NoLineaOC)
                    comenRM = Replace(Trim(rsDemas!comentariosRM), Chr(13) + Chr(10), "")
                    
                    '-------------------------------------------------
                    strNoPV = Trim(rsDemas!NUMPV)
                    strDatosAnexos = fnStrDatosAnexos(strNoPV) ' *** Trae Datos de la Orden de Venta
                    lsDatosAnexos = Split(strDatosAnexos, "~")
                    strNumPedido = strNoPV
                    strIDCliente = lsDatosAnexos(1)
                    strCliente = lsDatosAnexos(2)
                    strIDEjecutivo = lsDatosAnexos(3)
                    
                    'SalesO = lsDatosAnexos(4)
                    '-------------------------------------------------
                    
                    rsDemas.MoveNext
                    CostoOc = Format(CostoOc * rsBusca!Cantidad, "###,###.#0")
                    no_producto = encuentraproducto(Trim(rsBusca!articulo))
                    
                Loop
                
                
'''                If Trim(rsBusca!NoOC) = "OC00148976" Then
'''                    MsgBox "Revisar Caso"
'''''                End If

                '-------------------------------------------------
                If Trim(rsBusca!NoOC) <> "" Then
                    sDatosAgregados = fnDatosAgreados(Trim(rsBusca!NoOC))
                    lDatosAgregados = Split(sDatosAgregados, "~")

                    sDEALID = IIf(lDatosAgregados(0) = "", 0, lDatosAgregados(0))
                    sDEALID = Round(sDEALID, 0)
                    
'--------------------------------------
'--------------------------------------
            sIDENDUSER = lDatosAgregados(1)

            TipoArticulo = Left(Trim(rsBusca!articulo), 2)
            '' 20240116
            '' -- If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Then
            If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Or TipoArticulo = "CM" Then
                sIDENDUSER = lDatosAgregados(1)
                sENDUSER = strEndUserName(sIDENDUSER)
            Else
                If strNoPV <> "" Then
                        sIDENDUSER = fnBuscaIDEndUserXPV(strNoPV, Trim(rsBusca!articulo))
                        sENDUSER = strEndUserNameXPV(sIDENDUSER, Trim(rsBusca!articulo))
                Else
                        sIDENDUSER = 0
                        sENDUSER = ""
                End If
            End If
            
'--------------------------------------
'--------------------------------------
                    
                    sOCCLIENTE = lDatosAgregados(2)
                    
                    'EN LOS CASOS en los que la OC DEL CLIENTE ESTE EN BLANCO Y SE CUENTE CON NUMERO DE PV ENTONCES
                    If Trim(sOCCLIENTE) = "" And strNoPV <> "" Then
                        sOCCLIENTE = fnBuscaOCCLIENTE(strNoPV)
                    End If
                    
                    
                    sSALEORDER = lDatosAgregados(3)
                    sFechaOC = lDatosAgregados(4)
                    
                    sFechaOC = Mid(sFechaOC, 1, 4) & "-" & Mid(sFechaOC, 5, 2) & "-" & Mid(sFechaOC, 7, 2)
                    
                    
                    strIDEjecutivo = strIDEjecutivo
                    strEjecutivo = FnStrEjecutivo(strIDEjecutivo)
                Else
                    ' No deberia pasar por aqui
                    sDEALID = 0
                    sIDENDUSER = 0
                    sENDUSER = ""
                    sOCCLIENTE = ""
                    sSALEORDER = ""
                    strIDEjecutivo = 0
                    strEjecutivo = ""
                    sFechaOC = ""
                End If
                
                sFACTCOMER = fnBuscaFACTCOMER(Trim(rsBusca!NoRM))
                
                ''20240206
                sFACTPACK = fnBuscaFACTPACK(Trim(rsBusca!NoRM))
                
                '-------------------------------------------------
                                
''                strCadena = Trim(rsBusca!Sucursal) & vbTab & Trim(rsBusca!articulo) & vbTab & Trim(Codigo2) & vbTab & Trim(rsBusca!Cantidad) & vbTab & Trim(strSerLot) & vbTab _
''                            & Trim(FechaRM) & vbTab & Trim(rsBusca!NoRM) & vbTab & Trim(NoFacRM) & vbTab & Trim(comenRM) & vbTab & Trim(rsBusca!NoOC) & vbTab & Trim(NoLineaOC) & vbTab _
''                            & Trim(SalesO) & vbTab & Trim(NoLinea) & vbTab & Trim(CostoOc) & vbTab & Trim(FechaTRF) & vbTab & Trim(NoTFR) & vbTab & Trim(Costo) & vbTab _
''                            & Trim(PricoLista) & vbTab & Trim(Comentario) & vbTab & Trim(ComentarioTRF) & vbTab & strProveedor & vbTab & no_producto
                '----------------------------------------------------------
                strCadena = Trim(rsBusca!Sucursal) & vbTab & Trim(rsBusca!articulo) & vbTab & Trim(Codigo2) & vbTab & Trim(rsBusca!Cantidad) & vbTab & Trim(strSerLot) & vbTab _
                            & Trim(FechaRM) & vbTab & Trim(rsBusca!NoRM) & vbTab & Trim(NoFacRM) & vbTab & Trim(comenRM) & vbTab & Trim(rsBusca!NoOC) & vbTab & Trim(NoLineaOC) & vbTab _
                            & Trim(SalesO) & vbTab & Trim(NoLinea) & vbTab & Trim(CostoOc) & vbTab & Trim(FechaTRF) & vbTab & Trim(NoTFR) & vbTab & Trim(Costo) & vbTab _
                            & Trim(PricoLista) & vbTab & Trim(Comentario) & vbTab & Trim(ComentarioTRF) & vbTab & strProveedor & vbTab & no_producto & vbTab _
                            & sDEALID & vbTab & sIDENDUSER & vbTab & sENDUSER & vbTab & sOCCLIENTE & vbTab _
                            & Trim(strNumPedido) & vbTab & Trim(strIDCliente) & vbTab & Trim(strCliente) & vbTab _
                            & strIDEjecutivo & vbTab & strEjecutivo & vbTab & sFechaOC & vbTab & sFACTCOMER & vbTab & sFACTPACK
                            
                            ''20240206
                            '' strIDEjecutivo & vbTab & strEjecutivo & vbTab & sFechaOC & vbTab & sFACTCOMER
                '----------------------------------------------------------
                            
                SSOleDBGrid1.AddItem strCadena
                SumatoriaTotal = SumatoriaTotal + CostoOc
                strSerLot = ""
            End If
            rsDemas.Close
            rsBusca.MoveNext
        Loop

        strCadena = "TOTAL" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab _
                    & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab _
                    & vbTab & "" & vbTab & SumatoriaTotal & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
                    
                    
        SSOleDBGrid1.AddItem strCadena
        
        MsgBox "Se ha completado la busqueda de información con los parámetros establecidos", vbInformation + vbOKOnly, "Información"
        cmdExporta.Enabled = True
    Else
        MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
    End If
    rsBusca.Close
Else
    MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
End If

Set rsBusca = Nothing
Set rsSerLot = Nothing
Set rsRecibo = Nothing
Set rsTrans = Nothing
Set rsDemas = Nothing

Screen.MousePointer = vbNormal
frmExisteOC.Enabled = True

Exit Sub

RutinaError:
If Err.Number <> -2147217873 Then 'NO SE MUESTRA EL MENSAJE SI ES REGISTRO DUPLICADO
    MsgBox Err.Number & "  " & Err.Description, vbCritical + vbOK, "Error"
End If
Resume Next
End Sub

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'''Private Function fnStrGetNoPvC(strNoOc As String, strPORLSEQ As String) As String
'''
'''    Dim ssql As String
'''    Dim ssql2 As String
'''    Dim rsNoPV As Recordset
'''    Set rsNoPV = New Recordset
'''
'''    Dim rsNoPV2 As Recordset
'''    Set rsNoPV2 = New Recordset
'''
'''    Dim k As Integer
'''    fnStrGetNoPvC = ""
'''    k = 0
'''
'''    'ssql = "SELECT Top 1 OEONUMBER FROM POPORH1 as UNO left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ where PONUMBER = '" & Trim(strNoOc) & "' and OEONUMBER <> '' ORDER BY PORLREV DESC"
'''    ssql = "SELECT OEONUMBER FROM POPORH1 as UNO left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ where PONUMBER = '" & Trim(strNoOc) & "' and PORLSEQ = '" & Trim(strPORLSEQ) & "'"
'''    Debug.Print ssql
'''
'''    rsNoPV.Open ssql, cnDB, adOpenStatic, adLockReadOnly
'''    If rsNoPV.EOF = False And rsNoPV.BOF = False Then
'''
'''        If Trim(rsNoPV!OEONUMBER) = "" Then
'''
'''            ssql2 = "SELECT Top 1 OEONUMBER FROM POPORH1 as UNO left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ where PONUMBER = '" & Trim(strNoOc) & "' and OEONUMBER <> '' ORDER BY PORLREV DESC"
'''            Debug.Print ssql2
'''            rsNoPV2.Open ssql2, cnDB, adOpenStatic, adLockReadOnly
'''
'''            If rsNoPV2.EOF = False And rsNoPV2.BOF = False Then
'''                fnStrGetNoPvC = Mid(Trim(rsNoPV2!OEONUMBER), 1, 10)
'''            Else
'''                fnStrGetNoPvC = ""
'''            End If
'''
'''            rsNoPV2.Close
'''            Set rsNoPV2 = Nothing
'''
'''        Else
'''            fnStrGetNoPvC = Mid(Trim(rsNoPV!OEONUMBER), 1, 10)
'''        End If
'''    Else
'''        fnStrGetNoPvC = ""
'''    End If
'''
'''    rsNoPV.Close
'''    Set rsNoPV = Nothing
'''
'''End Function

Private Function fnStrGetNoPvC(strNoOc As String, strPORLSEQ As String) As String

    Dim strParteNumerica As String
    Dim ssql As String
    Dim ssql2 As String
    Dim ssql3 As String
    
    Dim rsNoPV As Recordset
    Set rsNoPV = New Recordset
    
    Dim rsNoPV2 As Recordset
    Set rsNoPV2 = New Recordset
    
    Dim rsNoPV3 As Recordset
    Set rsNoPV3 = New Recordset
    
    Dim k As Integer
    fnStrGetNoPvC = ""
    k = 0
    
'''
'''    If Trim(strNoOc) = "OC00150002" Then
'''        'MsgBox "Seguir"
'''    End If
    
    ssql = "SELECT OEONUMBER FROM POPORH1 as UNO left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ where PONUMBER = '" & Trim(strNoOc) & "' and PORLSEQ = '" & Trim(strPORLSEQ) & "'"
    Debug.Print ssql
    
    rsNoPV.Open ssql, cnDB, adOpenStatic, adLockReadOnly
    If rsNoPV.EOF = False And rsNoPV.BOF = False Then
       '**** SI ENCUENTRA LA PARTIDA EXACTA ENTONCES ...***************************************************
       '********************************************************
        If Trim(rsNoPV!OEONUMBER) = "" Then
                '**** SI PARA LA PARTIDA EXACTA EL DATO ESTA VACIO, ENTONCES ...***************************************************
                ' BUSCA PARA ESA OC LA PRIMERA PARTIDA QUE TIENE EL DATO
                        ssql2 = "SELECT Top 1 OEONUMBER FROM POPORH1 as UNO left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ where PONUMBER = '" & Trim(strNoOc) & "' and OEONUMBER <> '' ORDER BY PORLREV DESC"
                        Debug.Print ssql2
                        rsNoPV2.Open ssql2, cnDB, adOpenStatic, adLockReadOnly
                        
                        If rsNoPV2.EOF = False And rsNoPV2.BOF = False Then
                        ' PONE LA PRIMERA PARTIDA DE LA OC A LA PARTIDA QUE NO TIENE EL DATO DEL PV
                        ' 2022/02/22, SOLICITA MARICELA YA NO SE REPLIQUE LA PRIMERA PARTIDA EN LAS PARTIDAS SIN PV
                            '% fnStrGetNoPvC = Mid(Trim(rsNoPV2!OEONUMBER), 1, 10)
                        Else
                        '-------------------------------------
                            ' SI NO ENCUENTRA UNA PARTIDA QUE TENGA EL DATO DEL PV (NINGUNA PARTIDA TIENE PV),
                            ' ENTONCES LO BUSCA EN EL CAMPO DE REFERENCIA
                            
                                ssql3 = "select REFERENCE from POPORH1 as A left outer join POPORL as B on A.PORHSEQ=B.PORHSEQ WHERE A.PORHSEQ = B.PORHSEQ AND OEONUMBER = '' AND REFERENCE NOT Like '%STOCK%' AND PONUMBER = '" & Trim(strNoOc) & "'"
                                Debug.Print ssql3
                                rsNoPV3.Open ssql3, cnDB, adOpenStatic, adLockReadOnly
                            
                                If rsNoPV3.EOF = False And rsNoPV3.BOF = False Then
                                    '-- Poner dato que viene en referencia, SIEMPRE Y CUANDO INICE CON LA PALABR [PV]
                                    ' SI ENCUENTRA EN REFERENCIA UN NUMERO DE PV AL INICIO,
                                    ' Y EL CAMPO DE REFERNCIA NO CONTIENE LA PALABRA STOCK
                                    ' ENTONCES REPLICA EL PV EN LAS PARTIDAS QUE NO TIENEN EL DATO
                                    If Mid(Trim(rsNoPV3!REFERENCE), 1, 2) = "PV" Then
                                    
                                        strParteNumerica = Mid(Trim(rsNoPV3!REFERENCE), 3, 8)
                                        strParteNumerica = dejaNumeros(strParteNumerica)
                                    
                                        If Len(strParteNumerica) = 8 Then
                                            'SI TINE LA ESTRUCTURA CORRECTA, ENTONCES
                                            fnStrGetNoPvC = Mid(Trim(rsNoPV3!REFERENCE), 1, 10)
                                        Else
                                            'SI no TINE LA ESTRUCTURA CORRECTA, ENTONCES
                                            fnStrGetNoPvC = ""
                                        End If
                                    Else
                                        ' NUMERO DE REFERENCIA NO CONTIENE EL PREFIJO PV
                                        fnStrGetNoPvC = ""
                                    End If
                                Else
                                    'EL CAMPO DE REFERENCIA TIENE LA PALABRA: STOCK
                                    fnStrGetNoPvC = ""
                                End If
                        '-------------------------------------
                        End If
            
            rsNoPV2.Close
            Set rsNoPV2 = Nothing
           
        Else 'If Trim(rsNoPV!OEONUMBER) = "" Then
            fnStrGetNoPvC = Mid(Trim(rsNoPV!OEONUMBER), 1, 10)
        End If
       '********************************************************
       '********************************************************
        
    Else
       '**** SI no ENCUENTRA LA PARTIDA EXACTA ENTONCES ...***************************************************
       '********************************************************
        fnStrGetNoPvC = ""
    End If
    
    rsNoPV.Close
    Set rsNoPV = Nothing
    
End Function

Function dejaNumeros(cadenaTexto As String) As String
  Const listaNumeros = "0123456789"
  Dim cadenaTemporal As String
  Dim i As Integer

  cadenaTexto = Trim$(cadenaTexto)
  If Len(cadenaTexto) = 0 Then
    Exit Function
  End If
 
  cadenaTemporal = ""

  For i = 1 To Len(cadenaTexto)
    If InStr(listaNumeros, Mid$(cadenaTexto, i, 1)) Then
      cadenaTemporal = cadenaTemporal + Mid$(cadenaTexto, i, 1)
    End If
  Next
  dejaNumeros = cadenaTemporal
End Function
Private Function fnStrDatosAnexos(strNoPV As String) As String
    Dim ssql As String
    Dim rsPV As Recordset
    Set rsPV = New Recordset
    
    ssql = "SELECT ORDUNIQ, ORDNUMBER, CUSTOMER, (SELECT NAMECUST FROM ARCUS Where IDCUST = CUSTOMER) AS NAMECUST, SALESPER1 AS IDEJECUTIVO, "
    ssql = ssql & " PONUMBER "
    ssql = ssql & " FROM OEORDH where ORDNUMBER = '" & Trim(strNoPV) & "'"
        
    Debug.Print ssql
    
    rsPV.Open ssql, cnDB, adOpenStatic, adLockReadOnly
    If rsPV.EOF = False And rsPV.BOF = False Then
        strNumPedido = Trim(rsPV!ORDNUMBER)
        strIDCliente = Trim(rsPV!CUSTOMER)
        strCliente = Trim(rsPV!NAMECUST)
        strIDEjecutivo = Trim(rsPV!IDEjecutivo)
        strPONUMBER = Trim(rsPV!PONUMBER)
        
        fnStrDatosAnexos = strNumPedido & "~" & strIDCliente & "~" & strCliente & "~" & strIDEjecutivo & "~" & strPONUMBER
    Else
        fnStrDatosAnexos = "~~~~~"
    End If
    
    rsPV.Close
    Set rsPV = Nothing

End Function


Private Function fnDatosAgreados(sNoOC) As String
Dim ssql99 As String
Dim rsBusca01 As Recordset
Set rsBusca01 = New Recordset

fnDatosAgreados = ""

ssql99 = "select (Select Value FROM POPORHO Where PORHSEQ=UNO.PORHSEQ and OPTFIELD='DEALID') AS DEALID,  (Select Value FROM POPORHO Where PORHSEQ=UNO.PORHSEQ and OPTFIELD='ENDUSER') AS ENDUSER,  (Select Value FROM POPORHO Where PORHSEQ=UNO.PORHSEQ and OPTFIELD='OCCLIENTE') AS OCCLIENTE,  (Select Value FROM POPORHO Where PORHSEQ=UNO.PORHSEQ and OPTFIELD='SALEORDER') AS SALEORDER, [DATE] AS FECHAOC "
ssql99 = ssql99 & "from POPORH1 as UNO  left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ  left outer join ICITEM as TRES on TRES.ITEMNO=DOS.ITEMNO "
ssql99 = ssql99 & "left outer join PORCPL as CUATRO on CUATRO.PORHSEQ=DOS.PORHSEQ AND CUATRO.PORLSEQ=DOS.PORLSEQ  Where (DOS.OQRECEIVED>0 and CUATRO.OQRECEIVED>0)   and  (UNO.PONUMBER = '" & sNoOC & "')"

    Debug.Print ssql99
    rsBusca01.Open ssql99, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBusca01.EOF = False And rsBusca01.BOF = False Then
        fnDatosAgreados = Trim(rsBusca01!DealID) & "~" & Trim(rsBusca01!ENDUSER) & "~" & Trim(rsBusca01!OCCliente) & "~" & Trim(rsBusca01!SALEORDER) & "~" & Trim(rsBusca01!FechaOC)
    End If
 
End Function

Private Function fnBuscaIDEndUserXPV(strNoPV As String, strArticulo As String) As String

    Dim ssql As String
    Dim rsPVIDEndUser As Recordset
    Set rsPVIDEndUser = New Recordset
    
    Dim strEndUser, TipoArticulo As String
    
            TipoArticulo = Left(Trim(strArticulo), 2)
            Select Case TipoArticulo
            Case "FO"
                strEndUser = "EndUserFO"
            Case "RU"
                strEndUser = "EndUserRU"
            Case "XT"
                strEndUser = "EndUserXT"
            Case "DE"
                strEndUser = "EndUserDE"
            Case Else
                strEndUser = ""
            End Select
            
If strEndUser <> "" Then
    ssql = "SELECT B.VALUE as IDEndUser FROM OEORDH A left join "
    ssql = ssql & "(SELECT * FROM OEORDHO WHERE OPTFIELD = '" & strEndUser & "' ) AS B "
    ssql = ssql & " ON A.ORDUNIQ = B.ORDUNIQ  WHERE ORDNUMBER = '" & Trim(strNoPV) & "'"

    Debug.Print ssql
    
    rsPVIDEndUser.Open ssql, cnDB, adOpenStatic, adLockReadOnly
    If rsPVIDEndUser.EOF = False And rsPVIDEndUser.BOF = False Then
        If IsNull(rsPVIDEndUser!IDEndUser) Or Trim(rsPVIDEndUser!IDEndUser) = "" Then
            fnBuscaIDEndUserXPV = ""
        Else
            fnBuscaIDEndUserXPV = Trim(rsPVIDEndUser!IDEndUser)
        End If
        
    Else
        fnBuscaIDEndUserXPV = ""
    End If
    
    rsPVIDEndUser.Close
    Set rsPVIDEndUser = Nothing
Else
        fnBuscaIDEndUserXPV = ""
End If

End Function


Private Function strEndUserNameXPV(strIDEndUser As String, strArticulo As String) As String
Dim sSQLEndUserN As String
Dim rsEndUserN As Recordset
Set rsEndUserN = New Recordset

Dim strEndUser, TipoArticulo As String
    
            TipoArticulo = Left(Trim(strArticulo), 2)
            Select Case TipoArticulo
            Case "FO"
                strEndUser = "endcustomerFO"
            Case "RU"
                strEndUser = "endcustomerRU"
            Case "XT"
                strEndUser = "endcustomerXT"
            Case "DE"
                strEndUser = "endcustomerDE"
            Case Else
                strEndUser = ""
            End Select
            
If strEndUser <> "" Then
        strEndUserNameXPV = ""

        sSQLEndUserN = "select codigo, nombre from " & strEndUser & " where codigo = '" & strIDEndUser & "'"
        
            Debug.Print sSQLEndUserN
            rsEndUserN.Open sSQLEndUserN, cnDB, adOpenForwardOnly, adLockReadOnly
            If rsEndUserN.EOF = False And rsEndUserN.BOF = False Then
                strEndUserNameXPV = Trim(rsEndUserN!nombre)
            End If
Else
    strEndUserNameXPV = ""
End If

End Function


Private Function strEndUserName(strIDEndUser As String) As String
Dim sSQLEndUserN As String
Dim rsEndUserN As Recordset
Set rsEndUserN = New Recordset

strEndUserName = ""
sSQLEndUserN = "select codigo, nombre from endcustomer where codigo = '" & strIDEndUser & "'"

    Debug.Print sSQLEndUserN
    rsEndUserN.Open sSQLEndUserN, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsEndUserN.EOF = False And rsEndUserN.BOF = False Then
        strEndUserName = Trim(rsEndUserN!nombre)
    End If

End Function

Private Function FnStrEjecutivo(strIDEjecutivo As String) As String
Dim sSQLEjecutivo As String
Dim rsEjecutivo As Recordset
Set rsEjecutivo = New Recordset

FnStrEjecutivo = ""

sSQLEjecutivo = "SELECT NAMEEMPL As vDesc FROM ARSAP WHERE CODESLSP = '" & strIDEjecutivo & "'"

    Debug.Print sSQLEjecutivo
    rsEjecutivo.Open sSQLEjecutivo, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsEjecutivo.EOF = False And rsEjecutivo.BOF = False Then
        FnStrEjecutivo = Trim(rsEjecutivo!vDesc)
    End If

End Function

'-------------------------------------------------------------------
'-------------------------------------------------------------------

Private Function EncuentraComentarios(PORHSEQ As String) As String
Dim ssql As String
Dim rsComment As Recordset
Set rsComment = New Recordset

ssql = "select rtrim(UNO.COMMENT)+' '+rtrim(UNO.REFERENCE)+' '+rtrim(UNO.DESCRIPTIO) as COMMENT, ISNULL(TRES.COMMENT,' ')AS COMMENTD from POPORH1 as UNO left outer join "
ssql = ssql & "  POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ LEFT OUTER JOIN POPORC AS TRES ON DOS.PORHSEQ=TRES.PORHSEQ AND TRES.PORCSEQ=DOS.PORCSEQ where UNO.PORHSEQ='" & PORHSEQ & "'"
rsComment.Open ssql, cnDB, adOpenStatic, adLockReadOnly
If rsComment.EOF = False And rsComment.BOF = False Then
    EncuentraComentarios = Replace(Trim(rsComment!Comment), "'", "''")
    rsComment.MoveFirst
    Do While rsComment.EOF = False
        EncuentraComentarios = EncuentraComentarios & "  " & Replace(Trim(rsComment!Commentd), "'", "''")
        rsComment.MoveNext
    Loop
    'EncuentraComentarios = Left(EncuentraComentarios, 245)
Else
    EncuentraComentarios = ""
End If
rsComment.Close
Set rsComment = Nothing
End Function

Private Function ArticuloCon(itemno As String) As Integer
Dim ssql As String
Dim rsCon As Recordset
Set rsCon = New Recordset

ssql = "select SERIALNO,LOTITEM from ICITEM where ITEMNO='" & Trim(itemno) & "'"
Debug.Print ssql

rsCon.Open ssql, cnDB, adOpenStatic, adLockReadOnly
If rsCon.EOF = False And rsCon.BOF = False Then
    If rsCon!serialno = 1 Then
        ArticuloCon = 1 'SERIE
    ElseIf rsCon!LOTITEM = 1 Then
        ArticuloCon = 2 'LOTE
    Else
        ArticuloCon = 3 'NADA
    End If
Else
    ArticuloCon = 3 'NADA***
End If
rsCon.Close
Set rsCon = Nothing
End Function

Private Function ParametrosHoy() As Boolean
ParametrosHoy = True
If Check1.Value = 1 Then
    ParametrosHoy = False
End If
If Check2.Value = 1 Then
    ParametrosHoy = False
End If
If Check3.Value = 1 Then
    ParametrosHoy = False
End If
If Check4.Value = 1 Then
    ParametrosHoy = False
End If
If cmbDesdeSucursal.Text <> "TODAS" Then
    ParametrosHoy = False
End If
If cmbDesdeSucursal.Text <> "TODAS" Then
    ParametrosHoy = False
End If
End Function

Private Sub cmdBuscar_Click()
If optHoy.Value Then ' información del día de hoy
    If ParametrosHoy Then
        If MsgBox("No se recomienda ejecutar el reporte completo en horarios de alta demanda de los servidores." & vbCrLf & "¿Desea generar el reporte de cualquier forma?", vbExclamation + vbYesNo) = vbYes Then
            Call Buscar_informacion_del_dia
        End If
    Else
        Call Buscar_informacion_del_dia
    End If
Else ' información del repositorio de ayer
    Call Buscar_informacion_de_ayer
End If
End Sub

Private Sub cmdExporta_Click()
Screen.MousePointer = vbHourglass
frmExisteOC.Enabled = False

   Dim appExcel As Excel.Application
   Dim wrkExcel As Excel.Workbook
   Dim shtExcel As Excel.Worksheet
   
   Dim intNewSheets As Variant
   Dim bm As Variant
   
   Dim GridData As String
   Dim A As Double
   Dim J As Double
   Dim k As Double
   Dim intCol As Long
   Dim intRow As Long
   Dim strNombre As String
   Dim i As Double
   strNombre = "HistoricoVentas" & Date & Time
   
   Dim Leyenda As String
   
   On Error Resume Next
   Set appExcel = GetObject(, "Excel.Application")
   If appExcel Is Nothing Then
       Set appExcel = CreateObject("Excel.Application")
       If appExcel Is Nothing Then
           MsgBox "Cannot Open Microsoft Excel For Export", vbCritical
           Exit Sub
       End If
   End If
   SSOleDBGrid1.MoveFirst
   intCol = SSOleDBGrid1.Cols
   intRow = SSOleDBGrid1.Rows
   intNewSheets = appExcel.SheetsInNewWorkbook
   appExcel.SheetsInNewWorkbook = 1
   Set wrkExcel = appExcel.Workbooks.Add
   appExcel.SheetsInNewWorkbook = intNewSheets
   Set shtExcel = wrkExcel.Sheets(1)
   shtExcel.Cells(1, 1) = "Reporte de Existencias de OC"
   If Check1.Value = 1 Then
      shtExcel.Cells(2, 1) = "No OC: " & txtNoOC
   End If
   If Check2.Value = 1 Then
      shtExcel.Cells(2, 2) = "No Artículo: " & cmbArticulo
   End If
   If Check3.Value = 1 Then
      shtExcel.Cells(2, 3) = "Marca: " & cmbMarca
   End If
   If Check4.Value = 1 Then
      shtExcel.Cells(2, 4) = "Fecha OC del " & DTPicker1.Value & " al " & DTPicker2.Value
   End If
   
   For i = 0 To SSOleDBGrid1.Cols - 1
      shtExcel.Cells(3, i + 1) = SSOleDBGrid1.Columns(i).Caption
   Next
   
   With shtExcel
      For i = 0 To intRow - 1
           bm = SSOleDBGrid1.GetBookmark(i)
           For A = 0 To intCol
              GridData = SSOleDBGrid1.Columns(A).CellText(bm)
              J = i + 4
              k = A + 1
              
              .Cells(J, k).Value = CStr(Trim(GridData))
           Next A
      Next i
   End With
   shtExcel.Name = "Transacciones"
   With appExcel.Selection.Interior
       .ColorIndex = 15
       .Pattern = xlSolid
   End With
   shtExcel.Select
   shtExcel.Name = "DATOS"
   shtExcel.Range("A1").Select
   With shtExcel.PageSetup
       .PrintTitleRows = shtExcel.Rows(1).Address
       .PrintTitleColumns = ""
   End With
   shtExcel.PageSetup.PrintArea = ""
   With shtExcel.PageSetup
       .LeftHeader = ""
       .CenterHeader = strHeader
       .RightHeader = ""
       .LeftFooter = ""
       .CenterFooter = ""
       .RightFooter = ""
       .LeftMargin = Application.InchesToPoints(0.75)
       .RightMargin = Application.InchesToPoints(0.75)
       .TopMargin = Application.InchesToPoints(1)
       .BottomMargin = Application.InchesToPoints(1)
       .HeaderMargin = Application.InchesToPoints(0.5)
       .FooterMargin = Application.InchesToPoints(0.5)
       .PrintHeadings = False
       .PrintGridlines = True
       .PrintComments = xlPrintNoComments
       .PrintQuality = 600
       .CenterHorizontally = False
       .CenterVertically = False
       .Orientation = xlPortrait
       .Draft = False
       .PaperSize = xlPaperA4
       .FirstPageNumber = xlAutomatic
       .Order = xlDownThenOver
       .BlackAndWhite = False
       .Zoom = False
       .FitToPagesWide = 1
       .FitToPagesTall = False
   End With
   appExcel.DisplayFullScreen = True
   appExcel.DisplayFullScreen = False
   appExcel.Visible = True
   'appExcel.ActiveWorkbook.SaveAs ("D:\ATC\atc\TRANSACCIONES\" & txt_lote.Text)
   appExcel.ActiveWorkbook.SaveAs ("C:\" & strNombre & ".xls")
   ActiveWindow.Close
   appExcel.ActiveWorkbook.Close
   Set shtExcel = Nothing
   Set wrkExcel = Nothing
   Set appExcel = Nothing
   
Screen.MousePointer = vbNormal
frmExisteOC.Enabled = True
End Sub

Private Sub cmdRepositorio_Click()
On Error GoTo RutinaError

Screen.MousePointer = vbHourglass
frmExisteOC.Enabled = False

Dim PasaWhere As Boolean

Dim InfoDisponible As Boolean

Dim ArticulosSinSerieNiLote As String
Dim ssql As String
Dim ssql2 As String
Dim ssql3 As String
Dim ssql4 As String
Dim strCadena As String
Dim Serie As String

Dim rsBusca As Recordset
Set rsBusca = New Recordset

Dim rsSerLot As Recordset
Dim rsSerLotVal As Recordset

Set rsSerLot = New Recordset
Set rsSerLotVal = New Recordset

Dim rsRecibo As Recordset
Set rsRecibo = New Recordset
Dim rsTrans As Recordset
Set rsTrans = New Recordset
Dim rsDemas As Recordset
Set rsDemas = New Recordset

Dim strPORHSEQ As String
Dim ArticuloUsa As Integer
Dim strSerLot As String
Dim Codigo2 As String
Dim FechaRM As String
Dim comentariosRM As String
Dim comenRM As String
Dim SalesO As String
Dim NoLinea As String
Dim CostoOc As String
Dim FechaTRF As String
Dim NoTFR As String
Dim Costo As String
Dim PricoLista As String
Dim Comentario As String
Dim ComentarioTRF As String
Dim NoFacRM As String
Dim NoLineaOC As Long
Dim Cantidad As Double

Dim SumatoriaTotal As Double
Dim strProveedor As String

'---------------------------------
Dim strNoPV As String

Dim sDatosAgregados As String
Dim lDatosAgregados() As String

Dim sDEALID As String
Dim sENDUSER As String
Dim sOCCLIENTE As String
Dim sSALEORDER As String

Dim strEjecutivo As String
Dim strEndUser As String
Dim sIDENDUSER As String
Dim strIDEjecutivo As String
Dim strVendorID As String

''20240206
Dim sFACTCOMER As String
Dim sFACTPACK As String

strNoPV = ""
'---------------------------------

InfoDisponible = False
Generagrid
cmdExporta.Enabled = False
PasaWhere = True

sUserRepositorio = "REPOSITORI"

'---------------------------------------------------
'---------------------------------------------------
cnDB.Execute "DELETE FROM Existencias_OC where Usuario='" & sUserRepositorio & "'"
cnDB.Execute "DELETE FROM [Existencias_OC_Repositorio] where Usuario='" & sUserRepositorio & "'"
'---------------------------------------------------
'---------------------------------------------------
'MsgBox "Exixtencias_OC_DePrueba LIMPIA"

SumatoriaTotal = 0

If Check1.Value = 1 Or Check4.Value = 1 Then 'SE FILTRA POR OC O FECHA DE OC
            
            'CASO DE OC OC00036831
            'ssql2 = "select uno.VDCODE,UNO.PORHSEQ, DOS.ITEMNO, TRES.CATEGORY,DOS.LOCATION,"
            'ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=TRES.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
            'ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA,(select QTYONHAND from ICILOC where ICILOC.ITEMNO=DOS.ITEMNO and ICILOC.LOCATION=DOS.LOCATION) as QTYONHAND"
            'ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA,A.QTYONHAND"
            'ssql2 = ssql2 & "  from POPORH1 as UNO  left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ"
            'ssql2 = ssql2 & "  left outer join ICITEM as TRES on TRES.ITEMNO=DOS.ITEMNO"
            'ssql2 = ssql2 & "  left outer join ICILOC as A ON A.ITEMNO=DOS.ITEMNO and A.LOCATION=DOS.LOCATION"
            'ssql2 = ssql2 & "  Where DOS.OQRECEIVED>0 AND A.QTYONHAND>0"
            
            ssql2 = "select uno.VDCODE,UNO.PORHSEQ, DOS.ITEMNO, TRES.CATEGORY,/*C.LOCATION,*/CUATRO.RCPHSEQ,"
            

            ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=TRES.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
            ssql2 = ssql2 & "  DOS.PORLSEQ,DOS.PORLREV/1000 LINEA/*,A.QTYONHAND  */,"
            'CALSULAR EL COSTO DE LAS LÍNEAS DE LA OC CUANDO NO TENGAN NI SERIE NI LOTE
            ssql2 = ssql2 & "  case when CUATRO.DISCPCT=100 then 0 "
            ssql2 = ssql2 & "  Else ISNULL(CUATRO.EXTENDED-CUATRO.DISCOUNT,0)/CUATRO.OQRECEIVED  end as VALCOSTOOC  "
            'CALSULAR EL COSTO DE LAS LÍNEAS DE LA OC CUANDO NO TENGAN NI SERIE NI LOTE
            '------------------------------------
            ssql2 = ssql2 & ",  UNO.DATE AS FECHAOC"
            '------------------------------------
            ssql2 = ssql2 & "  from POPORH1 as UNO"
            ssql2 = ssql2 & "  left outer join POPORL as DOS on UNO.PORHSEQ=DOS.PORHSEQ"
            ssql2 = ssql2 & "  /*left outer join PORCPL as C on UNO.PORHSEQ=C.PORHSEQ AND C.PORLSEQ=DOS.PORLSEQ  */"
            ssql2 = ssql2 & "  left outer join ICITEM as TRES on TRES.ITEMNO=DOS.ITEMNO"
            ssql2 = ssql2 & "  /*left outer join ICILOC as A ON A.ITEMNO=DOS.ITEMNO and A.LOCATION=C.LOCATION  */"
            'SE AGREGA LA TABLA DE RECIBO PARA HACER EL CALCULO DEL COSTO DE LOS ITEMS DESDE LA OC
            ssql2 = ssql2 & "  left outer join PORCPL as CUATRO on CUATRO.PORHSEQ=DOS.PORHSEQ AND CUATRO.PORLSEQ=DOS.PORLSEQ"
            'SE AGREGA LA TABLA DE RECIBO PARA HACER EL CALCULO DEL COSTO DE LOS ITEMS DESDE LA OC
            ssql2 = ssql2 & "  Where (DOS.OQRECEIVED>0 and CUATRO.OQRECEIVED>0) /*AND A.QTYONHAND>0  */"
            
            If cmbDesdeSucursal.Text <> "TODAS" Or cmbHastaSucursal.Text <> "TODAS" Then 'SUCURSAL
                If Not PasaWhere Then
                    'ssql2 = ssql2 & "  where   DOS.LOCATION='" & cmbSucursal.Text & "'"
                    ssql2 = ssql2 & "  where  (DOS.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (DOS.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
                End If
            End If
            If Check2.Value = 1 Then 'ARTÍCULO
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (DOS.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and  '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "'("
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (DOS.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "')"
                End If
            End If
            If Check3.Value = 1 Then 'CATEGORÍA
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (TRES.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (TRES.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
                End If
            End If
            If Check1.Value = 1 Then 'NO OC
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (UNO.PONUMBER between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (UNO.PONUMBER  between '" & IIf(Trim(txtDesdeNoOC.Text) <> "", Trim(txtDesdeNoOC.Text), "") & "' and  '" & IIf(Trim(txtHastaNoOC.Text) <> "", Trim(txtHastaNoOC.Text), Trim(txtDesdeNoOC.Text)) & "')"
                    '" & txtNoOC.Text & "'"
                End If
            End If
            If Check4.Value = 1 Then 'RANGO FECHA
                If Not PasaWhere Then
                    ssql2 = ssql2 & "  where  (UNO.[DATE] between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "')"
                    PasaWhere = True
                Else
                    ssql2 = ssql2 & "  and  (UNO.[DATE] between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "')"
                End If
            End If
            Debug.Print ssql2
            
            rsBusca.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
            If rsBusca.EOF = False And rsBusca.BOF = False Then
               Do While rsBusca.EOF = False
                    ArticuloUsa = ArticuloCon(rsBusca!itemno)
                    
'''            If Trim(rsBusca!itemno) = "CPZP005" Then
''' '               MsgBox "Revisar este Caso"
'''            End If
                    
                    strProveedor = Trim(rsBusca!VDCODE)
                    
''                    If Trim(rsBusca!itemno) = "CM2WA000020" Then
''                        MsgBox "Aquí"
''                    End If
                    
                    Select Case ArticuloUsa
                        Case 1 'serie
                            ssql2 = "select uno.RCPNUMBER,UNO.[DATE] RMFECHA,uno.PORHSEQ, uno.PONUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, tres.SERIALNUMF,"
                            ssql2 = ssql2 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED,"
                            ssql2 = ssql2 & "  (Select Value FROM ICITEMO Where ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2 ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER "
                            '---------------------------
                            ssql2 = ssql2 & ",  DOS.PORLSEQ"
                            '---------------------------
                            ssql2 = ssql2 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ "
                            ssql2 = ssql2 & "  left outer join PORCPLS as TRES on dos.RCPHSEQ=tres.RCPHSEQ and dos.RCPLREV=tres.RCPLREV and dos.RCPLSEQ=tres.RCPLSEQ "
                            ssql2 = ssql2 & "  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            ssql2 = ssql2 & "  where UNO.RCPHSEQ='" & Trim(rsBusca!RCPHSEQ) & "' AND UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND dos.PORLSEQ='" & Trim(rsBusca!PORLSEQ) & "' AND dos.OQRECEIVED>0"
                            Debug.Print ssql2
                            rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                rsRecibo.MoveFirst
                                Do While rsRecibo.EOF = False
                                    FacturaRM = Trim(rsRecibo!INVNUMBER)
                                    comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                    comentariosRM = Replace(comentariosRM, Chr(13), "")
                                    comentariosRM = Replace(comentariosRM, "'", "''")
                                    
                                    ''20240212 -- quita tabulador de la cadena de texto
                                    comentariosRM = Replace(comentariosRM, vbTab, " ")
                                    
                                    
                                    '-----------------------
                                    strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                    '-----------------------
                                    
                                    'Serie = Replace(RTrim(rsRecibo!SERIALNUMF), "'", "''")
                                    'Correccion Porque Mandaba Error
                                    '-----------------------
                                    If IsNull(rsRecibo!SERIALNUMF) Then
                                        Serie = ""
                                    Else
                                        Serie = Replace(RTrim(rsRecibo!SERIALNUMF), "'", "")
                                    End If
                                    '-----------------------

                                    'ssql3 = "select * from ICXSER where SERIALNUMF='" & Serie & "' and [STATUS]=1 AND QTYONHAND>0"
                                    ssql3 = "select * from ICXSER as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.SERIALNUMF='" & Serie & "' AND ITEMNO='" & rsRecibo!itemno & "' and A.[STATUS]=1 and ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0)"  'B.QTYONHAND>0"
                                    Debug.Print ssql3
                                    rsSerLot.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                                        Serie = Replace(RTrim(rsSerLot!SERIALNUMF), "'", "''")
                                        ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                                        ssql4 = ssql4 & "  left outer join ICTREDS as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                                        ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsSerLot!Location) & "' AND dos.ITEMNO='" & Trim(rsRecibo!itemno) & "' and tres.SERIALNUMF='" & Serie & "' AND UNO.DOCTYPE=3"
                                        Debug.Print ssql4
                                        rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                                        If rsTrans.EOF = False And rsTrans.BOF = False Then
                                        
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                        & "'1','" & RTrim(Serie) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                        & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                        & "','" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsRecibo!itemno)) & "'" _
                                                        & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        ' & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                        Else
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                        & "'1','" & RTrim(Serie) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                        & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                        & "','01-01-1900',' ','" & ValorCosto(Trim(rsRecibo!itemno)) _
                                                        & "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & Trim(rsBusca!FechaOC) & "')" '  almacena cadena
                                                        '& "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                    
                                        End If
                                        Debug.Print strCadena
                                        cnDB.Execute strCadena
                                        rsTrans.Close
                                    End If
                                    rsSerLot.Close
                                    rsRecibo.MoveNext
                                Loop
                            Else
                            '---------------------------------------------------
                            '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'0',' ',' ',' ',' '" _
                                            & ",'" & ValorSALEORDER(Trim(rsBusca!PORHSEQ)) & "'," & "'0','" & 0 _
                                            & "',' ',' ','" & ValorCosto(Trim(rsBusca!itemno)) _
                                            & "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ',' ','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & Trim(rsBusca!FechaOC) & "')" '  almacena cadena **&&
                                            '& "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ',' ','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "')" 'almacena cadena
                            '---------------------------------------------------
                            '---------------------------------------------------
                                cnDB.Execute strCadena
                            End If
                            rsRecibo.Close
                        Case 2 'lote
                        
                            'ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            'ssql2 = ssql2 & " UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, tres.LOTNUMF,(Select Value FROM ICITEMO Where "
                            'ssql2 = ssql2 & " " ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
                            
                            '''*** 20230710
                            '''*** ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            ssql2 = "select DOS.LOCATION, uno.RCPNUMBER,uno.PONUMBER,uno.[DATE] as RMFECHA,uno.PORHSEQ,UNO.DESCRIPTIO, "
                            
                            ssql2 = ssql2 & " UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, REPLACE(tres.LOTNUMF,CHAR(39),'') AS LOTNUMF, "
                            ssql2 = ssql2 & " (Select Value FROM ICITEMO Where ITEMNO=DOS.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,"
                            
                            ssql2 = ssql2 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED"
                            ssql2 = ssql2 & "  ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER  "
                            
                            '---------------------------
                            ssql2 = ssql2 & ",  DOS.PORLSEQ"
                            '---------------------------

                            ssql2 = ssql2 & " from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ"
                            ssql2 = ssql2 & "  left outer join PORCPLL as TRES on dos.RCPHSEQ=tres.RCPHSEQ and dos.RCPLREV=tres.RCPLREV and dos.RCPLSEQ=tres.RCPLSEQ  "
                            ssql2 = ssql2 & "  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            
                            ssql2 = ssql2 & "  where UNO.RCPHSEQ='" & Trim(rsBusca!RCPHSEQ) & "' AND UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND dos.PORLSEQ='" & Trim(rsBusca!PORLSEQ) & "' and DOS.OQRECEIVED>0"
                            Debug.Print ssql2
                            rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                rsRecibo.MoveFirst
                                Do While rsRecibo.EOF = False
                                    FacturaRM = Trim(rsRecibo!INVNUMBER)
                                    comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                    comentariosRM = Replace(comentariosRM, Chr(13), "")
                                    comentariosRM = Replace(comentariosRM, "'", "''")
                                    
                                    ''20240212 -- quita tabulador de la cadena de texto
                                    comentariosRM = Replace(comentariosRM, vbTab, " ")
                                    
                                    
                                    '-----------------------
                                    strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                    '-----------------------
                                    
                                    'ssql3 = "select * from ICXLOT where LOTNUMF='" & Trim(rsRecibo!LOTNUMF) & "' and QTYAVAIL>0 AND QTYONHAND>0"
                                    
                                    If Check4.Value = 1 Then '--- si BUSCA por RANGO DE FECHAS de OC
                                        '''*** 20230710
                                        '''*** ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                         ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         
                                         Debug.Print ssql3
                                         
                                            rsSerLotVal.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsSerLotVal.EOF = False And rsSerLotVal.BOF = False Then
                                                '''---
                                            Else
                                                ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) AND STOCKDATE between '" & FormatoFecha(DTPicker1.Value) & "' and '" & FormatoFecha(DTPicker2.Value) & "' and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                                Debug.Print ssql3
                                            End If
                                            rsSerLotVal.Close
                                         
                                    Else
                                        '''*** 20230710
                                        'ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                        
                                        ''*** 20240131 -- Correccion, debido a que cuando hay transferencias cambia el Location y no muestra el renglon
                                         ''ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         
                                         ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "' and A.LOCATION='" & Trim(rsRecibo!Location) & "'" ' and B.QTYONHAND>0"
                                         Debug.Print ssql3
                                         
                                            rsSerLotVal.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsSerLotVal.EOF = False And rsSerLotVal.BOF = False Then
                                                '''---
                                            Else
                                                ssql3 = "select * from ICXLOT as A left outer join ICILOC as B on A.ITEMNUM = B.ItemNO And A.Location = B.Location where A.LOTNUMF='" & RTrim(rsRecibo!LOTNUMF) & "' and A.QTYAVAIL>0 AND ((B.QTYONHAND+B.QTYRENOCST+B.QTYADNOCST-B.QTYSHNOCST)>0) and A.ITEMNUM='" & Trim(rsRecibo!itemno) & "'" ' and B.QTYONHAND>0"
                                                Debug.Print ssql3
                                            End If
                                            rsSerLotVal.Close
                                    
                                    End If
                                    Debug.Print ssql3
                                    rsSerLot.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                                        rsSerLot.MoveFirst
                                        Do While rsSerLot.EOF = False
                                                                                
                                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                                            ssql4 = ssql4 & "  left outer join ICTREDL as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsSerLot!Location) & "' AND dos.ITEMNO='" & Trim(rsRecibo!itemno) & "' and tres.LOTNUMF='" & RTrim(rsSerLot!LOTNUMF) & "' AND UNO.DOCTYPE=3"
                                            Debug.Print ssql4
                                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                            & Trim(rsSerLot!QTYAVAIL) & ",'" & RTrim(rsSerLot!LOTNUMF) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                            & "','" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsRecibo!itemno)) & "'" _
                                                            & ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                            '& ",'" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                            
                                            Else
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsSerLot!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsRecibo!Codigo2) & "'," _
                                                            & Trim(rsSerLot!QTYAVAIL) & ",'" & RTrim(rsSerLot!LOTNUMF) & "','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                                            & "','01-01-1900',' ','" & ValorCosto(Trim(rsRecibo!itemno)) _
                                                            & "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                            '& "','" & ValorPL(Trim(rsRecibo!itemno)) & "','" & EncuentraComentarios(Trim(rsRecibo!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                                                '---------------------------------------------------
                                                '---------------------------------------------------
                                                
                                            End If
                                            rsTrans.Close
                                            cnDB.Execute strCadena
                                            rsSerLot.MoveNext
                                        Loop
                                    Else
                                        'strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!ITEMNO) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                        '            & "'0',' ',' ','" & Trim(rsRecibo!PONUMBER) & "',' '" _
                                        '            & ",'" & ValorSALEORDER(Trim(rsBusca!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) _
                                        '            & "',' ',' ','" & ValorCosto(Trim(rsBusca!ITEMNO)) _
                                        '            & "','" & ValorPL(Trim(rsBusca!ITEMNO)) & "','" & EncuentraComentarios(Trim(rsBusca!PORHSEQ)) & "',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "')" 'almacena cadena
                                        'cnDB.Execute strCadena
                                    End If
                                    rsSerLot.Close
                                    rsRecibo.MoveNext
                                Loop
                            End If
                            rsRecibo.Close
                        Case 3
                            'If ArticulosSinSerieNiLote <> "" Then
                            'Cantidad = Existencia(Trim(rsBusca!itemno))
                            Dim rsOpcional As Recordset
                            Set rsOpcional = New Recordset
                            ssql = "select SUM(QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST) EXISTENCIA,LOCATION from ICILOC where ITEMNO='" & Trim(rsBusca!itemno) & "' GROUP BY ITEMNO,LOCATION"
                            Debug.Print ssql
                            rsOpcional.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsOpcional.EOF = False And rsOpcional.BOF = False Then
                                rsOpcional.MoveFirst
                                Do While rsOpcional.EOF = False
                                    If rsOpcional!Existencia > 0 Then
                                    
                                        ssql2 = "select uno.RCPNUMBER,uno.PONUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,dos.ITEMNO,dos.UNITWEIGHT, uno.[DATE] as RMFECHA,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ "
                                        ssql2 = ssql2 & "   left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ  where UNO.PORHSEQ='" & Trim(rsBusca!PORHSEQ) & "' AND ITEMNO='" & Trim(rsBusca!itemno) & "'"
                                        
                                        Debug.Print ssql2
                                        rsRecibo.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                                        If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                            FacturaRM = Trim(rsRecibo!INVNUMBER)
                                            comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                            comentariosRM = Replace(comentariosRM, Chr(13), "")
                                            comentariosRM = Replace(comentariosRM, "'", "''")
                                            
                                            ''20240212 -- quita tabulador de la cadena de texto
                                            comentariosRM = Replace(comentariosRM, vbTab, " ")
                                            
                                            
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsOpcional!Location) & "','" & Trim(rsRecibo!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                                        & "'" & rsOpcional!Existencia & "',' ','" & Trim(rsRecibo!RMFECHA) & "','" & Trim(rsRecibo!PONUMBER) & "','" & Trim(rsRecibo!RCPNUMBER) & "','" _
                                                        & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "','" & Left(Trim(strNoPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        '& ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ','" & FacturaRM & "','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "')" 'sin datos
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                        
                                            Debug.Print strCadena
                                        Else
                                            FacturaRM = " "
                                            comentariosRM = " "
                                            
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                            strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "',' ','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                                        & "'" & rsOpcional!Existencia & "',' ','01-01-1900',' ',' ',' ',' ','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "','" & Left(Trim(rsBusca!NUMPV), 10) & "','" & rsBusca!FechaOC & "')" '  almacena cadena
                                                        '& "'" & rsOpcional!Existencia & "',' ','01-01-1900',' ',' ',' ',' ','" & Trim(rsBusca!VALCOSTOOC) & "','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','" & Trim(rsBusca!LINEA) & "','" & Trim(rsBusca!VDCODE) & "','" & comentariosRM & "')" 'sin datos
                                            '---------------------------------------------------
                                            '---------------------------------------------------
                                                        
                                            Debug.Print strCadena
                                        End If
                                        cnDB.Execute strCadena
                                        rsRecibo.Close
                                    End If
                                    rsOpcional.MoveNext
                                Loop
                                rsOpcional.Close
                            End If
                    End Select
                    rsBusca.MoveNext
                Loop
                InfoDisponible = True
            Else
                'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
                InfoDisponible = False
            End If
            
            rsBusca.Close
    
Else 'NO SE FILTRA POR OC O FECHA DE OC

    ssql = "select uno.LOCATION,uno.ITEMNO,(Select Value FROM ICITEMO Where ITEMNO=UNO.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,(QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST) QTYONHAND  from ICILOC AS UNO left outer join ICITEM as DOS on uno.ITEMNO=dos.ITEMNO where (QTYONHAND+QTYRENOCST+QTYADNOCST-QTYSHNOCST)>0"
    'ssql = "select uno.LOCATION,uno.ITEMNO,(Select Value FROM ICITEMO Where ITEMNO=UNO.ITEMNO and OPTFIELD='CODIGO2') AS CODIGO2,UNO.QTYONHAND  from ICILOC AS UNO left outer join ICITEM as DOS on uno.ITEMNO=dos.ITEMNO where (UNO.QTYONHAND+UNO.QTYRENOCST)>0"
    If cmbDesdeSucursal.Text <> "TODAS" Or cmbHastaSucursal.Text <> "TODAS" Then 'SUCURSAL
        If Not PasaWhere Then
            ssql = ssql & "  where  (UNO.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
            PasaWhere = True
        Else
            ssql = ssql & "  and  (UNO.LOCATION between '" & IIf(cmbDesdeSucursal.Text <> "TODAS", cmbDesdeSucursal.Text, "") & "' and '" & IIf(cmbHastaSucursal.Text <> "TODAS", cmbHastaSucursal.Text, "ZZZZZZZZZZ") & "')"
        End If
    End If
    If Check2.Value = 1 Then 'ARTÍCULO
        If Not PasaWhere Then
            ssql = ssql & "  where  (UNO.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and  '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "'("
            PasaWhere = True
        Else
            ssql = ssql & "  and  (UNO.ITEMNO between'" & IIf(txtDesdeArticulo.Text <> "TODAS", txtDesdeArticulo.Text, "") & "' and '" & IIf(txtHastaArticulo.Text <> "TODAS", txtHastaArticulo.Text, "") & "')"
        End If
    End If
    If Check3.Value = 1 Then 'CATEGORÍA
        If Not PasaWhere Then
            ssql = ssql & "  where  (DOS.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
            PasaWhere = True
        Else
            ssql = ssql & "  and  (DOS.CATEGORY  between'" & IIf(cmbDesdeMarca.Text <> "TODAS", cmbDesdeMarca.Text, "") & "' and  '" & IIf(cmbHastaMarca.Text <> "TODAS", cmbHastaMarca.Text, "ZZZZZZ") & "')"
        End If
    End If
    Debug.Print ssql
    rsBusca.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBusca.EOF = False And rsBusca.BOF = False Then
        rsBusca.MoveFirst
        Do While rsBusca.EOF = False
            ArticuloUsa = ArticuloCon(rsBusca!itemno)
            
''''            If Trim(rsBusca!itemno) = "ATCM010" Then
''''                                         'MK4LI000018
''''                                         'CB2YI000001
''''                                         'XT4SE000022
''''             MsgBox "Revisar este Caso"
''''            End If

'''If fnGetFirst() <> "CB2YI000001" Then
'''    'MsgBox "Revisar este Caso"
'''End If
                       
            Select Case ArticuloUsa
                Case 1 'series
                    ssql2 = "select SERIALNUMF,1 as CANTIDAD from ICXSER where ITEMNUM='" & Trim(rsBusca!itemno) & "' and [STATUS]=1 and LOCATION='" & Trim(rsBusca!Location) & "'" ' and LIFECONT1=0 REVISIÓN DE MARICELA TERRAZAS, NO PASA FILTRO DE LIFECONT
                    Debug.Print ssql2
                    rsSerLot.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                        Do While rsSerLot.EOF = False
                        
                            Serie = Replace(RTrim(rsSerLot!SERIALNUMF), "'", "''")
                            
                            ssql3 = "select uno.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,"
                            ssql3 = ssql3 & "  isnull(CUATRO.INVNUMBER,'') as INVNUMBER, DOS.UNITWEIGHT,"
                            ssql3 = ssql3 & "  case when DOS.DISCPCT=100 then  ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED else  ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED end as EXTENDED"
                            'ssql3 = ssql3 & "  (select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA"
                            '--------------------------------
                            ssql3 = ssql3 & ",  DOS.PORLSEQ"
                            '--------------------------------
                            ssql3 = ssql3 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ"
                            ssql3 = ssql3 & "  left outer join PORCPLS as TRES on DOS.RCPHSEQ=TRES.RCPHSEQ and DOS.RCPLSEQ=TRES.RCPLSEQ  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.SERIALNUMF='" & Serie & "'"
                            
                            ''20231002
                            TipoArticulo = Left(Trim(rsBusca!itemno), 2)
                            
                            ''-- 20240116
                            ''-- If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Then
                            If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Or TipoArticulo = "CM" Then
                                If Mid(Trim(rsBusca!itemno), 3, 1) = 5 Or Mid(Trim(rsBusca!itemno), 3, 1) = 6 Or Mid(Trim(rsBusca!itemno), 3, 1) = 7 Then
                                    ssql3 = ssql3 & "  ORDER BY UNO.DATE"
                                End If
                            End If
                            
                            Debug.Print ssql3
                            
                            rsRecibo.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                            'INSERT INTO Existencias_OC VALUES('sUserRepositorio','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                            'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                strProveedor = Trim(rsRecibo!VDCODE)
                                FacturaRM = Trim(rsRecibo!INVNUMBER)
                                comentariosRM = Trim(rsRecibo!DESCRIPTIO) & Trim(rsRecibo!REFERENCE)
                                comentariosRM = Replace(comentariosRM, Chr(13), "")
                                comentariosRM = Replace(comentariosRM, "'", "''")
                                
                                ''20240212 -- quita tabulador de la cadena de texto
                                comentariosRM = Replace(comentariosRM, vbTab, " ")
                                
                                
                                '-----------------------
                                strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                '-----------------------

                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsSerLot!Cantidad) & "','" & Trim(Serie) & "','" & Trim(rsRecibo!Date) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "','" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) & "'," 'datos RM
                                            strPORHSEQ = Trim(rsRecibo!PORHSEQ)
                                            NoLineaOC = EncuentraLineaOC(Trim(rsRecibo!PORHSEQ), Trim(rsBusca!itemno), Trim(rsRecibo!UNITWEIGHT))
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                            Else
                                strProveedor = " "
                                FacturaRM = " "
                                comentariosRM = " "
                                strNoPV = ""
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' " _
                                            & "',' '," & "' ',' '," 'sin datos de RM
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                                            strPORHSEQ = "0"
                                            NoLineaOC = 0
                            End If
                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                            ssql4 = ssql4 & "  left outer join ICTREDS as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsBusca!Location) & "' AND dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.SERIALNUMF='" & Serie & "' AND UNO.DOCTYPE=3"
                            Debug.Print ssql4
                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                strCadena = strCadena & "'" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')" 'sin datos de transferencia
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                            Else
                                strCadena = strCadena & "'01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) _
                                & "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')" 'almacena cadena
                                '& "','" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'sin datos de transferencia
                            End If
                            Debug.Print strCadena
                            cnDB.Execute strCadena
                            rsTrans.Close
                            rsRecibo.Close
                            
                            rsSerLot.MoveNext
                        Loop
                    Else
                    
                    End If
                    rsSerLot.Close
                Case 2 'lotes
                    '---------------------------------
                    strLote = ""
                    '---------------------------------
                    
                    Debug.Print strNoPV

                    ssql2 = "select LOTNUMF,QTYAVAIL from ICXLOT where ITEMNUM='" & Trim(rsBusca!itemno) & "' and QTYAVAIL>0 and LOCATION='" & Trim(rsBusca!Location) & "'"
                    Debug.Print ssql2
                    rsSerLot.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
                    If rsSerLot.EOF = False And rsSerLot.BOF = False Then
                        Do While rsSerLot.EOF = False
                            '---------------------------------
                            strLote = rsSerLot!LOTNUMF
                            '---------------------------------
                            
                            ' CON EL DESCUENTO DEL 100% EN LA OC EL COSTO DEL ARTICULO SERÁ 0 - MARICELA TERRAZAS 25 01 2016
                            'ssql3 = "select UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,case when DOS.DISCPCT=100 then ISNULL(DOS.EXTENDED,0)/DOS.OQRECEIVED "
                            'ssql3 = ssql3 & "  Else ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED  end as EXTENDED" ',(select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA
                            
                            '''*** 20230710
                            '''*** ssql3 = "select UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,case when DOS.DISCPCT=100 then 0 "
                            ssql3 = "select DOS.location, UNO.VDCODE,uno.[DATE],uno.PONUMBER,UNO.RCPNUMBER,UNO.DESCRIPTIO,UNO.REFERENCE,case when DOS.DISCPCT=100 then 0 "
                            
                            ssql3 = ssql3 & "  Else ISNULL(DOS.EXTENDED-DOS.DISCOUNT,0)/DOS.OQRECEIVED  end as EXTENDED" ',(select PORLREV/1000 from POPORL where PORHSEQ=DOS.PORHSEQ and ITEMNO=DOS.ITEMNO and UNITWEIGHT=DOS.UNITWEIGHT)LINEA
                            ssql3 = ssql3 & "  ,(select PORHSEQ from POPORH1 where PONUMBER=UNO.PONUMBER) as PORHSEQ,isnull(CUATRO.INVNUMBER,'') as INVNUMBER, DOS.UNITWEIGHT "
                            
                            '--------------------------------
                            ssql3 = ssql3 & ",  DOS.PORLSEQ"
                            '--------------------------------

                            ssql3 = ssql3 & "  from PORCPH1 as UNO left outer join PORCPL as DOS on UNO.RCPHSEQ=DOS.RCPHSEQ  "
                            ssql3 = ssql3 & "  left outer join PORCPLL as TRES on DOS.RCPHSEQ=TRES.RCPHSEQ and DOS.RCPLSEQ=TRES.RCPLSEQ  left outer join POINVH1 as CUATRO on UNO.RCPHSEQ=CUATRO.RCPHSEQ"
                            
                            '''*** 20230710
                            '''*** ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and TRES.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "'"
                            ssql3 = ssql3 & "  where dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and DOS.location='" & Trim(rsBusca!Location) & "' and TRES.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "'"
                            
                            Debug.Print ssql3
                            
                            rsRecibo.Open ssql3, cnDB, adOpenForwardOnly, adLockReadOnly
                            
                            'INSERT INTO Existencias_OC VALUES('sUserRepositorio','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                            'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                            
                            If rsRecibo.EOF = False And rsRecibo.BOF = False Then
                                strProveedor = Trim(rsRecibo!VDCODE)
                                FacturaRM = Trim(rsRecibo!INVNUMBER)
                                comentariosRM = Trim(rsRecibo!DESCRIPTIO) & " " & Trim(rsRecibo!REFERENCE)
                                comentariosRM = Replace(comentariosRM, Chr(13), "")
                                comentariosRM = Replace(comentariosRM, "'", "''")
                                
                                ''20240212 -- quita tabulador de la cadena de texto
                                comentariosRM = Replace(comentariosRM, vbTab, " ")
                                
                                
                                '-----------------------
                                strNoPV = fnStrGetNoPvC(Trim(rsRecibo!PONUMBER), Trim(rsRecibo!PORLSEQ))
                                '-----------------------
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsSerLot!QTYAVAIL) & "','" & Trim(strLote) & "','" & Trim(rsRecibo!Date) & "','" & Trim(rsRecibo!PONUMBER) & "','" _
                                            & Trim(rsRecibo!RCPNUMBER) & "','" & ValorSALEORDER(Trim(rsRecibo!PORHSEQ)) & "'," & "'" & Trim(rsRecibo!UNITWEIGHT) & "','" & Trim(rsRecibo!EXTENDED) & "'," 'datos RM
                                            strPORHSEQ = Trim(rsRecibo!PORHSEQ)
                                            NoLineaOC = EncuentraLineaOC(Trim(rsRecibo!PORHSEQ), Trim(rsBusca!itemno), Trim(rsRecibo!UNITWEIGHT))
                                '---------------------------------------------------
                                '---------------------------------------------------

                            Else
                                strProveedor = " "
                                FacturaRM = " "
                                comentariosRM = ""
                                
                                strNoPV = ""
                                
                                '---------------------------------------------------
                                '---------------------------------------------------
                                strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                            & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' " _
                                            & "',' '," & "' ',' '," 'sin datos de RM
                                '---------------------------------------------------
                                '---------------------------------------------------
                                            
                                            strPORHSEQ = "0"
                                            NoLineaOC = 0
                            End If
                            
                            ssql4 = "select uno.DOCNUM,uno.DATEBUS,DOS.COMMENTS from ICTREH as UNO left outer join ICTRED as DOS on uno.TRANFENSEQ=dos.TRANFENSEQ"
                            ssql4 = ssql4 & "  left outer join ICTREDL as TRES on uno.TRANFENSEQ=tres.TRANFENSEQ and dos.[LINENO]=tres.[LINENO]"
                            ssql4 = ssql4 & "  Where DOS.TOLOC='" & Trim(rsBusca!Location) & "' AND dos.ITEMNO='" & Trim(rsBusca!itemno) & "' and tres.LOTNUMF='" & Trim(rsSerLot!LOTNUMF) & "' AND UNO.DOCTYPE=3"
                            
                            Debug.Print ssql4
                            
                            rsTrans.Open ssql4, cnDB, adOpenForwardOnly, adLockReadOnly
                            If rsTrans.EOF = False And rsTrans.BOF = False Then
                                strCadena = strCadena & "'" & Trim(rsTrans!DATEBUS) & "','" & Trim(rsTrans!DOCNUM) & "','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')"
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "','" & Replace(Trim(rsTrans!Comments), "'", "''") & "','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'almacena cadena
                            Else
                                strCadena = strCadena & "'01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "'" _
                                & ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "','" & strNoPV & "','01-01-1900')"
                                '& ",'" & ValorPL(Trim(rsBusca!itemno)) & "','" & EncuentraComentarios(strPORHSEQ) & "',' ','" & FacturaRM & "','" & NoLineaOC & "','" & strProveedor & "','" & comentariosRM & "')" 'sin datos de transferencia
                            End If
                            Debug.Print strCadena
                            cnDB.Execute strCadena
                            rsTrans.Close
                            rsRecibo.Close
                        
                            rsSerLot.MoveNext
                        Loop
                    Else
                    
                    End If
                    rsSerLot.Close
                Case 3 'nada
                        'INSERT INTO Existencias_OC VALUES('sUserRepositorio','Sucursal','Artículo','Codigo2','Cantidad','Serie Lote','FechaRM','NoOC','NoRM','SalesOrd', _
                        'NoLinea','CostoOC','FechaTRF','NoTRF','Costo','Precio Lista','Comentarios')
                        
                        '---------------------------------------------------
                        '---------------------------------------------------
                        strCadena = "INSERT INTO Existencias_OC VALUES ('" & sUserRepositorio & "','" & Trim(rsBusca!Location) & "','" & Trim(rsBusca!itemno) & "','" & Trim(rsBusca!Codigo2) & "'," _
                                    & "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' ',' ',' ','0','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','0','" & strVendorID & "','','','')" 'sin datos
                                    '& "'" & Trim(rsBusca!QTYONHAND) & "',' ','01-01-1900',' ',' ',' ',' ','0','01-01-1900',' ','" & ValorCosto(Trim(rsBusca!itemno)) & "','" & ValorPL(Trim(rsBusca!itemno)) & "',' ',' ',' ','0','','')" 'sin datos
                        '---------------------------------------------------
                        '---------------------------------------------------
                                    
                        Debug.Print strCadena
                        cnDB.Execute strCadena
            End Select
            rsBusca.MoveNext
        Loop
        InfoDisponible = True
        cmdExporta.Enabled = True
    Else
        'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
        InfoDisponible = False
    End If
    rsBusca.Close
    
End If

If InfoDisponible Then
    '---------------------------------------------------
    '---------------------------------------------------
    ssql = "select Articulo,SUM(Cantidad) Cantidad,NoOC,NoRM,NoLinea,Sucursal from Existencias_OC where Usuario='" & sUserRepositorio & "' group by NoOC,NoRM,Sucursal,Articulo,NoLinea"
    '---------------------------------------------------
    '---------------------------------------------------
    Debug.Print ssql
    
    rsBusca.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBusca.EOF = False And rsBusca.BOF = False Then
        Do While rsBusca.EOF = False
        
''''            If Trim(rsBusca!articulo) = "MK4LI000018" Then
''''                 MsgBox "Revisar este Caso"
''''            End If
        
            '---------------------------------------------------
            '---------------------------------------------------
            ssql2 = "select * from Existencias_OC where Usuario='" & sUserRepositorio & "' and NoOC='" & Trim(rsBusca!NoOC) & "' and NoRM='" & Trim(rsBusca!NoRM) & "' and Sucursal='" & Trim(rsBusca!Sucursal) & "' and Articulo='" & Trim(rsBusca!articulo) & "' and NoLinea='" & Trim(rsBusca!NoLinea) & "'"
            '---------------------------------------------------
            '---------------------------------------------------
            Debug.Print ssql2
            
            rsDemas.Open ssql2, cnDB, adOpenForwardOnly, adLockReadOnly
            If rsDemas.EOF = False And rsDemas.BOF = False Then
                Do While rsDemas.EOF = False
                    If Trim(rsDemas!SerLot) <> "" Then strSerLot = Trim(rsDemas!SerLot) & "," & strSerLot
                    Codigo2 = Trim(rsDemas!Codigo2)
                    FechaRM = Trim(rsDemas!FechaRM)
                    
                    ''**  Sales Order se toma de la Orden de Compra
                    SalesO = Trim(rsDemas!SalesOrd)
                    
                    NoLinea = Trim(rsDemas!NoLinea)
                    CostoOc = rsDemas!CostoOc
                    FechaTRF = Trim(rsDemas!FechaTRF)
                    NoTFR = Trim(rsDemas!NoTRF)
                    strProveedor = Trim(rsDemas!Proveedor)
                    Costo = Trim(rsDemas!Costo) 'Format(Trim(rsDemas!Costo), "###,###.#0")
                    PricoLista = Trim(rsDemas!PrecioLista) 'Format(Trim(rsDemas!PrecioLista), "###,###.#0")
                    Comentario = Replace(Trim(rsDemas!Comentarios), Chr(9), " ")
                    ComentarioTRF = Replace(Trim(rsDemas!ComentariosTRF), Chr(9), " ")
                    NoFacRM = Trim(rsDemas!NoFactura)
                    NoLineaOC = Trim(rsDemas!NoLineaOC)
                    comenRM = Replace(Trim(rsDemas!comentariosRM), Chr(13) + Chr(10), "")
                    
                    '-------------------------------------------------
                    strNoPV = Trim(rsDemas!NUMPV)
                    strDatosAnexos = fnStrDatosAnexos(strNoPV) ' *** Trae Datos de la Orden de Venta
                    lsDatosAnexos = Split(strDatosAnexos, "~")
                    strNumPedido = strNoPV
                    strIDCliente = lsDatosAnexos(1)
                    strCliente = lsDatosAnexos(2)
                    strIDEjecutivo = lsDatosAnexos(3)
                    
                    'SalesO = lsDatosAnexos(4)
                    '-------------------------------------------------
                    
                    rsDemas.MoveNext
                    CostoOc = CostoOc * rsBusca!Cantidad 'Format(CostoOc * rsBusca!Cantidad, "###,###.#0")
                    no_producto = encuentraproducto(Trim(rsBusca!articulo))
                    
                Loop
                
                
'''                If Trim(rsBusca!NoOC) = "OC00148976" Then
'''                    MsgBox "Revisar Caso"
'''''                End If

                '-------------------------------------------------
                If Trim(rsBusca!NoOC) <> "" Then
                    sDatosAgregados = fnDatosAgreados(Trim(rsBusca!NoOC))
                    lDatosAgregados = Split(sDatosAgregados, "~")

                    sDEALID = IIf(lDatosAgregados(0) = "", 0, lDatosAgregados(0))
                    sDEALID = Round(sDEALID, 0)
                    
'--------------------------------------
'--------------------------------------
            sIDENDUSER = lDatosAgregados(1)

            TipoArticulo = Left(Trim(rsBusca!articulo), 2)
            '' 20240116
            '' -- If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Then
            If TipoArticulo = "CI" Or TipoArticulo = "CY" Or TipoArticulo = "MK" Or TipoArticulo = "CM" Then
                sIDENDUSER = lDatosAgregados(1)
                sENDUSER = strEndUserName(sIDENDUSER)
            Else
                If strNoPV <> "" Then
                        sIDENDUSER = fnBuscaIDEndUserXPV(strNoPV, Trim(rsBusca!articulo))
                        sENDUSER = strEndUserNameXPV(sIDENDUSER, Trim(rsBusca!articulo))
                Else
                        sIDENDUSER = 0
                        sENDUSER = ""
                End If
            End If
            
'--------------------------------------
'--------------------------------------
                    
                    sOCCLIENTE = lDatosAgregados(2)
                    
                    'EN LOS CASOS en los que la OC DEL CLIENTE ESTE EN BLANCO Y SE CUENTE CON NUMERO DE PV ENTONCES
                    If Trim(sOCCLIENTE) = "" And strNoPV <> "" Then
                        sOCCLIENTE = fnBuscaOCCLIENTE(strNoPV)
                    End If
                    
                    
                    sSALEORDER = lDatosAgregados(3)
                    sFechaOC = lDatosAgregados(4)
                    
                    sFechaOC = Mid(sFechaOC, 1, 4) & "-" & Mid(sFechaOC, 5, 2) & "-" & Mid(sFechaOC, 7, 2)
                    
                    
                    strIDEjecutivo = strIDEjecutivo
                    strEjecutivo = FnStrEjecutivo(strIDEjecutivo)
                Else
                    ' No deberia pasar por aqui
                    sDEALID = 0
                    sIDENDUSER = 0
                    sENDUSER = ""
                    sOCCLIENTE = ""
                    sSALEORDER = ""
                    strIDEjecutivo = 0
                    strEjecutivo = ""
                    sFechaOC = ""
                End If
                
                sFACTCOMER = fnBuscaFACTCOMER(Trim(rsBusca!NoRM))
                
                ''20240206
                sFACTPACK = fnBuscaFACTPACK(Trim(rsBusca!NoRM))
                
'                '----------------------------------------------------------
'                strCadena = Trim(rsBusca!Sucursal) & vbTab & Trim(rsBusca!articulo) & vbTab & Trim(Codigo2) & vbTab & Trim(rsBusca!Cantidad) & vbTab & Trim(strSerLot) & vbTab _
'                            & Trim(FechaRM) & vbTab & Trim(rsBusca!NoRM) & vbTab & Trim(NoFacRM) & vbTab & Trim(comenRM) & vbTab & Trim(rsBusca!NoOC) & vbTab & Trim(NoLineaOC) & vbTab _
'                            & Trim(SalesO) & vbTab & Trim(NoLinea) & vbTab & Trim(CostoOc) & vbTab & Trim(FechaTRF) & vbTab & Trim(NoTFR) & vbTab & Trim(Costo) & vbTab _
'                            & Trim(PricoLista) & vbTab & Trim(Comentario) & vbTab & Trim(ComentarioTRF) & vbTab & strProveedor & vbTab & no_producto & vbTab _
'                            & sDEALID & vbTab & sIDENDUSER & vbTab & sENDUSER & vbTab & sOCCLIENTE & vbTab _
'                            & Trim(strNumPedido) & vbTab & Trim(strIDCliente) & vbTab & Trim(strCliente) & vbTab _
'                            & strIDEjecutivo & vbTab & strEjecutivo & vbTab & sFechaOC & vbTab & sFACTCOMER & vbTab & sFACTPACK
'                SSOleDBGrid1.AddItem strCadena
'                SumatoriaTotal = SumatoriaTotal + CostoOc
'                '----------------------------------------------------------
                '----------------------------------------------------------
                strCadena = "INSERT INTO [dbo].[Existencias_OC_Repositorio]"
                strCadena = strCadena & " ([Usuario],[Sucursal],[Articulo],[Codigo2],[Cantidad],[SerLot],[FechaRM],[NoRM],[NoFacturaRM]"
                strCadena = strCadena & " ,[comentariosRM],[NoOC],[NoLineaOC],[SalesOrd],[NoLinea],[CostoOC],[FechaTRF],[NoTRF],[Costo]"
                strCadena = strCadena & " ,[PrecioLista],[Comentarios],[ComentariosTRF],[Proveedor],[NoProducto],[DealID],[IDEndUser]"
                strCadena = strCadena & " ,[EndUserName],[OCCliente],[NUMPV],[IDCliente],[NombreCliente],[IDEjecutivo],[EjecutivoName]"
                strCadena = strCadena & " ,[FechaOC],[FacturaComer],[FacturaPack],[Category])"
                strCadena = strCadena & " Values ('"
                strCadena = strCadena & sUserRepositorio & "','" & Trim(rsBusca!Sucursal) & "','" & Trim(rsBusca!articulo) & "','" & Trim(Codigo2) & "','" & Trim(rsBusca!Cantidad) & "','" & Trim(strSerLot) & "','" _
                            & Trim(FechaRM) & "','" & Trim(rsBusca!NoRM) & "','" & Trim(NoFacRM) & "','" & Replace(Trim(comenRM), "'", "''") & "','" & Trim(rsBusca!NoOC) & "','" & Trim(NoLineaOC) & "','" _
                            & Trim(SalesO) & "','" & Trim(NoLinea) & "','" & Val(CostoOc) & "','" & Trim(FechaTRF) & "','" & Trim(NoTFR) & "','" & Val(Costo) & "','" _
                            & Trim(PricoLista) & "','" & Replace(Trim(Comentario), "'", "''") & "','" & Replace(Trim(ComentarioTRF), "'", "''") & "','" & strProveedor & "','" & no_producto & "','" _
                            & sDEALID & "','" & sIDENDUSER & "','" & sENDUSER & "','" & sOCCLIENTE & "','" _
                            & Trim(strNumPedido) & "','" & Trim(strIDCliente) & "','" & Trim(strCliente) & "','" _
                            & strIDEjecutivo & "','" & strEjecutivo & "','" & sFechaOC & "','" & sFACTCOMER & "','" & sFACTPACK _
                            & "','" & EncuentraCategory(Trim(rsBusca!articulo)) & "')"
                '----------------------------------------------------------
                Debug.Print strCadena
                cnDB.Execute strCadena
                
                strSerLot = ""
            End If
            rsDemas.Close
            rsBusca.MoveNext
        Loop

        strCadena = "TOTAL" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab _
                    & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab _
                    & vbTab & "" & vbTab & SumatoriaTotal & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
                    
                    
        SSOleDBGrid1.AddItem strCadena
        
        'MsgBox "Se ha completado la busqueda de información con los parámetros establecidos", vbInformation + vbOKOnly, "Información"
        'cmdExporta.Enabled = True
    Else
        'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
    End If
    rsBusca.Close
Else
    'MsgBox "Sin información que mostrar con los parámetros definidos", vbExclamation + vbOKOnly, "Sin Datos"
End If

Set rsBusca = Nothing
Set rsSerLot = Nothing
Set rsRecibo = Nothing
Set rsTrans = Nothing
Set rsDemas = Nothing

Screen.MousePointer = vbNormal
frmExisteOC.Enabled = True

Exit Sub

RutinaError:
If Err.Number <> -2147217873 Then 'NO SE MUESTRA EL MENSAJE SI ES REGISTRO DUPLICADO
    'MsgBox Err.Number & "  " & Err.Description, vbCritical + vbOK, "Error"
End If
Resume Next

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
lSignonID = 0 ' MUST be initialized to 0 since you don't have a signon ID yet
Set mSessMgr = New AccpacSessionMgr
With mSessMgr
.AppID = "PO"
.AppVersion = "70A"
.ProgramName = "PO1210"
.ServerName = "" ' empty string if running on local computer
.CreateSession "", lSignonID, mSession ' first argument is the object handle (if you don't have one, pass "")
End With ' mSessMgr

Call mSession.GetSignonInfo(usuario, Empresa, EmpresaNombre)
    'MsgBox usuario & " " & empresa & " " & empresanombre
conecta
LlenaCombos
CrearTablaRepositorio
If Command <> "" Then
    Call cmdRepositorio_Click
    End
End If
End Sub

Public Sub IniciaSesionSage(EmpresaSesion As String)
Dim UserID As String
Dim Password As String
Dim OrgID As String
Dim OrgDesc As String
Dim SessionDate As Date
Dim sid As Object

Dim objSignOn As AccpacSignonManager.AccpacSignonMgr
Set objSession = AccpacCOMAPI.AccpacSession

OrgID = EmpresaSesion 'Gempresa
SessionDate = DTPicker1.Value

objSession.Init "", "AS", "AS1000", "70A"
objSession.Open sUserAcc, sPassAcc, OrgID, SessionDate, 0, ""

If objSession.IsOpened Then
    Set objCompany = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
    Set dbCmp = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
End If
End Sub

Private Sub CrearTablaRepositorio()
On Error GoTo RutinaError

Dim ssql As String

ssql = "CREATE TABLE [dbo].[Existencias_OC_Repositorio]("
ssql = ssql & "    [Usuario] [nvarchar](10) NOT NULL,"
ssql = ssql & "    [Sucursal] [nvarchar](10) NOT NULL,"
ssql = ssql & "    [Articulo] [nvarchar](11) NOT NULL,"
ssql = ssql & "    [Codigo2] [nvarchar](60) NULL,"
ssql = ssql & "    [Cantidad] [float] NULL,"
ssql = ssql & "    [SerLot] [nvarchar](max) NOT NULL,"
ssql = ssql & "    [FechaRM] [date] NULL,"
ssql = ssql & "    [NoRM] [nvarchar](15) NULL,"
ssql = ssql & "    [NoFacturaRM] [nvarchar](45) NULL,"
ssql = ssql & "    [comentariosRM] [nvarchar](max) NULL,"
ssql = ssql & "    [NoOC] [nvarchar](15) NOT NULL,"
ssql = ssql & "    [NoLineaOC] [float] NULL,"
ssql = ssql & "    [SalesOrd] [nvarchar](40) NULL,"
ssql = ssql & "    [NoLinea] [float] NULL,"
ssql = ssql & "    [CostoOC] [float] NULL,"
ssql = ssql & "    [FechaTRF] [date] NULL,"
ssql = ssql & "    [NoTRF] [nvarchar](15) NULL,"
ssql = ssql & "    [Costo] [float] NULL,"
ssql = ssql & "    [PrecioLista] [nvarchar](10) NULL,"
ssql = ssql & "    [Comentarios] [nvarchar](max) NULL,"
ssql = ssql & "    [ComentariosTRF] [nvarchar](max) NULL,"
ssql = ssql & "    [Proveedor] [nchar](60) NULL,"
ssql = ssql & "    [NoProducto] [nchar](20) NULL,"
ssql = ssql & "    [DealID] [nchar](30) NULL,"
ssql = ssql & "    [IDEndUser] [nchar](10) NULL,"
ssql = ssql & "    [EndUserName] [nchar](80) NULL,"
ssql = ssql & "    [OCCliente] [nchar](150) NULL,"
ssql = ssql & "    [NUMPV] [nchar](10) NULL,"
ssql = ssql & "    [IDCliente] [nchar](10) NULL,"
ssql = ssql & "    [NombreCliente] [nchar](60) NULL,"
ssql = ssql & "    [IDEjecutivo] [nchar](12) NULL,"
ssql = ssql & "    [EjecutivoName] [nchar](60) NULL,"
ssql = ssql & "    [FechaOC] [date] NULL,"
ssql = ssql & "    [FacturaComer] [nchar](100) NULL,"
ssql = ssql & "    [FacturaPack] [nchar](100) NULL,"
ssql = ssql & "    [Category] [nchar](10) NULL"
ssql = ssql & ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

cnDB.Execute ssql

Exit Sub

RutinaError:
If Err.Number = -2147217900 Then
    'YA EXISTE LA TABLA
Else
    MsgBox Err.Number + " - " + Err.Description, vbCritical + vbOKOnly, "Error"
End If
Resume Next
End Sub

Private Function EncuentraLineaOC(PORHSEQ As String, itemno As String, UNITWEIGHT As String) As Long
Dim ssql As String
Dim rsLinea As Recordset
Set rsLinea = New Recordset

ssql = "select PORLREV/1000 as LINEA from POPORL where PORHSEQ=" & PORHSEQ & " and ITEMNO='" & itemno & "' and UNITWEIGHT=" & UNITWEIGHT
Debug.Print ssql

rsLinea.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsLinea.EOF = False And rsLinea.BOF = False Then
    EncuentraLineaOC = rsLinea!LINEA
Else
    EncuentraLineaOC = 0
End If
rsLinea.Close
Set rsLinea = Nothing
End Function

Private Sub LlenaCombos()
Dim ssql As String
Dim rsDatos As Recordset
Set rsDatos = New Recordset

ssql = "select LOCATION from ICLOC order by LOCATION" 'SUCURSAL
rsDatos.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsDatos.EOF = False And rsDatos.BOF = False Then
    rsDatos.MoveFirst
    cmbDesdeSucursal.AddItem "TODAS"
    cmbHastaSucursal.AddItem "TODAS"
    Do While rsDatos.EOF = False
        cmbDesdeSucursal.AddItem Trim(rsDatos!Location)
        cmbHastaSucursal.AddItem Trim(rsDatos!Location)
        rsDatos.MoveNext
    Loop
    cmbDesdeSucursal.ListIndex = 0
    cmbHastaSucursal.ListIndex = 0
End If
rsDatos.Close

'''ssql = "SELECT ITEMNO FROM ICITEM order by ITEMNO" 'ARTÍCULO
'''rsDatos.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
'''If rsDatos.EOF = False And rsDatos.BOF = False Then
'''    rsDatos.MoveFirst
'''    Do While rsDatos.EOF = False
'''        txtDesdeArticulo.AddItem Trim(rsDatos!itemno)
'''        txtHastaArticulo.AddItem Trim(rsDatos!itemno)
'''        rsDatos.MoveNext
'''    Loop
'''    'txtDesdeArticulo.ListIndex = 0
'''    'txtHastaArticulo.ListIndex = txtHastaArticulo.ListCount - 1
'''End If
'''rsDatos.Close

ssql = "select CATEGORY from ICCATG order by CATEGORY" 'MARCAS
rsDatos.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsDatos.EOF = False And rsDatos.BOF = False Then
    rsDatos.MoveFirst
    Do While rsDatos.EOF = False
        cmbDesdeMarca.AddItem Trim(rsDatos!Category)
        cmbHastaMarca.AddItem Trim(rsDatos!Category)
        rsDatos.MoveNext
    Loop
    cmbDesdeMarca.ListIndex = 0
    cmbHastaMarca.ListIndex = cmbHastaMarca.ListCount - 1
End If
rsDatos.Close

DTPicker1.Value = Date - 30
DTPicker2.Value = Date

Set rsDatos = Nothing

optAyer.Caption = Date - 1
optHoy.Caption = Date

End Sub

Private Sub Form_Resize()
On Error Resume Next

Frame3.Top = frmExisteOC.Height - 1300
Frame3.Left = frmExisteOC.Width - 6000

cmdExporta.Top = frmExisteOC.Height - 1000
cmdSalir.Top = frmExisteOC.Height - 1000
cmdSalir.Left = frmExisteOC.Width - 1800
Frame2.Height = frmExisteOC.Height - 2950
SSOleDBGrid1.Height = frmExisteOC.Height - 3350
Frame2.Width = frmExisteOC.Width - 400
SSOleDBGrid1.Width = frmExisteOC.Width - 650

End Sub

Private Sub Form_Unload(Cancel As Integer)
Desconecta
End Sub

'---------------------
Private Sub Generagrid()
SSOleDBGrid1.Columns.RemoveAll
 
SSOleDBGrid1.Columns.Add 0
SSOleDBGrid1.Columns(0).Visible = True
SSOleDBGrid1.Columns(0).Caption = "Sucursal"
SSOleDBGrid1.Columns(0).Width = 1000

SSOleDBGrid1.Columns.Add 1
SSOleDBGrid1.Columns(1).Visible = True
SSOleDBGrid1.Columns(1).Caption = "Artículo"
SSOleDBGrid1.Columns(1).Width = 1200

SSOleDBGrid1.Columns.Add 2
SSOleDBGrid1.Columns(2).Visible = True
SSOleDBGrid1.Columns(2).Caption = "Código 2"
SSOleDBGrid1.Columns(2).Width = 1200

SSOleDBGrid1.Columns.Add 3
SSOleDBGrid1.Columns(3).Visible = True
SSOleDBGrid1.Columns(3).Caption = "Existencia"
SSOleDBGrid1.Columns(3).Width = 1000

SSOleDBGrid1.Columns.Add 4
SSOleDBGrid1.Columns(4).Visible = True
SSOleDBGrid1.Columns(4).Caption = "Serie/Lote"
SSOleDBGrid1.Columns(4).Width = 1200

SSOleDBGrid1.Columns.Add 5
SSOleDBGrid1.Columns(5).Visible = True
SSOleDBGrid1.Columns(5).Caption = "Fecha RM"
SSOleDBGrid1.Columns(5).Width = 1200

SSOleDBGrid1.Columns.Add 6
SSOleDBGrid1.Columns(6).Visible = True
SSOleDBGrid1.Columns(6).Caption = "No RM"
SSOleDBGrid1.Columns(6).Width = 1200

SSOleDBGrid1.Columns.Add 7
SSOleDBGrid1.Columns(7).Visible = True
SSOleDBGrid1.Columns(7).Caption = "No Factura en RM"
SSOleDBGrid1.Columns(7).Width = 1200

SSOleDBGrid1.Columns.Add 8
SSOleDBGrid1.Columns(8).Visible = True
SSOleDBGrid1.Columns(8).Caption = "Comentarios de RM"
SSOleDBGrid1.Columns(8).Width = 1200

SSOleDBGrid1.Columns.Add 9
SSOleDBGrid1.Columns(9).Visible = True
SSOleDBGrid1.Columns(9).Caption = "No OC"
SSOleDBGrid1.Columns(9).Width = 1200

SSOleDBGrid1.Columns.Add 10
SSOleDBGrid1.Columns(10).Visible = True
SSOleDBGrid1.Columns(10).Caption = "No de Linea de OC"
SSOleDBGrid1.Columns(10).Width = 1000

SSOleDBGrid1.Columns.Add 11
SSOleDBGrid1.Columns(11).Visible = True
SSOleDBGrid1.Columns(11).Caption = "Sales Order"
SSOleDBGrid1.Columns(11).Width = 1200

SSOleDBGrid1.Columns.Add 12
SSOleDBGrid1.Columns(12).Visible = True
SSOleDBGrid1.Columns(12).Caption = "No de Linea de CISCO"
SSOleDBGrid1.Columns(12).Width = 1200

SSOleDBGrid1.Columns.Add 13
SSOleDBGrid1.Columns(13).Visible = True
SSOleDBGrid1.Columns(13).Caption = "Costo OC"
SSOleDBGrid1.Columns(13).Width = 1200

SSOleDBGrid1.Columns.Add 14
SSOleDBGrid1.Columns(14).Visible = True
SSOleDBGrid1.Columns(14).Caption = "Fecha TRC"
SSOleDBGrid1.Columns(14).Width = 1200

SSOleDBGrid1.Columns.Add 15
SSOleDBGrid1.Columns(15).Visible = True
SSOleDBGrid1.Columns(15).Caption = "No TRC"
SSOleDBGrid1.Columns(15).Width = 1200

SSOleDBGrid1.Columns.Add 16
SSOleDBGrid1.Columns(16).Visible = True
SSOleDBGrid1.Columns(16).Caption = "Costo"
SSOleDBGrid1.Columns(16).Width = 1200

SSOleDBGrid1.Columns.Add 17
SSOleDBGrid1.Columns(17).Visible = True
SSOleDBGrid1.Columns(17).Caption = "Precio Lista"
SSOleDBGrid1.Columns(17).Width = 1200

SSOleDBGrid1.Columns.Add 18
SSOleDBGrid1.Columns(18).Visible = True
SSOleDBGrid1.Columns(18).Caption = "Comentarios OC"
SSOleDBGrid1.Columns(18).Width = 1200

SSOleDBGrid1.Columns.Add 19
SSOleDBGrid1.Columns(19).Visible = True
SSOleDBGrid1.Columns(19).Caption = "Comentarios TRF"
SSOleDBGrid1.Columns(19).Width = 1200

SSOleDBGrid1.Columns.Add 20
SSOleDBGrid1.Columns(20).Visible = True
SSOleDBGrid1.Columns(20).Caption = "Proveedor"
SSOleDBGrid1.Columns(20).Width = 1200

SSOleDBGrid1.Columns.Add 21
SSOleDBGrid1.Columns(21).Visible = True
SSOleDBGrid1.Columns(21).Caption = "N° Producto"
SSOleDBGrid1.Columns(21).Width = 1200

'--------------------------------------------------

    SSOleDBGrid1.Columns.Add 22
    SSOleDBGrid1.Columns(22).Visible = True
    SSOleDBGrid1.Columns(22).Caption = "DEALID"
    SSOleDBGrid1.Columns(22).Width = 1500
    
    SSOleDBGrid1.Columns.Add 23
    SSOleDBGrid1.Columns(23).Visible = True
    SSOleDBGrid1.Columns(23).Caption = "ID EndUser"
    SSOleDBGrid1.Columns(23).Width = 1500
    
    SSOleDBGrid1.Columns.Add 24
    SSOleDBGrid1.Columns(24).Visible = True
    SSOleDBGrid1.Columns(24).Caption = "EndUser"
    SSOleDBGrid1.Columns(24).Width = 1500
    
    SSOleDBGrid1.Columns.Add 25
    SSOleDBGrid1.Columns(25).Visible = True
    SSOleDBGrid1.Columns(25).Caption = "OC Cliente"
    SSOleDBGrid1.Columns(25).Width = 1500

    SSOleDBGrid1.Columns.Add 26
    SSOleDBGrid1.Columns(26).Visible = True
    SSOleDBGrid1.Columns(26).Caption = "Numero de Pedido"
    SSOleDBGrid1.Columns(26).Width = 1500
    
    SSOleDBGrid1.Columns.Add 27
    SSOleDBGrid1.Columns(27).Visible = True
    SSOleDBGrid1.Columns(27).Caption = "ID Cliente"
    SSOleDBGrid1.Columns(27).Width = 1500
    
    SSOleDBGrid1.Columns.Add 28
    SSOleDBGrid1.Columns(28).Visible = True
    SSOleDBGrid1.Columns(28).Caption = "Cliente"
    SSOleDBGrid1.Columns(28).Width = 1500
    
    SSOleDBGrid1.Columns.Add 29
    SSOleDBGrid1.Columns(29).Visible = True
    SSOleDBGrid1.Columns(29).Caption = "ID Ejecutivo"
    SSOleDBGrid1.Columns(29).Width = 1500

    SSOleDBGrid1.Columns.Add 30
    SSOleDBGrid1.Columns(30).Visible = True
    SSOleDBGrid1.Columns(30).Caption = "Ejecutivo"
    SSOleDBGrid1.Columns(30).Width = 1500

    SSOleDBGrid1.Columns.Add 31
    SSOleDBGrid1.Columns(31).Visible = True
    SSOleDBGrid1.Columns(31).Caption = "FechaOC"
    SSOleDBGrid1.Columns(31).Width = 1500

    SSOleDBGrid1.Columns.Add 32
    SSOleDBGrid1.Columns(32).Visible = True
    SSOleDBGrid1.Columns(32).Caption = "Factura Comercial"
    SSOleDBGrid1.Columns(32).Width = 1500

    ''--20240206
    SSOleDBGrid1.Columns.Add 33
    SSOleDBGrid1.Columns(33).Visible = True
    SSOleDBGrid1.Columns(33).Caption = "FACTPACK"
    SSOleDBGrid1.Columns(33).Width = 1500


End Sub

Private Function FormatoFecha(Fechas As Date) As String
FormatoFecha = Year(Fechas) & IIf(Month(Fechas) < 10, "0" & Month(Fechas), Month(Fechas)) & IIf(Day(Fechas) < 10, "0" & Day(Fechas), Day(Fechas))
End Function

Private Function FormatoFecha2(Fecha As String) As String
FormatoFecha2 = Right(Fecha, 2) & "-" & Mid(Fecha, 5, 2) & "-" & Left(Fecha, 4)
End Function

Private Function EncuentraCategory(sArticulo As String)
Dim ssql As String
Dim rsProducto As Recordset
Set rsProducto = New Recordset

ssql = "select CATEGORY from icitem where itemno='" & sArticulo & "' "
rsProducto.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsProducto.EOF = False And rsProducto.BOF = False Then
    EncuentraCategory = RTrim(rsProducto!Category)
Else
    EncuentraCategory = ""
End If
rsProducto.Close
Set rsProducto = Nothing
End Function

Private Function encuentraproducto(articulo As String)
Dim ssql As String
Dim rsProducto As Recordset
Set rsProducto = New Recordset

ssql = "select commodim from icitem where itemno='" & articulo & "' "
rsProducto.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsProducto.EOF = False And rsProducto.BOF = False Then
    encuentraproducto = RTrim(rsProducto!commodim)
Else
    encuentraproducto = ""
End If
rsProducto.Close
Set rsProducto = Nothing
End Function


Function fnGetFirst() As String

    Dim ssqlExistencias_OC As String
    Dim rsExistencias_OC As Recordset
    Set rsExistencias_OC = New Recordset
    
    ssqlExistencias_OC = "select top 1 Articulo from Existencias_OC"
    Debug.Print ssqlExistencias_OC
    
    rsExistencias_OC.Open ssqlExistencias_OC, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsExistencias_OC.EOF = False And rsExistencias_OC.BOF = False Then
        fnGetFirst = RTrim(rsExistencias_OC!articulo)
    Else
        fnGetFirst = ""
    End If
    
    rsExistencias_OC.Close
    Set rsExistencias_OC = Nothing


End Function


Function fnBuscaOCCLIENTE(strPV) As String
'
    Dim ssqlBuscaOCCLIENTE As String
    Dim rsBuscaOCCLIENTE As Recordset
    Set rsBuscaOCCLIENTE = New Recordset
    
    ssqlBuscaOCCLIENTE = "select PONUMBER from OEORDH  WHERE ORDNUMBER = '" & strPV & "'"
    Debug.Print ssqlBuscaOCCLIENTE
    
    rsBuscaOCCLIENTE.Open ssqlBuscaOCCLIENTE, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBuscaOCCLIENTE.EOF = False And rsBuscaOCCLIENTE.BOF = False Then
        fnBuscaOCCLIENTE = RTrim(rsBuscaOCCLIENTE!PONUMBER)
    Else
        fnBuscaOCCLIENTE = ""
    End If
    
    rsBuscaOCCLIENTE.Close
    Set rsBuscaOCCLIENTE = Nothing
    
End Function



Function fnBuscaFACTCOMER(sNoRM As String) As String

    Dim sSQLBuscaFACTCOMER As String
    Dim rsBuscaFACTCOMER As Recordset
    Set rsBuscaFACTCOMER = New Recordset
    
    sSQLBuscaFACTCOMER = "select VALUE  FROM PORCPH1 A INNER JOIN PORCPHO B on a.RCPHSEQ = b.RCPHSEQ WHERE A.RCPNUMBER = '" & sNoRM & "' AND OPTFIELD = 'FACTCOMER'"
    Debug.Print sSQLBuscaFACTCOMER
    
    rsBuscaFACTCOMER.Open sSQLBuscaFACTCOMER, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBuscaFACTCOMER.EOF = False And rsBuscaFACTCOMER.BOF = False Then
        fnBuscaFACTCOMER = RTrim(rsBuscaFACTCOMER!Value)
    Else
        fnBuscaFACTCOMER = ""
    End If
    
    rsBuscaFACTCOMER.Close
    Set rsBuscaFACTCOMER = Nothing

End Function

''20240206
Function fnBuscaFACTPACK(sNoRM As String) As String

    Dim sSQLBuscaFACTPACK As String
    Dim rsBuscaFACTPACK As Recordset
    Set rsBuscaFACTPACK = New Recordset
    
    sSQLBuscaFACTPACK = "select VALUE FROM PORCPH1 A INNER JOIN PORCPHO B on a.RCPHSEQ = b.RCPHSEQ WHERE A.RCPNUMBER = '" & sNoRM & "' AND OPTFIELD = 'FACTPACK'"
    Debug.Print sSQLBuscaFACTPACK
    
    rsBuscaFACTPACK.Open sSQLBuscaFACTPACK, cnDB, adOpenForwardOnly, adLockReadOnly
    If rsBuscaFACTPACK.EOF = False And rsBuscaFACTPACK.BOF = False Then
        fnBuscaFACTPACK = RTrim(rsBuscaFACTPACK!Value)
    Else
        fnBuscaFACTPACK = ""
    End If
    
    rsBuscaFACTPACK.Close
    Set rsBuscaFACTPACK = Nothing

End Function

