Attribute VB_Name = "Funciones"
 
Public cnDB As Connection

Public Function ChecaNullInt(cadena As Variant) As String
ChecaNullInt = IIf(IsNull(cadena), 0, cadena)
End Function

Public Function ChecaNullStr(cadena As Variant) As String
ChecaNullStr = IIf(IsNull(cadena), "", cadena)
End Function
Public Function Formato(cadena As Currency) As Currency
Formato = Format(Round(cadena, 2), "$###,##0.00")
End Function

Public Function conecta()

Set cnDB = New Connection

Dim sINIFile As String
n = False
Dim sschema As String
Dim sTipoDataBase As String
Dim sDataBase As String
Dim sServer As String
Dim sPassword As String

'sschema = "PRUSER"
sINIFile = App.Path & "\dice.ini"

'leer el nombre del archivo ini

   sDataBase = sGetINI(sINIFile, "settings", "databaseAccpac", "?")
   If sDataBase = "?" Then
      MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
      Exit Function
   End If
   sServer = sGetINI(sINIFile, "settings", "servidor", "?")
   If sServer = "?" Then
      MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
      Exit Function
   End If
   sPassword = sGetINI(sINIFile, "settings", "password", "?")
   If sPassword = "?" Then
      MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
      Exit Function
   End If
   
   Do Until cnDB.State = 1
      With cnDB
      .Provider = "SQLOLEDB.1"
      .ConnectionString = "Password='" & sPassword & "';Persist Security Info=True;User ID=sa;Initial Catalog='" & sDataBase & "';Data Source='" & sServer & "'"

      '.ConnectionString = "User ID=sa;" & _
                      "Initial Catalog= '" & sDataBase & "';" & _
                     "Data Source='" & sServer & "'"
      .ConnectionTimeout = 50
      .Open
      .CursorLocation = adUseClient
      End With
   Loop
End Function
Public Function Desconecta()
cnDB.Close
Set cnDB = Nothing
End Function

Public Function Restringido(usuario As String, modulo As String, pantalla As String) As Boolean
Dim rsRestringido As Recordset
Set rsRestringido = New Recordset
Dim scadena As String

conecta (0)
scadena = "SELECT aplica FROM PRIVILEGIOS WHERE UACCPAC = '" & usuario & "' AND MODULO = '" & modulo & "' AND Pantalla = '" & pantalla & "'"
Debug.Print scadena
rsRestringido.Open scadena, cnDB, adOpenForwardOnly, adLockReadOnly
If rsRestringido.EOF = False And rsRestringido.BOF = False Then
   If rsRestringido!aplica = 0 Then
      Restringido = False
   Else
      Restringido = True
   End If
Else
   Restringido = True
End If
rsRestringido.Close
Set rsRestringido = Nothing
cnDB.Close
End Function

'***************************************+
Public Function conecta2(tipo As Integer)

Set cnDB = New Connection

Dim sINIFile As String
Dim sDataBase As String
Dim sServer As String
n = False
sINIFile = App.Path & "\dice.ini"

'leer el nombre del archivo ini
If tipo = 0 Then ' datos
   sDataBase = sGetINI(sINIFile, "settings", "database", "?")
   If sDataBase = "?" Then
      MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
      Exit Function
   End If
Else ' sistema
   sDataBase = sGetINI(sINIFile, "settings", "databasesys", "?")
   If sDataBase = "?" Then
      MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
      Exit Function
   End If
End If
sServer = sGetINI(sINIFile, "settings", "servidor", "?")
If sServer = "?" Then
   MsgBox "No existe el archivo INI, favor de llamar al Administrador del sistema"
   Exit Function
End If
Do Until cnDB.State = 1
   With cnDB
   .Provider = "SQLOLEDB"
   .ConnectionString = "User ID=usuario;" & _
                   "Initial Catalog= '" & sDataBase & "';" & _
                   "Data Source='" & sServer & "'"
   .ConnectionTimeout = 50
   .Open
   End With
Loop
End Function

'********************************************



Public Function Sucursal(usuario As String) As String
Dim rsSucursal As Recordset
Set rsSucursal = New Recordset
Dim scadena As String

conecta (0)

scadena = "SELECT sucursal FROM USUARIOS WHERE uaccpac = '" & usuario & "'"
rsSucursal.Open scadena, cnDB, adOpenForwardOnly, adLockReadOnly
If rsSucursal.BOF = False And rsSucursal.EOF = False Then
   Sucursal = rsSucursal!Sucursal
Else
   Sucursal = "02MEX"
End If
rsSucursal.Close
Set rsSucursal = Nothing
cnDB.Close
End Function
