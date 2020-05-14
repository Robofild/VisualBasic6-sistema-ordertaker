Attribute VB_Name = "conexaobrid"
 Option Explicit

Public con As ADODB.Connection
Public conl As ADODB.Connection
Function ConServer()

Set con = New ADODB.Connection
con.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=216.172.173.216;Database=robofi61_order_taker;User=robofi61_gloria;Password=rarianepgloria;Option=3;"
con.Open





End Function
Function ConServerloc()

Set conl = New ADODB.Connection
conl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Order_Taker\CEP.MDB;Persist Security Info=False"
conl.Open

End Function
Public Function LocalizarCoordenadasGeograficasRelevantes()




'INSERT INTO `ProntoAtendimentoTelemarket` (`id`, `NumeroPrato`, `obs`) VALUES (NULL, '52', 'preparo simples');

'ConServer

'Dim sql As String
'Dim rs As New ADODB.Recordset
'
'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = con

'sql = "SELECT * FROM `coordenadasgeonow` WHERE `Latitude` BETWEEN " & latidudinalMinimo & " AND " & latitudinalMaximos & " AND `Longitude` BETWEEN " & LongitudeMinima & "  AND  " & LongitudeMaxima & "  ORDER BY `idcoordenadasGeoNow` ASC"

'rs.Open sql
'Form50.Text13.Text = Adodc1.Recordset("idCardapio_medidas").Value



'If Adodc1.Recordset.EOF = False Then
'Adodc1.Refresh
'rs.Fields("idingredientes_por_id_anexos").Value
'rs.Close sql
'identificadorIndicesTabela
'SituaçãoPrimeiraLinha

'End If



 '  Set rs = Nothing

End Function
