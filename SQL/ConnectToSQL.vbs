Const adOpenStatic = 3 
Const adLockOptimistic = 3 
 
Set objConnection = CreateObject("ADODB.Connection") 
Set objRecordSet = CreateObject("ADODB.Recordset") 
 
objConnection.Open _ 
    "Provider=SQLOLEDB;Data Source=atl-sql-01;" & _ 
        "Trusted_Connection=Yes;Initial Catalog=Northwind;" & _ 
             "User ID=fabrikam\kenmyer;Password=34DE6t4G!;" 
 
objRecordSet.Open "SELECT * FROM Customers", _ 
        objConnection, adOpenStatic, adLockOptimistic 
 
objRecordSet.MoveFirst 
 
Wscript.Echo objRecordSet.RecordCount 
