Attribute VB_Name = "DatabaseConnectionModul"
Const CONFIG_MYSQL_USERNAME_RANGE As String = "B6"
Const CONFIG_MYSQL_PASSWORD_RANGE As String = "B7"
Const CONFIG_MYSQL_HOSTNAME_RANGE As String = "B4"
Const CONFIG_MYSQL_DATABASENAME_RANGE As String = "B5"
Const CONFIG_MYSQL_PORT_RANGE As String = "B8"
' return mySqlConfigObject with DB configuration from Settings sheet

Function getConnection(ByVal mySqlConfig As mySqlConfigObject)

    Dim conn As ADODB.connection
    Set conn = New ADODB.connection
    conn.connectionString = getConnectionStringFromMySqlConfigObject(mySqlConfig)
    conn.Open
    Set getConnection = conn
End Function

Function getConnectionStringFromMySqlConfigObject(ByVal mySqlConfig As mySqlConfigObject) As String
    Dim connectionStringTemplate As String
    
    connectionStringTemplate = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" _
        & "SERVER=#HOSTNAME#;" _
        & " DATABASE=#DATABASE#;" _
        & "UID=#USERNAME#;PWD=#PWD#; OPTION=3;PORT=#PORT#"
    Dim connectionString
    connectionString = Replace(connectionStringTemplate, "#HOSTNAME#", mySqlConfig.hostname)
    connectionString = Replace(connectionString, "#DATABASE#", mySqlConfig.database)
    connectionString = Replace(connectionString, "#USERNAME#", mySqlConfig.username)
    connectionString = Replace(connectionString, "#PWD#", mySqlConfig.password)
    connectionString = Replace(connectionString, "#PORT#", mySqlConfig.port)
    getConnectionStringFromMySqlConfigObject = connectionString
End Function


Function getMySqlConfigObjectFromConfigSheet() As mySqlConfigObject
    Dim mySqlConfigObject As mySqlConfigObject
    Set mySqlConfigObject = New mySqlConfigObject
    
    mySqlConfigObject.hostname = ConfigTable.Range(CONFIG_MYSQL_HOSTNAME_RANGE).value
    mySqlConfigObject.database = ConfigTable.Range(CONFIG_MYSQL_DATABASENAME_RANGE).value
    mySqlConfigObject.username = ConfigTable.Range(CONFIG_MYSQL_USERNAME_RANGE).value
    mySqlConfigObject.password = ConfigTable.Range(CONFIG_MYSQL_PASSWORD_RANGE).value
    mySqlConfigObject.port = ConfigTable.Range(CONFIG_MYSQL_PORT_RANGE).value
    
    Set getMySqlConfigObjectFromConfigSheet = mySqlConfigObject
End Function

