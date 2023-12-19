Option Compare Database

Public Function GetFieldType(fd) As String
    
    'See reference for field type numbers at https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum?view=sql-server-ver16
      
'     Const dbBigInt = 20
'    Const dbBinary = 128
'    Const dbBoolean = 11
'    'Const dbBSTR = 8
'    Const dbChar = 129
'    'Const dbCurrency = 6
'    Const dbDate = 7
'    Const dbDBDate = 133
'    Const dbTime = 134
'    Const dbTimeStamp = 135
'    Const dbDecimal = 14
'
'    Const dbEmpty = 0
'
'    Const dbFileTime = 64
'    Const dbGUID = 72
'    Const dbInteger = 3
'    Const dbLongVarChar = 201
'    Const dbNumeric = 131
'     'Const dbSmallInt = 2
'    Const dbTinyInt = 16
'    Const dbVarBinary = 204
'    Const dbVarChar = 200
'    Const dbVarNumeric = 139
'
    
    Const dbAttachment = 101
    'Const dbAutoIncrement = 4
    Const dbBoolean = 1
    Const dbCurrency = 5
    Const dbDateTime = 8
    
    Const dbInteger = 3
    
    Const dbNumber = 2
    
    Const dbLongInteger = 4
    
    Const dbLongText = 12
    Const dbShortText = 10
      
      
'    Const dbBigInt = 20
'    Const dbBinary = 128
'    Const dbBoolean = 11
'    Const dbBSTR = 8
'    Const dbChar = 129
'    Const dbCurrency = 6
'    Const dbDate = 7
'    Const dbDBDate = 133
'    Const dbTime = 134
'    Const dbTimeStamp = 135
'    Const dbDecimal = 14
'    Const dbDouble = 5
'    Const dbEmpty = 0
'    Const dbError = 10
'    Const dbFileTime = 64
'    Const dbGUID = 72
'    Const dbInteger = 3
'    Const dbLongVarChar = 201
'    Const dbNumeric = 131
'    Const dbSingle = 4
'    Const dbSmallInt = 2
'    Const dbTinyInt = 16
'    Const dbVarBinary = 204
'    Const dbVarChar = 200
'    Const dbVarNumeric = 139
       
    Dim a
    
    
     Select Case fd.Type
        Case dbAttachment
         a = "dbAttachment"
        Case dbBoolean
         a = "dbBoolean"
        Case dbLongInteger
         a = "dbLongInteger"
        Case dbCurrency
         a = "dbCurrency"
        Case dbDateTime
         a = "dbDateTime"
        Case dbInteger
         a = "dbInteger"
        Case dbShortText
         a = "dbShortText"
        Case dbLongText
         a = "dbLongText"
        Case dbNumber
         a = "dbNumber"
             
'       Case dbBigInt
'        a = "dbBigInt"
'       Case dbBinary
'        a = "dbBinary"
'       Case dbBoolean
'        a = "dbBoolean"
'       Case dbBSTR
'        a = "dbBSTR"
'       Case dbChar
'        a = "dbChar"
'       Case dbDate
'        a = "dbDate"
'       Case dbDBDate
'        a = "dbDBDate"
'       Case dbTime
'        a = "dbTime"
'       Case dbTimeStamp
'        a = "dbTimeStamp"
'       Case dbDecimal
'        a = "dbDecimal"
'       Case dbDouble
'        a = "dbDouble"
'       Case dbEmpty
'        a = "dbEmpty"
'       Case dbError
'        a = "dbError"
'       Case dbFileTime
'        a = "dbFileTime"
'       Case dbGUID
'        a = "dbGUID"
'       Case dbInteger
'        a = "dbInteger"
'       Case dbLongVarChar
'        a = "dbLongVarChar"
'       Case dbNumeric
'        a = "dbNumeric"
'       Case dbSingle
'        a = "dbSingle"
'       Case dbSmallInt
'        a = "dbSmallInt"
'       Case dbTinyInt
'        a = "dbTinyInt"
'       Case dbVarBinary
'        a = "dbVarBinary"
'       Case dbVarChar
'        a = "dbVarChar"
'       Case dbVarNumeric
'        a = "dbVarNumeric"
        
       Case Else
          '>>> raise error
          a = "Field " & fd.Name & _
                " of type " & fd.Type & " has been ignored!!!"
       End Select
    
    
    
'    Select Case fd.Type
'       Case dbBigInt
'        a = "dbBigInt"
'       Case dbBinary
'        a = "dbBinary"
'       Case dbBoolean
'        a = "dbBoolean"
'       Case dbBSTR
'        a = "dbBSTR"
'       Case dbChar
'        a = "dbChar"
'       Case dbCurrency
'        a = "dbCurrency"
'       Case dbDate
'        a = "dbDate"
'       Case dbDBDate
'        a = "dbDBDate"
'       Case dbTime
'        a = "dbTime"
'       Case dbTimeStamp
'        a = "dbTimeStamp"
'       Case dbDecimal
'        a = "dbDecimal"
'       Case dbDouble
'        a = "dbDouble"
'       Case dbEmpty
'        a = "dbEmpty"
'       Case dbError
'        a = "dbError"
'       Case dbFileTime
'        a = "dbFileTime"
'       Case dbGUID
'        a = "dbGUID"
'       Case dbInteger
'        a = "dbInteger"
'       Case dbLongVarChar
'        a = "dbLongVarChar"
'       Case dbNumeric
'        a = "dbNumeric"
'       Case dbSingle
'        a = "dbSingle"
'       Case dbSmallInt
'        a = "dbSmallInt"
'       Case dbTinyInt
'        a = "dbTinyInt"
'       Case dbVarBinary
'        a = "dbVarBinary"
'       Case dbVarChar
'        a = "dbVarChar"
'       Case dbVarNumeric
'        a = "dbVarNumeric"
'       Case Else
'          '>>> raise error
'          a = "Field " & fd.Name & _
'                " of type " & fd.Type & " has been ignored!!!"
'       End Select

    GetFieldType = a

            
'       Case dbLong  'Long
          'test if counter, doesn't detect random property if set
'          If (fld.Attributes And dbAutoIncrField) Then
'             a = "COUNTER"
'          Else
'            a = "LONG"
'          End If
                   
'       Case dbText 'Text
'         a = "VARCHAR(" & fld.Size & ")"
          
  
End Function



Public Sub MyQuery()

  Dim db As DAO.Database
  Dim tdf As DAO.TableDef
  Dim fd As DAO.Field
  
  Dim fso As Object
  Dim oFile As Object

  Set db = CurrentDb()
  Set objRecordSet = CreateObject("ADODB.Recordset")
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set oFile = fso.CreateTextFile("C:\Users\eyw7\OneDrive - CDC\My Documents\OIDM\MMP Database Specifications\MMP_db\db_clouds\test2.txt")
  'oFile.WriteLine "test"

  'Open "C:\Users\eyw7\OneDrive - CDC\My Documents\OIDM\MMP Database Specifications\MMP_db\db_clouds\test.txt" For Output As dataFil
  'Print dataFile, "This is a test"
  
  'only print headers in first row of file
  Debug.Print "Table | Name | DataType | DataSize | OrdinalPosition"
  oFile.WriteLine "Table | Name | DataType | DataSize | OrdinalPosition"
  
  For Each tdf In db.TableDefs
    If Not (tdf.Name Like "MSys*") Then
        For Each fd In tdf.Fields
          Debug.Print tdf.Name & " | " & fd.Name & " | " & GetFieldType(fd) & " | " & fd.Size & " | " & fd.OrdinalPosition
          oFile.WriteLine tdf.Name & " | " & fd.Name & " | " & GetFieldType(fd) & " | " & fd.Size & " | " & fd.OrdinalPosition
          'Debug.Print vbNewLine
        Next fd
    End If
    'Debug.Print vbNewLine
    'oFile.WriteLine vbNewLine
    
    'Set fd = tdf.Fields
  Next tdf
  
  oFile.Close
  
  Set tdf = Nothing
  Set db = Nothing
  Set fso = Nothing
  Set oFile = Nothing
  
  'For Each qdf In db.QueryDefs
    'Debug.Print qdf.sql
  'Next qdf
  'Set qdf = Nothing
  'Set db = Nothing


End Sub

Public Sub Data_Dictionary()

    Const adSchemaTables = 60
    Const adSchemaColumns = 4
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")
    
    objConnection.Open _
        "Provider = Microsoft.ACE.OLEDB.12.0; " & _
            "Data Source = ""C:\Users\eyw7\OneDrive - CDC\Desktop\Database1.accdb"""
                  
  'objConnection.Open _
    '"Provider = Microsoft.Jet.OLEDB.4.0; " & _
        '"Data Source = 'C:\Scripts\Test.mdb'"
            
            
    Set objRecordSet = objConnection.OpenSchema(adSchemaTables)
    Do Until objRecordSet.EOF
        strTableName = objRecordSet("Table_Name")
        Set objFieldSchema = objConnection.OpenSchema(adSchemaColumns, _
            Array(Null, Null, strTableName))
        Wscript.Echo UCase(objRecordSet("Table_Name"))
        Do While Not objFieldSchema.EOF
            Wscript.Echo objFieldSchema("Column_Name") & ", " & objFieldSchema("Data_Type")
            objFieldSchema.MoveNext
        Loop
        Wscript.Echo
        objRecordSet.MoveNext
    Loop

End Sub