#Selecionando um funcionário de uma BaseDeDados

public sub getEmployee()
    dim rs as DAO.Recordset
    dim lngEmployeeID as Long 
    dim sSQL as String 
    
    lngEmployeeID = 5 'ID DO FUNCIONÁRIO

    sSQL = "SELECT FirstName, LastName, JobTitle FROM tblEmployess"_
    & "WHERE EmployeeID =" & lngEmployeeID & ";" 'CONSULTA SQL 

    Debug.Print sSQL

    Set rs = CurrentDd.OpenRecordset (sSQL)

    Debug.print rs!FirstName & " " & rs!LastName & " " & rs!JobTitle

    Set rs = Nothing 
end sub 


#Inserindo valores

sub insertEmployee()
    Dim sFName As String
    Dim sSname As String
    Dim sJobTitle As String
    Dim sAddress As String
    Dim sCity As String
    Dim sStateProvince As String
    Dim sZipPostal As String
    Dim sCountryRegion As String
    Dim sSQL As Strng     

    sFName = "Roy"
    sSname = "Greene"
    sJobTitle = "Director"
    sAddress = "123 11th Street"
    sCity = "Los Angeles"
    sStateProvince = "California"
    sZipPostal = "99999"
    sCountryRegion = "USA"

    sSQL = "INSERT INTO tblEmployess ("  &_ 'Instrução para inserir as informações 
    "FirsName, LastName, JobTitle, Address, City, " &_
    "StateProvince, ZipPostal, CountryRegion) " &_
    "VALUES (" &_ 'Inserindo as informações
    "'"  & sFName & "' ," &_
    "'"  & sSname & "' ," &_
    "'"  & sJobTitle & "' ," &_
    "'"  & sAddress & "' ," &_
    "'"  & sCity & "' ," &_
    "'"  & sStateProvince & "' ," &_
    "'"  & sZipPostal & "' ," &_
    "'"  & sCountryRegion & "' ," &_

    Debug.Print sSQL

    CurrentDb.Execute sSQL

end sub 

#Atualizando valores
sub UpdateEmployee()
    Dim sFName As String
    Dim sSname As String
    Dim sJobTitle As String
    Dim sAddress As String
    Dim sCity As String
    Dim sStateProvince As String
    Dim sZipPostal As String
    Dim sCountryRegion As String
    Dim sSQL As Strng     

    lngEmployeeID = 4
    sFName = "Ray"
    sSname = "Brown"
    sJobTitle = "Admin"
    sAddress = "123 11th Street"
    sCity = "New York"
    sZipPostal = "99999"
    sCountryRegion = "USA"

    sSQL = "UPDATE tblEmployees SET " &_
    "FirstName = '" & sFName & "', " &_
    "LastName = '" & sSname & "' ," &_
    "JobTitle= '" & sJobTitle & "' ," &_
    "Address= '" & sAddress & "' ,"  &_
    "City = '" & sCity & "' ,"  &_
    "StateProvince= '" & sStateProvince & "' ,"  &_
    "ZIPPostal= '" & sZipPostal & "' "  &_
    "CountryRegion= '" & sCountryRegion & "' "  &_
    "WHERE EmployeeID=" & lngEmployeeID

    Debug.Print sSql
    
    CurrentDb.Execute sSQL

End Sub


#Deletando 
Public Sub DeleteEmployee()
    Dim lngEmployeeID As Long
    Dim sSQL As String

    lngEmployeeID =3 

    sSql = "DELETE * FROM tblEmployess " &_
        "WHERE EmployeeID " & lngEmployeeID

    Debug.Print sSQL

    CurrentDb.Execute sSQL, dbFailOnError
end Sub 