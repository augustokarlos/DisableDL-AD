
'Constante que define inicio da Leitura do arquivo
'
Const ForReading = 1

'Variavel que define o caminho do arquivo
'
strCSV = "c:\userlist.csv"

strCSV2 = "c:\userlistdisable.csv"
'Instancia objeto para abertura do arquivo com a lista de usuários
'
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strCSV,ForReading)
'Instancia objeto para gravação do log de usuários desabilitados
Set objFSO2 = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile(strCSV2,true)

'Função IsInArray para checar se uma string está no array
'Será utilizada para checar se usuários do Dominio estão na lista de usuários ativos

Function IsInArray(strIn, arrCheck)
bFlag  = False
 
    If IsArray(arrCheck) AND Not IsNull(strIn) Then
        Dim i
        For i = 0 to UBound(arrCheck)
            If LCase(arrcheck(i)) = LCase(strIn) Then
                bFlag = True
                Exit For
            End If
        Next
    End If
    IsInArray = bFlag
End Function

' Get OU
'
strOU = "DC=seudominio,DC=local"

' Create connection to AD
'
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

' Create command
'
Set objCommand = CreateObject("ADODB.Command")

objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000

' Execute command to get all users in OU
' Filtrando somente objetos do ad de usuários ativos com excessão do administrador
objCommand.CommandText = _
  "<LDAP://" & strOU & ">;" & _
  "(&(objectclass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!sAMAccountName=administrador));" & _
  "description,distinguishedname,sAMAccountName;subtree"
Set objRecordSet = objCommand.Execute



'Array instanciado para armazenar os cpf de usuários consultados na lista de funcionarios ativos 
a=Array()
ReDim Preserve a(0)
'Cria uma exceção para usuários que tem o cpf 99999999999
ReDim Preserve a(UBound(a)+1) 
a(UBOUND(a)) = "99999999999"

' Percorre o resultado da consulta do arquivo de funcionarios ativos
Do Until objFile.AtEndOfStream

line=objFile.ReadLine
'Separa somente a primeira coluna antes separada por ,
arrLine = Split(line,",")
account = arrLine(0) 
'Adiciona os usuários do arquivo no array para posteriormente serem comparados com usuários do AD
'Obs  a função string abaixo faz o strpad adicionando 0  a esquerda quando o cpf tem menos de 11 digitos
ReDim Preserve a(UBound(a)+1) 
a(UBOUND(a)) = String(11 - Len(CStr(account)), "0") & CStr(account)
	
Loop


'Percorre o resultado da consulta do Ldap
Do Until objRecordSet.EOF

strDesc = objRecordSet.Fields("description").Value
descricao = ""
'O campo description é retornado em forma de array, por isso deve ser convertido em string
If IsNull(strDesc) Then
descricao = "none"
Else
For Each StrLine In strDesc
descricao =+  StrLine 
Next
End if

'WScript.Echo descricao + objRecordSet.Fields("sAMAccountName").Value
'Obs  a função string abaixo faz o strpad adicionando 0  a esquerda quando o cpf tem menos de 11 digitos
descricao = String(11 - Len(CStr(descricao)), "0") & CStr(descricao)
If IsInArray(descricao,a) Then
		'objLog.writeline objRecordSet.Fields("sAMAccountName").Value + " ok"
	Else
		objLog.writeline objRecordSet.Fields("sAMAccountName").Value + " desabilitado"
	End If



objRecordSet.MoveNext

Loop





objFile.Close


objRecordSet.Close

Set objRecordSet = Nothing
Set objCommand = Nothing
objConnection.Close
Set objConnection = Nothing
