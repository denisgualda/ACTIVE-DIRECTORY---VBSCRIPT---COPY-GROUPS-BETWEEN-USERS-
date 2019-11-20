

strUser1 = InputBox ("Introdueix Usuari Origen (DNI)")
strUser2 = InputBox ("Introdueix Usuari Desti (DNI)")

wscript.echo vbNullString
wscript.echo "*****************************"
wscript.echo "*                           *"
wscript.echo "* COPIA GRUPS ENTRA USUARIS *"
wscript.echo "*                           *"
wscript.echo "*****************************"
wscript.echo vbNullString
wscript.echo " >> Usuari origen:    " & strUser1
wscript.echo " >> Usuari desti:     " & strUser2
wscript.echo vbNullString
wscript.echo "--------------------------------"
wscript.echo vbNullString
wscript.echo "Grups Copiats:"

Set ObjRootDSE = GetObject("LDAP://RootDSE") 
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext")) 
Set ObjRootDSE = Nothing 

'*****************************************************************************************************************************************
'*****************************************************************************************************************************************
'ESBORRA GRUPS USUARI 2         - RECORREGUT PER USUARI DESTÍ
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************
StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & strUser2 & "'" 
 
Set ObjConn = CreateObject("ADODB.Connection") 
ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
Set ObjRS = CreateObject("ADODB.Recordset") 
ObjRS.Open StrSQL, ObjConn 
If Not ObjRS.EOF Then 
    ObjRS.MoveLast:    ObjRS.MoveFirst 
    WScript.Echo vbNullString 
    Set ObjUser = GetObject (Trim(ObjRS.Fields("ADsPath").Value)) 
    Set GroupCollection = ObjUser.Groups 
    For Each ObjGroup In GroupCollection 


            set objSystemInfo = CreateObject("ADSystemInfo") 
            strDomain = objSystemInfo.DomainShortName
            '------------------------------------------------------
                'OBTENIR DISTINGUISHED NAME DEL GRUP D'USUARI AMB LA FUNCIO GetUserDN
                grupc = GetUserDN(objGroup.CN,strDomain)
                dsgrupusuari2 = "LDAP://" & grupc 'CAL AFEGIR LDAP AL DISTINGUESHED NAME
    
                'OBTENIR DISTINGUISHED NAME DEL GRUP D'USUARI AMB LA FUNCIO GetUserDN
                userc = GetUserDN(strUser2,strDomain)
                dsusuari2 = "LDAP://" & userc

            'dsgrupusuari2 = "LDAP://CN=" & objGroup.CN & ",OU=Grups,DC=domini,DC=lab"
            'dsusuari2 = "LDAP://CN=" & StrUserName & ",OU=Usuaris,DC=domini,DC=lab"
    
            'CRIDA LA FUNCIONA PER ESBORRAR GRUP
            removeFromGroup dsusuari2,dsgrupusuari2	
            
            'Escriu per pantalla grups ESBORRATS (objGroup.CN)
            'WScript.Echo "  GRUP ESBORRAT: " &  ObjGroup.CN 
    Next 
        Set ObjGroup = Nothing:    Set GroupCollection = Nothing:    Set ObjUser = Nothing 
    Else 
        WScript.Echo "L'usuari: " & StrUserName & " no s'ha trobat al domini" 
    End If 
    ObjRS.Close:    Set ObjRS = Nothing 
    ObjConn.Close:    Set ObjConn = Nothing 

'*****************************************************************************************************************************************
'*****************************************************************************************************************************************



'*****************************************************************************************************************************************
'*****************************************************************************************************************************************
'RECORREGUT USUARI 1 - LLISTA I COPIA GRUPS D'USUARI ORIGEN A USUARI DESTÍ 
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************
StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & strUser1 & "'" 
 
Set ObjConn = CreateObject("ADODB.Connection") 
ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
Set ObjRS = CreateObject("ADODB.Recordset") 
ObjRS.Open StrSQL, ObjConn 
If Not ObjRS.EOF Then 
    ObjRS.MoveLast:    ObjRS.MoveFirst 
    WScript.Echo vbNullString 
    Set ObjUser = GetObject (Trim(ObjRS.Fields("ADsPath").Value)) 
    Set GroupCollection = ObjUser.Groups 
    For Each ObjGroup In GroupCollection 

    'Escriu per pantalla grups trobat (objGroup.CN)
    'WScript.Echo "  >> " &  ObjGroup.CN
    '*******************************************************************************************

	'*******************************************************************************************	
	'AFEGIM GRUPS D'USUARI1 A USUARI2
        'OBTENIM EL DISTINGUISHED NAME DE GRUPS D'USUARI1 QUE S'OBTENEN DURANT EL RECORREGUT


        '------------------------------------------------------
        'AFEGIT PER UTILIZAR LA FUNCIO: GET DISTINGUESHED NAME
        set objSystemInfo = CreateObject("ADSystemInfo") 
        strDomain = objSystemInfo.DomainShortName
        '------------------------------------------------------

'---> CRIDA LA FUNCIO OBTENIR DISTINGUESHED NAME DE GRUPS USUARI ORIGEN I DISTINGUESHED NAME D'USUARI DESTÍ
            'OBTENIR DISTINGUISHED NAME DEL GRUP D'USUARI AMB LA FUNCIO GetUserDN
            
            'LLISTA GRUPS A AFEGIR
            wscript.echo ">> " & objGroup.CN
            
            grupc = GetUserDN(objGroup.CN,strDomain)
            disgrup = "LDAP://" & grupc 'CAL AFEGIR LDAP AL DISTINGUESHED NAME
            

            'OBTENIR DISTINGUISHED NAME DEL GRUP D'USUARI AMB LA FUNCIO GetUserDN
            userc = GetUserDN(strUser2,strDomain)
            disusuari = userc
            

'---> CRIDA LA FUNCIO AFEGIR GRUPS A USUARI 2    
            addusertogroup disgrup, disusuari


	'strGrup = "LDAP://CN="Grup1,OU=Grups,DC=domini,DC=lab"
    'strUsuari = "CN=Usuari01,OU=Usuaris,DC=domini,DC=lab"
	
	'*******************************************************************************************	

    Next 
    Set ObjGroup = Nothing:    Set GroupCollection = Nothing:    Set ObjUser = Nothing 
    
    WScript.Echo vbNullString 
Else 
    WScript.Echo "L'usuari: " & strUser1 & " no s'ha trobat al domini" 
End If 
ObjRS.Close:    Set ObjRS = Nothing 
ObjConn.Close:    Set ObjConn = Nothing 
wscript.echo "--------------------------------"
wscript.sleep 1000000000


'*******************************************************************************************	
'FUNCIO AFAGEIX USUARI A GRUP

sub addusertogroup (dsgrup,dsusuari)
	Const ADS_PROPERTY_APPEND = 3
	
	Set objGroup = GetObject (dsgrup)
	objGroup.PutEx ADS_PROPERTY_APPEND,"member", Array(dsusuari)
	objGroup.SetInfo
	'wscript.echo "Afegit:" & dsgrup & "a usuari: " & dsusuari
	
    End sub
'*******************************************************************************************	

'*******************************************************************************************	
'FUNCIO OBTENIR DISTINGUESHED NAME
Function GetUserDN(byval strUserName,byval strDomain)

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

End function
'*******************************************************************************************	

'*******************************************************************************************	
'FUNCIO ELIMINA GRUP
sub removeFromGroup(userPath, groupPath)

	dim objGroup
	set objGroup = getobject(groupPath)
	
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			objGroup.Remove(userPath)
			exit sub
		end if
	next
end sub
'**************************************************************
