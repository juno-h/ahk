; LADP control

Class LDAP
{
        __new()
        {
                this.Connection := ComObjCreate("ADODB.Connection")
                this.Command := ComObjCreate("ADODB.Command")
                this.Recordset := ComObjCreate("ADODB.Recordset")
                this.Domain := ComObjGet("LDAP://rootDSE").Get("defaultNamingContext") 
                this.Connection.Open("Provider=ADsDSOObject;")
                this.Command.ActiveConnection := this.Connection
        }
        GetUserList( limit := "" )
        {                
                fieldList                    := "givenName,sn,displayName,SAMAccountName,mail,telephoneNumber,department,st,employeeID,distinguishedName"
                this.Command.CommandText     := "SELECT " fieldList " From 'LDAP://" this.Domain "' WHERE samAccountType='805306368'" 
                this.Recordset := this.Command.Execute()
                array := object()
                i := 0
                loop
                {
                        
                        if (this.Recordset.eof)
                                break
                        if instr(this.Recordset.Fields("distinguishedName").value , limit) and this.Recordset.Fields("distinguishedName").value
                        {
                                i += 1
                                loop % this.Recordset.Fields.count
                                {                       
                                       array["Fields" , this.Recordset.Fields[A_index-1].name , i] := this.Recordset.Fields(this.Recordset.Fields[A_index-1].name).value                                       
                                       array["record" , i , A_index] := this.Recordset.Fields[this.Recordset.Fields[A_index-1].name].value 
                                }       
                                data := this.Recordset.Fields("SAMAccountName").value
                                array["ID", data] := this.Recordset.Fields("distinguishedName").value
                        }
                        this.Recordset.MoveNext()
                }          
                return array
        }
        GetDomain()
        {
                return ComObjGet("LDAP://rootDSE").Get("defaultNamingContext") 
        }
        MoveHere(UserName , NewOU , OldOU , AccountDisabled , description)
        {        
                Try 
                {            
                        UserName := "CN=" . UserName
                        TagetOU := ComObjGet("LDAP://" NewOU)      
                        objUser := ComObjGet("LDAP://" OldOU)               
                        objUser.AccountDisabled 	:= AccountDisabled
                        objUser.description 	        := description
                        objUser.setinfo
                        TagetOU.MoveHere("LDAP://" OldOU , UserName)           ; 객체 이동
                        return true
                } catch e {
                        msgbox ,16 ,ERROR , % e.Message
                        exit
                }       
        }
        AddOrganizationalUnit(DN , OUname)
        {
                objADSI	:= ComObjGet("LDAP://" DN)
                objOU   := objADSI.create("organizationalUnit" , OUname)
                try
                        objOU.setInfo
                catch e
                {
                        ; 0x80071392 - 개체가 이미 있습니다.
                        ErrorMsg := ""
                        if e.Message in 0x80071392
                                return
                        msgbox % e.Message
                        exit
                }
                return
        }
}
