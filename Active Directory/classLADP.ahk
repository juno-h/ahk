; LADP control

; 268435456 SAM_GROUP_OBJECT
; 268435457 SAM_NON_SECURITY_GROUP_OBJECT
; 536870912 SAM_ALIAS_OBJECT
; 536870913 SAM_NON_SECURITY_ALIAS_OBJECT
; 805306368 SAM_NORMAL_USER_ACCOUNT
; 805306369 SAM_MACHINE_ACCOUNT
; 805306370 SAM_TRUST_ACCOUNT
; 1073741824 SAM_APP_BASIC_GROUP
; 1073741825 SAM_APP_QUERY_GROUP
; 2147483647 SAM_ACCOUNT_TYPE_MAX


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
                this.DomainR := StrReplace(StrReplace(this.Domain,"DC=",""),",",".")
        }
        GetUserList(limit := "" )
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
        
        
        GetMachineList(limit := "" )
        {                
                ; description           - 설명
                fieldList                    := "SAMAccountName,distinguishedName"
                this.Command.CommandText     := "SELECT " fieldList " From 'LDAP://" this.Domain "' WHERE samAccountType='805306369'" 
                this.Recordset := this.Command.Execute()
                
                array := object()
                i := 0
                Critical on
                loop
                {
                        if (this.Recordset.eof)
                                break
                        if instr(this.Recordset.Fields("distinguishedName").value , limit) and this.Recordset.Fields("distinguishedName").value
                        {
                                try
                                        objADSI	:= ComObjGet("LDAP://" this.Recordset.Fields("distinguishedName").value)
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
                Critical off
                return array
        }
        
                                ; mk .=            objADSI.cn 
                                ;         . "`t" . objADSI.description 
                                ;         . "`t" . this.Recordset.Fields("distinguishedName").value "`n"
        GetGroupList(limit := "" )
        {                
                fieldList                    := "givenName,sn,SAMAccountName,distinguishedName"
                this.Command.CommandText     := "SELECT " fieldList " From 'LDAP://" this.Domain "' WHERE samAccountType='268435456'"   ; GROUP_OBJECT
                this.Recordset := this.Command.Execute()
                array := object()
                i := 0
                loop
                {
                        if (this.Recordset.eof)
                                break
                        distinguishedName := this.Recordset.Fields("distinguishedName").value
                        GroupName := this.Recordset.Fields("SAMAccountName").value
                        if instr(distinguishedName , limit) and distinguishedName
                        {           
                                try
                                        objADSI	:= ComObjGet("LDAP://" distinguishedName)
                                this.Command.CommandText     := "SELECT givenName , distinguishedName , SAMAccountName From 'LDAP://" this.Domain "' WHERE objectCategory='user' and memberof='" . this.Recordset.Fields("distinguishedName").value . "'"
                                this.Recordset_member := this.Command.Execute()
                                tempName := ""
                                loop
                                {
                                        if (this.Recordset_member.eof)
                                                break
                                                
                                        Member_ID      := this.Recordset_member.Fields("SAMAccountName").value
                                        Member_name    := this.Recordset_member.Fields("givenName").value
                                        Member_distinguishedName    := this.Recordset_member.Fields("distinguishedName").value
                                        isobject(array["user", Member_ID]) ? "" :  (array["user", Member_ID] := [])
                                        array["user", Member_ID].push(GroupName)
                                        tempName .= (tempName ? "," : "") . Member_name "(" Member_ID ")"
                                        this.Recordset_member.MoveNext()

                                }
                                ; msgbox % objADSI.member.count
                                ; mk .=            objADSI.cn 
                                ;         . "`t" . objADSI.description 
                                ;         . "`t" . tempName
                                ;         . "`t" . distinguishedName "`n"   

                                ; member 가 없을 시 삭제
                                ; if !tempName ; and !instr(this.Recordset.Fields("distinguishedName").value , "N2020")
                                ; {
                                ;         del_list .= objADSI.cn . " member 없음 " tempName " 삭제 `n"
                                ;         objADSI.DeleteObject(0)
                                ; }

                                objADSI := ""     

                                i += 1
                                loop % this.Recordset.Fields.count
                                {                       
                                       array["Fields" , this.Recordset.Fields[A_index-1].name , i] := this.Recordset.Fields(this.Recordset.Fields[A_index-1].name).value                                       
                                       array["record" , i , A_index] := this.Recordset.Fields[this.Recordset.Fields[A_index-1].name].value 
                                }       
                                data := this.Recordset.Fields("SAMAccountName").value
                                array["ID", data] := distinguishedName
                        }
                        this.Recordset.MoveNext()
                }       
                return array
        }
        addGroupMember(groupdistinguishedName, userdistinguishedName )    
        {
                group	:= ComObjGet("LDAP://" groupdistinguishedName)
                user	:= ComObjGet("LDAP://" userdistinguishedName)
                group.add("LDAP://" userdistinguishedName)
                user.memberof
                return this
        }
        removeGroupmember(groupdistinguishedName, userdistinguishedName )    
        {
                group	:= ComObjGet("LDAP://" groupdistinguishedName)
                group.Remove("LDAP://" userdistinguishedName)
                group   := ""
                return this
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
                        if e.Message in 0x80071392
                                return
                        msgbox % Clipboard := e.Message
                        exit
                }
                return
        }
        AddUser(object)
        {
                ; object.ID
                ; object.name
                ; object.strDN
                ; object.description
                ; object.dept
                description             := object.description ? object.description : "신규등록" A_YYYY "-" A_MM "-" A_DD
                strCLass 	        := "User"
                strNAME 	        := "CN=" . object.name
                objADSI	                := ComObjGet("LDAP://" object.strDN)
                objUser                 := objADSI.create(strCLass , strNAME)
                objUser.put("SAMAccountName" ,  object.ID)
                objUser.givenName       := object.name			; 이름
                objUser.displayName     := object.name		        ; 출력이름
                objUser.userPrincipalName := object.ID . "@" . this.DomainR
                objUser.setInfo
                objUser := ""
                objUser := ComObjGet("LDAP://" strNAME . "," . object.strDN)
                objUser.SetPassword(object.ID)
                objUser.AccountDisabled         := false
                objUser.userAccountControl      := 512
                objUser.description := description
                objUser.setInfo
                
                return this
        }

}




; ##### 조직도 정렬
;   ERP 데이터 기반으로 OU 생성 

SetOrganizationalUnitTree(dept , TopOrganizationalUnit)
{
    Tree    := object()
    loop % dept.record.count()
    {
        Tree["deptName" , dept.record[A_index].2] := (dept.record[A_index].3 = "부사장") ? "[" TopOrganizationalUnit "]"  : dept.record[A_index].3
        Tree["deptName" , dept.record[A_index].2] := StrReplace(Tree["deptName" , dept.record[A_index].2] , "/" , "")       ; / 제거, AD 오류 발생
        Tree["parent"   , dept.record[A_index].2] := dept.record[A_index].4                  
    }
    return Tree
}
