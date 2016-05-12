'*********************************************************
' LDAPQuery
' Emmanuel Tsouris
' Query AD with an LDAP string.
'
' Reference:
'
'   http://www.msexchange.org/articles/scripting-exchange-vbscript-adsi-part1.html
'*********************************************************

' Could also use GC:// for a global catalog query
Set CNUsers = GetObject ("LDAP://OU=Users,OU=SomeOU,DC=DOMAIN,DC=SOMEPLACE,DC=COM")
CNUsers.Filter = Array("user")
For Each User in CNUsers
     DisplayName = User.displayName
     WScript.Echo DisplayName
Next


