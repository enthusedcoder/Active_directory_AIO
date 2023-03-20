Dim objUser
objArg = Wscript.Arguments.Item(0)
Set objUser = GetObject _
  ("LDAP://" & objArg)

objUser.IsAccountLocked = False
objUser.SetInfo
