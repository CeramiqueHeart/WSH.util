// search_computer_account_creator.js

var objRootDSE;
var baseDN;

if (WScript.Arguments.count() == 0) {
  WScript.Echo("Usage: " + WScript.ScriptName + " <username1> [username2 ...]");
  WScript.Quit(1);
}

try {
  objRootDSE = GetObject("LDAP://rootDSE");
  baseDN = objRootDSE.Get("defaultNamingContext");
} catch(e) {
  WScript.Echo("Can't connect LDAP server.");
  WScript.Quit(1);
}

for (var i = 0; i < WScript.Arguments.count(); i++) {
  var username;
  var sid;
  var objRecordSet;
  
  username = WScript.Arguments(i);
  sid      = get_aduser_sid(username);
  if (sid == null) {
    WScript.Echo(username + " is not found in directory.");
    continue;
  }
  
  try {
    objRecordSet = search_computer_account_by_creator_sid(sid);
  } catch(e) {
    WScript.Echo("Can't connect Active Directory server.");
    WScript.Quit(1);
  }
  
  if (objRecordSet.RecordCount > 0) {
    objRecordSet.MoveFirst;
    while (! objRecordSet.EOF) {
      WScript.Echo(objRecordSet.Fields("cn").Value);
      objRecordSet.MoveNext;
    }
  }

  WScript.Echo();
}


function get_aduser_sid(username) {
  var objWMI;
  var result;
  var sid;
  
  objWMI = GetObject("winmgmts:");
  result = objWMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount = False AND Name = '" + username + "'");

  // 'result' should have 1 (exist) or 0 (not exist) item.
  if (result.count == 0) {
    sid = null;
  } else {
    sid = new Enumerator(result).item().SID;
  }
  return sid;
}


function search_computer_account_by_creator_sid(sid) {
  var objConnection;
  var objCommand;

  objConnection = new ActiveXObject("ADODB.Connection");
  objConnection.Provider = "ADsDSOObject";
  objConnection.Open     = "Active Directory Provider";

  objCommand = new ActiveXObject("ADODB.Command");
  objCommand.ActiveConnection = objConnection;
  objCommand.CommandText      = "<LDAP://cn=Computers," + baseDN + ">;(mS-DS-CreatorSID=" + sid + ");cn;Subtree";
  
  return objCommand.Execute;
}
