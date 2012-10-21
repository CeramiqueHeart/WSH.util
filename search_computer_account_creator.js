var objRootDSE;
var baseDN;
var objConnection;
var objCommand;

if (WScript.Arguments.count() == 0) {
  WScript.Echo("Usage: " + WScript.ScriptName + " <username1> [username2 ...]");
  WScript.Quit();
}

try {
  objRootDSE = GetObject("LDAP://rootDSE");
  baseDN = objRootDSE.Get("defaultNamingContext");
} catch(e) {
  WScript.Echo("Can't connect LDAP server.");
  WScript.Quit();
}

objConnection = new ActiveXObject("ADODB.Connection");
objConnection.Provider = "ADsDSOObject";
objConnection.Open     = "Active Directory Provider";

objCommand = new ActiveXObject("ADODB.Command");
objCommand.ActiveConnection = objConnection;

for (var i = 0; i < WScript.Arguments.count(); i++) {
  search_computer_account(WScript.Arguments(i));
  WScript.Echo();
}


function search_computer_account(username) {
  var objWMI;
  var result;
  var objRecordSet;
  
  objWMI = GetObject("winmgmts:");
  result = objWMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount = False AND Name = '" + username + "'");

  if (result.count == 0) {
    WScript.Echo (username + " is not found in directory.");
    return;
  } else {
    var e = new Enumerator(result);
    var sid = e.item().SID;
  }

  objCommand.CommandText = "<LDAP://cn=Computers," + baseDN + ">;(mS-DS-CreatorSID=" + sid + ");cn;Subtree";
  objRecordSet = objCommand.Execute;

  WScript.Echo(username + " (" + sid + ") joined " + objRecordSet.RecordCount + " computer(s).");

  if (objRecordSet.RecordCount > 0) {
    objRecordSet.MoveFirst;
    while (! objRecordSet.EOF) {
      WScript.Echo(objRecordSet.Fields("cn").Value);
      objRecordSet.MoveNext;
    }
  }
}
