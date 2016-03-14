//To add an user 
function insertRecord_Users(first, last, ldap, role, userType, programs_allowed)
{
  Logger.log('programs_allowedof user'+programs_allowed);
  var conn;
  try
  {
    conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
    var insertOrUpdate;
         Logger.log('ldap'+ldap);
    var  searchOrModifyLDAPQuery = Utilities.formatString("SELECT count(*) FROM users where ldap='"+ldap+"'");
    var resultSet = getQueryResultSet(searchOrModifyLDAPQuery);
    insertOrUpdate = resultSet.records;
     Logger.log('insertOrUpdate'+insertOrUpdate);
    if (insertOrUpdate > 0) // update
    {
      var stmt = conn
      .prepareStatement('UPDATE users SET lastname=?, firstname=?, role=?, usertype=?, programs_allowed=? , modified_time=?, modified_ldap=? WHERE ldap = ?');
      stmt.setString(1, last);
      stmt.setString(2, first);
      stmt.setString(3, role);
      stmt.setString(4, userType);
      stmt.setString(5, programs_allowed);
      stmt.setString(6, arpTime);
      stmt.setString(7, USER_NAME);
      stmt.setString(8, ldap);
      stmt.execute();
    }
    else
    {
      var stmt1 = conn
      .prepareStatement('INSERT INTO users (lastname, firstname, ldap, role, usertype, programs_allowed,created_time,created_ldap) VALUES (?,?,?,?,?,?,?,?)');
      stmt1.setString(1, last);
      stmt1.setString(2, first);
      stmt1.setString(3, ldap);
      stmt1.setString(4, role);
      stmt1.setString(5, userType);
      stmt1.setString(6, programs_allowed);
      stmt1.setString(7, arpTime);
      stmt1.setString(8, USER_NAME);
      stmt1.execute();
    }
  }
  catch (e)
  {
    Logger.log(e);
  }
  finally
  {
    conn.close();
  }
}
// to search for an agent
function getModifiedUser(searchLdap)
{
  var searchOrModifyLDAPQuery = Utilities.formatString("SELECT lastname,firstname,ldap,role,usertype,programs_allowed FROM users where ldap='%s'", searchLdap);
  var resultSet = getQueryResultSet(searchOrModifyLDAPQuery);
  return resultSet.records;
}