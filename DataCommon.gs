
function getQueryResultSet(query)
{
	var resultSet = new Object();
	resultSet.columns = "";
	resultSet.records = new Array();
	
	var jdbcConnection;
	var connectionStatement;
	var jdbcResultSet;
	
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		connectionStatement = jdbcConnection.createStatement();
		jdbcResultSet = connectionStatement.executeQuery(query);
		var jdbcResultSetMetaData = jdbcResultSet.getMetaData();
		var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();
		
		for (var i = 0; i < jdbcResultSetMetaDataColumnCount; i++)
		{
			var columnName = jdbcResultSetMetaData.getColumnName(i + 1);
			resultSet.columns += columnName + ",";
		}
        
		while (jdbcResultSet.next())
		{
			var rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
                var repString = replaceTilde(jdbcResultSet.getString(col + 1));
		    	rowString += repString + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	resultSet.records.push(rowString);
	    	}
		}
	}
	catch (exception)
	{
      Logger.log("getQueryResultSet(): %s", exception);
	}
	finally
	{
      try
      {
		connectionStatement.close();
		jdbcResultSet.close();
		jdbcConnection.close();
      }
      catch (closeException)
      {
        Logger.log("getQueryResultSet():closeException: %s", closeException);
        throw Error("getQueryResultSet(): Please verify your query: [[" + query + "]]");
      }
	}
	
	if (resultSet.columns != "")
	{
		resultSet.columns = resultSet.columns.substring(0, resultSet.columns.length - 1);
	}
	
	return resultSet;
}

function callSPGetRecord(programBatch, ldap)
{
    var jdbcConnection;
    var jdbcCallableStatement;
    var jdbcResultSet;
    var rowString;
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
        jdbcCallableStatement = jdbcConnection.prepareCall('CALL sp_get_record(?,?,?)');
        jdbcCallableStatement.setString(1, programBatch); // param_program_batch
        jdbcCallableStatement.setString(2, ldap); // param_ldap
        jdbcCallableStatement.setString(3, arpTime); // param_start_time
        jdbcResultSet = jdbcCallableStatement.executeQuery();
		var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();

		while (jdbcResultSet.next())
		{
			rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
		    	rowString += jdbcResultSet.getString(col + 1) + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	Logger.log("rowString: %s", rowString);
	    	}
		}
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
        jdbcResultSet.close();
        jdbcCallableStatement.close();
		jdbcConnection.close();
	}
    
    return rowString;
}

function iqmsCategorize()
{
    var jdbcConnection;
    var jdbcCallableStatement;
    var jdbcResultSet;
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
        jdbcCallableStatement = jdbcConnection.prepareCall('CALL sp_iqms_categorize(?)');
        jdbcCallableStatement.setString(1, arpTime);
        jdbcResultSet = jdbcCallableStatement.executeQuery();
		var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();

		while (jdbcResultSet.next())
		{
			var rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
		    	rowString += jdbcResultSet.getString(col + 1) + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	Logger.log("rowString: %s", rowString);
	    	}
		}
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
        jdbcResultSet.close();
        jdbcCallableStatement.close();
		jdbcConnection.close();
	}
    
     var email = Session.getActiveUser().getEmail();
     var email_Split = email.split("@");
     var userName = email_Split[0];
     updateIQMSCategoryHistory(userName);
  
    return rowString;
}

function callSP_IQMSGetRecord(programBatch, ldap)
{
    var jdbcConnection;
    var jdbcCallableStatement;
    var jdbcResultSet;
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
        jdbcCallableStatement = jdbcConnection.prepareCall('CALL sp_iqms_get_next_case(?,?,?)');
        jdbcCallableStatement.setString(1, programBatch); // param_program_batch
        jdbcCallableStatement.setString(2, ldap); // param_ldap
        jdbcCallableStatement.setString(3, arpTime); // param_start_time
        jdbcResultSet = jdbcCallableStatement.executeQuery();
		var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();

		while (jdbcResultSet.next())
		{
			var rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
		    	rowString += jdbcResultSet.getString(col + 1) + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	Logger.log("rowString: %s", rowString);
	    	}
		}
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
        jdbcResultSet.close();
        jdbcCallableStatement.close();
		jdbcConnection.close();
	}
    
    return rowString.split("~");
}


function sendNotificationToAgent(agentLdap,cId,remarksComments)
{
  var toList = agentLdap + "@google.com";
  var ccList = USER_NAME + "@google.com";
  var bccList = "msahadevan@google.com";
  var htmlMail = Utilities.formatString("<br/>Hi %s,<br/><br/>Your <b>CID: %s</b> has been rejected.<br/>For reason: %s<br/><br/>Thanks,<br/>%s", agentLdap, cId, remarksComments, USER_NAME);

  try
  {
   MailApp.sendEmail({
     to:toList.toString(), 
     cc:ccList.toString(), 
     bcc:bccList.toString(), 
     subject:"ARP: Your Case Rejected - Automatic Email Notification",
     htmlBody:htmlMail,
     });
  }
  catch (exception)
  {
    Logger.log("sendNotificationToAgent(): %s", exception);
  }
}


function replaceTilde(oriString)
{

  var repString = null;
  if(oriString != null)
  {
  repString = oriString.replace(/~/g,'&#126');

  }
  return repString;
  
}  

function LogMessage(message)
{
 
  var logTrixId = "170J_9v5xpn-Xz19JaFgwnVgeL9Rk8zcIBPtNsbLqcKc";
  var logSheet;
  logSheet = SpreadsheetApp.openById(logTrixId);
  var userType = getUserType(USER_NAME);
  if(userType == 'MTV')
  {
     logSheet.getActiveSheet().appendRow([message]);
  }
  else
  {
    // For Debug Purpose Only
    logSheet.getActiveSheet().appendRow([message]);
  }
 
}

function callSPTest(sp)
{
    var resultSet = new Object();
	resultSet.columns = "";
	resultSet.records = new Array();
  
    var jdbcConnection;
    var jdbcCallableStatement;
    var jdbcResultSet;
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
        jdbcCallableStatement = jdbcConnection.prepareCall('CALL '+sp);
        jdbcResultSet = jdbcCallableStatement.executeQuery();
        var jdbcResultSetMetaData = jdbcResultSet.getMetaData();
        var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();
		
		for (var i = 0; i < jdbcResultSetMetaDataColumnCount; i++)
		{
			var columnName = jdbcResultSetMetaData.getColumnName(i + 1);
			resultSet.columns += columnName + ",";
         
		}
        
		while (jdbcResultSet.next())
		{
			var rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
		    	rowString += jdbcResultSet.getString(col + 1) + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	resultSet.records.push(rowString);
	    	}
		}
	}
	catch (exception)
	{
		Logger.log("callSPTest(): %s", exception);

	}
	finally
	{
      try
      {
		jdbcCallableStatement.close();
		jdbcResultSet.close();
		jdbcConnection.close();
      }
      catch (closeException)
      {
        Logger.log("callSPTest():closeException: %s", closeException);
        throw Error("callSPTest(): [[" + closeException + "]]");
      }
	}
	
	if (resultSet.columns != "")
	{
		resultSet.columns = resultSet.columns.substring(0, resultSet.columns.length - 1);
	}
  
	return resultSet;
}

var QUERY = "normalQuery";
var PROCEDURE = "storedProcedure";

function getDynamicQueryResult(titleQuery,type,tname)
{
  var str = titleQuery.split('*~*');
  var title = str[0];
  var queryString = str[1];
  var arrayString = new Array();
  if(type == QUERY)
  {
    var resultSet = getQueryResultSet(queryString);
  }
  else if(type == PROCEDURE)
  {
     var resultSet = callSPTest(queryString);
  }
  
    arrayString.push(resultSet.columns);
  
    if (resultSet.records != '')
    {
      for(var i=0;i< resultSet.records.length;i++)
      {
        var data = resultSet.records[i].split("~");
        arrayString.push(data);
      }
    }
    else
    {
       arrayString.push("empty");
    }
    arrayString.push(title);
    arrayString.push(tname);
    return arrayString;
}

function getDynamicQueryResultNew(queryString,type,sdate,edate)
{
  Logger.log(type)
  var arrayString = new Array();
  if(type == QUERY)
  {
    var resultSet = getQueryResultSet(queryString);
  }
  else if(type == PROCEDURE)
  {
     var resultSet = callSPTestNew(queryString,sdate,edate);
  }
  
    arrayString.push(resultSet.columns);
  
    if (resultSet.records != '')
    {
      for(var i=0;i< resultSet.records.length;i++)
      {
        var data = resultSet.records[i].split("~");
        arrayString.push(data);
      }
    }
    else
    {
       arrayString.push("empty");
    }
    return arrayString;
}

function callSPTestNew(sp,sdate,edate)
{
    var resultSet = new Object();
	resultSet.columns = "";
	resultSet.records = new Array();
  
    var jdbcConnection;
    var jdbcCallableStatement;
    var jdbcResultSet;
	try
	{
		jdbcConnection = Jdbc.getConnection(DB_URL, USER, USER_PWD);
        jdbcCallableStatement = jdbcConnection.prepareCall('CALL '+sp+'(?,?)');
        jdbcCallableStatement.setString(1,sdate);  
        jdbcCallableStatement.setString(2,edate); 
        jdbcResultSet = jdbcCallableStatement.executeQuery();
        var jdbcResultSetMetaData = jdbcResultSet.getMetaData();
        var jdbcResultSetMetaDataColumnCount = jdbcResultSet.getMetaData().getColumnCount();
		
		for (var i = 0; i < jdbcResultSetMetaDataColumnCount; i++)
		{
			var columnName = jdbcResultSetMetaData.getColumnName(i + 1);
			resultSet.columns += columnName + ",";
         
		}
        
		while (jdbcResultSet.next())
		{
			var rowString = "";
		    for (var col = 0; col < jdbcResultSetMetaDataColumnCount; col++)
		    {
		    	rowString += jdbcResultSet.getString(col + 1) + "~";
		    }
		    
		    if (rowString != "")
	    	{
		    	rowString = rowString.substring(0, rowString.length - 1);
		    	resultSet.records.push(rowString);
	    	}
		}
	}
	catch (exception)
	{
		Logger.log("callSPTestNew(): %s", exception);

	}
	finally
	{
      try
      {
		jdbcCallableStatement.close();
		jdbcResultSet.close();
		jdbcConnection.close();
      }
      catch (closeException)
      {
        Logger.log("callSPTestNew():closeException: %s", closeException);
        throw Error("callSPTestNew(): [[" + closeException + "]]");
      }
	}
	
	if (resultSet.columns != "")
	{
		resultSet.columns = resultSet.columns.substring(0, resultSet.columns.length - 1);
	}
  
	return resultSet;
}

function getDistinctProgramBatch()
{

	var distinctProgramQuery =  Utilities.formatString("SELECT DISTINCT(program) AS program_batch FROM upload_history ORDER BY program");
	var resultSet = getQueryResultSet(distinctProgramQuery);
	Logger.log("resultSet:"+resultSet.records);
	return resultSet.records;
}

function updateIQMSCategoryHistory(ldap)
{
    var conn;
    try
    {
         conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
         conn.setAutoCommit(false);
         var stmt = conn.prepareStatement('INSERT INTO iqms_category_history (ldap,timestamp) VALUES (?,?)');
         stmt.setString(1,ldap);
         stmt.setString(2,arpTime);
         stmt.execute();
         conn.commit();
         conn.close();
      
    }
    catch(e)
    {
      Logger.log("Exception has occured in updateIQMSCategoryHistory(): %s",e);
      throw Error("updateIQMSCategoryHistory() :"+e);
    }
} 

function getUserType(ldap)
{
 
  var userTypeProgramQuery = Utilities.formatString("SELECT usertype FROM `users` WHERE ldap='%s'",ldap);
  var resultSet = getQueryResultSet(userTypeProgramQuery);
  var data = resultSet.records[0].split("~"); 
  return data;                                    
}
  