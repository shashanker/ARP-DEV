//To insert records into sfdc when the lead uploads a trix file
function insertRecord_sfdc(dynamicProgramValue, compny_type, pod, cid, proposed_company_name, oppurt, website, rows, inputrixpathidd, progname, batchId, countOfrowsInserted, startordone)
{
	var returnStatus = new Array();
	returnStatus[3] = "inProgress";
	returnStatus[0] = countOfrowsInserted;
	returnStatus[2] = " ";
	var message = '';
	var conn;
	var date_time = arp_makeTimestamp();
	var cidNum = parseInt(cid);
	var sheet = SpreadsheetApp.openById(inputrixpathidd);
	var sheet1 = sheet.getSheets()[1];
	var starttime = "Start time: " + date_time;
	var totalRecords = "Total records: " + rows;
	var errorMessages = "Error occurred record:";
	var insertionEndTime = "End time: " + date_time;
	if (startordone == START)
	{
		sheet1.appendRow([starttime]);
		sheet1.appendRow([totalRecords]);
		sheet1.appendRow([errorMessages]);
	}
	if (startordone == SINGLEROW)
	{
		sheet1.appendRow([starttime]);
		sheet1.appendRow([totalRecords]);
		sheet1.appendRow([errorMessages]);
	}
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn
				.prepareStatement("INSERT INTO sfdc (program_batch, received_on, company_type, pod, cid, proposed_company_name, oppurtunity_name, website, status) VALUES (?,?,?,?,?,?,?,?,?)");
		stmt.setString(1, dynamicProgramValue);
        stmt.setString(2, date_time);
		stmt.setString(3, compny_type);
		stmt.setString(4, pod);
		stmt.setInt(5, cidNum);
		stmt.setString(6, proposed_company_name);
		stmt.setString(7, oppurt);
		stmt.setString(8, website);
		stmt.setString(9, UNASSIGNED);
		stmt.execute();
		message = SUCCESS_MESSAGE;
		SUCCESS_TRANSACTION++;
      Logger.log('SUCCESS_TRANSACTION from fo'+SUCCESS_TRANSACTION)
	}
	catch (e)
	{
		Logger.log(e);
		var exceptio = e;
		sheet1.appendRow([exceptio]);
		message = FAILURE_MESSAGE;
		returnStatus[2] = SpreadsheetApp.openById(inputrixpathidd).getUrl();
	}
	finally
	{
		if (startordone == DONE)
		{
			//var updatedRecords = "Total Uploaded records:" + SUCCESS_TRANSACTION;
			returnStatus[3] = DONE;
			//sheet1.appendRow([updatedRecords]);
			sheet1.appendRow([insertionEndTime]);
		}
		if (startordone == SINGLEROW)
		{
			//var updatedRecords = "Total Uploaded records:	" + SUCCESS_TRANSACTION;
			returnStatus[3] = DONE;
			//sheet1.appendRow([updatedRecords]);
			sheet1.appendRow([insertionEndTime]);
		}
		conn.close();
	}
	returnStatus[1] = message;
	return returnStatus;
}
// To insert a record into upload_history table when a trix is uploaded first to
// get the last batchid +1
function getLastProgramIdex_fromhistory(ProgramSelectedGeneric)
{
  var searchOrModifyLDAPQuery = Utilities.formatString("SELECT ((batch_id) + 1) FROM upload_history WHERE program = '"+ ProgramSelectedGeneric +"' ORDER BY batch_id DESC LIMIT 1 ");
  var resultSet = getQueryResultSet(searchOrModifyLDAPQuery);
  if (resultSet.records == '')
  {
    returnLastProgramId = ProgramSelectedGeneric + "_1";
  }
  else
  {
    returnLastProgramId = ProgramSelectedGeneric + "_" + resultSet.records;
  }
  return returnLastProgramId; 
}
// insert a record into upload_history when a trix is uploaded
function insertRecord_userHistory(lastProgramName, lastprogbatchid, trixpathid)
{
	var conn;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt;
		stmt = conn.prepareStatement("INSERT INTO upload_history(program,trix_id, batch_id, ldap, uploaded_on) VALUES (?,?, ?, ?,?);");
		stmt.setString(1, lastProgramName);
		stmt.setString(2, trixpathid);
		stmt.setString(3, lastprogbatchid);
		stmt.setString(4, USER_NAME);
		stmt.setString(5, arpTime);
		stmt.execute();
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
// To log the details into a inputtrix sheet when a trix is uploaded
function insertUploadTicketsLog(inputrixpathidd, countOfrowsInserted, completed)
{
	var sheet = SpreadsheetApp.openById(inputrixpathidd);
	var sheet1 = sheet.getSheets()[1];
	var insertionEndTime = "End time: " + arpTime;
	var updateTickets = "Updated Tickets: ";
	sheet1.appendRow([updateTickets]);
	sheet1.appendRow([insertionEndTime]);
}