//Constant Variables for connecting database
var ADDRESS = '173.194.86.135';
var ROOT_PWD = 'test';
var USER = 'testuser';
var USER_PWD = 'testuser';
var DB = 'dev_work_assignment_db';
var ROOT = 'root';
var INSTANCE_URL = 'jdbc:mysql://' + ADDRESS;
var DB_URL = INSTANCE_URL + '/' + DB;
// Constant Variables using for Leadpage/Agent/QA page starts
var USER_ROLE_AGENT = 'Agent';
var USER_ROLE_QA = 'QA';
var USER_ROLE_LEAD = 'Lead';
var UNASSIGNED = 'unassigned';
var WIP_AGENT = 'wip_agent';
var WIP_QA = 'wip_qa';
var QA = 'qa';
var IDLE = 'Idle';
var CLARIFICATION_AGENT = 'clarification_agent';
var CLARIFICATION_QA = 'clarification_qa';
var WIP_AGENTREWORK = 'wip_agent_rework';
var WIP_QAREWORK = 'wip_qa_rework';
var COMPLETED = 'completed';
var PAUSE_STATUS = 'Pause';
var PLAY_STATUS = 'none';
var SUCCESS_TRANSACTION = 0;
// Constant Variables using for Leadpage/Agent/QA page Ends
// global variables used for Leadpage/Agent/QA page starts
var email = Session.getActiveUser().getEmail();
var email_Split = email.split("@");
var USER_NAME = email_Split[0];
var returnLastProgramId = '';
var exportResults = new Array();
var spreadSheetURL;
var role;
var arpTime = arp_makeTimestamp();
var SUCCESS_MESSAGE = 'success';
var FAILURE_MESSAGE = 'fail';
var START = 'start';
var SINGLEROW = 'onlyOne';
var DONE = 'done';
var RELEASE = 'release';
var REASSIGN = 'reassign';

// global variables used for Leadpage/Agent/QA page Ends
// init method
function doGet()
{

    try
    {
      var roleArray = getUserRole();
      role = roleArray[1];
      
      if (role == USER_ROLE_LEAD)
      {
          return HtmlService.createTemplateFromFile('LeadView').evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
      }
      else if (role == USER_ROLE_AGENT || role == USER_ROLE_QA)
      {
          return HtmlService.createTemplateFromFile('AgentQAView').evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
      }
      else
      {
          return HtmlService.createTemplateFromFile('UnknownUserView').evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
      }
    }
    catch (errorMessage)
    {
      //return HtmlService.createTemplateFromFile('UnknownUserView').evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
    }
    finally
    {
       
    }

}
// to include other files in html
function include(filename)
{
	return HtmlService.createHtmlOutputFromFile(filename).setSandboxMode(HtmlService.SandboxMode.NATIVE).getContent();
}

// to get system timestamp
function arp_makeTimestamp()
{
	var date = new Date();
	var yyyy = date.getFullYear();
	var mm = date.getMonth() + 1;
	var dd = date.getDate();
	var hh = date.getHours();
	var min = date.getMinutes();
	var ss = date.getSeconds();
	var mysqlDateTime = yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + min + ':' + ss;
 
	return mysqlDateTime;
}

// to get distinct programs based on user LDAP
function getProgramBatch(userName, programsAllowed)
{
    var allowedPrograms = programsAllowed.split(",");
    var programBatchQuery = Utilities.formatString("SELECT distinct program_batch from sfdc where status <> 'completed' and program_batch regexp '(%s)_[0-9]+' order by program_batch", allowedPrograms.join('|'));
	var resultSet = getQueryResultSet(programBatchQuery);
    return resultSet.records;
}

// For Loading Respective screens and getting user related info used in
// Lead/Agent/QA screens
function getUserRole()
{
	var userDetailsQuery = Utilities.formatString("SELECT role, programs_allowed from users WHERE ldap = '%s' limit 1", USER_NAME);
	var resultSet = getQueryResultSet(userDetailsQuery);
	var userRecord = resultSet.records[0].split("~");
	
	var array = new Array(3);
	array[0] = USER_NAME;
	array[1] = userRecord[0];
    array[2] = userRecord[1];

    return array;
}

// To get Ticket Related Info in Agent/QA
function getUserTicketCount(userName, userRole)
{
    var ticketCountQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT count(*) from sfdc WHERE ldap_qa='%s'", userName) : Utilities.formatString("SELECT count(*) from sfdc WHERE ldap_agent='%s'", userName);
	var resultSet = getQueryResultSet(ticketCountQuery);
    return resultSet.records[0];
}

// to print date in Lead Report screen/Agent/QA screen
function getDate()
{
	var date = new Date();
	var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
	var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
	var result = Utilities.formatString("%s, %s %s, %s", days[date.getDay()], months[date.getMonth()], date.getDate(), date.getFullYear());

	return result;
}

// updating sfdc relating to Agent/QA screen
function userAssigning(userName, rowNum, userRole)
{
    var conn;
    var stmt;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		if (userRole == USER_ROLE_AGENT)
		{
			stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=?, agent_start_time=? WHERE id = ?');
			stmt.setString(1, WIP_AGENT);
			stmt.setString(2, userName);
			stmt.setString(3, arpTime);
			stmt.setString(4, rowNum);
		}
		else
		{
			stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=?, qa_start_time=? WHERE id = ?');
			stmt.setString(1, WIP_QA);
			stmt.setString(2, userName);
			stmt.setString(3, arpTime);
			stmt.setString(4, rowNum);
		}
		stmt.execute();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
        stmt.close();
		conn.close();
	}
}
