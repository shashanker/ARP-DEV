var newFileId = "";
function createSpreadSheetTemplate()
{
    var root = DriveApp.getRootFolder();
	var date1 = Utilities.formatDate(new Date(), "IST", "MMddyyyy_HH_mm_ss");
	var createSheetName = 'ARP_Report_' + date1;
	newFileId = SpreadsheetApp.create(createSheetName);
	var sheets = SpreadsheetApp.openById(newFileId.getId());
	var templateSheet = sheets.getSheetByName('Export');
    var sheet1 = sheets.getSheets()[0];
    sheets.insertSheet(1, {template:templateSheet});
    return sheet1;
}

function getCountOfReport(leadProgrm, date, hourAssignment, statusAssignment,fromSubOrExport, ldapsId)
{
   Logger.log("date:"+date+"leadProgrm:"+leadProgrm+"hourAssignment:"+hourAssignment+"statusAssignment:"+statusAssignment+"fromSubOrExport:"+fromSubOrExport+"ldapsId:"+ldapsId);
   var count ="" ;
   if(hourAssignment == "select")
    {
        var startDate = date+" 00:00:00";
        var endDate   = date+" 23:59:59";
    }
    else
    {
        var hour = hourAssignment.split("-");
        var startHour = hour[0].trim()+":00";
        var endHour =   hour[1].trim()+":00";
        var startDate = date+" "+startHour;
        var endDate   = date+" "+endHour;
      
    }
  if(date == "" && statusAssignment == "select" && ldapsId == "select")
  {
    count = getReportCountAllAndProgramWise(leadProgrm);
  }
  else if(date == "" && statusAssignment == "select" && ldapsId != "select")
  {
    count = getReportCountAllAndProgramldapWise(leadProgrm,ldapsId);
  }
  else if(date == "" && statusAssignment != "select" && ldapsId == "select")
  {
    if(leadProgrm == "select")
    {
      count = getALLCountReportStatusWise(statusAssignment);
    }
    else
    {
      count = getReportCountProgramStatus(leadProgrm,statusAssignment);
    }
  }
  else if(date == "" && statusAssignment != "select" && ldapsId != "select")
  {
  
    if(leadProgrm == "select")
    {
      count = getALLCountReportStatusLdapWise(statusAssignment,ldapsId);
    }
    else
    {
      count = getReportCountProgramStatusLdap(leadProgrm,statusAssignment,ldapsId);
    }
  }
  else if(date != "" && statusAssignment == "select" && ldapsId == "select")
  {     
   
        count = getReportCountProgramDate(leadProgrm,startDate,endDate);
  
  }
  else if(date != "" && statusAssignment == "select" && ldapsId != "select")
  {
    
        count = getReportCountProgramDateLdap(leadProgrm,startDate,endDate,ldapsId);
  
  }
  else if(date != "" && statusAssignment != "select" && ldapsId == "select")
  {
  
         if(leadProgrm == "select")
         {
             count = getAllReportCountDateStatus(leadProgrm,startDate,endDate,statusAssignment);
         }
         else
         {
            count = getReportCountProgramDateStatus(leadProgrm,startDate,endDate,statusAssignment);
         }
       
  }
  else if(date != "" && statusAssignment != "select" && ldapsId != "select")
  {
    
         if(leadProgrm == "select")
         {
             count = getAllReportCountDateStatusLdap(leadProgrm,startDate,endDate,statusAssignment,ldapsId);
         }
         else
         {
            count = getReportCountProgramDateStatusLdap(leadProgrm,startDate,endDate,statusAssignment,ldapsId);
         }
       
  }
  return count;
}


function getReportCountAllAndProgramWise(programBatch)//1-->
{
    Logger.log("getReportCountAllAndProgramWise....");
    if(programBatch == "select")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid ";
    }
    else
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"'";
    }
    Logger.log(countQuery);
    var resultSet = getQueryResultSet(countQuery);
    var data = resultSet.records[0].split("~");
  Logger.log("data:"+data);
    return data;
  }
  
function getReportCountAllAndProgramldapWise(programBatch,ldap)//2-->
{
    Logger.log("getReportCountAllAndProgramldapWise....");
   
    if(programBatch == "select")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') ";
    }
    else
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') ";
    }
    Logger.log(countQuery);
    var resultSet = getQueryResultSet(countQuery);
    var data = resultSet.records[0].split("~");
    return data;
} 

function getALLCountReportStatusWise(status)//3-->
{
   Logger.log("getALLCountReportStatusWise....");
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"'";
  }
  else if(status == "wip_agent")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"'";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"'";
  }
  else if(status == "wip_qa")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"'";
  }
  else if(status == "unassigned")
  {
    var countQuery = "select COUNT(*) from sfdc s where status='"+status+"'";
  }
  else if(status == "Idle")
  {
    var countQuery = "select COUNT(*) from sfdc s where status='"+status+"'";
  }
  
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  Logger.log("data:"+data);
  return data;
}

function getReportCountProgramStatus(programBatch,status)//4-->
{
  Logger.log("getReportCountProgramStatus....");
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
  else if(status == "wip_agent")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
  else if(status == "wip_qa")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
  else if(status == "unassigned")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
  else if(status == "Idle")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"'";
  }
    Logger.log(countQuery);
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  return data;
 }
 
 function getALLCountReportStatusLdapWise(status,ldap)//5-->
{
    Logger.log("getALLCountReportStatusLdapWise....");
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "wip_agent")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"'  AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "wip_qa")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "unassigned")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "Idle")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
    Logger.log(countQuery);
   var resultSet = getQueryResultSet(countQuery);
   var data = resultSet.records[0].split("~");
   return data;
  }
  
  
function getReportCountProgramStatusLdap(programBatch,status,ldap)//6-->
{

    Logger.log("getReportCountProgramStatusLdap....");
  
    if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    else if(status == "wip_agent")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    else if(status == "wip_qa")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    else if(status == "unassigned")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    else if(status == "Idle")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
    }
    Logger.log(countQuery);
    var resultSet = getQueryResultSet(countQuery);
    var data = resultSet.records[0].split("~");
    return data;
  
  }


function getReportCountProgramDate(programBatch,startDate,endDate)//7-->
{
 
  Logger.log("getReportCountProgramDate....");
  if(programBatch == "select")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
 
  }
  else
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
    Logger.log(countQuery);
   var resultSet = getQueryResultSet(countQuery);
   var data = resultSet.records[0].split("~");
   return data;
}

function getReportCountProgramDateLdap(programBatch,startDate,endDate,ldap)//8 -->
{
 
  Logger.log("getReportCountProgramDateLdap....");
  
  if(programBatch == "select")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
 
  }
  else
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
  }
    Logger.log(countQuery);
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  return data;
 }
 
 function getAllReportCountDateStatus(programBatch,startDate,endDate,status)//9 -->
{
    Logger.log("getAllReportCountDateStatus....");
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' ";
  }
  else if(status == "wip_agent")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' ";
  }
  else if(status == "wip_qa")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "unassigned")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "Idle")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  
    Logger.log(countQuery);
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  return data;
 }
 
function getReportCountProgramDateStatus(programBatch,startDate,endDate,status)//10 -->
{
    Logger.log("getReportCountProgramDateStatus....");

  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' ";
  }
  else if(status == "wip_agent")
  {
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
   var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' ";
  }
  else if(status == "wip_qa")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "unassigned")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
  else if(status == "Idle")
  {
    var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' "; 
  }
    Logger.log(countQuery);
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  return data;
 }
 
function getAllReportCountDateStatusLdap(programBatch,startDate,endDate,status,ldap)//11-->
{
    Logger.log("getAllReportCountDateStatusLdap....");
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "wip_agent")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')";
  }
  else if(status == "wip_qa")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
  }
   else if(status == "unassigned")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
  }
   else if(status == "Idle")
  {
     var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')"; 
  }
    Logger.log(countQuery);
  var resultSet = getQueryResultSet(countQuery);
  var data = resultSet.records[0].split("~");
  return data;
 }
 
 function getReportCountProgramDateStatusLdap(programBatch,startDate,endDate,status,ldap)//12 -->
 {
  
    Logger.log("getReportCountProgramDateStatusLdap....");
  
    if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') ";
    }
    else if(status == "wip_agent")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') "; 
    }
    else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
	{
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') ";
    }
    else if(status == "wip_qa")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') "; 
    }
    else if(status == "unassigned")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') "; 
    }
    else if(status == "Idle")
    {
      var countQuery = "select COUNT(*) from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') "; 
    }
    Logger.log(countQuery);
    var resultSet = getQueryResultSet(countQuery);
    var data = resultSet.records[0].split("~");
	 return data;
 }

/* Added By Mousumi Hazarika on 27-07-2015 */
function getDailyReport()
{
    Logger.clear();
    Logger.log("getDailyReport....");
//    var url = "";
    var array = new Array();
    var sheet1 = createSpreadSheetTemplate();
    var countQuery = "SELECT COUNT(*) FROM sfdc as mtv LEFT JOIN agent as a ON mtv.cid = a.cid LEFT JOIN qa as q ON a.cid = q.cid where (mtv.status = 'qa' or mtv.status = 'wip_qa' or mtv.status = 'completed') and ((DATE(mtv.agent_end_time) = CURRENT_DATE) or (DATE(mtv.qa_end_time) = CURRENT_DATE))";
  
    var resultSet = getQueryResultSet(countQuery);
    var data = resultSet.records[0].split("~");
  
  Logger.log("Count data:"+data);
  
  if(data != 0)
  {
     var limit = 500;
     var breakLimit = 20;
     var realLimit = 0;
     var modulus = data % limit;
     var noOfLoops = parseInt(data/limit);
     Logger.log("NoOfloops:"+noOfLoops); 
     Logger.log("Modulus:"+modulus); 
  
    if(noOfLoops < breakLimit)
    {
      realLimit = noOfLoops;
    }
    else
    {
     realLimit = breakLimit;
    }
  
   for(var a=0;a <= noOfLoops+1;a++)
   {
     Logger.log("loops:"+a);
     var offset = (limit * a);
  
     Logger.log("offset:"+offset);
     var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
     var stmt = conn.createStatement();
     stmt.setMaxRows(1000);

     var query = " SELECT company_type, pod, mtv.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa FROM sfdc as mtv LEFT JOIN agent as a ON mtv.cid = a.cid LEFT JOIN qa as q ON a.cid = q.cid where (mtv.status = 'qa' or mtv.status = 'wip_qa' or mtv.status = 'completed') and ((DATE(mtv.agent_end_time) = CURRENT_DATE) or (DATE(mtv.qa_end_time) = CURRENT_DATE)) LIMIT "+limit+" OFFSET "+offset+" ";
 
     Logger.log(query);
     stmt = conn.prepareStatement(query);
     var results = stmt.executeQuery();	

     var numCols = results.getMetaData();
     var columnCount =  numCols.getColumnCount();  
     var columnArray = [];
     for (var i = 1; i <= columnCount; i++)
     {
        var colAt = numCols.getColumnName(i);
        columnArray.push(colAt);
     }

     var rowst1 = '';
     while (results.next())
     {
   
       var rowst ='';
       for (var i = 0; i < columnCount; i++)
       {
         rowst += results.getString(i + 1) + '~^~';
       }
        rowst1 += rowst + '#^#';
        
     }
   
     rowst1 = rowst1.replace(/null/g,"");
//     rowst1 = rowst1.replace(/qa/g,"agentCompleted");
//     rowst1 = rowst1.replace(/completed/g,"qaAudited");
     insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
//     Logger.log("rowst1:"+rowst1);
    if(a == realLimit){break;}
  }

    array[0] = newFileId.getUrl();
    Logger.log(array[0]);
//    return array[0];
 }
 else
 {
   array[0] = "No records found.";
//   return array[0];
 }
   return array[0];
   Logger.log(array[0]);
}

/* Export functionality code */
var limit = 500;
var breakLimit = 16;

function getUrlExportLink(leadProgrm, date, hourAssignment, statusAssignment,fromSubOrExport, ldapsId)
{
  Logger.log("date:"+date+"leadProgrm:"+leadProgrm+"hourAssignment:"+hourAssignment+"statusAssignment:"+statusAssignment+"fromSubOrExport:"+fromSubOrExport+"ldapsId:"+ldapsId);
  
  var linkUrl = "" ;
  
   if(hourAssignment == "select")
    {
        var startDate = date+" 00:00:00";
        var endDate   = date+" 23:59:59";
    }
    else
    {
        var hour = hourAssignment.split("-");
        var startHour = hour[0].trim()+":00";
        var endHour =   hour[1].trim()+":00";
        var startDate = date+" "+startHour;
        var endDate   = date+" "+endHour;
      
    }
  if(date == "" && statusAssignment == "select" && ldapsId == "select")
  {
    linkUrl = getReportAllAndProgramWise(leadProgrm);
  }
  else if(date == "" && statusAssignment == "select" && ldapsId != "select")
  {
    linkUrl = getReportAllAndProgramldapWise(leadProgrm,ldapsId);
  }
  else if(date == "" && statusAssignment != "select" && ldapsId == "select")
  {
    if(leadProgrm == "select")
    {
      linkUrl = getALLReportStatusWise(statusAssignment);
    }
    else
    {
      linkUrl = getReportProgramStatus(leadProgrm,statusAssignment);
    }
  }
  else if(date == "" && statusAssignment != "select" && ldapsId != "select")
  {
  
    if(leadProgrm == "select")
    {
      linkUrl = getALLReportStatusLdapWise(statusAssignment,ldapsId);
    }
    else
    {
      linkUrl = getReportProgramStatusLdap(leadProgrm,statusAssignment,ldapsId);
    }
  }
  else if(date != "" && statusAssignment == "select" && ldapsId == "select")
  {

        linkUrl = getReportProgramDate(leadProgrm,startDate,endDate);
  
  }
  else if(date != "" && statusAssignment == "select" && ldapsId != "select")
  {
//        var startDate = date+" 00:00:00";
//        var endDate   = date+" 23:59:59";
        linkUrl = getReportProgramDateLdap(leadProgrm,startDate,endDate,ldapsId);
  
  }
  else if(date != "" && statusAssignment != "select" && ldapsId == "select")
  {

         if(leadProgrm == "select")
         {
             linkUrl = getAllReportDateStatus(leadProgrm,startDate,endDate,statusAssignment);
         }
         else
         {
            linkUrl = getReportProgramDateStatus(leadProgrm,startDate,endDate,statusAssignment);
         }
       
  }
  else if(date != "" && statusAssignment != "select" && ldapsId != "select")
  {

         if(leadProgrm == "select")
         {
             linkUrl = getAllReportDateStatusLdap(leadProgrm,startDate,endDate,statusAssignment,ldapsId);
         }
         else
         {
            linkUrl = getReportProgramDateStatusLdap(leadProgrm,startDate,endDate,statusAssignment,ldapsId);
         }
       
  }
  return linkUrl;
}

function insertSpreadSheetData(rowString,columnCount,columnArray,offset,sheet1)
{
    var resB = rowString.split("#^#");
    var latArr = new Array();
    var resC;
 
  for(var k =0;k < resB.length-1;k++)
  {
     latArr[k] = new Array();
     resC = resB[k].split("~^~");
   
     for(var l=0;l< columnCount;l++)
      {
         latArr[k][l]=resC[l];
       
      }
  }
    sheet1.getRange(1, 1, 1,columnArray.length).setValues([columnArray]);
    sheet1.getRange(offset+2,1,latArr.length,latArr[0].length).setValues(latArr);
} 

function getReportAllAndProgramWise(programBatch)//1
{
  Logger.log("getReportAllAndProgramWise....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountAllAndProgramWise(programBatch);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
//  Logger.log("NoOfloops:"+noOfLoops); 
//  Logger.log("Modulus:"+modulus); 
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
//  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  if(programBatch == "select")
  {
  var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' LIMIT "+limit+" OFFSET "+offset+" ";
 }
   
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
   
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
   
   rowst1 = rowst1.replace(/null/g,"");
   insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
   if(a == realLimit){break;}
  }
  
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
  }
  else
  {
    
     url = "No records found.";
     return url;
  }
}

function getReportProgramDate(programBatch,startDate,endDate)//2
{
  Logger.log("getReportProgramDate....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramDate(programBatch,startDate,endDate);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  if(programBatch == "select")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
 
  }
  else
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
 
  }
   
  Logger.log(query); 
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
 
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
   else
  {
    
     url = "No records found.";
     return url;
  }
} 

function getReportProgramDateStatus(programBatch,startDate,endDate,status)//3
{
  Logger.log("getReportProgramDateStatus....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramDateStatus(programBatch,startDate,endDate,status);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  { 
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
  
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}  
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
   }
   else
   {
     url = "No records found.";
     return url;
   
   }
}  

function getReportProgramStatus(programBatch,status)//4
{
  Logger.log("getReportProgramStatus....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramStatus(programBatch,status);
  if(data !=0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
   else if(status == "unassigned")
  { 
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
}
  else if(status == "Idle")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
 } 
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
   }
    
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
   
    }
     url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else
	{
	 url = "No records found.";
     return url;
	}
} 

function getALLReportStatusWise(status)//5
{
  Logger.log("getALLReportStatusWise....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getALLCountReportStatusWise(status);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  { 
   var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
     if(a == realLimit){break;}
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else{
	 url = "No records found.";
     return url;
	}
}  

function getAllReportDateStatus(programBatch,startDate,endDate,status)//6
{
  Logger.log("getAllReportDateStatus....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getAllReportCountDateStatus(programBatch,startDate,endDate,status);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
//  Logger.log("NoOfloops:"+noOfLoops); 
//  Logger.log("Modulus:"+modulus); 
//  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND "+endDate+" LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  { 
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' LIMIT "+limit+" OFFSET "+offset+" ";
  }
   
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
   while (results.next())
   {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
   }
    Logger.log(rowst1);
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else{
	 url = "No records found.";
     return url;
	}
}

function getReportAllAndProgramldapWise(programBatch,ldap)//7
{
  Logger.log("getReportAllAndProgramldapWise....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountAllAndProgramldapWise(programBatch,ldap);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
//  Logger.log("NoOfloops:"+noOfLoops); 
//  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  if(programBatch == "select")
  {
  var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid WHERE (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
 }
   
   Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
   
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
   rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
   if(a == realLimit){break;}
  }
   url = newFileId.getUrl();
    Logger.log(url);
    return url;
  }
  else{
   url = "No records found.";
   return url;
  }
} 

function getALLReportStatusLdapWise(status,ldap)//8
{
  Logger.log("getALLReportStatusLdapWise....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getALLCountReportStatusLdapWise(status,ldap);
  if(data != 0){
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')LIMIT "+limit+" OFFSET "+offset+" "; 
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')LIMIT "+limit+" OFFSET "+offset+" "; 
  }
  else if(status == "Idle")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"')LIMIT "+limit+" OFFSET "+offset+" "; 
  }
   
 Logger.log(query);

  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
     if(a == realLimit){break;}
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else{
	
     url = "No records found.";
     return url;
	}
}  

function getReportProgramStatusLdap(programBatch,status,ldap)//9
{
  Logger.log("getReportProgramStatusLdap....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramStatusLdap(programBatch,status,ldap);
  if(data != 0){
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
   else if(status == "unassigned")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }  
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
    
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
   
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
  }
  else{
     url = "No records found.";
     return url;
  
  }
} 
function getReportProgramDateLdap(programBatch,startDate,endDate,ldap)//10
{
  Logger.log("getReportProgramDateLdap....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramDateLdap(programBatch,startDate,endDate,ldap);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  if(programBatch == "select")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
 
  }
  else
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"'  AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
 
  }
   
  Logger.log(query); 
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
 
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
   
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else{
	 url = "No records found.";
     return url;
	}
} 

function getAllReportDateStatusLdap(programBatch,startDate,endDate,status,ldap)//11
{
  Logger.log("getAllReportDateStatusLdap....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getAllReportCountDateStatusLdap(programBatch,startDate,endDate,status,ldap);
  if(data != 0){
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
  Logger.log("loops:"+a);
  var offset = (limit * a);
  
  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND "+endDate+" AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  {
   var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
   var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
   
   
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
    Logger.log(rowst1);
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
  }
    url = newFileId.getUrl();
    Logger.log(url);
    return url;
	}
	else{
	
	 url = "No records found.";
     return url;
	}
}  

function getReportProgramDateStatusLdap(programBatch,startDate,endDate,status,ldap)//12
{
  Logger.log("getReportProgramDateStatusLdap....");
  var url = "";
  var sheet1 = createSpreadSheetTemplate();
  var data = getReportCountProgramDateStatusLdap(programBatch,startDate,endDate,status,ldap);
  if(data != 0)
  {
  var realLimit = 0;
  var modulus = data % limit;
  var noOfLoops = parseInt(data/limit);
  Logger.log("NoOfloops:"+noOfLoops); 
  Logger.log("Modulus:"+modulus); 
  
   if(noOfLoops < breakLimit)
   {
     realLimit = noOfLoops;
   
   }
   else
   {
     realLimit = breakLimit;
   
   }
  
 for(var a=0;a <= noOfLoops+1;a++)
 {
//  Logger.log("loops:"+a);
  var offset = (limit * a);
  
//  Logger.log("offset:"+offset);
  var  conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
  var stmt = conn.createStatement();
  stmt.setMaxRows(1000);
  
  if(status == "qa" || status == "clarification_agent" || status == "wip_agent_rework")
  {
     var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_agent")
  {
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "completed" || status == "clarification_qa" || status == "wip_qa_rework"){
    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "wip_qa")
  {
      var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "unassigned")
  {
   var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
  else if(status == "Idle")
  {
   var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where program_batch='"+programBatch+"' and status='"+status+"' and qa_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND (ldap_agent='"+ldap+"' OR ldap_qa='"+ldap+"') LIMIT "+limit+" OFFSET "+offset+" ";
  }
   
 
  Logger.log(query);
  stmt = conn.prepareStatement(query);
  var results = stmt.executeQuery();	

  var numCols = results.getMetaData();
  var columnCount =  numCols.getColumnCount();  
  var columnArray = [];
  for (var i = 1; i <= columnCount; i++)
  {
    var colAt = numCols.getColumnName(i);
    columnArray.push(colAt);
  }

  var rowst1 = '';
  while (results.next())
  {
     var output = '';
     var rowst ='';
    for (var i = 0; i < columnCount; i++)
    {
        rowst += results.getString(i + 1) + '~^~';
    }
        rowst1 += rowst + '#^#';
  }
  
    rowst1 = rowst1.replace(/null/g,"");
    insertSpreadSheetData(rowst1,columnCount,columnArray,offset,sheet1);
    if(a == realLimit){break;}
  }
   
    url = newFileId.getUrl();
    Logger.log(url);
//    return url;
	}
	else{
     url = "No records found.";
//     return url;
	}
    return url;
} 

function clearAll()
{
  
  Logger.clear();
}
