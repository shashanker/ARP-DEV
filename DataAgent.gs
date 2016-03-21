// function using for loading the Unassigned status Ticket Information of Agent  and qa status ticket.
function getDataFromDb(program, userRole)
{
    var data;
    
    if (userRole == USER_ROLE_QA)
    {
      var unassignedTicketQuery = Utilities.formatString("SELECT sfdc.id, sfdc.program_batch, sfdc.received_on, sfdc.company_type, sfdc.pod, sfdc.cid, sfdc.proposed_company_name, sfdc.oppurtunity_name, sfdc.website, sfdc.status, sfdc.ldap_agent, sfdc.agent_start_time, sfdc.agent_end_time, sfdc.ldap_qa, sfdc.qa_start_time, sfdc.qa_end_time, agent.id, agent.cid, agent.website_from_ics, agent.current_company_rollup, agent.correct_company, agent.correct_company_id, agent.mcc_id, agent.agency_name, agent.require_escalation, agent.parent_on_npl, agent.flag_reason, agent.parent_company, agent.parent_company_in_grm, agent.child_company, agent.child_company_in_grm, agent.division_company, agent.division_company_in_grm, agent.childrdivision_company_id, agent.action_on_mapping, agent.merge_exists, agent.merge_status, agent.source_id, agent.destination_id, agent.merge_flag, agent.clarification_required_agent, agent.remarks_agent from sfdc as sfdc LEFT OUTER JOIN agent as agent ON sfdc.cid = agent.cid where sfdc.program_batch='%s' and sfdc.status='%s' LIMIT 1;", program, QA);
      var resultSet = getQueryResultSet(unassignedTicketQuery);
      data = resultSet.records[0].split("~");
    }
    else
    {
      var rowString = callSPGetRecord(program, USER_NAME);
      data = rowString.split("~");
    }
    
    return data;
}
// function for saving the Agent data into table on clciking submit or Pause
// button
function saveAgentData(cId, websiteIcsId, rollUpId, correctCompNameId, corectCompId, mccId, agencyId, escalationSelection, ParentSelection, flagSelection, parentCompId, lowestParentCompId, childCompId, lowestChildCompId, divCompId, lowestDivCompId, lowestCompId, actionId, sourceId, destId, mergeSelection, clarSelection, remarksId, playStatus, userName, userRole, select_merge_exists, select_merge_status)
{

    var erroString = "No Error";
    var programBatchQuery = Utilities.formatString("SELECT count(*) FROM agent where cid='%s'", cId);
    var resultSet = getQueryResultSet(programBatchQuery);
   
    if(resultSet.records[0] > 0)
	{
		try
		{
			conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
			var stmt = conn
					.prepareStatement('UPDATE agent SET website_from_ics=?, current_company_rollup=?, correct_company=?, correct_company_id=?, mcc_id=?,agency_name=?, require_escalation=?,parent_on_npl=?, flag_reason=?,parent_company=?,parent_company_in_grm=?,child_company=?,child_company_in_grm=?,division_company=?,division_company_in_grm=?,childrdivision_company_id=?,action_on_mapping=?,source_id=?,destination_id=?,merge_flag=?,clarification_required_agent=?,remarks_agent=?,merge_exists=?,merge_status=? WHERE cid = ?');
			stmt.setString(1, websiteIcsId);
			stmt.setString(2, rollUpId);
			stmt.setString(3, correctCompNameId);
			stmt.setString(4, corectCompId);
			stmt.setString(5, mccId);
			stmt.setString(6, agencyId);
			stmt.setString(7, escalationSelection);
			stmt.setString(8, ParentSelection);
			stmt.setString(9, flagSelection);
			stmt.setString(10, parentCompId);
			stmt.setString(11, lowestParentCompId);
			stmt.setString(12, childCompId);
			stmt.setString(13, lowestChildCompId);
			stmt.setString(14, divCompId);
			stmt.setString(15, lowestDivCompId);
			stmt.setString(16, lowestCompId);
			stmt.setString(17, actionId);
			stmt.setString(18, sourceId);
			stmt.setString(19, destId);
			stmt.setString(20, mergeSelection);
			stmt.setString(21, clarSelection);
			stmt.setString(22, remarksId);
            stmt.setString(23, select_merge_exists);
            stmt.setString(24, select_merge_status);
			stmt.setString(25, cId);
			Logger.log(stmt);
			stmt.execute();
            erroString = "No Error";
            if (playStatus == PLAY_STATUS)
            {
              updateStatusonSubmit(cId, userRole, clarSelection);
            }
            else
            {
              updateStatusHistory(cId, playStatus, userName);
            }
		}
		catch (e)
		{
			Logger.log(e);
            erroString = e.toString();
		}
		finally
		{
			conn.close();
		}
	}
	else
	{

		try
		{
			conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
			var stmt = conn
					.prepareStatement("INSERT INTO agent (cid, website_from_ics, current_company_rollup, correct_company, correct_company_id, mcc_id, agency_name, require_escalation, parent_on_npl, flag_reason, parent_company, parent_company_in_grm, child_company, child_company_in_grm, division_company, division_company_in_grm, childrdivision_company_id, action_on_mapping, source_id, destination_id, merge_flag, clarification_required_agent, remarks_agent, merge_exists, merge_status) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
			stmt.setString(1, cId);
			stmt.setString(2, websiteIcsId);
			stmt.setString(3, rollUpId);
			stmt.setString(4, correctCompNameId);
			stmt.setString(5, corectCompId);
			stmt.setString(6, mccId);
			stmt.setString(7, agencyId);
			stmt.setString(8, escalationSelection);
			stmt.setString(9, ParentSelection);
			stmt.setString(10, flagSelection);
			stmt.setString(11, parentCompId);
			stmt.setString(12, lowestParentCompId);
			stmt.setString(13, childCompId);
			stmt.setString(14, lowestChildCompId);
			stmt.setString(15, divCompId);
			stmt.setString(16, lowestDivCompId);
			stmt.setString(17, lowestCompId);
			stmt.setString(18, actionId);
			stmt.setString(19, sourceId);
			stmt.setString(20, destId);
			stmt.setString(21, mergeSelection);
			stmt.setString(22, clarSelection);
			stmt.setString(23, remarksId);
            stmt.setString(24, select_merge_exists);
            stmt.setString(25, select_merge_status);
			Logger.log(stmt);
			stmt.execute();
          
            erroString = "No Error";
            if (playStatus == PLAY_STATUS)
            {
              erroString = updateStatusonSubmit(cId, userRole, clarSelection);
            }
            else
            {
              updateStatusHistory(cId, playStatus, userName);
            }
		}
		catch (e)
		{
			Logger.log(e);
            erroString = e.toString();
		}
		finally
		{
			conn.close();
		}
	}

    return erroString;
}
// function using for updating the status of the ticket after clicking submit
function updateStatusonSubmit(cId, userRole, clarSelection)
{
    var updateStatus;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		if (userRole == USER_ROLE_AGENT)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, agent_end_time=? WHERE cid = ?');
			if (clarSelection == "Yes")
			{
				stmt.setString(1, CLARIFICATION_AGENT);
			}
            else if(clarSelection == WIP_QAREWORK)
			{
				stmt.setString(1, WIP_QAREWORK);
			}
			else
			{
				stmt.setString(1, QA);
			}
			stmt.setString(2, arpTime);
			stmt.setString(3, cId);
		}
		else
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, qa_end_time=? WHERE cid = ?');
			if (clarSelection == "Yes")
			{
				stmt.setString(1, CLARIFICATION_QA);
			}
            else if(clarSelection == WIP_AGENTREWORK)
            {
               stmt.setString(1, WIP_AGENTREWORK);
            } 
			else
			{
				stmt.setString(1, COMPLETED);
			}
			stmt.setString(2, arpTime);
			stmt.setString(3, cId);
         
		}
		stmt.execute();
        updateStatus = "No Error";
	}
	catch (e)
	{
		Logger.log(e);
        updateStatus = "failure:" + e.toString();
	}
	finally
	{
		conn.close();
	}
   return updateStatus;
}
// function using for updating the history table after selecting the Pause button
function updateStatusHistory(cId, playStatus, userName)
{
    var updateStatus;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		if (playStatus == PAUSE_STATUS)
		{
			var stmt = conn.prepareStatement('INSERT INTO history (ldap,cid,status,pause_time) VALUES (?,?,?,?)');
			stmt.setString(1, userName);
			stmt.setString(2, cId);
			stmt.setString(3, WIP_AGENT);
			stmt.setString(4, arpTime);
		}
		else
		{
			var stmt = conn.prepareStatement('UPDATE history SET play_time=? WHERE cid = ? and play_time IS NULL');
			stmt.setString(1, arpTime);
			stmt.setString(2, cId);
		}
		Logger.log(stmt);
		stmt.execute();
        updateStatus="No Error";
	}
	catch (e)
	{
		Logger.log(e);
        updateStatus = "failure:"+e.toString();
	}
	finally
	{
		conn.close();
	}
	updateBreakTime(cId);
    return updateStatus;
}
// function using for updating the break time between play and pause time after clicking play button
function updateBreakTime(cId)
{
  var  updateStatus;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn
				.prepareStatement('UPDATE history SET break_time=MINUTE(TIMEDIFF(play_time,pause_time)) WHERE cid = ? and break_time IS NULL and play_time IS NOT NULL');
		stmt.setString(1, cId);
		Logger.log(stmt);
		stmt.execute();
         updateStatus = "No Error";
	}
	catch (e)
	{
		Logger.log(e);
        updateStatus = "failure:"+e.toString();
	}
	finally
	{
		conn.close();
	}
  
   return updateStatus;

}
// function using for loading the work in progress tickets after clicking getNext button
function loadPendingTicket(userName, userRole, pid)
{
    var pendingTicketQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT cid from sfdc WHERE ldap_qa='%s' and (status='%s' or status='%s' ) and program_batch='%s' LIMIT 1", userName,WIP_QA,WIP_QAREWORK,pid) : Utilities.formatString("SELECT cid from sfdc WHERE ldap_agent='%s' and (status='%s' or status='%s' ) and program_batch='%s' LIMIT 1", userName,WIP_AGENT,WIP_AGENTREWORK,pid); 
    var resultSet = getQueryResultSet(pendingTicketQuery);
  
    if(resultSet.records[0])
    {
      return resultSet.records[0];
    }
    else
    {
      return 0;
    }
 
}

// function using for pulling the work in progress ticket information
function AgentDataJoin(pid, cid, userRole, userName)
{
  var unassignedTicketQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT sfdc.id, sfdc.program_batch, sfdc.received_on, sfdc.company_type, sfdc.pod, sfdc.cid, sfdc.proposed_company_name, sfdc.oppurtunity_name, sfdc.website, sfdc.status, sfdc.ldap_agent, sfdc.agent_start_time, sfdc.agent_end_time, sfdc.ldap_qa, sfdc.qa_start_time, sfdc.qa_end_time, agent.id, agent.cid, agent.website_from_ics, agent.current_company_rollup, agent.correct_company, agent.correct_company_id, agent.mcc_id, agent.agency_name, agent.require_escalation, agent.parent_on_npl, agent.flag_reason, agent.parent_company, agent.parent_company_in_grm, agent.child_company, agent.child_company_in_grm, agent.division_company, agent.division_company_in_grm, agent.childrdivision_company_id, agent.action_on_mapping, agent.merge_exists, agent.merge_status, agent.source_id, agent.destination_id, agent.merge_flag, agent.clarification_required_agent, agent.remarks_agent, qa.id, qa.cid, qa.website_from_ics, qa.website_from_ics_cmt, qa.current_company_rollup, qa.current_company_rollup_cmt, qa.correct_company, qa.correct_company_cmt, qa.correct_company_id, qa.correct_company_id_cmt, qa.mcc_id, qa.mcc_id_cmt, qa.agency_name, qa.agency_name_cmt, qa.require_escalation, qa.require_escalation_cmt, qa.parent_on_npl, qa.parent_on_npl_cmt, qa.flag_reason, qa.flag_reason_cmt, qa.parent_company, qa.parent_company_cmt, qa.parent_company_in_grm, qa.parent_company_in_grm_cmt, qa.child_company, qa.child_company_cmt, qa.child_company_in_grm, qa.child_company_in_grm_cmt, qa.division_company, qa.division_company_cmt, qa.division_company_in_grm, qa.division_company_in_grm_cmt, qa.childrdivision_company_id, qa.childrdivision_company_id_cmt, qa.action_on_mapping, qa.action_on_mapping_cmt, qa.merge_exists, qa.merge_exists_cmt, qa.merge_status, qa.merge_status_cmt, qa.source_id, qa.source_id_cmt, qa.destination_id, qa.destination_id_cmt, qa.merge_flag, qa.merge_flag_cmt, qa.clarification_required_qa, qa.remarks_qa from sfdc AS sfdc LEFT OUTER JOIN agent AS agent on sfdc.cid=agent.cid LEFT OUTER JOIN qa AS qa on agent.cid=qa.cid where sfdc.program_batch='%s' and (sfdc.status='%s' or sfdc.status='%s') and sfdc.ldap_qa='%s' order by sfdc.cid, agent.cid LIMIT 1;",pid,WIP_QA,WIP_QAREWORK,userName) : Utilities.formatString("SELECT sfdc.id, sfdc.program_batch, sfdc.received_on, sfdc.company_type, sfdc.pod, sfdc.cid, sfdc.proposed_company_name, sfdc.oppurtunity_name, sfdc.website, sfdc.status, sfdc.ldap_agent, sfdc.agent_start_time, sfdc.agent_end_time, sfdc.ldap_qa, sfdc.qa_start_time, sfdc.qa_end_time, agent.id, agent.cid, agent.website_from_ics, agent.current_company_rollup, agent.correct_company, agent.correct_company_id, agent.mcc_id, agent.agency_name, agent.require_escalation, agent.parent_on_npl, agent.flag_reason, agent.parent_company, agent.parent_company_in_grm, agent.child_company, agent.child_company_in_grm, agent.division_company, agent.division_company_in_grm, agent.childrdivision_company_id, agent.action_on_mapping, agent.merge_exists, agent.merge_status, agent.source_id, agent.destination_id, agent.merge_flag, agent.clarification_required_agent, agent.remarks_agent, qa.id, qa.cid, qa.website_from_ics, qa.website_from_ics_cmt, qa.current_company_rollup, qa.current_company_rollup_cmt, qa.correct_company, qa.correct_company_cmt, qa.correct_company_id, qa.correct_company_id_cmt, qa.mcc_id, qa.mcc_id_cmt, qa.agency_name, qa.agency_name_cmt, qa.require_escalation, qa.require_escalation_cmt, qa.parent_on_npl, qa.parent_on_npl_cmt, qa.flag_reason, qa.flag_reason_cmt, qa.parent_company, qa.parent_company_cmt, qa.parent_company_in_grm, qa.parent_company_in_grm_cmt, qa.child_company, qa.child_company_cmt, qa.child_company_in_grm, qa.child_company_in_grm_cmt, qa.division_company, qa.division_company_cmt, qa.division_company_in_grm, qa.division_company_in_grm_cmt, qa.childrdivision_company_id, qa.childrdivision_company_id_cmt, qa.action_on_mapping, qa.action_on_mapping_cmt, qa.merge_exists, qa.merge_exists_cmt, qa.merge_status, qa.merge_status_cmt, qa.source_id, qa.source_id_cmt, qa.destination_id, qa.destination_id_cmt, qa.merge_flag, qa.merge_flag_cmt, qa.clarification_required_qa, qa.remarks_qa from sfdc AS sfdc LEFT OUTER JOIN agent AS agent ON sfdc.cid = agent.cid LEFT OUTER JOIN qa AS qa on agent.cid=qa.cid where sfdc.program_batch='%s' and (sfdc.status='%s' or sfdc.status='%s') and sfdc.ldap_agent='%s' LIMIT 1;", pid, WIP_AGENT,WIP_AGENTREWORK,userName);
  var resultSet = getQueryResultSet(unassignedTicketQuery);
  var data = resultSet.records[0].split("~");
  return data;
}

// for updating status of the tickets after click on submit
function updatedTicketCount(userName, userRole, dateTime)
{
  var ticketCountQuery = "";
  var completedTicketQuery ="";
  var clarificationTicketQuery="";
   if (userRole == USER_ROLE_QA)
   {
     ticketCountQuery = Utilities.formatString("select count(*) from sfdc where ldap_qa='%s' and DATE(qa_end_time) = DATE('%s');", userName, arpTime);
     completedTicketQuery = Utilities.formatString("select count(*) from sfdc where ldap_qa='%s' and DATE(qa_end_time) = DATE('%s') and status in ('completed','wip_agent_rework');", userName, arpTime);
     clarificationTicketQuery = Utilities.formatString("select count(*) from sfdc where ldap_qa='%s' and DATE(qa_end_time) = DATE('%s') and status = 'clarification_qa';", userName, arpTime);
   }
   else
   {
     ticketCountQuery = Utilities.formatString("select count(*) from sfdc where ldap_agent='%s' and DATE(agent_end_time) = DATE('%s');", userName, arpTime);
     completedTicketQuery = Utilities.formatString("select count(*) from sfdc where ldap_agent='%s' and DATE(agent_end_time) = DATE('%s') and status in ('completed','qa','wip_qa_rework','wip_qa');", userName, arpTime);
     clarificationTicketQuery = Utilities.formatString("select count(*) from sfdc where ldap_agent='%s' and DATE(agent_end_time) = DATE('%s') and status = 'clarification_agent';", userName, arpTime);
   }

//   var ticketCountQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("select count(*) from sfdc where ldap_qa='%s' and DATE(qa_end_time) = DATE('%s');", userName, dateTime) : Utilities.formatString("select count(*) from sfdc where ldap_agent='%s' and DATE(agent_end_time) = DATE('%s');", userName, dateTime);

   var resultSet = getQueryResultSet(ticketCountQuery);
   var completedCount = getQueryResultSet(completedTicketQuery);
   var clarificationCount = getQueryResultSet(clarificationTicketQuery);
   var result = resultSet.records[0] +"/80   (Completed : "+ completedCount.records[0]+ " , Clarification : "  +clarificationCount.records[0]+")";
//   LogMessage(resultSet.records[0] + "::" + resultSet.records[1]);
   return result;
  
}
// function for updating playtime of wip tickets
function statusinHistory(cid, userName, userRole)
{

    var ticketCountQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT id from history WHERE ldap='%s' and status='%s' and cid='%s' and play_time IS NULL", userName,WIP_QA,cid) : Utilities.formatString("SELECT id from history WHERE ldap='%s' and status='%s' and cid='%s' and play_time IS NULL", userName,WIP_AGENT,cid);
    var resultSet = getQueryResultSet(ticketCountQuery);
  
    if(resultSet.records[0])
    {
      return resultSet.records[0];
    }
    else
    {
      return 0;
    } 
}
// getting number of tickets with status work in progress
function getWipCount(pid, userName, userRole)
{
    var ticketCountQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT count(*) from sfdc WHERE ldap_qa='%s' and (status='%s' OR status = '%s') and program_batch='%s'", userName,WIP_QA,WIP_QAREWORK,pid) : Utilities.formatString("SELECT count(*) from sfdc WHERE ldap_agent='%s' and (status='%s' or status='%s') and program_batch='%s'", userName,WIP_AGENT,WIP_AGENTREWORK,pid);
    var resultSet = getQueryResultSet(ticketCountQuery);
 
    if(resultSet.records[0] > 0)
    {
      return resultSet.records[0];
    }
    else
    {
     return 0;
    }
}

function totalProgramtickets(pid)
{
    var totalProgramQuery = Utilities.formatString("SELECT count(*) FROM sfdc where program_batch='%s'", pid);
	var resultSet = getQueryResultSet(totalProgramQuery);
    
    return resultSet.records[0];
}

function eligibleForQaStatus(pid, ELIGIBLE_FOR_QA, userRole)
{

	var totaltickets = totalProgramtickets(pid);

    var eligibleQAQuery = Utilities.formatString("SELECT count(*) FROM sfdc where (status='%s' or status='%s') and program_batch='%s'", COMPLETED,WIP_QA,pid);
	
    var resultSet = getQueryResultSet(eligibleQAQuery);
    
    var countQa = (ELIGIBLE_FOR_QA/100) * totaltickets;
    
    if(resultSet.records[0] >= countQa)
    {
       return 0; 
    }
    else
    {
       return 1;
    }
}

function unassignedTicketCount(pid, userRole)
{
    var unassignedTicketQuery = (userRole == USER_ROLE_QA) ? Utilities.formatString("SELECT COUNT(*) FROM sfdc where program_batch='%s' and status='%s'", pid,QA) : Utilities.formatString("SELECT COUNT(*) FROM sfdc where program_batch='%s' and status='%s'", pid,UNASSIGNED);
    var resultSet = getQueryResultSet(unassignedTicketQuery);
    return resultSet.records[0];
}

function barchartData(program_batch)
{

	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);

		if (program_batch == "none")
		{
			var query = "SELECT ldap_agent,"
				    +"COUNT(CASE WHEN status='clarification_agent' THEN 1 ELSE NULL END) AS clarification_agent,"
				    +"COUNT(CASE WHEN status='wip_agent' THEN 1 ELSE NULL END) AS wip_agent,"
				    +"COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa," 
				    +"COUNT(CASE WHEN status='wip_qa' THEN 1  ELSE NULL END) AS wip_qa,"
				    +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed_Agent, "
				    +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa, "
				    +"COUNT(CASE WHEN status='wip_agent_rework' THEN 1 ELSE NULL END) AS wip_agent_rework, " 
				    +"COUNT(CASE WHEN status='wip_qa_rework' THEN 1 ELSE NULL END) AS wip_qa_rework, " 
				    +"COUNT(CASE WHEN INSTR(CONCAT(parent_on_npl,parent_company,parent_company_in_grm,child_company,child_company_in_grm,division_company,division_company_in_grm,childrdivision_company_id,action_on_mapping,flag_reason,source_id,destination_id,merge_flag), 'rejected') > 0 THEN 1 ELSE NULL END) AS externalUnits,"
				    +"IFNULL((CONCAT((ROUND((((COUNT(t2.cid) - COUNT(CASE WHEN INSTR(CONCAT(parent_on_npl,parent_company,parent_company_in_grm,child_company,child_company_in_grm,division_company,division_company_in_grm,childrdivision_company_id,action_on_mapping,flag_reason,source_id,destination_id,merge_flag), 'rejected') > 0 THEN 1 ELSE NULL END))/COUNT(t2.cid))*100), 2)), '%')), '0.00%') as 'qaScore %' "
				    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid GROUP BY ldap_agent";

			var stmt = conn.prepareStatement(query);
		}
		else
		{
			var query = "SELECT ldap_agent,"
				    +"COUNT(CASE WHEN status='clarification_agent' THEN 1 ELSE NULL END) AS clarification_agent,"
				    +"COUNT(CASE WHEN status='wip_agent' THEN 1 ELSE NULL END) AS wip_agent,"
				    +"COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa," 
				    +"COUNT(CASE WHEN status='wip_qa' THEN 1  ELSE NULL END) AS wip_qa,"
				    +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed_Agent, "
				    +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa, "
				    +"COUNT(CASE WHEN status='wip_agent_rework' THEN 1 ELSE NULL END) AS wip_agent_rework, " 
				    +"COUNT(CASE WHEN status='wip_qa_rework' THEN 1 ELSE NULL END) AS wip_qa_rework, " 
				    +"COUNT(CASE WHEN INSTR(CONCAT(parent_on_npl,parent_company,parent_company_in_grm,child_company,child_company_in_grm,division_company,division_company_in_grm,childrdivision_company_id,action_on_mapping,flag_reason,source_id,destination_id,merge_flag), 'rejected') > 0 THEN 1 ELSE NULL END) AS externalUnits,"
				    +"IFNULL((CONCAT((ROUND((((COUNT(t2.cid) - COUNT(CASE WHEN INSTR(CONCAT(parent_on_npl,parent_company,parent_company_in_grm,child_company,child_company_in_grm,division_company,division_company_in_grm,childrdivision_company_id,action_on_mapping,flag_reason,source_id,destination_id,merge_flag), 'rejected') > 0 THEN 1 ELSE NULL END))/COUNT(t2.cid))*100), 2)), '%')), '0.00%') as 'qaScore %' "
				    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid WHERE program_batch = ? GROUP BY ldap_agent";

			var stmt = conn.prepareStatement(query);
			stmt.setString(1, program_batch);
		}
		var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
		while (results.next())
		{
			var rowString = '';
			for (var col = 0; col < numCols; col++)
			{
				rowString += results.getString(col + 1) + ',';
			}
			data[i] = rowString;
			i++;
		}
		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	Logger.log("data:" + data);
	return data;

}

function barchartDataQA(program_batch)
{

	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);

		if (program_batch == "none")
		{
			var stmt = conn.prepareStatement("SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa, COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed, COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc where ldap_agent is not null  GROUP BY ldap_qa");
		}
		else
		{
			var stmt = conn.prepareStatement("SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa, COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed, COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc where ldap_agent is not null and program_batch=? GROUP BY ldap_qa");
			stmt.setString(1, program_batch);
		}
		var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
		while (results.next())
		{
			var rowString = '';
			for (var col = 0; col < numCols; col++)
			{
				rowString += results.getString(col + 1) + ',';
			}
			data[i] = rowString;
			i++;
		}
		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	Logger.log("data1:" + data);
	return data;

}

function externalQualityScoreQA()
{

	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn
				.prepareStatement("select ldap_agent, COUNT(CASE WHEN correct_company_id='rejected' THEN 1 ELSE NULL END) AS correct_company_id,COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason,COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id,COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid group by ldap_agent");
		var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
		while (results.next())
		{
			var rowString = '';
			for (var col = 0; col < numCols; col++)
			{
				rowString += results.getString(col + 1) + ',';
			}
			data[i] = rowString;
			i++;
		}
		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	Logger.log("data1:" + data);
	return data;

}

function totalCidCountQA(value)
{
	Logger.log("value from daywise:" + value);
	Logger.log("value from daywise:" + parseInt(value));

	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		if (value == 0)
		{
			var stmt = conn.prepareStatement('select count(*) from qa');
		}
		else
		{
			var stmt = conn.prepareStatement('select count(*) from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid where month(qa_end_time)=?');
			stmt.setString(1, parseInt(value));
		}

		var results1 = stmt.executeQuery();
		var program = new Array();
		var count;
		while (results1.next())
		{
			count = results1.getString("count(*)");

		}
		results1.close();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	Logger.log(count);
	return count;

}

function dashboardData(programPie, month, programWise)
{
	var finalArray = new Array();
	var resultArray = new Array();
	var typeArray = new Array();
	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		Logger.log(programPie);
		if (programPie == "program wise")
		{
			var stmt = conn.prepareStatement('select count(status),status from sfdc group by status');
			typeArray.push('program');
		}
		else if (programPie == "flag wise")
		{
			var query = "SELECT COUNT(CASE WHEN strcmp(a.flag_reason, '[0] - LCS Pod Assignment') = 0 THEN 1 ELSE NULL END) as flag0 ,"
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[1] - Rollup and Name Correct') = 0 THEN 1 ELSE NULL END) as flag1 ,"
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[2] - Company Renamed') = 0 THEN 1 ELSE NULL END) as flag2 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[3] - Company Created & CID Linked') = 0 THEN 1 ELSE NULL END) as flag3 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[4] - Relinked to Existing Company') = 0 THEN 1 ELSE NULL END) as flag4 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[5] - Relinked to Existing Company; Renamed to Follow CHIEF Guideline') = 0 THEN 1 ELSE NULL END) as flag5 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[6] - ICS Not Built Out') = 0 THEN 1 ELSE NULL END) as flag6 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[7] - Invalid CID; CID is MCC') = 0 THEN 1 ELSE NULL END) as flag7 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[8] - Invalid URL') = 0 THEN 1 ELSE NULL END) as flag8 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[9] - Multiple Ads') = 0 THEN 1 ELSE NULL END) as flag9 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[01] - LCS Division Creation Escalation') = 0 THEN 1 ELSE NULL END) as flag01 "
				    +"FROM sfdc s inner join agent a on s.cid = a.cid "
				    var stmt = conn.prepareStatement(query);
			typeArray.push('flag');
		}

		if (month != "none")
		{
			var stmt = conn
				   .prepareStatement("select date(qa_end_time), COUNT(CASE WHEN correct_company_id='rejected' THEN 1 ELSE NULL END) AS correct_company_id,COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason,COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id,COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid where month(qa_end_time)=? group by date(qa_end_time)");
			stmt.setString(1, month);
		}
		if (programWise != "none")
		{
			var stmt = conn
				   .prepareStatement("select program_batch, COUNT(CASE WHEN correct_company_id='rejected' THEN 1 ELSE NULL END) AS correct_company_id,COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason,COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id,COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid group by program_batch");

		}
		Logger.log(query);
		var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var resultSetMetaData = results.getMetaData();
		var numCols = results.getMetaData().getColumnCount();

		while (results.next())
		{
			var rowString = '';
			for (var col = 0; col < numCols; col++)
			{
				rowString += results.getString(col + 1) + ',';
			}
			data[i] = rowString;
			i++;
		}
		for (var i = 0; i < numCols; i++)
		{
			var columnName = resultSetMetaData.getColumnName(i + 1);
			resultArray.push(columnName);

		}
		results.close();

	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}

	finalArray.push(typeArray);
	finalArray.push(resultArray);
	finalArray.push(data);

	return finalArray;

}

function dashboardDataDateRange(programPie,programType,startDate,endDate,ldaps)
{
	Logger.clear();
	var finalArray = new Array();
	var resultArray = new Array();
	var typeArray = new Array();

	if(startDate != ""){startDate =startDate+" 00:00:00";}
	if(endDate != ""){endDate =endDate+" 23:59:59";}

	Logger.log("Execute Query:"+startDate+"....");
	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);

		if (programPie == "program wise") 
		{
			if(ldaps = "empty")
			{
				if(programType == 'select')
				{ 
					if(startDate != "")
						{var query = "select count(status),status from sfdc where agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' group by status";}
					else{ var query = "select count(status),status from sfdc group by status";}

				}
				else
				{
					if(startDate != "")
					{
						var query = "select count(status),status from sfdc where program_batch='"+programType+"' AND agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' group by status";
					}
					else
					{
						var query = "select count(status),status from sfdc where program_batch='"+programType+"' group by status";
					} 
				}
				typeArray.push('program');
			}
			else
			{
				typeArray.push('program');
	//var query = "select count(status),status from sfdc where  agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND ldap_agent IN ("+ldaps+") group by status";
			}
		}
		else if (programPie == "flag wise")
		{

			var query = "SELECT COUNT(CASE WHEN strcmp(a.flag_reason, '[0] - LCS Pod Assignment') = 0 THEN 1 ELSE NULL END) as flag0 ,"
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[1] - Rollup and Name Correct') = 0 THEN 1 ELSE NULL END) as flag1 ,"
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[2] - Company Renamed') = 0 THEN 1 ELSE NULL END) as flag2 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[3] - Company Created & CID Linked') = 0 THEN 1 ELSE NULL END) as flag3 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[4] - Relinked to Existing Company') = 0 THEN 1 ELSE NULL END) as flag4 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[5] - Relinked to Existing Company; Renamed to Follow CHIEF Guideline') = 0 THEN 1 ELSE NULL END) as flag5 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[6] - ICS Not Built Out') = 0 THEN 1 ELSE NULL END) as flag6 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[7] - Invalid CID; CID is MCC') = 0 THEN 1 ELSE NULL END) as flag7 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[8] - Invalid URL') = 0 THEN 1 ELSE NULL END) as flag8 ," 
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[9] - Multiple Ads') = 0 THEN 1 ELSE NULL END) as flag9 ,"
				    +"COUNT(CASE WHEN strcmp(a.flag_reason, '[01] - LCS Division Creation Escalation') = 0 THEN 1 ELSE NULL END) as flag01 "
				    +"FROM sfdc s inner join agent a on s.cid = a.cid " 

				    if(startDate != "")
			{      
				if(programType == 'select'){ query = query +" WHERE agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"'";}
				else{ query = query +" WHERE program_batch='"+programType+"' AND agent_end_time BETWEEN '"+startDate+"' AND '"+endDate+"'";}
			}
				    else
				    {
					    if(programType != 'select'){query = query +" WHERE program_batch='"+programType+"'";}
				    }
			typeArray.push('flag');
		}

		Logger.log(query);

		var stmt = conn.prepareStatement(query);
		var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var resultSetMetaData = results.getMetaData();
		var numCols = results.getMetaData().getColumnCount();

		while (results.next())
		{
			var rowString = '';
			for (var col = 0; col < numCols; col++)
			{
				rowString += results.getString(col + 1) + ',';
			}
			data[i] = rowString;
			i++;
		}

		for (var i = 0; i < numCols; i++)
		{
			var columnName = resultSetMetaData.getColumnName(i + 1);
			resultArray.push(columnName);

		}
		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}

	finalArray.push(typeArray);
	finalArray.push(resultArray);
	finalArray.push(data);

	Logger.log("Ends dashboardDataDateRange");
	return finalArray;

}

function totalCidCountQAAgent(sampleValue)
{
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement('select count(*) from qa T1 LEFT OUTER JOIN sfdc T2 on T1.cid = T2.cid  WHERE T2.ldap_agent=? LIMIT 1');
        stmt.setString(1,sampleValue);
		var results1 = stmt.executeQuery();
		var program = new Array();
		var count;
		while (results1.next())
		{
			count = results1.getString("count(*)");

		}
		results1.close();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	return count;

}

function totalCidCountQAProgram(sampleValue)
{
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement('select count(*) from qa T1 LEFT OUTER JOIN sfdc T2 on T1.cid = T2.cid  WHERE T2.program_batch=? LIMIT 1');
        stmt.setString(1,sampleValue);
		var results1 = stmt.executeQuery();
		var program = new Array();
		var count;
		while (results1.next())
		{
			count = results1.getString("count(*)");

		}
		results1.close();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	return count;

}

function totalCidCountQADaywsie(sampleValue)
{
  Logger.log("totalCidCountQADaywsie :"+sampleValue);
  try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement('select count(*) from qa T1 LEFT OUTER JOIN sfdc T2 on T1.cid = T2.cid  WHERE month(T2.qa_end_time)=month(?) LIMIT 1');
        stmt.setString(1,sampleValue);
		var results1 = stmt.executeQuery();
		var program = new Array();
		var count;
		while (results1.next())
		{
			count = results1.getString("count(*)");

		}
		results1.close();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
	return count;
  
}  

function externalQualityScoreQAAgentWise()
{

	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement("select ldap_agent, COUNT(CASE WHEN parent_on_npl='rejected' THEN 1 ELSE NULL END) AS parent_on_npl,"
                   +"COUNT(CASE WHEN parent_company='rejected' THEN 1 ELSE NULL END) AS parent_company,"
                   +"COUNT(CASE WHEN parent_company_in_grm='rejected' THEN 1 ELSE NULL END) AS parent_company_in_grm,"
                   +"COUNT(CASE WHEN child_company='rejected' THEN 1 ELSE NULL END) AS child_company,"
                   +"COUNT(CASE WHEN child_company_in_grm='rejected' THEN 1 ELSE NULL END) AS child_company_in_grm,"
                   +"COUNT(CASE WHEN division_company='rejected' THEN 1 ELSE NULL END) AS division_company," 
                   +"COUNT(CASE WHEN division_company_in_grm='rejected' THEN 1 ELSE NULL END) AS division_company_in_grm," 
                   +"COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id," 
                   +"COUNT(CASE WHEN action_on_mapping='rejected' THEN 1 ELSE NULL END) AS action_on_mapping,"
                   +"COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason," 
                   +"COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,"
                   +"COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,"
                   +"COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag "
                   +"from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid group by ldap_agent");
        var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
        var rowString = '';
        var sampleValue ;
        var totalCidAgent;
        var cidAgentArray = [];
       
		while (results.next())
		{
            var rowArray = [];
			for (var col = 0; col < numCols; col++)
			{
              rowArray.push(results.getString(col + 1));
              rowString += results.getString(col + 1) + ',';
			}
          
          var sampleValue = rowArray[0];
          totalCidAgent = totalCidCountQAAgent(sampleValue);
          cidAgentArray.push(totalCidAgent);
          rowString = rowString+'&';
		}

		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}

  
  var finalStr = rowString + "Count:" + cidAgentArray.toString();
  Logger.log(finalStr);
  return finalStr;

}


function externalQualityScoreQAProgramWise()
{

	try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement("select program_batch, COUNT(CASE WHEN parent_on_npl='rejected' THEN 1 ELSE NULL END) AS parent_on_npl,"
                   +"COUNT(CASE WHEN parent_company='rejected' THEN 1 ELSE NULL END) AS parent_company,"
                   +"COUNT(CASE WHEN parent_company_in_grm='rejected' THEN 1 ELSE NULL END) AS parent_company_in_grm,"
                   +"COUNT(CASE WHEN child_company='rejected' THEN 1 ELSE NULL END) AS child_company,"
                   +"COUNT(CASE WHEN child_company_in_grm='rejected' THEN 1 ELSE NULL END) AS child_company_in_grm,"
                   +"COUNT(CASE WHEN division_company='rejected' THEN 1 ELSE NULL END) AS division_company," 
                   +"COUNT(CASE WHEN division_company_in_grm='rejected' THEN 1 ELSE NULL END) AS division_company_in_grm," 
                   +"COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id," 
                   +"COUNT(CASE WHEN action_on_mapping='rejected' THEN 1 ELSE NULL END) AS action_on_mapping,"
                   +"COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason," 
                   +"COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,"
                   +"COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,"
                   +"COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag "
                   +"from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid group by program_batch");
        var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
        var rowString = '';
        var sampleValue ;
        var totalCidProgram;
        var cidProgramArray = [];
       
		while (results.next())
		{
            var rowArray = [];
			for (var col = 0; col < numCols; col++)
			{
              rowArray.push(results.getString(col + 1));
              rowString += results.getString(col + 1) + ',';
			}
          
          var sampleValue = rowArray[0];
          totalCidProgram = totalCidCountQAProgram(sampleValue);
          cidProgramArray.push(totalCidProgram);
          rowString = rowString+'&';
		}
		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}

  
  var finalStr = rowString + "Count:" + cidProgramArray.toString();
  Logger.log(finalStr);
  return finalStr;

}

function externalQualityScoreQADayWise(month)
{
  
  Logger.log("externalQualityScoreQADayWise :"+month);
  try
	{

		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var stmt = conn.prepareStatement("select date(qa_end_time), COUNT(CASE WHEN parent_on_npl='rejected' THEN 1 ELSE NULL END) AS parent_on_npl,"
                   +"COUNT(CASE WHEN parent_company='rejected' THEN 1 ELSE NULL END) AS parent_company,"
                   +"COUNT(CASE WHEN parent_company_in_grm='rejected' THEN 1 ELSE NULL END) AS parent_company_in_grm,"
                   +"COUNT(CASE WHEN child_company='rejected' THEN 1 ELSE NULL END) AS child_company,"
                   +"COUNT(CASE WHEN child_company_in_grm='rejected' THEN 1 ELSE NULL END) AS child_company_in_grm,"
                   +"COUNT(CASE WHEN division_company='rejected' THEN 1 ELSE NULL END) AS division_company," 
                   +"COUNT(CASE WHEN division_company_in_grm='rejected' THEN 1 ELSE NULL END) AS division_company_in_grm," 
                   +"COUNT(CASE WHEN childrdivision_company_id='rejected' THEN 1 ELSE NULL END) AS childrdivision_company_id," 
                   +"COUNT(CASE WHEN action_on_mapping='rejected' THEN 1 ELSE NULL END) AS action_on_mapping,"
                   +"COUNT(CASE WHEN flag_reason='rejected' THEN 1 ELSE NULL END) AS flag_reason," 
                   +"COUNT(CASE WHEN source_id='rejected' THEN 1 ELSE NULL END) AS source_id,"
                   +"COUNT(CASE WHEN destination_id='rejected' THEN 1 ELSE NULL END) AS destination_id,"
                   +"COUNT(CASE WHEN merge_flag='rejected' THEN 1 ELSE NULL END) AS merge_flag "
                   +"from qa T1 INNER JOIN sfdc T2 ON T2.cid = T1.cid where month(qa_end_time)=? group by date(qa_end_time)");
        stmt.setString(1,month);
        var data = new Array();
		var i = 0;
		var results = stmt.executeQuery();
		var numCols = results.getMetaData().getColumnCount();
        var rowString = '';
        var sampleValue ;
        var totalCidDaywise;
        var cidDaywiseArray = [];
       
		while (results.next())
		{
            var rowArray = [];
			for (var col = 0; col < numCols; col++)
			{
              rowArray.push(results.getString(col + 1));
              rowString += results.getString(col + 1) + ',';
			}
          
          var sampleValue = rowArray[0];
          totalCidDaywise = totalCidCountQADaywsie(sampleValue);
          cidDaywiseArray.push(totalCidDaywise);
          rowString = rowString+'&';
		}

		results.close();

	}

	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}

  
  var finalStr = rowString + "Count:" +  cidDaywiseArray.toString();
  Logger.log(finalStr);
  return finalStr;

}