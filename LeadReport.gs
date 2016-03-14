/*  Lead  Report Screen */
// To get next hour when Hour Filter is applied 
function getNextHour(dateStr)
{
	var nextHour = '';
	var year = +dateStr.substring(0, 4);
	var month = +dateStr.substring(5, 7);
	var day = +dateStr.substring(8, 10);
	var hour = +dateStr.substring(11, 13);
	var pubdate = new Date();
	var date = new Date();
	var yyyy = year;
	var mm = month;
	var dd = day;
	var hh = hour + 1;
	nextHour = yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + "00" + ':' + "00";
	return nextHour;
}
// Main function in Report Screen to display results based on selection
function getLeadResults(leadProgrm, date, selectedHour, statusAssignment, subOrExport, ldapsId)
{
	var conn, stmt;
	var arrayOfLeadResults = new Array();
	var countOfParameters = 0;
	var sys_time = arp_makeTimestamp();
	var hourSelected;
	var hour;
	if (typeof selectedHour != 'undefined')
	{
		hourSelected = selectedHour.split("-");
		hour = hourSelected[0].substring(0, hourSelected[0].length - 1);
	}
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		if(leadProgrm != 'select')
        {  
		    var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
        }
        else
        {
           var query = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid"
        } 
        var queryForIdle;
		if (statusAssignment != 'Idle')
		{
			if (leadProgrm != 'select')
			{
				query = query + " program_batch = " + "'" + leadProgrm + "'";
			}
            else
			{
                 if((ldapsId == 'select') && (statusAssignment == 'select') && (date == ''))
                 {
				 
                 }
                 else
                 {
                    query = query + " WHERE ";
                 } 
             
			}
			if (ldapsId != 'select')
			{
                if(leadProgrm != 'select')
                {
				  if ((statusAssignment == 'wip_agent') || (statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework'))
                  {
                    query = query + " and ldap_agent = " + "'" + ldapsId + "'";
                  }
                  if ((statusAssignment == 'clarification_qa') || (statusAssignment == 'wip_qa') || (statusAssignment == 'wip_qa_rework')) 
                  {
                    query = query + " and ldap_qa = " + "'" + ldapsId + "'";
                  }
                  else
                  {
                    query = query + " and (ldap_agent = " + "'" + ldapsId + "'" + " or ldap_qa = " + "'" + ldapsId + "')";
                  }
                }
              else
              {
                  if ((statusAssignment == 'wip_agent') || (statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework'))
                  {
                    query = query + " ldap_agent = " + "'" + ldapsId + "'";
                  }
                  else if ((statusAssignment == 'clarification_qa') || (statusAssignment == 'wip_qa') || (statusAssignment == 'wip_qa_rework'))
                  {
                    query = query + " ldap_qa = " + "'" + ldapsId + "'";
                  }
                  else
                  {
                    query = query + "(ldap_agent = " + "'" + ldapsId + "'" + " or ldap_qa = " + "'" + ldapsId + "')";
                  }
                
              }
			}
         
           
		   if ((date.length > 1) && (statusAssignment != 'select') && (selectedHour == 'select'))
			{
			if(leadProgrm != 'select')
              {
                if ((date.length > 1) && (statusAssignment == 'wip_agent'))
				{
					var query1 = "select  company_type, pod,s.cid, proposed_company_name,oppurtunity_name, website, status, ldap_agent, agent_start_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where";
					var splitWipAgentQuery = query.split("where")
					query = query1 + splitWipAgentQuery[1] + " and status ='wip_agent' and date(agent_start_time)=" + "'" + date + "'";
				}
				if ((date.length > 1) && (statusAssignment == 'wip_qa'))
				{
					var queryForWIP_qa = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time, a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
					var splitWipQAQuery = query.split("where")
					query = queryForWIP_qa + splitWipQAQuery[1] + " and status ='wip_qa' and date(qa_start_time)=" + "'" + date + "'";
				}
				if ((date.length > 1) && ((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework')))
				{
					var queryForClari_agent1 = "select  company_type, pod,s.cid,  proposed_company_name,oppurtunity_name, website,status, ldap_agent, agent_start_time,agent_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
					var splitclariAgentQuery = query.split("where")
					query = queryForClari_agent1 + splitclariAgentQuery[1] + " and status ='" + statusAssignment + "' and date(agent_end_time)=" + "'" + date + "'";
				}
				if ((date.length > 1) && ((statusAssignment == 'clarification_qa') || (statusAssignment == 'completed') || (statusAssignment == 'wip_qa_rework')))
				{
					var queryForClar_QA = "select  company_type, pod,s.cid, proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
					var splitclariQAQuery = query.split("where")
					query = queryForClar_QA + splitclariQAQuery[1] + " and status ='" + statusAssignment + "' and date(qa_end_time)=" + "'" + date + "'";
				}
				if ((selectedHour.length < 1))
				{
					query = query + " and date(agent_start_time)=" + "'" + date + "'";
				}
				else
				{
					if ((statusAssignment == 'select'))
					{
						var hour = hourSelected[0].substring(0, hourSelected[0].length - 1);
						var date_Hour = date + " " + hour + ":00";
						var nextHr = getNextHour(date_Hour);
						query = query + " and(agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
					}
				}
              }
              else
              {
                
                if(ldapsId != "select")
                {
                  if ((date.length > 1) && (statusAssignment == 'wip_agent'))
                  {
                    query = query + " and status ='wip_agent' and date(agent_start_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && (statusAssignment == 'wip_qa'))
                  {
                    query = query + " and status ='wip_qa' and date(qa_start_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && ((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework')))
                  {
                    query = query + " and status ='" + statusAssignment + "' and date(agent_end_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && ((statusAssignment == 'clarification_qa') || (statusAssignment == 'completed') || (statusAssignment == 'wip_qa_rework')))
                  {
                    query = query + " and status ='" + statusAssignment + "' and date(qa_end_time)=" + "'" + date + "'";
                  }
                  if ((selectedHour.length < 1))
                  {
                    query = query + " and date(agent_start_time)=" + "'" + date + "'";
                  }
                  else
                  {
                    if ((statusAssignment == 'select'))
                    {
                      query = query + " and (agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
                    }
                  }
                
               } 
              else
              {
                 if ((date.length > 1) && (statusAssignment == 'wip_agent'))
                  {
                    query = query + " status ='wip_agent' and date(agent_start_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && (statusAssignment == 'wip_qa'))
                  {
                    query = query + " status ='wip_qa' and date(qa_start_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && ((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework')))
                  {
                    query = query + " status ='" + statusAssignment + "' and date(agent_end_time)=" + "'" + date + "'";
                  }
                  if ((date.length > 1) && ((statusAssignment == 'clarification_qa') || (statusAssignment == 'completed') || (statusAssignment == 'wip_qa_rework')))
                  {
                    query = query + " status ='" + statusAssignment + "' and date(qa_end_time)=" + "'" + date + "'";
                  }
                  if ((selectedHour.length < 1))
                  {
                    query = query + " date(agent_start_time)=" + "'" + date + "'";
                  }
                  else
                  {
                    if ((statusAssignment == 'select'))
                    {
                      query = query + " (agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
                    }
                  }
                }
              }
			}
			if ((statusAssignment != 'select'))
			{
				if ((date.length > 1) && (selectedHour != 'select') && (statusAssignment == 'wip_agent'))
				{
                  if (leadProgrm != 'select')
                  {
					var query1 = "select  company_type, pod,s.cid, proposed_company_name,oppurtunity_name, website, status, ldap_agent, agent_start_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where";
					var date_Hour = date + " " + hour + ":00";
					var nextHr = getNextHour(date_Hour);
					var splitWipAgentQuery = query.split("where")
					query = query1 + splitWipAgentQuery[1] + " and status ='wip_agent'" + " and(agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
                  }
                }
				if ((date.length > 1) && (selectedHour != 'select') && (statusAssignment == 'wip_qa'))
				{
                  if (leadProgrm != 'select')
                  {
					var queryForWIP_qa = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time, a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
					var date_Hour = date + " " + hour + ":00";
					var nextHr = getNextHour(date_Hour);
					var splitWipQAQuery = query.split("where")
					query = queryForWIP_qa + splitWipQAQuery[1] + " and status ='wip_qa' " + " and(qa_start_time >= " + "'" + date_Hour + "') and(qa_start_time <= " + "'" + nextHr + "')";
                  }
                }
				if ((date.length > 1) && (selectedHour != 'select') && ((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa') || (statusAssignment == 'wip_agent_rework')))
				{
                   if (leadProgrm != 'select')
                   {
                     var date_Hour = date + " " + hour + ":00";
                     var nextHr = getNextHour(date_Hour);
                     var queryForClari_agent1 = "select  company_type, pod,s.cid,  proposed_company_name,oppurtunity_name, website,status, ldap_agent, agent_start_time,agent_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
                     var splitclariAgentQuery = query.split("where")
                     query = queryForClari_agent1 + splitclariAgentQuery[1] + " and status ='" + statusAssignment + "' " + " and(agent_end_time >= " + "'" + date_Hour + "') and(agent_end_time <= " + "'" + nextHr + "')";
                   }
                }
				if ((date.length > 1) && (selectedHour != 'select') && ((statusAssignment == 'clarification_qa') || (statusAssignment == 'completed') || (statusAssignment == 'wip_qa_rework')))
				{
					if (leadProgrm != 'select')
                    {
                      var date_Hour = date + " " + hour + ":00";
                      var nextHr = getNextHour(date_Hour);
                      var queryForClar_QA = "select  company_type, pod,s.cid, proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
                      var splitclariQAQuery = query.split("where")
                      query = queryForClar_QA + splitclariQAQuery[1] + " and status ='" + statusAssignment + "' " + " and(qa_end_time >= " + "'" + date_Hour + "') and(qa_end_time <= " + "'" + nextHr + "')";
                    }
                }
				else
				{
                  if (leadProgrm != 'select')
                  {
					if (((statusAssignment == 'wip_agent')) && (date.length < 1))
					{
						var splitCols = query.split("where");
						var queryForWIP_agent = "select  company_type, pod,s.cid, proposed_company_name,oppurtunity_name, website, status, ldap_agent, agent_start_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
						query = queryForWIP_agent + splitCols[1] + " and status=" + "'" + statusAssignment + "'";
					}
					if (((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa')) && (date.length < 1))
					{
						var splitCols = query.split("where");
						var queryForClari_agent1 = "select  company_type, pod,s.cid,  proposed_company_name,oppurtunity_name, website,status, ldap_agent, agent_start_time,agent_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
						query = queryForClari_agent1 + splitCols[1] + " and status =" + "'" + statusAssignment + "'";
					}
					if (statusAssignment == 'unassigned')
					{
						var unassignedQuery = "select  company_type, pod,cid,  proposed_company_name,oppurtunity_name, website  from sfdc where ";
						var splitCols = query.split("where");
						query = unassignedQuery + splitCols[1] + " and status =" + "'" + statusAssignment + "'";;
					}
					if ((statusAssignment == 'wip_qa') && (date.length < 1))
					{
						var splitCols = query.split("where");
						var queryForWIP_qa = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time, a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
						query = queryForWIP_qa + splitCols[1] + " and status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'clarification_qa') && (date.length < 1))
					{
						var splitCols = query.split("where");
						var queryForClar_QA = "select  company_type, pod,s.cid, proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time,qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
						query = queryForClar_QA + splitCols[1] + " and status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'completed') && (date.length < 1))
					{
						var splitCols = query.split("where");
						var queryForWIP_agent = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
						query = queryForWIP_agent + splitCols[1] + " and status = " + "'" + statusAssignment + "'";
					}
                  if ((statusAssignment == 'wip_agent_rework') && (date.length < 1))
                  {
                    var splitCols = query.split("where");
                    var queryForWIP_agent = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
                    query = queryForWIP_agent + splitCols[1] + " and status = " + "'" + statusAssignment + "'";
                  }
                  if ((statusAssignment == 'wip_qa_rework') && (date.length < 1))
                  {
                    var splitCols = query.split("where");
                    var queryForWIP_agent = "select  company_type, pod, s.cid,proposed_company_name, oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where"
                    query = queryForWIP_agent + splitCols[1] + " and status = " + "'" + statusAssignment + "'";
                  }
				}
                else
                {
                 if(ldapsId != 'select')
                 {
                  if (((statusAssignment == 'wip_agent')) && (date.length < 1))
					{
						query = query + " and status=" + "'" + statusAssignment + "'";
					}
					if (((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa')) && (date.length < 1))
					{
						query = query + " and status =" + "'" + statusAssignment + "'";
					}
					if (statusAssignment == 'unassigned')
					{
					    query = query + " and status =" + "'" + statusAssignment + "'";;
					}
					if ((statusAssignment == 'wip_qa') && (date.length < 1))
					{
						query = query + " and status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'clarification_qa') && (date.length < 1))
					{
						query = query + " and status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'completed') && (date.length < 1))
					{
						query = query + " and status = " + "'" + statusAssignment + "'";
					}
                   if ((statusAssignment == 'wip_agent_rework') && (date.length < 1))
                   {
                     query = query + " and status = " + "'" + statusAssignment + "'";
                   }
                   if ((statusAssignment == 'wip_qa_rework') && (date.length < 1))
                   {
                     query = query + " and status = " + "'" + statusAssignment + "'";
                   }
                 }
                 else
                  {
                    if (((statusAssignment == 'wip_agent')) && (date.length < 1))
					{
						query = query + " status=" + "'" + statusAssignment + "'";
					}
					if (((statusAssignment == 'clarification_agent') || (statusAssignment == 'qa')) && (date.length < 1))
					{
						query = query + " status =" + "'" + statusAssignment + "'";
					}
					if (statusAssignment == 'unassigned')
					{
					    query = query + " status =" + "'" + statusAssignment + "'";;
					}
					if ((statusAssignment == 'wip_qa') && (date.length < 1))
					{
						query = query + " status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'clarification_qa') && (date.length < 1))
					{
						query = query + " status = " + "'" + statusAssignment + "'";
					}
					if ((statusAssignment == 'completed') && (date.length < 1))
					{
						query = query + " status = " + "'" + statusAssignment + "'";
					}
                   if ((statusAssignment == 'wip_agent_rework') && (date.length < 1))
                   {
                     query = query + " status = " + "'" + statusAssignment + "'";
                   }
                   if ((statusAssignment == 'wip_qa_rework') && (date.length < 1))
                   {
                     query = query + " status = " + "'" + statusAssignment + "'";
                   }
                    
                  }
                }
                }
                  
			}
            if(leadProgrm != 'select')
            {
              
              if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour != 'select'))
              {
                var hour = hourSelected[0].substring(0, hourSelected[0].length - 1);
                var date_Hour = date + " " + hour + ":00";
                var nextHr = getNextHour(date_Hour);
                query = query + " and(agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
              }
              if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour == 'select'))
              {
                query = query + "  and date(agent_start_time)=" + "'" + date + "'";
              }
            } 
            else
            {
              if(ldapsId != 'select')
              {
                 if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour != 'select'))
                 {
                   var hour = hourSelected[0].substring(0, hourSelected[0].length - 1);
                   var date_Hour = date + " " + hour + ":00";
                   var nextHr = getNextHour(date_Hour);
                   query = query + " and(agent_start_time >= " + "'" + date_Hour + "') and(agent_start_time <= " + "'" + nextHr + "')";
                 }
                if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour == 'select'))
                {
                  query = query + "  and date(agent_start_time)=" + "'" + date + "'";
                }
              }
              else
              {
                if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour != 'select'))
                {
                  var hour = hourSelected[0].substring(0, hourSelected[0].length - 1);
                  var date_Hour = date + " " + hour + ":00";
                  var nextHr = getNextHour(date_Hour);
                  query = query + "agent_start_time >= " + "'" + date_Hour + "' and agent_start_time <= " + "'" + nextHr + "'";
                }
                if ((statusAssignment == 'select') && (date.length > 1) && ( selectedHour == 'select'))
                {
                  query = query + "  date(agent_start_time)=" + "'" + date + "'";
                }
                
              }
              
              
            }
            
          Logger.log('query' + query);
			stmt = conn.prepareStatement(query);
			var results = stmt.executeQuery();
			var j = 1;
			var numCols = results.getMetaData();
			var colnames = '';
			for (var i = 1; i <= numCols.getColumnCount(); i++)
			{
				var colAt = numCols.getColumnName(i);
				colnames = colnames + colAt + ',';
			}
			arrayOfLeadResults[0] = colnames;
			while (results.next())
			{
				var rowst = '';
				for (var i = 1; i <= numCols.getColumnCount(); i++)
				{
					var output = results.getString(i);
					if (output == null)
					{
						rowst += ' ' + ',';
					}
					else
					{
                        output = output.replace(/,/g, "~"); // Encrypting
						rowst += output + ',';
					}
				}
				rowst = rowst.substring(0, rowst.length - 1);
				arrayOfLeadResults[j] = rowst;
				j++;
			}
			results.close();
		}
		else
		{
			if ((statusAssignment == 'Idle') && (leadProgrm != 'select'))
			{
    
    	//queryForIdle = "select  company_type, pod, s.cid, proposed_company_name,oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where  ((status='wip_agent' and "+sys_time+"-agent_start_time>12) or(status='wip_qa' and "+ sys_time+"-qa_start_time>12)) order by a.cid";
				queryForIdle = "select  company_type, pod, s.cid, proposed_company_name,oppurtunity_name, website, status, ldap_agent, agent_start_time, agent_end_time, ldap_qa, qa_start_time, qa_end_time,a.website_from_ics, a.current_company_rollup, a.correct_company, a.correct_company_id, a.mcc_id, a.agency_name, a.require_escalation, a.parent_on_npl, a.flag_reason, a.parent_company, a.parent_company_in_grm, a.child_company, a.child_company_in_grm, a.division_company, a.division_company_in_grm, a.childrdivision_company_id, a.action_on_mapping, a.source_id, a.destination_id, a.clarification_required_agent, a.remarks_agent,q.website_from_ics,q. website_from_ics_cmt,q. current_company_rollup,q. current_company_rollup_cmt,q. correct_company,q. correct_company_cmt,q. correct_company_id,q. correct_company_id_cmt,q. mcc_id,q. mcc_id_cmt,q. agency_name,q. agency_name_cmt,q. require_escalation,q. require_escalation_cmt,q. parent_on_npl,q. parent_on_npl_cmt,q. flag_reason,q. flag_reason_cmt,q. parent_company,q. parent_company_cmt,q. parent_company_in_grm,q. parent_company_in_grm_cmt,q. child_company,q. child_company_cmt,q. child_company_in_grm,q. child_company_in_grm_cmt,q. division_company,q. division_company_cmt,q. division_company_in_grm,q. division_company_in_grm_cmt,q. childrdivision_company_id,q. childrdivision_company_id_cmt,q. action_on_mapping,q. action_on_mapping_cmt,q. source_id,q. source_id_cmt,q. destination_id,q. destination_id_cmt,q. merge_flag,q. merge_flag_cmt,q. clarification_required_qa,q. remarks_qa from sfdc s  left outer join agent a on s.cid=a.cid  left outer join  qa  q on a.cid=q.cid where  ((status='wip_agent' and now()-agent_start_time>12) or(status='wip_qa' and now()-qa_start_time>12)) order by a.cid";
				var splitIdle = queryForIdle.split("where");
				var whereQuery = "program_batch =" + "'" + leadProgrm + "'";
				if (ldapsId != 'select')
				{
					whereQuery = whereQuery + " and ldap_agent = " + "'" + ldapsId + "'";
				}
				queryForIdle = splitIdle[0] + " where " + whereQuery + " and" + splitIdle[1];
              Logger.log('queryForIdle'+queryForIdle)
				stmt = conn.prepareStatement(queryForIdle);
				var results = stmt.executeQuery();
				var j = 1;
				var numCols = results.getMetaData();
				var colnames = '';
				for (var i = 1; i <= numCols.getColumnCount(); i++)
				{
					var colAt = numCols.getColumnName(i);
					colnames = colnames + colAt + ',';
				}
				arrayOfLeadResults[0] = colnames;
				while (results.next())
				{
					var rowst = '';
					for (var i = 0; i < numCols.getColumnCount(); i++)
					{
						var output = results.getString(i + 1);
						if (output == null)
						{
							rowst += ' ' + ',';
						}
						else
						{
							rowst += output + ',';
						}
					}
					rowst = rowst.substring(0, rowst.length - 1);
					arrayOfLeadResults[j] = rowst;
					j++;
				}
				results.close();
			}
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
	return arrayOfLeadResults;
}

// To display distinct programs in Lead Report screen/ Reused in DashBoard
// Screen too
function getDistinctProgramData(userName)
{
  var  searchOrModifyLDAPQuery = Utilities.formatString("SELECT DISTINCT(program_batch) from sfdc order by program_batch");
  var resultSet = getQueryResultSet(searchOrModifyLDAPQuery);
  return resultSet.records;
   Logger.log(resultSet.records);
}

// To Release the ticket in Report Screen
function releaseUserAgent(cid, status)
{
	var conn;
	var status_agent = 'unassigned';
	var status_qa = 'qa';
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		if (status == "wip_agent")
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=?  WHERE cid = ?');
			stmt.setString(1, status_agent);
			stmt.setString(2, "null");
			stmt.setString(3, cid);
		}
		else
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=? WHERE cid = ?');
			stmt.setString(1, status_qa);
			stmt.setString(2, "null");
			stmt.setString(3, cid);
		}
		stmt.execute();
	}
	catch (e)
	{
		Logger.log(e);
	}
	finally
	{
		conn.close();
		// While relasing the ticket validating whether play time is null for
		// the corresponding cid or not
		selectHistoryPlay(cid)
	}
}
// To release the ticket in Report Screen validating whether play time is null
// for the corresponding cid or not if null update the play time
function updateHistoryPlay(cid)
{
	var conn;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var date_time = arp_makeTimestamp();
		var stmt = conn.prepareStatement("update history set  play_time ='" + date_time + "' where cid='" + cid + "'");
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

function selectHistoryPlay(cid)
{
	var conn;
	try
	{
		var count = 0;
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		stmt.setMaxRows(1000);
		var results = stmt.executeQuery("select id,cid from history where play_time is null and cid ='" + 10063 + "'");
		var numCols = results.getMetaData().getColumnCount();
		while (results.next())
		{
			count++;
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
		if (count > 0)
		{
			updateHistoryPlay(cid);
		}
	}
}

// To get the count of completed tickets in Report Screen based on program
// selection
function getStatusTickets(progSelected)
{
  var program = new Array();
  var  completedTicketsQuery = Utilities.formatString("select count(*) from sfdc where status='completed' and program_batch='%s'", progSelected);
  var resultSet = getQueryResultSet(completedTicketsQuery);
  program[0] = resultSet.records;
  var  totalTicketsQuery = Utilities.formatString("select count(*) from sfdc where program_batch ='%s'", progSelected);
  var resultSet = getQueryResultSet(totalTicketsQuery);
  program[1] = resultSet.records;
  return program;
}
// To get distinct LDAP's in Report screen filter and reused in Reassign too..
function getLDAP()
{
  var  searchOrModifyLDAPQuery = Utilities.formatString("SELECT distinct(ldap) from users order by ldap");
  var resultSet = getQueryResultSet(searchOrModifyLDAPQuery);
  return resultSet.records;
}
// To Reassign the tickets in Report Screen
function reassignToWIP(cid, status, ldap)
{
	var conn;
	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		if (status == CLARIFICATION_AGENT)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=? WHERE cid = ?');
			stmt.setString(1, WIP_AGENT);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
		}
		if (status == WIP_AGENT)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=? WHERE cid = ?');
			stmt.setString(1, WIP_AGENT);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
		}
        if (status == WIP_AGENTREWORK)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=? WHERE cid = ?');
			stmt.setString(1, WIP_AGENTREWORK);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
          Logger.log(WIP_AGENTREWORK);
		}
		if (status == WIP_QA)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=? WHERE cid = ?');
			stmt.setString(1, WIP_QA);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
		}
		if (status == IDLE) // doubt based on front end selection
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_agent=? WHERE cid = ?');
			if (status == WIP_AGENT) stmt.setString(1, WIP_AGENT);
			if (status == WIP_QA) stmt.setString(1, WIP_QA);
			if (status == CLARIFICATION_AGENT) stmt.setString(1, WIP_AGENT);
			if (status == CLARIFICATION_QA) stmt.setString(1, QA);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
		}
		if (status == CLARIFICATION_QA)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=?  WHERE cid = ?');
			stmt.setString(1, QA);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
        }
        if (status == QA)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=?  WHERE cid = ?');
			stmt.setString(1, WIP_QA);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
          
          /*var ldapAgent = getLdapAgentfromCid(cid);
          var sSize = getSamplingSize(ldapAgent);
          var sampleSize = parseInt(sSize);
          var newSampleSize = (sampleSize - 1);
          updateSamplingSize(newSampleSize,ldapAgent);*/
		}
         if (status == WIP_QAREWORK)
		{
			var stmt = conn.prepareStatement('UPDATE sfdc SET status=?, ldap_qa=?  WHERE cid = ?');
			stmt.setString(1, WIP_QAREWORK);
			stmt.setString(2, ldap);
			stmt.setString(3, cid);
          
          /* var ldapAgent = getLdapAgentfromCid(cid);
           var sSize = getSamplingSize(ldapAgent);
          var sampleSize = parseInt(sSize);
          var newSampleSize = (sampleSize - 1);
          updateSamplingSize(newSampleSize,ldapAgent);*/
		}
		stmt.execute();
	}
	catch (e)
	{
      throw Error("reassignToWIP :"+e);
		Logger.log(e);
	}
	finally
	{
		conn.close();
	}
}
// To Export Report results into a spreadsheet
function exportToExcel_Report(leadProgrm, date, hourAssignment, statusAssignment, fromSubOrExport, ldapsId)
{
	if (typeof date == 'undefined')
	{
		date = '';
	}
	if (typeof hourAssignment == 'undefined')
	{
		hourAssignment = '';
	}
	if (typeof date == 'undefined')
	{
		date = '';
	}
	if (typeof statusAssignment == 'undefined')
	{
		statusAssignment = '';
	}
	if (typeof fromSubOrExport == 'undefined')
	{
		fromSubOrExport = '';
	}
	if (typeof ldapsId == 'undefined')
	{
		ldapsId = '';
	}
	var arrayOfLeadResults = getLeadResults(leadProgrm, date, hourAssignment, statusAssignment, fromSubOrExport, ldapsId);
	var array = new Array();
	var root = DriveApp.getRootFolder();
	var date1 = Utilities.formatDate(new Date(), "IST", "MMddyyyy_HH_mm_ss");
	var createSheetName = 'ARP_Report_' + date1;
	var newFileId = SpreadsheetApp.create(createSheetName);
	var sheet1 = SpreadsheetApp.openById(newFileId.getId());
	var templateSheet = sheet1.getSheetByName('Export');
	sheet1.insertSheet(1, {template:templateSheet});
	var sheet2 = sheet1.getSheets()[0];
	var row;
	var conn;
	var stmt;
	var sData;
    var headerCols = [];
    var allCols = [];
	array[0] = newFileId.getUrl();
	var cols = arrayOfLeadResults[0].split(",");
	for (var i = 0; i < cols.length; i++)
	{
		sheet2.getRange(1, i + 1).setValue(cols[i]);
	}
	for (var j = 1; j < arrayOfLeadResults.length; j++)
	{
		row = sheet2.getLastRow();
		var cols = arrayOfLeadResults[j].split(",");
		for (var i = 0; i < cols.length; i++)
		{
            cols[i] = cols[i].replace(/~/g, ","); // Decrypting
			sheet2.getRange(row + 1, i + 1).setValue(cols[i]);
		}
	}
	array[0] = newFileId.getUrl();
	var sheetlog = SpreadsheetApp.openById(newFileId.getId());
	if (typeof date == 'undefined')
	{
		date = '';
	}
	var sheet_log = sheetlog.getSheets()[1];
	var date_time = arp_makeTimestamp();
	var starttime = "Start time: " + date_time;
	var programFilter = "Program Filter: " + leadProgrm;
	var dateFilter = "Date Filter: " + date;
	var hourFilter = "Hour Filter: " + hourAssignment;
	var ldapFilter = "LDAP Filter: " + ldapsId;
	var statusFilter = "Status Filter: " + statusAssignment;
	sheet_log.appendRow([starttime]);
	sheet_log.appendRow([programFilter]);
	sheet_log.appendRow([dateFilter]);
	sheet_log.appendRow([hourFilter]);
	sheet_log.appendRow([ldapFilter]);
	sheet_log.appendRow([statusFilter]);
	return array;
}

function barchartDataDateRange(program_batch,ldaps,startDate,endDate)
{
	startDate = startDate+" 00:00:00";
	endDate = endDate+" 23:59:59";

	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
		var stmt = conn.createStatement();
		if(ldaps == "empty")
		{
			if(program_batch == "none")
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
					    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid WHERE " // WHERE ldap_qa is null AND  
					    +"agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' GROUP BY ldap_agent";
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
					    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid  WHERE program_batch ='"+program_batch+"' " //WHERE ldap_qa is null AND
					    +"AND agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' GROUP BY ldap_agent";

			}  

		}
		else
		{
			if(program_batch == "none")
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
					    +"IFNULL((CONCAT((ROUND((((COUNT(t2.cid) - COUNT(CASE WHEN INSTR(CONCAT(parent_on_npl,parent_company,parent_company_in_grm,child_company,child_company_in_grm,division_company,division_company_in_grm,childrdivision_company_id,action_on_mapping,flag_reason,source_id,destination_id,merge_flag), 'rejected') > 0 THEN 1 ELSE NULL END))/COUNT(t2.cid))*100), 2)), '%')), '0.00%') as 'qaScore %'"
					    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid WHERE " // WHERE ldap_qa is null AND 
					    +"agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND ldap_agent IN ("+ldaps+") GROUP BY ldap_agent";
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
					    +"FROM sfdc AS t1 LEFT OUTER JOIN qa AS t2 ON t1.cid=t2.cid WHERE program_batch ='"+program_batch+"' " //WHERE ldap_qa is null AND 
					    +"AND agent_start_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND ldap_agent IN ("+ldaps+") GROUP BY ldap_agent";
			}  

		}
		Logger.log(query);
		var stmt = conn.prepareStatement(query);

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

function barchartDataQADateRange(program_batch,ldaps,startDate,endDate)		
{		
    Logger.log("Started barchartDataQADateRange......");		
    startDate = startDate+" 00:00:00";		
    endDate = endDate+" 23:59:59";		
	try		
	{		
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);		
		var stmt = conn.createStatement();		
		stmt.setMaxRows(1000);		
        if(ldaps == "empty")		
        {		
		 if (program_batch == "none")		
		 {		
			var query = "SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa,"		
                        +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,"		
                        +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed," 		
                        +"COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc where ldap_agent is not null "		
                        +"AND qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' GROUP BY ldap_qa";		
		 }		
		else		
		 {		
			var query = "SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa,"		
                        +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,"		
                        +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed," 		
                        +"COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc " 		
                        +"where ldap_agent is not null AND program_batch='"+program_batch+"' AND qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' GROUP BY ldap_qa";		
		 }		
        }		
        else		
        {		
          if (program_batch == "none")		
		  {		
			var query = "SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa,"		
                        +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,"		
                        +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed," 		
                        +"COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc where ldap_agent is not null "		
                        +"AND qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND ldap_agent IN ("+ldaps+") GROUP BY ldap_qa";		
		 }		
		 else		
		 {		
			var query = "SELECT ldap_qa,COUNT(CASE WHEN status='qa' THEN 1  ELSE NULL END) AS qa,"		
                        +"COUNT(CASE WHEN status='clarification_qa' THEN 1 ELSE NULL END) AS clarification_qa,"		
                        +"COUNT(CASE WHEN status='completed' THEN 1 ELSE NULL END) AS completed," 		
                        +"COUNT(CASE WHEN status='wip_qa' THEN 1 ELSE NULL END) AS wip_qa FROM sfdc " 		
                        +"where ldap_agent is not null AND program_batch='"+program_batch+"' AND qa_end_time BETWEEN '"+startDate+"' AND '"+endDate+"' AND ldap_agent IN ("+ldaps+") GROUP BY ldap_qa";		
		 }		
        }		
		Logger.log(query);		
		var stmt = conn.prepareStatement(query);		
				
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
//	Logger.log("data1:" + data);		
	return data;		
}

function getLdapAgentfromCid(cId)
{
  var programBatchQuery = Utilities.formatString("SELECT ldap_agent FROM sfdc where cid='%s'", cId);
  var resultSet = getQueryResultSet(programBatchQuery);
  var data = resultSet.records[0].split("~");
  var ldapName = data.toString();
  return ldapName;
}

var FOLDER_ID = '0B_7Lh_yQhmWjfjU5LVFtZ2xfWV9IQ1I0NUd2a2MzX0tId3E2M0staVNoOGhqT2RzYXZ1aWs';
var XML_FILE = "Prod_DashBoard.xml";

function parseXml(sdate,edate) {
	Logger.clear();
	Logger.log("parseXml");
	var newArray = new Array();
	var files = DriveApp.getFolderById(FOLDER_ID).getFiles();
	while (files.hasNext()) {

		var file = files.next();
		if(file.getName() == XML_FILE)
		{  
			var xml =  file.getBlob().getDataAsString();
			var document = XmlService.parse(xml);
			var root = document.getRootElement();
			var children = root.getChildren();
			for(var i=0;i < children.length ;i++)
			{ 
				var arrayString = new Array();
				var subChildren = children[i].getChildren();
				Logger.log("length:"+subChildren.length);
				var legendArray = new Array();
				for(var j=0;j < subChildren.length ;j++)
				{ 
					var ssubChildren = subChildren[j].getChildren();
					var title = ssubChildren[0].getText();
					var sp = ssubChildren[1].getText();
					var legends = ssubChildren[2].getChildren();
					for(var k=0;k < legends.length;k++)
					{ 
						legendArray.push(legends[k].getText());
					}
	     //arrayString = getDynamicQueryResult(sp,"storedProcedure");
					arrayString = getDynamicQueryResultNew(sp,"storedProcedure",sdate,edate);
					arrayString.push(title);
					if(legends.length > 0)
					{
						arrayString.push(legendArray);
					}
					else
					{
						var legendArray = new Array();
						arrayString.push(legendArray);
					}
					newArray.push(arrayString);

				}

			}

		} 
	}
	Logger.log(newArray);
	Logger.log(newArray.length);
	return newArray;
}