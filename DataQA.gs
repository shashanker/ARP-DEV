// function using for saving QA ticket information after clicking the submit or pause button
function saveQAdata(cId, websiteIcsIdAprove, websiteIcsIdComments, rollUpIdAprove, rollUpIdComments, correctCompNameAprove, correctCompNameComments, correctCompIDAprove, correctCompIDComments, mccAprove, mccComments, agencyAprove, agencyComments, escalationAprove, escalationComments, parentNPLAprove, parentNPLComments, FlagAprove, FlagComments, parentCmpAprove, parentCmpComments, parentLowestAprove, parentLowestComments, childAprove, childComments, childLowestAprove, childLowestComments, divisionAprove, divisionComments, divisionLowestAprove, divisionLowestComments, lowestAprove, lowestComments, actionAprove, actionComments, sourceAprove, sourceComments, destinationtAprove, destinationComments, mergeAprove, mergeComments, clarComments, remarksComments, playStatus, userName, userRole, mergeExistsAprove_22, mergeExistsComments, mergeStatusAprove_21, mergeStatusComments)
{
   Logger.log('saveQAdata');
    var errorString = "No Error";
	var programBatchQuery = Utilities.formatString("SELECT count(*) FROM qa where cid='%s'", cId);
    var resultSet = getQueryResultSet(programBatchQuery);
  Logger.log("resultSet.records[0]:"+resultSet.records[0]);
    if (resultSet.records[0] > 0)
	{
		try
		{
			conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
			var stmt = conn
					.prepareStatement('UPDATE qa SET website_from_ics=?, website_from_ics_cmt=?, current_company_rollup=?, current_company_rollup_cmt=?, correct_company=?,correct_company_cmt=?, correct_company_id=?,correct_company_id_cmt=?, mcc_id=?,mcc_id_cmt=?,agency_name=?,agency_name_cmt=?,require_escalation=?,require_escalation_cmt=?,parent_on_npl=?,parent_on_npl_cmt=?,flag_reason=?,flag_reason_cmt=?,parent_company=?,parent_company_cmt=?,parent_company_in_grm=?,parent_company_in_grm_cmt=?, child_company=?,child_company_cmt=?,child_company_in_grm=?,child_company_in_grm_cmt=?,division_company=?,division_company_cmt=?,division_company_in_grm=?,division_company_in_grm_cmt=?,childrdivision_company_id=?,childrdivision_company_id_cmt=?,action_on_mapping=?,action_on_mapping_cmt=?,source_id=?,source_id_cmt=?,destination_id=?,destination_id_cmt=?,merge_flag=?,merge_flag_cmt=?,clarification_required_qa=?,remarks_qa=?,merge_exists=?,merge_exists_cmt=?,merge_status=?,merge_status_cmt=? WHERE cid = ?');
			stmt.setString(1, websiteIcsIdAprove);
			stmt.setString(2, websiteIcsIdComments);
			stmt.setString(3, rollUpIdAprove);
			stmt.setString(4, rollUpIdComments);
			stmt.setString(5, correctCompNameAprove);
			stmt.setString(6, correctCompNameComments);
			stmt.setString(7, correctCompIDAprove);
			stmt.setString(8, correctCompIDComments);
			stmt.setString(9, mccAprove);
			stmt.setString(10, mccComments);
			stmt.setString(11, agencyAprove);
			stmt.setString(12, agencyComments);
			stmt.setString(13, escalationAprove);
			stmt.setString(14, escalationComments);
			stmt.setString(15, parentNPLAprove);
			stmt.setString(16, parentNPLComments);
			stmt.setString(17, FlagAprove);
			stmt.setString(18, FlagComments);
			stmt.setString(19, parentCmpAprove);
			stmt.setString(20, parentCmpComments);
			stmt.setString(21, parentLowestAprove);
			stmt.setString(22, parentLowestComments);
			stmt.setString(23, childAprove);
			stmt.setString(24, childComments);
			stmt.setString(25, childLowestAprove);
			stmt.setString(26, childLowestComments);
			stmt.setString(27, divisionAprove);
			stmt.setString(28, divisionComments);
			stmt.setString(29, divisionLowestAprove);
			stmt.setString(30, divisionLowestComments);
			stmt.setString(31, lowestAprove);
			stmt.setString(32, lowestComments);
			stmt.setString(33, actionAprove);
			stmt.setString(34, actionComments);
			stmt.setString(35, sourceAprove);
			stmt.setString(36, sourceComments);
			stmt.setString(37, destinationtAprove);
			stmt.setString(38, destinationComments);
			stmt.setString(39, mergeAprove);
			stmt.setString(40, mergeComments);
			stmt.setString(41, clarComments);
			stmt.setString(42, remarksComments);
			stmt.setString(43, mergeExistsAprove_22);
			stmt.setString(44, mergeExistsComments);
			stmt.setString(45, mergeStatusAprove_21);
			stmt.setString(46, mergeStatusComments);
			stmt.setString(47, cId);
			stmt.execute();
            errorString = "No Error";
          
            if (playStatus == "none")
            {
              updateStatusonSubmit(cId, userRole, clarComments);
            }
            else
            {
              updateQaStatusHistory(cId, playStatus, userName);
            }
		}
		catch (e)
		{
			Logger.log("saveQAdata"+e);
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
					.prepareStatement("INSERT INTO qa (cid, website_from_ics, website_from_ics_cmt, current_company_rollup, current_company_rollup_cmt, correct_company, correct_company_cmt, correct_company_id, correct_company_id_cmt, mcc_id, mcc_id_cmt, agency_name, agency_name_cmt, require_escalation, require_escalation_cmt, parent_on_npl, parent_on_npl_cmt, flag_reason, flag_reason_cmt, parent_company, parent_company_cmt, parent_company_in_grm, parent_company_in_grm_cmt, child_company, child_company_cmt, child_company_in_grm, child_company_in_grm_cmt, division_company, division_company_cmt, division_company_in_grm, division_company_in_grm_cmt, childrdivision_company_id, childrdivision_company_id_cmt, action_on_mapping, action_on_mapping_cmt, source_id, source_id_cmt, destination_id, destination_id_cmt, merge_flag, merge_flag_cmt, clarification_required_qa, remarks_qa, merge_exists, merge_exists_cmt, merge_status, merge_status_cmt) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
			stmt.setString(1, cId);
			stmt.setString(2, websiteIcsIdAprove);
			stmt.setString(3, websiteIcsIdComments);
			stmt.setString(4, rollUpIdAprove);
			stmt.setString(5, rollUpIdComments);
			stmt.setString(6, correctCompNameAprove);
			stmt.setString(7, correctCompNameComments);
			stmt.setString(8, correctCompIDAprove);
			stmt.setString(9, correctCompIDComments);
			stmt.setString(10, mccAprove);
			stmt.setString(11, mccComments);
			stmt.setString(12, agencyAprove);
			stmt.setString(13, agencyComments);
			stmt.setString(14, escalationAprove);
			stmt.setString(15, escalationComments);
			stmt.setString(16, parentNPLAprove);
			stmt.setString(17, parentNPLComments);
			stmt.setString(18, FlagAprove);
			stmt.setString(19, FlagComments);
			stmt.setString(20, parentCmpAprove);
			stmt.setString(21, parentCmpComments);
			stmt.setString(22, parentLowestAprove);
			stmt.setString(23, parentLowestComments);
			stmt.setString(24, childAprove);
			stmt.setString(25, childComments);
			stmt.setString(26, childLowestAprove);
			stmt.setString(27, childLowestComments);
			stmt.setString(28, divisionAprove);
			stmt.setString(29, divisionComments);
			stmt.setString(30, divisionLowestAprove);
			stmt.setString(31, divisionLowestComments);
			stmt.setString(32, lowestAprove);
			stmt.setString(33, lowestComments);
			stmt.setString(34, actionAprove);
			stmt.setString(35, actionComments);
			stmt.setString(36, sourceAprove);
			stmt.setString(37, sourceComments);
			stmt.setString(38, destinationtAprove);
			stmt.setString(39, destinationComments);
			stmt.setString(40, mergeAprove);
			stmt.setString(41, mergeComments);
			stmt.setString(42, clarComments);
			stmt.setString(43, remarksComments);
			stmt.setString(44, mergeExistsAprove_22);
			stmt.setString(45, mergeExistsComments);
			stmt.setString(46, mergeStatusAprove_21);
			stmt.setString(47, mergeStatusComments);

			stmt.execute();
            errorString = "No Error";
            if (playStatus == "none")
            {
              updateStatusonSubmit(cId, userRole, clarComments);
            }
            else
            {
              updateQaStatusHistory(cId, playStatus, userName);
            }
		}
		catch (e)
		{
			Logger.log("saveQAdata"+e);
            errorString = e.toString();
		}
		finally
		{
			conn.close();
		}
	}

    return  errorString;
}
// function using for updating the ticket afer selecting the pause button for QA
function updateQaStatusHistory(cId, playStatus, userName)
{

	try
	{
		conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);

		if (playStatus == "Pause")
		{
			var stmt = conn.prepareStatement('INSERT INTO history (ldap,cid,status,pause_time) VALUES (?,?,?,?)');
			stmt.setString(1, userName);
			stmt.setString(2, cId);
			stmt.setString(3, WIP_QA);
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
	}
	catch (e)
	{
		Logger.log(e);
      
	}
	finally
	{
		conn.close();
	}
	updateBreakTime(cId);
}

function updateAgentHistory(cId,ldapAgent,ldapQA)
{
  Logger.log("updateAgentHistory"+cId+ldapAgent+ldapQA);
  var errorMsg = "Ok";
  var query = "SELECT * FROM agent where cid="+cId+"";
  Logger.log(query);
  var programBatchQuery = Utilities.formatString("SELECT * FROM agent where cid='%s'", cId);
  var resultSet = getQueryResultSet(programBatchQuery);

  var data = resultSet.records[0].split("~");

  var websiteFrIcs=data[2]; 
  var currentCompanyRollup=data[3]; 
  var correctCompany=data[4];
  var correctCompanyId=data[5]; 
  var mccId=data[6]; 
  var agencyName=data[7]; 
  var requireEscalation=data[8]; 
  var parentOnNpl=data[9]; 
  var flagReason=data[10]; 
  var parentCompany=data[11]; 
  var parentCompanyIngrm=data[12];
  var childCompany=data[13]; 
  var childCompanyIngrm=data[14];
  var divisionCompany=data[15]; 
  var divisionCompanyIngrm=data[16]; 
  var childrdivisionCompanyId=data[17]; 
  var actionOnMapping=data[18]; 
  var mergeExists=data[19]; 
  var mergeStatus=data[20];
  var sourceId=data[21];
  var destinationId=data[22]; 
  var mergeFlag=data[23]; 
  var clarificationRequiredAgent=data[24]; 
  var remarksAgent=data[25]; 
  
     
   try
	{
            conn = Jdbc.getConnection(DB_URL, USER, USER_PWD);
			var stmt = conn.prepareStatement("INSERT INTO agent_history (cid, website_from_ics, current_company_rollup, correct_company, correct_company_id, mcc_id, agency_name, require_escalation, parent_on_npl, flag_reason, parent_company, parent_company_in_grm, child_company, child_company_in_grm, division_company, division_company_in_grm, childrdivision_company_id, action_on_mapping, merge_exists, merge_status,source_id, destination_id,merge_flag, clarification_required_agent, remarks_agent,ldap_agent,ldap_qa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
			stmt.setString(1, cId);
			stmt.setString(2, websiteFrIcs);
			stmt.setString(3, currentCompanyRollup);
			stmt.setString(4, correctCompany);
			stmt.setString(5, correctCompanyId);
			stmt.setString(6, mccId);
			stmt.setString(7, agencyName);
			stmt.setString(8, requireEscalation);
			stmt.setString(9, parentOnNpl);
			stmt.setString(10, flagReason);
			stmt.setString(11, parentCompany);
			stmt.setString(12, parentCompanyIngrm);
			stmt.setString(13, childCompany);
			stmt.setString(14, childCompanyIngrm);
			stmt.setString(15, divisionCompany);
			stmt.setString(16, divisionCompanyIngrm);
			stmt.setString(17, childrdivisionCompanyId);
			stmt.setString(18, actionOnMapping);
            stmt.setString(19, mergeExists);
            stmt.setString(20, mergeStatus);
			stmt.setString(21, sourceId);
			stmt.setString(22, destinationId);
			stmt.setString(23, mergeFlag);
			stmt.setString(24, clarificationRequiredAgent);
			stmt.setString(25, remarksAgent);
            stmt.setString(26, ldapAgent);
            stmt.setString(27, ldapQA);
           
			Logger.log(stmt);
			stmt.execute();
      errorMsg = "Ok";
    }
    catch(e)
    {
      Logger.log("updateAgentHistory:"+e);
      errorMsg = e.toString();
    }
    finally
    {
        conn.close();
    }
   return errorMsg;
}  