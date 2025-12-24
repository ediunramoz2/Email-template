const ss = SpreadsheetApp.getActiveSpreadsheet();

// --- CONFIGURATION ---
const CONFIG = {
  URL_LOGO: 'https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/Cognizant_logo_2022.svg/768px-Cognizant_logo_2022.svg.png',
  SHEET_URL_BASE: ss.getUrl()
};

// --- WEB APP SERVING & INITIAL DATA ---
function doGet(e) {
  return HtmlService.createTemplateFromFile('MainIndex').evaluate().setTitle('Meeting Summary Generator for Admins');
}

function getInitialData() {
  try {
    const workflows = getSheetData('Workflows', 4).map(row => ({ name: row[0], recipientSpreadsheetId: row[1], recipientSheetName: row[2], recipientRange: row[3] }));
    let templates = getSheetData('EmailTemplates', 5).map(row => ({ name: row[0], greeting: row[1], intro: row[2], closing: row[3], footer: row[4] }));

    // Ensure MMI template exists with its own closing/footer (same as Support Team)
    const hasMMITemplate = templates.some(t => t.name === 'MMI');
    if (!hasMMITemplate) {
      const saTemplate = templates.find(t => t.name === 'SA');
      if (saTemplate) {
        const mmiTemplate = { 
          name: 'MMI', 
          greeting: 'Hi Team,', 
          intro: '', 
          closing: 'For any questions, please contact the QA PoCs.',
          footer: 'Thank you,\n\nThis auto-generated summary was created by the QA Team.'
        };
        templates.push(mmiTemplate);
      } else {
        const defaultTemplate = templates.find(t => t.name === 'Default');
        if (defaultTemplate) {
          const mmiTemplate = { 
            name: 'MMI', 
            greeting: 'Hi Team,', 
            intro: '', 
            closing: 'For any questions, please contact the QA PoCs.',
            footer: 'Thank you,\n\nThis auto-generated summary was created by the QA Team.'
          };
          templates.push(mmiTemplate);
        }
      }
    }

    return { workflows: workflows, templates: templates };
  } catch (e) {
    console.error('Error in getInitialData:', e.toString());
    return { error: e.toString() };
  }
}

function getSupportInitialData() {
  try {
    const pems = getPemsFromSheet();
    const pocs = getPocsFromSheet();
    const managers = getManagersFromSheet();

    return {
      pems: pems,
      pocs: pocs,
      managers: managers,
    };
  } catch (e) {
    console.error(' Error in getSupportInitialData:', e.toString());
    return { error: e.toString() };
  }
}

// --- CALENDAR & EMAIL LOGIC ---
function getEventsForDate(dateString) {
  try {
    const targetDate = new Date(dateString + 'T00:00:00');
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEventsForDay(targetDate);
    Logger.log('Found ' + events.length + ' events for date: ' + dateString);
    return events.map(event => ({
      id: event.getId(),
      title: event.getTitle(),
      time: event.getStartTime().toLocaleTimeString('en-GB', {timeZone: "Europe/Lisbon", hour: '2-digit', minute:'2-digit'})
    }));
  } catch(err) { 
    Logger.log("Error in getEventsForDate: " + err.message); 
    throw new Error('Failed to load calendar events: ' + err.message);
  }
}

// --- RECIPIENT & PEM MANAGEMENT ---
function getPemsFromSheet() {
  return getSheetData('PEM ldaps', 3).map(row => ({ ldap: row[0], fullName: row[1], workflow: row[2] }));
}

function getMMILdapsFromExternalSheet() {
  try {
    const externalSheetId = '1Nsjc-tI8UEoQs29t0zyEnmzti_FF4DhYQ9iPlAwqW4k';
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    
    const tabNames = ['MMI Lis Combined', 'MMI KRK', 'MMI KL'];
    const allLdaps = new Set();
    
    tabNames.forEach(tabName => {
      try {
        const sheet = externalSheet.getSheetByName(tabName);
        if (!sheet) {
          console.log(` Tab '${tabName}' not found in external sheet`);
          return;
        }
        
        const lastColumn = sheet.getLastColumn();
        
        if (lastColumn < 4) { // Column D is the 4th column
          console.log(` Tab '${tabName}' has less than 4 columns (no data from D onwards)`);
          return;
        }
        
        // Read row 2 starting from column D (column 4) to the last column
        // Range: D2:LastColumn2
        const ldapRange = sheet.getRange(2, 4, 1, lastColumn - 3); // Row 2, Column D, 1 row, (lastColumn - 3) columns
        const ldapValues = ldapRange.getValues()[0]; // Get the first (and only) row
        
        let loadedCount = 0;
        ldapValues.forEach(ldap => {
          if (ldap && typeof ldap === 'string' && ldap.trim() !== '') {
            allLdaps.add(ldap.trim());
            loadedCount++;
          }
        });
        
        console.log(` Loaded ${loadedCount} LDAPs from tab '${tabName}' (row 2, columns D:${String.fromCharCode(64 + lastColumn)})`);
      } catch (tabError) {
        console.error(` Error reading tab '${tabName}':`, tabError.message);
      }
    });
    
    const result = Array.from(allLdaps);
    console.log(` Total MMI LDAPs loaded: ${result.length}`);
    return result;
  } catch (err) {
    console.error(' Error in getMMILdapsFromExternalSheet:', err.message);
    return [];
  }
}

function getSALdapsFromExternalSheet() {
  try {
    const externalSheetId = '1Nsjc-tI8UEoQs29t0zyEnmzti_FF4DhYQ9iPlAwqW4k';
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    
    const teamConfigs = {
      'MSAB': [
        { tab: 'Impersonation Lis', startCol: 'D' },
        { tab: 'Deceptive identity Lis', startCol: 'D' }
      ],
      'Impersonation Lis': [
        { tab: 'Impersonation Lis', startCol: 'D' }
      ],
      'Civics News': [
        { tab: 'Civics news Impersonation', startCol: 'D' }
      ]
    };
    
    const allLdaps = {
      'MSAB': new Set(),
      'Impersonation Lis': new Set(),
      'Civics News': new Set()
    };
    
    // Load LDAPs for each team
    Object.keys(teamConfigs).forEach(teamName => {
      teamConfigs[teamName].forEach(config => {
        try {
          const sheet = externalSheet.getSheetByName(config.tab);
          if (!sheet) {
            console.log(` Tab '${config.tab}' not found for team ${teamName}`);
            return;
          }
          
          const lastColumn = sheet.getLastColumn();
          const startColNum = config.tab === 'Civics news Impersonation' ? 4 : 4; // D=4 for Civics (D5), D=4 for others (D8)
          const startRow = config.tab === 'Civics news Impersonation' ? 5 : 8; // Row 5 for Civics, Row 8 for others
          
          if (lastColumn < startColNum) {
            console.log(` Tab '${config.tab}' has insufficient columns`);
            return;
          }
          
          // Read from start row, column D onwards
          const ldapRange = sheet.getRange(startRow, startColNum, 1, lastColumn - startColNum + 1);
          const ldapValues = ldapRange.getValues()[0];
          
          let loadedCount = 0;
          ldapValues.forEach(ldap => {
            if (ldap && typeof ldap === 'string' && ldap.trim() !== '') {
              allLdaps[teamName].add(ldap.trim());
              loadedCount++;
            }
          });
          
          console.log(` Loaded ${loadedCount} LDAPs from '${config.tab}' for team '${teamName}' (row ${startRow}, col D onwards)`);
        } catch (tabError) {
          console.error(` Error reading tab '${config.tab}' for team ${teamName}:`, tabError.message);
        }
      });
    });
    
    // Convert Sets to Arrays
    const result = {
      'MSAB': Array.from(allLdaps['MSAB']),
      'Impersonation Lis': Array.from(allLdaps['Impersonation Lis']),
      'Civics News': Array.from(allLdaps['Civics News'])
    };
    
    console.log(` Total SA LDAPs loaded - MSAB: ${result['MSAB'].length}, Impersonation Lis: ${result['Impersonation Lis'].length}, Civics News: ${result['Civics News'].length}`);
    return result;
  } catch (err) {
    console.error(' Error in getSALdapsFromExternalSheet:', err.message);
    return { 'MSAB': [], 'Impersonation Lis': [], 'Civics News': [] };
  }
}

function getSARecipientsByTeam(selectedTeams) {
  try {
    const allSALdaps = getSALdapsFromExternalSheet();
    const uniqueLdaps = new Set();
    
    // Collect LDAPs from all selected teams
    selectedTeams.forEach(team => {
      if (allSALdaps[team]) {
        allSALdaps[team].forEach(ldap => uniqueLdaps.add(ldap));
      }
    });
    
    const ldaps = Array.from(uniqueLdaps);
    console.log(` Total unique LDAPs for selected SA teams: ${ldaps.length}`);
    
    // Return emails
    return ldaps.map(ldap => ldap + '@google.com');
  } catch (err) {
    console.error(' Error in getSARecipientsByTeam:', err.message);
    return [];
  }
}

function getMMIRecipientsByTeam(selectedTeams) {
  try {
    const externalSheetId = '1Nsjc-tI8UEoQs29t0zyEnmzti_FF4DhYQ9iPlAwqW4k';
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    
    const teamMapping = {
      'All MMI': ['MMI Lis Combined', 'MMI KRK', 'MMI KL'],
      'Lisbon': ['MMI Lis Combined'],
      'Krakow': ['MMI KRK'],
      'Kuala Lumpur': ['MMI KL']
    };
    
    const allLdaps = new Set();
    const tabsToRead = new Set();
    
    // Determine which tabs to read based on selected teams
    selectedTeams.forEach(team => {
      const tabs = teamMapping[team];
      if (tabs) {
        tabs.forEach(tab => tabsToRead.add(tab));
      }
    });
    
    tabsToRead.forEach(tabName => {
      try {
        const sheet = externalSheet.getSheetByName(tabName);
        if (!sheet) return;
        
        const lastColumn = sheet.getLastColumn();
        if (lastColumn < 4) return;
        
        const ldapRange = sheet.getRange(2, 4, 1, lastColumn - 3);
        const ldapValues = ldapRange.getValues()[0];
        
        ldapValues.forEach(ldap => {
          if (ldap && typeof ldap === 'string' && ldap.trim() !== '') {
            allLdaps.add(ldap.trim());
          }
        });
      } catch (tabError) {
        console.error(` Error reading tab '${tabName}':`, tabError.message);
      }
    });
    
    return Array.from(allLdaps).map(ldap => ldap + '@google.com');
  } catch (err) {
    console.error(' Error in getMMIRecipientsByTeam:', err.message);
    return [];
  }
}

function getMMIRecipientsFromExternalSheet() {
  try {
    const externalSheetId = '1Nsjc-tI8UEoQs29t0zyEnmzti_FF4DhYQ9iPlAwqW4k';
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    
    const tabNames = ['MMI Lis Combined', 'MMI KRK', 'MMI KL'];
    const allRecipients = new Set();
    
    tabNames.forEach(tabName => {
      try {
        const sheet = externalSheet.getSheetByName(tabName);
        if (!sheet) {
          console.log(` Tab '${tabName}' not found in external sheet`);
          return;
        }
        
        const lastColumn = sheet.getLastColumn();
        
        if (lastColumn < 5) { // Column E is the 5th column
          console.log(` Tab '${tabName}' has less than 5 columns (no data from E onwards)`);
          return;
        }
        
        // Read row 2 starting from column E (column 5) to the last column
        // Range: E2:LastColumn2
        const recipientRange = sheet.getRange(2, 5, 1, lastColumn - 4); // Row 2, Column E, 1 row, (lastColumn - 4) columns
        const recipientValues = recipientRange.getValues()[0]; // Get the first (and only) row
        
        let loadedCount = 0;
        recipientValues.forEach(recipient => {
          if (recipient && typeof recipient === 'string' && recipient.trim() !== '') {
            allRecipients.add(recipient.trim());
            loadedCount++;
          }
        });
        
        console.log(` Loaded ${loadedCount} recipients from tab '${tabName}' (row 2, columns E:${String.fromCharCode(64 + lastColumn)})`);
      } catch (tabError) {
        console.error(` Error reading tab '${tabName}':`, tabError.message);
      }
    });
    
    const result = Array.from(allRecipients);
    console.log(` Total MMI Recipients loaded: ${result.length}`);
    return result;
  } catch (err) {
    console.error(' Error in getMMIRecipientsFromExternalSheet:', err.message);
    return [];
  }
}

function getPocsFromSheet() {
  return getSheetData('POCs', 3).map(row => ({ email: row[0], workflow: row[1], vertical: row[2] || '' }));
}

function getManagersFromSheet() {
  return getSheetData('Managers', 1).map(row => ({ email: row[0] }));
}

function getStakeholdersFromSheet() {
  try {
    const externalSheetId = '1u8zIOBivGxjYOOtmtQ4DO-fRAMquB93-7m8orN_FJnI';
    const externalSheet = SpreadsheetApp.openById(externalSheetId);
    const stakeholdersTab = externalSheet.getSheetByName('Stakeholders 🫱🏼‍🫲🏽');
    
    if (!stakeholdersTab) {
      Logger.log('Stakeholders tab not found');
      return [];
    }
    
    const data = stakeholdersTab.getRange('A2:F26').getValues(); // Skip header row
    Logger.log('📋 Raw stakeholder data loaded, parsing workflows...');
    
    return data.filter(row => row[1]).map(row => {
      // Parse workflows (MMI/GMI/SA format or "SA MSAB" format)
      const workflowsRaw = (row[5] || '').toString().trim();
      Logger.log('  Raw workflow string: "' + workflowsRaw + '" for ' + row[0]);
      
      // Split by '/' to get workflow groups
      const workflowGroups = workflowsRaw.split('/').map(w => w.trim()).filter(w => w);
      
      // Extract base workflows (MMI, GMI, SA) and subteams
      const workflows = [];
      const subTeams = [];
      
      workflowGroups.forEach(group => {
        // Check if group contains subteam identifiers
        if (group.includes('MSAB')) {
          workflows.push('SA'); // Base workflow
          subTeams.push('MSAB');
        } else if (group.includes('SSL')) {
          workflows.push('SA'); // Assuming SSL is under SA
          subTeams.push('SSL');
        } else if (group.includes('DP')) {
          workflows.push('SA');
          subTeams.push('DP');
        } else if (group.includes('CNI')) {
          workflows.push('SA');
          subTeams.push('CNI');
        } else if (group.includes('IMP')) {
          workflows.push('SA');
          subTeams.push('IMP');
        } else {
          // Plain workflow (MMI, GMI, SA without subteam)
          workflows.push(group);
        }
      });
      
      // Remove duplicates
      const uniqueWorkflows = [...new Set(workflows)];
      const uniqueSubTeams = [...new Set(subTeams)];
      
      Logger.log('    → Workflows: [' + uniqueWorkflows.join(', ') + '], SubTeams: [' + uniqueSubTeams.join(', ') + ']');
      
      return {
        name: row[0],
        email: row[1],
        ldap: row[2],
        designation: row[3] || '',
        location: row[4] || '',
        workflows: uniqueWorkflows,
        subTeams: uniqueSubTeams,
        rawWorkflow: workflowsRaw
      };
    });
  } catch (e) {
    Logger.log('Error reading Stakeholders sheet: ' + e.toString());
    return [];
  }
}

function getOpsLeadsForSupport(workflow, vertical) {
  const stakeholders = getStakeholdersFromSheet();
  Logger.log('🔍 getOpsLeadsForSupport called with: workflow="' + workflow + '", vertical="' + vertical + '"');
  Logger.log('📊 Total stakeholders loaded: ' + stakeholders.length);
  
  const filtered = stakeholders.filter(s => {
    // Must be Ops Lead
    if (!s.designation.includes('Ops Lead')) return false;
    
    // Must have matching workflow
    if (!s.workflows.includes(workflow)) return false;
    
    // If there's a vertical/subteam, must match
    if (vertical && s.subTeams.length > 0) {
      return s.subTeams.includes(vertical);
    }
    
    return true;
  });
  
  Logger.log('✅ Found ' + filtered.length + ' matching Ops Leads:');
  filtered.forEach(s => {
    Logger.log('  - ' + s.name + ' (' + s.email + ') | Location: ' + s.location + ' | Workflows: [' + s.workflows.join(', ') + '] | SubTeams: [' + s.subTeams.join(', ') + ']');
  });
  
  const emails = filtered.map(s => s.email);
  Logger.log('📧 Ops Leads emails to add to CC: [' + emails.join(', ') + ']');
  
  return emails;
}

function getStakeholdersForAdmin(workflow, selectedTeams) {
  const stakeholders = getStakeholdersFromSheet();
  
  // Determine site filter based on selected teams
  let siteFilter = null;
  if (selectedTeams && selectedTeams.length > 0 && !selectedTeams.includes('All MMI')) {
    // Extract sites from team names
    const sites = [];
    selectedTeams.forEach(team => {
      if (team.includes('Lisbon') || team.includes('Lis')) sites.push('LIS');
      if (team.includes('Krakow') || team.includes('KRK')) sites.push('KRK');
      if (team.includes('Kuala Lumpur') || team.includes('KL')) sites.push('KUL');
      if (team.includes('Hyderabad') || team.includes('HYD')) sites.push('HYD');
    });
    siteFilter = sites.length > 0 ? sites : null;
  }
  
  // Only filter for Ops Leads - QA/Trainers are already in PoCs or Managers lists
  const filtered = stakeholders.filter(s => {
    // Must have matching workflow
    if (!s.workflows.includes(workflow)) return false;
    
    // Only include Ops Leads (not Ops TM, Ops SDM, etc.)
    const isOpsLead = s.designation.includes('Ops Lead');
    
    if (!isOpsLead) return false;
    
    Logger.log('  Checking Ops Lead: ' + s.name + ' (' + s.designation + ') | Location: ' + s.location);
    
    // Filter by site if applicable
    if (siteFilter) {
      const included = siteFilter.includes(s.location);
      Logger.log('    → Site filter: [' + siteFilter.join(', ') + '] | Location: ' + s.location + ' | Included: ' + included);
      return included;
    }
    
    // No site filter = include all Ops Leads
    return true;
  });
  
  Logger.log('📊 Filtered ' + filtered.length + ' Ops Leads from ' + stakeholders.length + ' total stakeholders');
  
  return filtered.map(s => s.email);
}

function getAdminEmailRecipients(formData) {
  try {
    Logger.log('🔍 Calculating Admin email recipients...');
    
    // Calculate TO list
    let toList = [];
    if (formData.recipients && formData.recipients.length > 0) {
      toList = formData.recipients;
    } else {
      const allWorkflows = getInitialData().workflows;
      const workflow = allWorkflows.find(wf => wf.name === formData.workflowName);
      if (workflow) {
        toList = getRecipients(workflow);
      }
    }
    
    // Calculate CC list
    const ccList = [];
    
    // Auto-add stakeholders
    const stakeholderEmails = getStakeholdersForAdmin(formData.workflowName, formData.selectedTeams);
    ccList.push(...stakeholderEmails);
    
    // Add managers if requested
    if (formData.includeManagers) {
      const managers = getManagersFromSheet().map(m => m.email);
      ccList.push(...managers);
    }
    
    const uniqueCcList = [...new Set(ccList)];
    
    return {
      to: toList,
      cc: uniqueCcList,
      totalCount: toList.length + uniqueCcList.length
    };
  } catch (e) {
    Logger.log('Error in getAdminEmailRecipients: ' + e.toString());
    return { to: [], cc: [], totalCount: 0 };
  }
}

function savePocsToSheet(pocs) {
  return saveDataToSheet('POCs', pocs.map(p => [p.email, p.workflow, p.vertical]), ['Email', 'Workflow', 'Vertical']);
}

function saveManagersToSheet(managers) {
  return saveDataToSheet('Managers', managers.map(m => [m.email]), ['Email']);
}

// --- GENERIC HELPER FUNCTIONS FOR SHEETS ---
function getSheetData(sheetName, numColumns) {
  try {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      console.error(` Sheet '${sheetName}' not found in spreadsheet '${ss.getName()}'`);
      console.error(` Available sheets: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2) {
      console.warn(` Sheet '${sheetName}' has no data rows`);
      return [];
    }

    if (lastColumn < numColumns) {
      console.warn(` Sheet '${sheetName}' has only ${lastColumn} columns but expected ${numColumns}`);
    }

    const data = sheet.getRange(2, 1, lastRow - 1, Math.min(numColumns, lastColumn)).getValues();
    return data;
  } catch (err) {
    console.error(` Error reading sheet ${sheetName}:`, err.message);
    return [];
  }
}

function saveDataToSheet(sheetName, data, headers) {
  const sheet = ss.getSheetByName(sheetName);
  sheet.clearContents();
  sheet.appendRow(headers);
  if (data && data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  if (sheetName === 'PEM ldaps') return getPemsFromSheet();
  if (sheetName === 'POCs') return getPocsFromSheet();
  if (sheetName === 'Managers') return getManagersFromSheet();
}

function getHtmlPreview(formData) {
  Logger.log('📧 ============ PREVIEW - CALCULATING RECIPIENTS ============');
  Logger.log('Workflow: ' + formData.workflow + ', Vertical: ' + formData.vertical);
  Logger.log('Recipients config: ' + JSON.stringify(formData.recipients));
  
  // Calculate TO list (PoCs)
  let toList = [];
  if (formData.recipients && formData.recipients.pocs) {
    const allPocs = getPocsFromSheet();
    const targetWorkflow = (formData.workflow || '').toUpperCase();
    const targetVertical = (formData.vertical || '').toUpperCase();
    toList = allPocs.filter(p => {
      const pocWorkflows = (p.workflow || '').split(',').map(wf => wf.trim().toUpperCase());
      const pocVerticals = (p.vertical || '').split(',').map(v => v.trim().toUpperCase());
      if (!pocWorkflows.includes(targetWorkflow)) return false;
      if (pocVerticals.includes('ALL') || pocVerticals.includes('') || pocVerticals.includes(targetVertical) || !targetVertical) return true;
      return false;
    }).map(p => p.email);
  }
  Logger.log('📨 TO (PoCs): [' + toList.join(', ') + '] (' + toList.length + ' recipients)');
  
  // Calculate CC list
  let ccList = [];
  
  // Add Managers if requested
  if (formData.recipients && formData.recipients.managers) {
    const managers = getManagersFromSheet().map(m => m.email);
    ccList = ccList.concat(managers);
    Logger.log('👔 Managers added: [' + managers.join(', ') + '] (' + managers.length + ' recipients)');
  }
  
  // Add Ops Leads if requested
  if (formData.recipients && formData.recipients.opsLeads) {
    Logger.log('📋 Ops Leads checkbox is CHECKED - fetching Ops Leads...');
    const opsLeads = getOpsLeadsForSupport(formData.workflow, formData.vertical);
    Logger.log('➕ Adding ' + opsLeads.length + ' Ops Leads to CC list');
    ccList = ccList.concat(opsLeads);
  } else {
    Logger.log('⬜ Ops Leads checkbox is NOT checked - skipping');
  }
  
  // Remove duplicates
  ccList = [...new Set(ccList)];
  Logger.log('📧 CC (Final): [' + ccList.join(', ') + '] (' + ccList.length + ' recipients)');
  Logger.log('✅ Total recipients: TO=' + toList.length + ', CC=' + ccList.length);
  Logger.log('📧 ======================================================');
  
  const event = CalendarApp.getEventById(formData.eventId);
  const eventTitle = event.getTitle();
  const eventDate = new Date(formData.eventDate + 'T00:00:00').toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const pems = getPemsFromSheet();
  return buildModernEmailHtml(formData, eventTitle, eventDate, pems);
}

function sendFinalEmail(emailData) {
  try {
    // Validate required fields
    if (!emailData) {
      throw new Error('No email data provided');
    }
    
    if (!emailData.eventId) {
      throw new Error('Event ID is missing');
    }
    
    if (!emailData.workflow) {
      throw new Error('Workflow is missing');
    }
    
    let toList = [];
    if (emailData.recipients && emailData.recipients.pocs) {
      const allPocs = getPocsFromSheet();
      const targetWorkflow = (emailData.workflow || '').toUpperCase();
      const targetVertical = (emailData.vertical || '').toUpperCase();
      toList = allPocs.filter(p => {
        const pocWorkflows = (p.workflow || '').split(',').map(wf => wf.trim().toUpperCase());
        const pocVerticals = (p.vertical || '').split(',').map(v => v.trim().toUpperCase());
        if (!pocWorkflows.includes(targetWorkflow)) return false;
        if (pocVerticals.includes('ALL') || pocVerticals.includes('') || pocVerticals.includes(targetVertical) || !targetVertical) return true;
        return false;
      }).map(p => p.email);
    }
    
    let ccList = [];
    if (emailData.recipients && emailData.recipients.managers) {
      ccList = getManagersFromSheet().map(m => m.email);
    }
    
    // Add Ops Leads if requested (Support Team only)
    if (emailData.recipients && emailData.recipients.opsLeads) {
      Logger.log('📋 Ops Leads checkbox is CHECKED - fetching Ops Leads...');
      const opsLeads = getOpsLeadsForSupport(emailData.workflow, emailData.vertical);
      Logger.log('➕ Adding ' + opsLeads.length + ' Ops Leads to CC list');
      ccList = ccList.concat(opsLeads);
    } else {
      Logger.log('⬜ Ops Leads checkbox is NOT checked - skipping');
    }
    
    // Remove duplicates from CC list
    const beforeDedup = ccList.length;
    ccList = [...new Set(ccList)];
    Logger.log('🔄 CC list deduplicated: ' + beforeDedup + ' → ' + ccList.length + ' emails');
    Logger.log('📧 Final CC list: [' + ccList.join(', ') + ']');
    
    const uniqueToList = [...new Set(toList)];
    Logger.log('📧 Final TO list: [' + uniqueToList.join(', ') + ']');
    if (uniqueToList.length === 0) {
      throw new Error('No primary recipients (PoCs) found for the selected Workflow and Vertical.');
    }

    const event = CalendarApp.getEventById(emailData.eventId);
    const eventTitle = event ? event.getTitle() : 'Meeting';
    const eventDate = emailData.eventDate ? 
      new Date(emailData.eventDate + 'T00:00:00').toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }) :
      new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    
    const pems = getPemsFromSheet();
    const finalHtmlBody = buildModernEmailHtml(emailData, eventTitle, eventDate, pems);
    
    const subject = emailData.subject || `${emailData.workflow} Summary - ${eventDate}`;
    
    console.log(` Sending final email to: ${uniqueToList.join(', ')}`);
    console.log(` Subject: ${subject}`);
    
    GmailApp.sendEmail(uniqueToList.join(','), subject, '', {
      htmlBody: finalHtmlBody,
      cc: ccList.length > 0 ? ccList.join(',') : ''
    });

    logSentEmail(emailData, uniqueToList.join(', '), ccList.join(', '), finalHtmlBody);
    console.log(' Email sent successfully!');
    return 'Email sent successfully and logged!';
  } catch (error) {
    console.error(' Error in sendFinalEmail:', error.message);
    Logger.log('Error in sendFinalEmail: ' + error.message);
    throw new Error('Failed to send email: ' + error.message);
  }
}

function sendTestEmailToServer(emailData) {
  try {
    const testRecipient = Session.getActiveUser().getEmail();

    if (!testRecipient) {
      throw new Error("Could not get the current user's email address.");
    }

    const event = CalendarApp.getEventById(emailData.eventId);
    const eventTitle = event.getTitle();
    const eventDate = new Date(emailData.eventDate + 'T00:00:00').toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    const pems = getPemsFromSheet();
    const finalHtmlBody = buildModernEmailHtml(emailData, eventTitle, eventDate, pems);

    GmailApp.sendEmail(testRecipient, `[TEST] ${emailData.subject}`, '', {
      htmlBody: finalHtmlBody
    });

    return `Test email sent successfully to ${testRecipient}!`;

  } catch (err) {
    throw new Error(`Failed to send test email: ${err.message}`);
  }
}

// --- LOGGING AND HISTORY ---
function logSentEmail(emailData, to, cc, finalHtmlBody) {
  try {
    const sheet = ss.getSheetByName('Meeting record2');
    const headers = ['Timestamp', 'Workflow', 'Subject', 'To', 'CC', 'HTML Body', 'Sent By (LDAP)'];
    
    // Initialize headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }
    
    // Get the user's LDAP (email address without @google.com or @domain)
    const userEmail = Session.getEffectiveUser().getEmail();
    const ldap = userEmail.split('@')[0]; // Extract username part before @
    
    // Limit HTML to prevent Google Sheets cell size limit (50,000 chars)
    let htmlToStore = finalHtmlBody || '';
    const maxLength = 45000; // Leave some buffer
    if (htmlToStore.length > maxLength) {
      htmlToStore = htmlToStore.substring(0, maxLength) + '\n\n<!-- HTML truncated due to size limits -->';
    }
    
    Logger.log('Logging email - HTML body length: ' + htmlToStore.length + ', Sent by: ' + ldap);
    sheet.appendRow([
      new Date(),
      `${emailData.workflow} - ${emailData.vertical || 'General'}`,
      emailData.subject || 'No Subject',
      to || 'N/A',
      cc || '',
      htmlToStore,
      ldap
    ]);
    
    Logger.log('Email logged successfully at row ' + sheet.getLastRow());
  } catch (err) {
    Logger.log('Error in logSentEmail: ' + err.message);
    throw new Error('Failed to log email: ' + err.message);
  }
}

function getSentHistory() {
  const sheet = ss.getSheetByName('Meeting record2'); // Support Team history
  if (!sheet) {
    Logger.log('getSentHistory: Meeting record2 sheet not found');
    return [];
  }
  
  if (sheet.getLastRow() < 2) {
    Logger.log('getSentHistory: No data rows in sheet (only headers or empty)');
    return [];
  }
  
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = Math.min(7, sheet.getLastColumn()); // Get up to 7 columns, or max available
    Logger.log('getSentHistory: Reading ' + (lastRow - 1) + ' rows and ' + lastCol + ' columns');
    
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    const result = data.map((row, index) => ({
      id: index + 2,
      timestamp: row[0] ? new Date(row[0]).toLocaleString() : 'N/A',
      workflow: row[1] || 'N/A',
      subject: row[2] || 'N/A',
      to: row[3] || 'N/A',
      cc: row[4] || '',
      ldap: row[6] ? row[6] : 'Unknown' // Column 7 - Sent By (LDAP), if exists
    })).reverse();
    
    Logger.log('getSentHistory: Returning ' + result.length + ' emails');
    return result;
  } catch (err) {
    Logger.log('Error in getSentHistory: ' + err.message + ' | Stack: ' + err.stack);
    return [];
  }
}

// Quick test function to verify getSentHistory works
function testGetSentHistory() {
  const result = getSentHistory();
  Logger.log('Test result: ' + JSON.stringify(result));
  return result;
}

// Debug function to check both sheets
function debugBothSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('=== DEBUGGING BOTH SHEETS ===');
  
  // Check Meeting record (Admin)
  const adminSheet = ss.getSheetByName('Meeting record');
  if (adminSheet) {
    Logger.log('Admin Sheet "Meeting record":');
    Logger.log('  - Last Row: ' + adminSheet.getLastRow());
    Logger.log('  - Last Column: ' + adminSheet.getLastColumn());
    if (adminSheet.getLastRow() >= 2) {
      const lastRowData = adminSheet.getRange(adminSheet.getLastRow(), 1, 1, adminSheet.getLastColumn()).getValues()[0];
      Logger.log('  - Last row data: ' + JSON.stringify(lastRowData.map((v, i) => 'Col' + (i+1) + ': ' + (v ? v.toString().substring(0, 50) : 'empty'))));
    }
  } else {
    Logger.log('Admin Sheet "Meeting record" NOT FOUND');
  }
  
  Logger.log('');
  
  // Check Meeting record2 (Support)
  const supportSheet = ss.getSheetByName('Meeting record2');
  if (supportSheet) {
    Logger.log('Support Sheet "Meeting record2":');
    Logger.log('  - Last Row: ' + supportSheet.getLastRow());
    Logger.log('  - Last Column: ' + supportSheet.getLastColumn());
    if (supportSheet.getLastRow() >= 2) {
      const lastRowData = supportSheet.getRange(supportSheet.getLastRow(), 1, 1, supportSheet.getLastColumn()).getValues()[0];
      Logger.log('  - Last row data:');
      lastRowData.forEach((val, idx) => {
        const preview = val ? val.toString().substring(0, 100) : 'EMPTY';
        Logger.log('    Col ' + (idx + 1) + ': ' + preview);
      });
    }
  } else {
    Logger.log('Support Sheet "Meeting record2" NOT FOUND');
  }
  
  Logger.log('=== END DEBUG ===');
  return 'Check logs above';
}

function getEmailHtmlById(rowId, sheetName) {
  try {
    // If sheetName not provided, try Support first, then Admin
    let sheet;
    if (sheetName) {
      sheet = ss.getSheetByName(sheetName);
    } else {
      // Try Support Team first (most common)
      sheet = ss.getSheetByName('Meeting record2');
      if (!sheet || sheet.getLastRow() < 2) {
        // Fallback to Admin if Support is empty
        sheet = ss.getSheetByName('Meeting record');
      }
    }
    
    if (!sheet) {
      Logger.log('Sheet not found');
      return createErrorHtml('Sheet not found');
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log('getEmailHtmlById - Sheet: ' + sheet.getName() + ', Row: ' + rowId + ', Last Row: ' + lastRow);
    
    if (rowId < 2 || rowId > lastRow) {
      Logger.log('Row ID out of bounds: ' + rowId + ' (valid range: 2-' + lastRow + ')');
      return createErrorHtml('Email not found (invalid row number: ' + rowId + ', valid: 2-' + lastRow + ')');
    }
    
    // Get all columns to check for HTML (need 7 for LDAP column)
    const maxCols = Math.min(7, sheet.getLastColumn());
    Logger.log('Reading ' + maxCols + ' columns from row ' + rowId);
    
    const rowData = sheet.getRange(rowId, 1, 1, maxCols).getValues()[0];
    Logger.log('Row data: Timestamp=' + rowData[0] + ', Workflow=' + rowData[1] + ', Subject=' + rowData[2]);
    Logger.log('Column 6 (HTML) length: ' + (rowData[5] ? rowData[5].toString().length : 0));
    if (maxCols >= 7) {
      Logger.log('Column 7 (LDAP): ' + (rowData[6] || 'N/A'));
    }
    
    // For Support Team (Meeting record2), HTML is in column 6
    if (sheet.getName() === 'Meeting record2' && maxCols >= 6) {
      const htmlValue = rowData[5];
      const htmlLength = htmlValue ? htmlValue.toString().length : 0;
      Logger.log('Support HTML - Type: ' + typeof htmlValue + ', Length: ' + htmlLength);
      
      if (htmlValue && htmlValue.toString().trim().length > 0) {
        Logger.log('Returning Support HTML body (' + htmlLength + ' characters)');
        return htmlValue.toString();
      } else {
        Logger.log('HTML is empty or missing in column 6');
      }
    }
    
    // For Admin (Meeting record), also check for HTML in column 6 (now stored since update)
    if (sheet.getName() === 'Meeting record' && maxCols >= 6) {
      const htmlValue = rowData[5];
      const htmlLength = htmlValue ? htmlValue.toString().length : 0;
      Logger.log('Admin HTML - Type: ' + typeof htmlValue + ', Length: ' + htmlLength);
      
      if (htmlValue && htmlValue.toString().trim().length > 0) {
        Logger.log('Returning Admin HTML body (' + htmlLength + ' characters)');
        return htmlValue.toString();
      } else {
        Logger.log('HTML is empty or missing in column 6 - will show summary');
      }
    }
    
    // Fallback: Create summary for old records without HTML
    Logger.log('Creating email summary for sheet: ' + sheet.getName());
    return createEmailSummaryHtml(rowData, sheet.getName());
    
  } catch (err) {
    Logger.log('Error in getEmailHtmlById: ' + err.message + ' | Stack: ' + err.stack);
    return createErrorHtml('Error: ' + err.message);
  }
}

function createErrorHtml(message) {
  return '<html><body style="font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5;">' +
    '<div style="background-color: #ffebee; border: 1px solid #ef5350; padding: 15px; border-radius: 4px; color: #c62828;">' +
    '<h2 style="margin-top: 0; color: #b71c1c;">⚠️ Error</h2>' +
    '<p>' + message + '</p>' +
    '</div></body></html>';
}

// Helper function to debug Support history HTML storage
function debugSupportHistory() {
  try {
    const sheet = ss.getSheetByName('Meeting record2');
    if (!sheet) {
      Logger.log('DEBUG: Meeting record2 sheet not found');
      return 'Sheet not found';
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log('DEBUG: Sheet "Meeting record2" - Rows: ' + lastRow + ', Columns: ' + lastCol);
    
    if (lastRow < 2) {
      Logger.log('DEBUG: No data in sheet');
      return 'No data in sheet (only headers or empty)';
    }
    
    // Check last row for HTML content
    const lastRowData = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
    Logger.log('DEBUG: Last row data:');
    Logger.log('  Column 1 (Timestamp): ' + lastRowData[0]);
    Logger.log('  Column 2 (Workflow): ' + lastRowData[1]);
    Logger.log('  Column 3 (Subject): ' + lastRowData[2]);
    Logger.log('  Column 4 (To): ' + lastRowData[3]);
    Logger.log('  Column 5 (CC): ' + lastRowData[4]);
    
    if (lastCol >= 6) {
      const htmlCol = lastRowData[5];
      Logger.log('  Column 6 (HTML) Length: ' + (htmlCol ? htmlCol.toString().length : 0));
      Logger.log('  Column 6 (HTML) Has Content: ' + (htmlCol && htmlCol.toString().trim().length > 0 ? 'YES' : 'NO'));
      
      if (htmlCol && htmlCol.toString().length > 0) {
        // Show first 200 chars
        const preview = htmlCol.toString().substring(0, 200);
        Logger.log('  Column 6 (HTML) Preview: ' + preview + '...');
      }
    } else {
      Logger.log('  Column 6 (HTML): NOT PRESENT - Sheet only has ' + lastCol + ' columns');
    }
    
    return 'DEBUG: Check logs above for details on last row in Meeting record2';
  } catch (err) {
    Logger.log('DEBUG Error: ' + err.message);
    return 'Error: ' + err.message;
  }
}

function createEmailSummaryHtml(rowData, sheetName) {
  const isAdmin = sheetName === 'Meeting record';
  const message = isAdmin ? 
    '(Older Admin emails - HTML not stored)' :
    '(HTML not available for this email)';
  
  return '<html><body style="font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5;">' +
    '<div style="background-color: #fff3e0; border: 1px solid #ffb74d; padding: 15px; border-radius: 4px; color: #e65100;">' +
    '<h2 style="margin-top: 0;">📝 Email Summary ' + message + '</h2>' +
    '<table style="width: 100%; border-collapse: collapse; margin-top: 15px;">' +
    '<tr style="background-color: #fff9e6;"><td style="padding: 8px; border: 1px solid #ffb74d; font-weight: bold;">Timestamp:</td><td style="padding: 8px; border: 1px solid #ffb74d;">' + rowData[0] + '</td></tr>' +
    '<tr><td style="padding: 8px; border: 1px solid #ffb74d; font-weight: bold;">Workflow:</td><td style="padding: 8px; border: 1px solid #ffb74d;">' + (rowData[1] || 'N/A') + '</td></tr>' +
    '<tr style="background-color: #fff9e6;"><td style="padding: 8px; border: 1px solid #ffb74d; font-weight: bold;">Subject:</td><td style="padding: 8px; border: 1px solid #ffb74d;">' + (rowData[2] || 'N/A') + '</td></tr>' +
    '<tr><td style="padding: 8px; border: 1px solid #ffb74d; font-weight: bold;">To:</td><td style="padding: 8px; border: 1px solid #ffb74d;">' + (rowData[3] || 'N/A') + '</td></tr>' +
    '<tr style="background-color: #fff9e6;"><td style="padding: 8px; border: 1px solid #ffb74d; font-weight: bold;">CC:</td><td style="padding: 8px; border: 1px solid #ffb74d;">' + (rowData[4] || 'None') + '</td></tr>' +
    '</table>' +
    '<p style="color: #666; font-size: 12px; margin-top: 20px;">💡 Note: Emails sent after this update will include full HTML content. This is an older record.</p>' +
    '</div></body></html>';
}

// --- ROBUST YURT LINK CREATION ---
function createYurtLink(inputId) {
  if (!inputId) return { link: '', displayId: '' };
  const cleanId = inputId.trim();
  const displayId = cleanId;
  const encodedId = encodeURIComponent(cleanId);
  const longUrl = `https://yurt.corp.google.com/?entity_id=${encodedId}&entity_type=CLUSTER&config_id=prod%2Freview_session%2Fcluster%2Fimpersonation_lookup&jt=buganizer_id&jv=237313745&ds_id=YURT_LOOKUP!15549525059342401415&de_id=2025-07-15T14%3A3A53.325459422%2B00%3A00#lookup-v2`;
  return { link: longUrl, displayId: displayId };
}

// Extract video ID from Yurt URL
function extractVideoIdFromYurt(url) {
  if (!url || typeof url !== 'string') return url;
  
  // If it's not a Yurt URL, return as-is
  if (!url.includes('yurt.corp.google.com')) return url;
  
  let extractedId = '';
  
  // Method 1: Check for ?q= parameter (encoded deeplink)
  if (url.includes('?q=')) {
    try {
      const urlParts = url.split('?');
      if (urlParts.length > 1) {
        const queryString = urlParts[1].split('#')[0];
        const match = queryString.match(/q=([^&]+)/);
        if (match) {
          const qParam = decodeURIComponent(match[1]);
          const jsonData = JSON.parse(Utilities.newBlob(Utilities.base64Decode(qParam)).getDataAsString());
          if (jsonData.entityIds && jsonData.entityIds.length > 0) {
            extractedId = jsonData.entityIds[0];
            Logger.log(`Extracted video ID from ?q= parameter: ${extractedId}`);
          }
        }
      }
    } catch (e) {
      Logger.log('Could not parse encoded deeplink: ' + e.message);
    }
  }
  
  // Method 2: Check for entity_id parameter
  if (!extractedId && url.includes('entity_id=')) {
    const match = url.match(/entity_id=([^&]+)/);
    if (match) {
      extractedId = decodeURIComponent(match[1]);
      Logger.log(`Extracted video ID from entity_id: ${extractedId}`);
    }
  }
  
  return extractedId || url;
}

// Helper function to get policy area color styling
function getPolicyAreaStyle(policyArea) {
  if (!policyArea) return { bg: '#e0e0e0', text: '#333' };
  
  // ALL Policy Areas → Yellow
  return { bg: '#fff9c4', text: '#f57f17' }; // Yellow background, dark yellow text
}

// Helper function to get final decision color styling
function getFinalDecisionStyle(decision) {
  if (!decision) return { bg: '#fce8e6', text: '#d93025' }; // Default red
  
  const decisionStr = String(decision).trim();
  
  // Green for 9008 or 11800
  if (decisionStr === '9008' || decisionStr === '11800') {
    return { bg: '#e8f5e9', text: '#2e7d32', borderColor: '#4caf50' }; // Green
  }
  
  // Red for everything else
  return { bg: '#fce8e6', text: '#d93025', borderColor: '#d93025' }; // Red
}

// New function to create link based on ID type (for SA workflow)
function createSALink(caseInfo) {
  if (!caseInfo || !caseInfo.id) return { link: '', displayId: '' };
  
  const cleanId = caseInfo.id.trim();
  const idType = caseInfo.idType || 'simple';
  
  if (idType === 'deeplink' && caseInfo.originalDeeplink) {
    // Use the original deeplink as the hyperlink
    return { link: caseInfo.originalDeeplink, displayId: cleanId };
  } else if (idType === 'simple') {
    // Simple ID - no link, just display the ID
    return { link: '', displayId: cleanId };
  } else {
    // Lookup or other - create lookup link
    const encodedId = encodeURIComponent(cleanId);
    const lookupUrl = `https://yurt.corp.google.com/?entity_id=${encodedId}&entity_type=CLUSTER&config_id=prod%2Freview_session%2Fcluster%2Fimpersonation_lookup&jt=buganizer_id&jv=237313745&ds_id=YURT_LOOKUP!15549525059342401415&de_id=2025-07-15T14%3A3A53.325459422%2B00%3A00#lookup-v2`;
    return { link: lookupUrl, displayId: cleanId };
  }
}

// --- DATA & EMAIL LOGIC ---
function getRecipients(workflow) {
  // MMI workflow: read recipients from external sheet tabs
  if (workflow.name === 'MMI') {
    return getMMIRecipientsFromExternalSheet();
  }
  
  // All other workflows: read from configured sheet
  if (!workflow.recipientSpreadsheetId || !workflow.recipientSheetName || !workflow.recipientRange) {
    Logger.log(`Workflow '${workflow.name}' is missing recipient configuration.`);
    return [];
  }
  try {
    const targetSpreadsheet = SpreadsheetApp.openById(workflow.recipientSpreadsheetId);
    const targetSheet = targetSpreadsheet.getSheetByName(workflow.recipientSheetName);
    if (!targetSheet) return [];
    const range = targetSheet.getRange(workflow.recipientRange);
    return range.getValues().flat().filter(String);
  } catch (e) {
    Logger.log(`Error getting recipients for workflow '${workflow.name}': ${e.message}`);
    return [];
  }
}

function generateEmailPreview(formData) {
  try {
    Logger.log('📧 ============ ADMIN PREVIEW - CALCULATING RECIPIENTS ============');
    Logger.log('Workflow: ' + formData.workflowName);
    Logger.log('Selected Teams: ' + (formData.selectedTeams ? formData.selectedTeams.join(', ') : 'N/A'));
    Logger.log('Include Managers: ' + (formData.includeManagers ? 'YES' : 'NO'));
    
    // Calculate TO list (workflow recipients)
    let toList = [];
    if (formData.recipients && formData.recipients.length > 0) {
      toList = formData.recipients;
      Logger.log('📨 TO (Custom recipients): [' + toList.join(', ') + '] (' + toList.length + ' recipients)');
    } else {
      const allWorkflows = getInitialData().workflows;
      const workflow = allWorkflows.find(wf => wf.name === formData.workflowName);
      if (workflow) {
        toList = getRecipients(workflow);
        Logger.log('📨 TO (Workflow recipients): [' + toList.join(', ') + '] (' + toList.length + ' recipients)');
      }
    }
    
    // Calculate CC list (stakeholders)
    const ccList = [];
    
    // Auto-add stakeholders (QA + Trainers + Ops Leads)
    Logger.log('📋 Calculating stakeholders (QA + Trainers + Ops Leads)...');
    const stakeholderEmails = getStakeholdersForAdmin(formData.workflowName, formData.selectedTeams);
    ccList.push(...stakeholderEmails);
    Logger.log('✅ Auto-added stakeholders to CC: [' + stakeholderEmails.join(', ') + '] (' + stakeholderEmails.length + ' stakeholders)');
    
    // Add managers if requested
    if (formData.includeManagers) {
      const managers = getManagersFromSheet().map(m => m.email);
      ccList.push(...managers);
      Logger.log('👔 Managers added: [' + managers.join(', ') + '] (' + managers.length + ' managers)');
    } else {
      Logger.log('⬜ Managers NOT requested - skipping');
    }
    
    // Remove duplicates
    const uniqueCcList = [...new Set(ccList)];
    Logger.log('📧 CC (Final): [' + uniqueCcList.join(', ') + '] (' + uniqueCcList.length + ' recipients)');
    Logger.log('✅ Total recipients: TO=' + toList.length + ', CC=' + uniqueCcList.length);
    Logger.log('📧 ======================================================');
    
    // Ensure template exists - create MMI if missing
    if (!formData.template) {
      const allData = getInitialData();
      const foundTemplate = allData.templates.find(t => t.name === formData.workflowName);
      if (foundTemplate) {
        formData.template = foundTemplate;
      } else {
        console.error(' Template still not found after reload for workflow:', formData.workflowName);
        throw new Error('Email template not found for workflow: ' + formData.workflowName);
      }
    }
    
    const html = buildAdminModernEmailHtml(formData);
    
    // CRITICAL VALIDATION: Ensure HTML is not corrupted
    if (!html || html.length === 0) {
      console.error(' CRITICAL: HTML is empty!');
      throw new Error('HTML generation returned empty string');
    }
    
    if (!html.includes('<!DOCTYPE') || !html.includes('</html>')) {
      console.error(' CRITICAL: HTML structure incomplete!');
      throw new Error('HTML structure is incomplete');
    }
    
    return html;
  } catch (error) {
    console.error(' Error in generateEmailPreview:', error.message);
    Logger.log('Error in generateEmailPreview: ' + error.message);
    throw new Error('Failed to generate email preview: ' + error.message);
  }
}

// --- LIVE SEND FUNCTION FOR ADMINS ---

function sendAdminTestEmail(formData, ldap) {
  try {
    // Validate and clean LDAP
    const cleanLdap = (ldap || '').trim().replace(/@google\.com$/, '');
    
    if (!cleanLdap) {
      throw new Error('LDAP is empty');
    }
    
    // Validate LDAP format (alphanumeric, dots, hyphens, underscores)
    if (!/^[a-zA-Z0-9._-]+$/.test(cleanLdap)) {
      throw new Error('Invalid LDAP format. Use only letters, numbers, dots, hyphens, or underscores.');
    }
    
    const testRecipient = cleanLdap + '@google.com';
    console.log(` Sending Admin test email to: ${testRecipient}`);
    
    // Generate email HTML
    const html = generateEmailPreview(formData);
    
    if (!html || html.length === 0) {
      throw new Error('Email HTML is empty');
    }
    
    // Send email using GmailApp instead of MailApp for better compatibility
    GmailApp.sendEmail(
      testRecipient,
      '[TEST - ADMIN] ' + formData.subject,
      '', // Plain text body (empty)
      {
        htmlBody: html,
        name: 'Meeting Summary System (TEST)'
      }
    );
    
    console.log(` Test email sent successfully to ${testRecipient}`);
    return { success: true, message: `Test email sent to ${testRecipient}` };
  } catch (error) {
    console.error(' Error sending Admin test email:', error.message);
    Logger.log('Error in sendAdminTestEmail: ' + error.message);
    throw new Error('Failed to send test email: ' + error.message);
  }
}

function sendSupportTestEmail(formData, ldap) {
  try {
    // Validate form data
    if (!formData) {
      throw new Error('Form data is missing');
    }
    
    if (!formData.eventId) {
      throw new Error('Event ID is missing. Please select a meeting first.');
    }
    
    if (!formData.workflow) {
      throw new Error('Workflow is missing. Please select a workflow first.');
    }
    
    // Validate and clean LDAP
    const cleanLdap = (ldap || '').trim().replace(/@google\.com$/, '');
    
    if (!cleanLdap) {
      throw new Error('LDAP is empty');
    }
    
    // Validate LDAP format (alphanumeric, dots, hyphens, underscores)
    if (!/^[a-zA-Z0-9._-]+$/.test(cleanLdap)) {
      throw new Error('Invalid LDAP format. Use only letters, numbers, dots, hyphens, or underscores.');
    }
    
    const testRecipient = cleanLdap + '@google.com';
    console.log(` Sending Support test email to: ${testRecipient}`);
    console.log(` Form data:`, { eventId: formData.eventId, workflow: formData.workflow, eventDate: formData.eventDate });
    
    // Generate email HTML (using getHtmlPreview logic)
    let html;
    try {
      html = getHtmlPreview(formData);
    } catch (htmlError) {
      console.error(' Error generating HTML preview:', htmlError.message);
      throw new Error('Failed to generate email: ' + htmlError.message);
    }
    
    if (!html || html.length === 0) {
      throw new Error('Email HTML is empty');
    }
    
    // Get event title for subject
    let eventTitle = 'Meeting';
    let eventDate = new Date().toLocaleDateString('es-ES', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    });
    
    try {
      const event = CalendarApp.getEventById(formData.eventId);
      if (event) {
        eventTitle = event.getTitle();
      }
    } catch (e) {
      console.warn(' Could not get event title:', e.message);
    }
    
    if (formData.eventDate) {
      try {
        eventDate = new Date(formData.eventDate + 'T00:00:00').toLocaleDateString('es-ES', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        });
      } catch (e) {
        console.warn(' Could not parse event date:', e.message);
      }
    }
    
    const subject = `[TEST - SUPPORT] ${eventTitle} | ${formData.workflow} Summary (${eventDate})`;
    
    // Send email using GmailApp instead of MailApp for better compatibility
    GmailApp.sendEmail(
      testRecipient,
      subject,
      '', // Plain text body (empty)
      {
        htmlBody: html,
        name: 'Meeting Summary System (TEST)'
      }
    );
    
    console.log(` Test email sent successfully to ${testRecipient}`);
    return { success: true, message: `Test email sent to ${testRecipient}` };
  } catch (error) {
    console.error(' Error sending Support test email:', error.message);
    Logger.log('Error in sendSupportTestEmail: ' + error.message);
    throw new Error('Failed to send test email: ' + error.message);
  }
}

// --- LIVE SEND FUNCTION FOR ADMINS ---
function sendAdminFinalEmail(emailData) {
  try {
    Logger.log('=== sendAdminFinalEmail called ===');
    Logger.log('Email data received: ' + JSON.stringify({
      workflowName: emailData.workflowName,
      subject: emailData.subject,
      hasRecipients: !!(emailData.recipients && emailData.recipients.length),
      recipientCount: emailData.recipients ? emailData.recipients.length : 0,
      hasHtmlBody: !!emailData.htmlBody,
      htmlBodyLength: emailData.htmlBody ? emailData.htmlBody.length : 0
    }));
    
    let toList = [];
    
    // For MMI with custom recipients, use those
    if (emailData.recipients && emailData.recipients.length > 0) {
      toList = emailData.recipients;
      Logger.log('Using provided recipients: ' + toList.join(', '));
    } else {
      // Otherwise, use workflow recipients from sheet
      Logger.log('Fetching recipients from sheet for workflow: ' + emailData.workflowName);
      const allWorkflows = getInitialData().workflows;
      const workflow = allWorkflows.find(wf => wf.name === emailData.workflowName);
      if (!workflow) {
        throw new Error(`Workflow "${emailData.workflowName}" not found in configuration.`);
      }
      toList = getRecipients(workflow);
      Logger.log('Recipients from sheet: ' + toList.join(', '));
    }
    
    // Ensure all recipients have @google.com domain
    toList = toList.map(recipient => {
      const clean = recipient.trim();
      if (!clean) return null;
      // If it's already an email, return as-is
      if (clean.includes('@')) return clean;
      // Otherwise, append @google.com
      return clean + '@google.com';
    }).filter(email => email !== null);
    Logger.log('Emails after adding @google.com: ' + toList.join(', '));
    
    const ccList = [];
    
    // 1. Auto-add Ops Leads (from Stakeholders sheet, filtered by workflow/site)
    const opsLeadEmails = getStakeholdersForAdmin(emailData.workflowName, emailData.selectedTeams);
    ccList.push(...opsLeadEmails);
    Logger.log('Auto-added Ops Leads to CC: ' + opsLeadEmails.join(', '));
    
    // 2. Auto-add Support Team PoCs (from PoCs sheet, filtered by workflow)
    const allPocs = getPocsFromSheet();
    const targetWorkflow = (emailData.workflowName || '').toUpperCase();
    const supportTeamPocs = allPocs.filter(p => {
      const pocWorkflows = (p.workflow || '').split(',').map(wf => wf.trim().toUpperCase());
      return pocWorkflows.includes(targetWorkflow);
    }).map(p => p.email);
    ccList.push(...supportTeamPocs);
    Logger.log('Auto-added Support Team PoCs to CC: ' + supportTeamPocs.join(', '));
    
    // 3. Add managers if requested (optional)
    if (emailData.includeManagers) {
      const managers = getManagersFromSheet().map(m => m.email);
      ccList.push(...managers);
      Logger.log('Added Managers to CC: ' + managers.join(', '));
    }
    
    // Remove duplicates from CC list
    const uniqueCcList = [...new Set(ccList)];
    Logger.log('Final CC list (' + uniqueCcList.length + ' recipients): ' + uniqueCcList.join(', '));

    if (toList.length === 0) {
      throw new Error("No recipients found for this workflow. Check your configuration sheet.");
    }
    
    Logger.log('Sending email to: ' + toList.join(','));
    Logger.log('Subject: ' + emailData.subject);
    
    GmailApp.sendEmail(toList.join(','), emailData.subject, '', { htmlBody: emailData.htmlBody, cc: uniqueCcList.join(',') });
    
    Logger.log('Email sent successfully via GmailApp');
    
    logAdminSentEmail(emailData.workflowName, emailData.subject, toList.join(', '), uniqueCcList.join(', '), emailData.htmlBody);
    
    Logger.log('Email logged successfully');
    
    return 'Success! The email has been sent to ' + toList.length + ' recipient(s).';

  } catch (e) {
    Logger.log("ERROR in sendAdminFinalEmail: " + e.message);
    Logger.log("Error stack: " + e.stack);
    throw new Error('Failed to send email: ' + e.message);
  }
}

function logAdminSentEmail(workflowName, subject, to, cc, htmlBody) {
  const logSheetName = 'Meeting record';
  let sheet = ss.getSheetByName(logSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(logSheetName);
    const headers = ['Timestamp', 'Workflow', 'Subject', 'To', 'CC', 'HTML Body', 'Sent By'];
    sheet.appendRow(headers);
    sheet.getRange("A1:G1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() === 0) {
      const headers = ['Timestamp', 'Workflow', 'Subject', 'To', 'CC', 'HTML Body', 'Sent By'];
      sheet.appendRow(headers);
  }
  
  // Get user LDAP
  const userEmail = Session.getEffectiveUser().getEmail();
  const ldap = userEmail.split('@')[0];
  
  // Store HTML body (same as Support - with size limit)
  let htmlToStore = htmlBody || '';
  const maxLength = 45000; // Leave some buffer
  if (htmlToStore.length > maxLength) {
    htmlToStore = htmlToStore.substring(0, maxLength) + '\n\n<!-- HTML truncated due to size limits -->';
  }
  
  Logger.log('Logging Admin email - HTML body length: ' + htmlToStore.length + ', Sent by: ' + ldap);
  sheet.appendRow([new Date(), workflowName, subject, to, cc, htmlToStore, ldap]);
}

// --- EMAIL HTML BUILDER FOR ADMINS ---
function buildAdminModernEmailHtml(data) {
    try {
      // Validate template exists
      if (!data.template) {
        console.error(' Template is null for workflow:', data.workflowName);
        throw new Error('Email template not found for workflow: ' + data.workflowName);
      }
      const meetingDate = new Date(data.calendarEvent.date + 'T00:00:00').toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
      const internalTrixLink = "https://docs.google.com/spreadsheets/d/1vQ0kPp9UbGoobQxP89uvHxQ798xkEmfhGKLEB8r_wbM/edit?gid=0#gid=0";
      let headerTitle = 'Feedback Summary';
      if (data.workflowName === 'SA') {
        headerTitle = 'Impersonation Youtube T&S';
      } else if (data.workflowName === 'MMI') {
        headerTitle = 'MMI - Meeting Summary';
      }
      const corporateBlue = '#5479A5';
    const createSectionCard = (title, content) => {
        if (!content || content.trim() === '' || content.trim() === '<div><br></div>' || content.trim() === '<br>') return '';
        return `<div style="background-color: #ffffff; border-radius: 8px; margin-top: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border: 1px solid #e0e0e0;"><div style="background-color: ${corporateBlue}; color: #ffffff; padding: 12px 18px; border-radius: 8px 8px 0 0; font-weight: 500; font-size: 16px;"><span>${title}</span></div><div style="padding: 18px 20px; color: #3c4043; font-size: 14px; line-height: 1.6;">${content}</div></div>`;
    };
    
    if (data.workflowName === 'SA') {
      // SA WORKFLOW EMAIL BUILDING
      let goldenSetContent = '';
      if (data.goldenSet.preAppeal && data.goldenSet.postAppeal) {
          const getScoreStyle = (score) => {
              const numScore = parseFloat(score);
              if (numScore >= 95) return { bg: '#e6f4ea', text: '#188038' };
              if (numScore >= 90) return { bg: '#fff7e0', text: '#b06000' };
              return { bg: '#fce8e6', text: '#d93025' };
          };
          const preStyle = getScoreStyle(data.goldenSet.preAppeal);
          const postStyle = getScoreStyle(data.goldenSet.postAppeal);
          const cycleDate = new Date(data.goldenSet.date + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
          const optionalCommentHtml = data.goldenSet.optionalComment ? `<div style="margin-top: 15px; padding: 10px; background-color: #f8f9fa; border-left: 3px solid #ccc; border-radius: 4px; white-space: pre-wrap;">${data.goldenSet.optionalComment}</div>` : '';
          goldenSetContent = `<p style="margin: 0 0 15px 0; text-align: center; color: #5f6368; font-size: 17px;">Results for the cycle from <strong>${cycleDate}</strong></p><table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 14px;"><tr><td align="center" width="48%" style="background-color: ${preStyle.bg}; border-radius: 8px; padding: 8px;"><p style="margin: 0; font-weight: 500; color: ${preStyle.text}; font-size: 12px;">Pre-appeal score</p><p style="margin: 2px 0 0 0; font-weight: bold; color: ${preStyle.text}; font-size: 20px;">${data.goldenSet.preAppeal}%</p></td><td width="4%"></td><td align="center" width="48%" style="background-color: ${postStyle.bg}; border-radius: 8px; padding: 8px;"><p style="margin: 0; font-weight: 500; color: ${postStyle.text}; font-size: 12px;">Post-appeal score</p><p style="margin: 2px 0 0 0; font-weight: bold; color: ${postStyle.text}; font-size: 20px;">${data.goldenSet.postAppeal}%</p></td></tr></table>${optionalCommentHtml}`;
      }
      const createCaseRows = (cases, rationaleLabel) => {
          if (!cases || cases.length === 0) return '';
          return cases.map((caseInfo, index) => {
              const linkInfo = createSALink(caseInfo);
              const borderStyle = index > 0 ? 'border-top: 1px solid #e8eaed; padding-top: 15px; margin-top: 15px;' : '';
              
              // If there's a link, make it clickable; otherwise just show the ID
              const idDisplay = linkInfo.link 
                ? `<a href="${linkInfo.link}" style="color: #1a73e8; text-decoration: none;">${linkInfo.displayId}</a>`
                : `<span style="color: #3c4043;">${linkInfo.displayId}</span>`;
              
              return `<div style="${borderStyle}"><p style="margin: 5px 0;"><strong>ID:</strong> ${idDisplay}</p><p style="margin: 10px 0 5px 0;"><strong>${rationaleLabel}:</strong></p><div style="padding: 12px; background-color: #f7f9fc; border-left: 3px solid ${corporateBlue}; white-space: pre-wrap; font-family: 'Menlo', 'Courier New', monospace; font-size: 13px; border-radius: 4px;">${caseInfo.rationale}</div>${(caseInfo.additionalInfo && caseInfo.additionalInfo.trim() !== '<div><br></div>' && caseInfo.additionalInfo.trim() !== '' && caseInfo.additionalInfo.trim() !== '<br>') ? `<p style="margin: 10px 0 5px 0;"><strong>Additional Information:</strong></p><div style="padding: 10px; background-color: #f8f9fa; border-left: 3px solid #ccc; border-radius: 4px;">${caseInfo.additionalInfo}</div>` : ''}</div>`;
          }).join('');
      };
      let qaClarificationContent = '';
      if (data.qaClarification && data.qaClarification.trim() !== '') qaClarificationContent += `<p><strong>QA question:</strong><br><div style="padding: 12px; background-color: #f8f9fa; border-radius: 4px; margin-top: 5px; white-space: pre-wrap;">${data.qaClarification}</div></p>`;
      if (data.fteAnswer && data.fteAnswer.trim() !== '' && data.fteAnswer.trim() !== '<div><br></div>' && data.fteAnswer.trim() !== '<br>') {
          qaClarificationContent += `<p style="margin-top: 15px;"><strong>FTE answer:</strong></p><div style="padding: 12px; background-color: #f7f9fc; border: 1px solid #e0e0e0; border-radius: 4px; margin-top: 5px;">${data.fteAnswer}</div>`;
      }
      const additionalMessageContent = data.additionalMessage.text;
      const createCalibrationCaseRows = (cases) => {
          if (!cases || cases.length === 0) return '';
          return cases.map((caseInfo, index) => {
              const linkInfo = createSALink(caseInfo);
              const borderStyle = index > 0 ? 'border-top: 1px solid #e8eaed; padding-top: 15px; margin-top: 15px;' : '';
              
              // If there's a link, make it clickable; otherwise just show the ID
              const idDisplay = linkInfo.link 
                ? `<a href="${linkInfo.link}" style="color: #1a73e8; text-decoration: none;">${linkInfo.displayId}</a>`
                : `<span style="color: #3c4043;">${linkInfo.displayId}</span>`;
              
              return `<div style="${borderStyle}"><p style="margin: 5px 0;"><strong>ID:</strong> ${idDisplay}</p><p style="margin: 10px 0 5px 0;"><strong>Analysis:</strong></p><div style="padding: 12px; background-color: #f7f9fc; border: 1px solid #e0e0e0; border-radius: 4px;">${caseInfo.analysis}</div></div>`;
          }).join('');
      };
      return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>body { font-family: 'Roboto', Arial, sans-serif; margin: 0; padding: 0; }</style></head><body><div style="font-family: 'Roboto', Arial, sans-serif; background-color: #f4f7f6; padding: 20px;"><div style="max-width: 1100px; margin: auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #dadce0;"><div style="text-align: center; padding: 15px; background-color: #f8f9fa; border-bottom: 1px solid #dadce0;"><img src="${CONFIG.URL_LOGO}" alt="Cognizant Logo" style="height: 20px; width: auto; margin-bottom: 10px;"><h2 style="margin: 0; color: #3c4043; font-weight: 500; font-size: 20px;">${headerTitle}</h2></div><div style="padding: 20px 30px 30px 30px;"><p style="color: #3c4043; font-size: 15px;">${data.template.greeting}</p><p style="color: #5f6368; font-size: 15px; line-height: 1.6;">This email contains a summary from the meeting held on <strong>${meetingDate}</strong> regarding "<strong>${data.calendarEvent.title}</strong>".</p><p style="color: #5f6368; font-size: 15px; line-height: 1.6;">${data.template.intro}</p>${createSectionCard('Golden Set Results', goldenSetContent)}${createSectionCard('FTE Calibrated', createCaseRows(data.fteCases, 'FTE Rationale'))}${createSectionCard('Clarifications', qaClarificationContent)}${createSectionCard('Calibration Cases', createCalibrationCaseRows(data.calibrationCases))}${createSectionCard(data.additionalMessage.title || 'Additional Message', additionalMessageContent)}<div style="border-top: 1px solid #e8eaed; margin-top: 30px; padding-top: 20px; color: #5f6368; font-size: 14px;"><p style="margin: 0;">${data.template.closing.replace('Internal Trix', `<a href="${internalTrixLink}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">Internal Trix</a>`)}</p><p style="margin: 5px 0 0 0;">${data.template.footer}</p></div></div></div></div></body></html>`;
    } else if (data.workflowName === 'MMI') {
      // MMI WORKFLOW EMAIL BUILDING
      let goldenSetContent = '';
      if (data.goldenSet.preAppeal && data.goldenSet.postAppeal) {
          const getScoreStyle = (score) => {
              const numScore = parseFloat(score);
              if (numScore >= 95) return { bg: '#e6f4ea', text: '#188038' };
              if (numScore >= 90) return { bg: '#fff7e0', text: '#b06000' };
              return { bg: '#fce8e6', text: '#d93025' };
          };
          const preStyle = getScoreStyle(data.goldenSet.preAppeal);
          const postStyle = getScoreStyle(data.goldenSet.postAppeal);
          const cycleDate = new Date(data.goldenSet.date + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
          const optionalCommentHtml = data.goldenSet.optionalComment ? `<div style="margin-top: 15px; padding: 10px; background-color: #f8f9fa; border-left: 3px solid #ccc; border-radius: 4px; white-space: pre-wrap;">${data.goldenSet.optionalComment}</div>` : '';
          goldenSetContent = `<p style="margin: 0 0 15px 0; text-align: center; color: #5f6368; font-size: 17px;">Results for the cycle from <strong>${cycleDate}</strong></p><table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-size: 14px;"><tr><td align="center" width="48%" style="background-color: ${preStyle.bg}; border-radius: 8px; padding: 8px;"><p style="margin: 0; font-weight: 500; color: ${preStyle.text}; font-size: 12px;">Pre-appeal score</p><p style="margin: 2px 0 0 0; font-weight: bold; color: ${preStyle.text}; font-size: 20px;">${data.goldenSet.preAppeal}%</p></td><td width="4%"></td><td align="center" width="48%" style="background-color: ${postStyle.bg}; border-radius: 8px; padding: 8px;"><p style="margin: 0; font-weight: 500; color: ${postStyle.text}; font-size: 12px;">Post-appeal score</p><p style="margin: 2px 0 0 0; font-weight: bold; color: ${postStyle.text}; font-size: 20px;">${data.goldenSet.postAppeal}%</p></td></tr></table>${optionalCommentHtml}`;
      }
      const createMMIKeyCaseRows = (cases) => {
          if (!cases || cases.length === 0) return '';
          return cases.map((caseInfo, index) => {
              // Handle video ID/URL
              let videoId = caseInfo.id;
              let originalUrl = caseInfo.originalUrl || '';
              let idPill = '';
              let idTypeIndicator = '';
              
              // Priority 1: If we have an original URL, use it for the hyperlink
              if (originalUrl && originalUrl.includes('yurt.corp.google.com')) {
                idPill = `<a href="${originalUrl}" style="color: #1a73e8; text-decoration: none; font-weight: 500; font-family: monospace;">${videoId}</a>`;
                idTypeIndicator = '<span style="font-size: 11px; color: #188038; margin-left: 6px;">Yurt Link</span>';
              } else if (videoId.includes('yurt.corp.google.com')) {
                // User provided a full Yurt URL directly in the ID field
                const match = videoId.match(/entity_id=([^&]+)/);
                if (match) {
                  const extractedId = decodeURIComponent(match[1]);
                  idPill = `<a href="${videoId}" style="color: #1a73e8; text-decoration: none; font-weight: 500; font-family: monospace;">${extractedId}</a>`;
                  idTypeIndicator = '<span style="font-size: 11px; color: #188038; margin-left: 6px;">Yurt Link</span>';
                } else {
                  // Fallback - use entire URL as display text
                  idPill = `<a href="${videoId}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">${videoId}</a>`;
                  idTypeIndicator = '<span style="font-size: 11px; color: #188038; margin-left: 6px;">Link</span>';
                }
              } else if (videoId.includes('http') || (originalUrl && originalUrl.includes('http'))) {
                // Some other HTTP link - create clickable link
                const linkUrl = originalUrl || videoId;
                idPill = `<a href="${linkUrl}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">${videoId}</a>`;
                idTypeIndicator = '<span style="font-size: 11px; color: #1967d2; margin-left: 6px;">External Link</span>';
              } else {
                // Simple ID (from sheet or manual entry) - display as plain text, NO LINK
                idPill = `<span style="font-family: monospace; color: #5f6368; font-weight: 500;">${videoId}</span>`;
                idTypeIndicator = '<span style="font-size: 11px; color: #5f6368; margin-left: 6px;">Video ID</span>';
              }
              
              // Build details HTML with better horizontal space usage
              let detailsHtml = '';
              
              // ID Pill at top
              detailsHtml += `<div style="background-color: #e8f0fe; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 13px; margin-bottom: 10px; border: 1px solid #d2e3fc;">${idPill}${idTypeIndicator}</div>`;
              
              if (caseInfo.policyArea) {
                const style = getPolicyAreaStyle(caseInfo.policyArea);
                detailsHtml += `<span style="background-color: ${style.bg}; display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; color: ${style.text}; font-weight: 500; margin-left: 8px; margin-bottom: 10px;">Policy: ${caseInfo.policyArea}</span>`;
              }
              detailsHtml += '<br>';
              
              // Build metadata box from structured data (policyId, timestamp - NOT from question text)
              const metadataParts = [];
              if (caseInfo.policyId) metadataParts.push(`Policy ID: ${caseInfo.policyId}`);
              if (caseInfo.timestamp) metadataParts.push(`Timestamp: ${caseInfo.timestamp}`);
              if (caseInfo.whatPart) metadataParts.push(`What part: ${caseInfo.whatPart}`);
              
              if (metadataParts.length > 0) {
                const metadataContent = metadataParts.join('<br>');
                detailsHtml += `<span style="background-color: #e8f0fe; display: inline-block; padding: 8px 12px; border-radius: 6px; border: 1px solid #d2e3fc; margin-bottom: 10px; font-size: 13px; color: #1967d2; line-height: 1.6;">${metadataContent}</span><br>`;
              }
              
              if (caseInfo.question) {
                  // Convert \n to <br> and remove escaped backslashes
                  const formattedQuestion = caseInfo.question.replace(/\n/g, '<br>').replace(/\\/g, '');
                  detailsHtml += `<div style="background-color: #f8f9fa; padding: 12px; border-radius: 6px; border-left: 4px solid #3498db; margin-bottom: 10px;"><strong style="color: #2c3e50; display: block; margin-bottom: 6px;">Question:</strong><span style="color: #5f6368; line-height: 1.6;">${formattedQuestion}</span></div>`;
              }
              if (caseInfo.answer) {
                  // Convert \n to <br> and remove escaped backslashes
                  const formattedAnswer = caseInfo.answer.replace(/\n/g, '<br>').replace(/\\/g, '');
                  detailsHtml += `<div style="background-color: #f8f9fa; padding: 12px; border-radius: 6px; border-left: 4px solid #34a853; margin-bottom: 10px;"><strong style="color: #2c3e50; display: block; margin-bottom: 6px;">Answer:</strong><span style="color: #5f6368; line-height: 1.6;">${formattedAnswer}</span></div>`;
              }
              if (caseInfo.decision) {
                  const style = getFinalDecisionStyle(caseInfo.decision);
                  detailsHtml += `<p style="margin: 5px 0 0 0;"><strong style="color: #1a73e8;">Decision:</strong> <span style="background-color: ${style.bg}; color: ${style.text}; padding: 3px 8px; border-radius: 4px; font-weight: 600;">${caseInfo.decision}</span></p>`;
              }
              return `<tr style="border-bottom: 2px solid #e8eaed;">
                        <td style="padding: 16px; vertical-align: top;">${detailsHtml || '<span style="color: #9aa0a6;">No details provided</span>'}</td>
                      </tr>`;
          }).join('');
      };
      
      const createGoldenSetCaseRows = (cases) => {
          if (!cases || cases.length === 0) return '';
          return cases.map((caseInfo, index) => {
              const caseNumber = index + 1;
              let videoId = caseInfo.videoId;
              let videoUrl = '';
              
              // Create video URL
              if (caseInfo.originalUrl) {
                videoUrl = caseInfo.originalUrl;
              } else if (videoId) {
                videoUrl = `https://yurt.corp.google.com/?entity_id=${encodeURIComponent(videoId)}&entity_type=VIDEO&config_id=prod%2Freview_session%2Fvideo%2Fstandard_lookup`;
              }
              
              const videoLink = videoUrl ? `<a href="${videoUrl}" style="color: #1a73e8; text-decoration: none; font-weight: 600; font-size: 15px;">[Video] ${videoId}</a>` : `<span style="color: #5f6368; font-weight: 600;">[Video] ${videoId}</span>`;
              
              const policyAreaHtml = caseInfo.policyArea ? (() => {
                const style = getPolicyAreaStyle(caseInfo.policyArea);
                return `<div style="background-color: ${style.bg}; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 13px; color: ${style.text}; font-weight: 500; margin-top: 8px;">Policy: ${caseInfo.policyArea}</div>`;
              })() : '';
              
              const adminDecisionHtml = caseInfo.adminDecision ? `<div style="background-color: #fff3e0; padding: 12px; border-radius: 8px; border-left: 4px solid #ff9800; margin-top: 12px;"><strong style="color: #e65100; font-size: 13px;">[!] Admin Decision:</strong> <span style="color: #202124; font-weight: 600; font-size: 14px;">${caseInfo.adminDecision}</span></div>` : '';
              
              const correctDecisionHtml = caseInfo.correctDecision ? (() => {
                const style = getFinalDecisionStyle(caseInfo.correctDecision);
                return `<div style="background-color: ${style.bg}; padding: 12px; border-radius: 8px; border-left: 4px solid ${style.borderColor}; margin-top: 8px;"><strong style="color: ${style.text}; font-size: 13px;">[✓] Correct Decision:</strong> <span style="color: #202124; font-weight: 600; font-size: 14px;">${caseInfo.correctDecision}</span></div>`;
              })() : '';
              
              const rationaleHtml = caseInfo.rationale ? `<div style="background-color: #fafafa; padding: 14px; border-radius: 8px; border: 1px solid #e8eaed; margin-top: 12px;"><strong style="color: #3c4043; font-size: 13px; display: block; margin-bottom: 8px;">Rationale:</strong><div style="color: #5f6368; font-size: 14px; line-height: 1.6;">${caseInfo.rationale}</div></div>` : '';
              
              return `<div style="background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); border: 2px solid #1a73e8; border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(26,115,232,0.15);">
                        <div style="border-bottom: 2px solid #1a73e8; padding-bottom: 12px; margin-bottom: 16px;">
                          <h4 style="margin: 0; color: #1a73e8; font-size: 16px; font-weight: 600;">[GS] Golden Set Case #${caseNumber}</h4>
                        </div>
                        <div style="margin-bottom: 12px;">
                          ${videoLink}
                          ${policyAreaHtml}
                        </div>
                        ${adminDecisionHtml}
                        ${correctDecisionHtml}
                        ${rationaleHtml}
                      </div>`;
          }).join('');
      };
      
      let goldenSetCasesSectionHtml = '';
      if (data.goldenSetCases && data.goldenSetCases.length > 0) {
          goldenSetCasesSectionHtml = `<h4 style="color: #202124; margin-top: 20px; font-size: 18px; border-bottom: 2px solid #1a73e8; padding-bottom: 8px;">Golden Set Cases</h4><div style="margin-top: 16px;">${createGoldenSetCaseRows(data.goldenSetCases)}</div>`;
      }
      
      let keyCasesSectionHtml = '';
      if (data.keyCases && data.keyCases.length > 0) {
          keyCasesSectionHtml = `<h4 style="color: #202124; margin-top: 20px; font-size: 18px; border-bottom: 2px solid #34a853; padding-bottom: 8px;">Key Cases Discussed</h4><table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 14px; margin-top: 16px; border: 1px solid #e8eaed; border-radius: 8px; overflow: hidden;"><tbody>${createMMIKeyCaseRows(data.keyCases)}</tbody></table>`;
      }
      
      let whatDiscussedContent = data.whatDiscussed || '';
      let concerningIssuesContent = data.concerningIssues ? `<p style="white-space: pre-wrap; margin: 0;">${data.concerningIssues}</p>` : '';
      
      // Additional Message (if provided)
      const additionalMessageContent = (data.additionalMessage && data.additionalMessage.text) ? data.additionalMessage.text : '';
      
      // For MMI, closing doesn't have Internal Trix link
      const closingText = data.template.closing;
      
      return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>body { font-family: 'Roboto', Arial, sans-serif; margin: 0; padding: 0; }</style></head><body><div style="font-family: 'Roboto', Arial, sans-serif; background-color: #f4f7f6; padding: 20px;"><div style="max-width: 1100px; margin: auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #dadce0;"><div style="text-align: center; padding: 15px; background-color: #f8f9fa; border-bottom: 1px solid #dadce0;"><img src="${CONFIG.URL_LOGO}" alt="Cognizant Logo" style="height: 20px; width: auto; margin-bottom: 10px;"><h2 style="margin: 0; color: #3c4043; font-weight: 500; font-size: 20px;">Meeting Summary: MMI</h2></div><div style="padding: 20px 30px 30px 30px;"><p style="color: #3c4043; font-size: 15px;">${data.template.greeting}</p><p style="color: #5f6368; font-size: 15px; line-height: 1.6;">This email summarizes the meeting held on <strong>${meetingDate}</strong> regarding "<strong>${data.calendarEvent.title}</strong>".</p>${createSectionCard('Golden Set Results', goldenSetContent)}${createSectionCard('What was discussed?', whatDiscussedContent)}${createSectionCard('Golden Set Cases', goldenSetCasesSectionHtml)}${createSectionCard('Key Cases Discussed', keyCasesSectionHtml)}${createSectionCard('Concerning Issues / Comments', concerningIssuesContent)}${createSectionCard(data.additionalMessage.title || 'Additional Message', additionalMessageContent)}<div style="border-top: 1px solid #e8eaed; margin-top: 30px; padding-top: 20px; color: #5f6368; font-size: 14px;"><p style="margin: 0;">${closingText}</p><p style="margin: 5px 0 0 0;">Thank you,</p></div><div style="background-color: #f8f9fa; padding: 16px; text-align: center; font-size: 12px; color: #5f6368; border-top: 1px solid #ddd;"><p style="margin: 0;">This auto-generated summary was created by the QA Team.</p></div></div></div></div></body></html>`;
    }
    
    console.error(' CRITICAL: Workflow not recognized! workflowName = "' + data.workflowName + '"');
    return '';
    } catch (error) {
      console.error(' Error in buildAdminModernEmailHtml:', error.message);
      Logger.log('Error in buildAdminModernEmailHtml: ' + error.message);
      throw error;
    }
}

// --- EMAIL HTML BUILDER ---
function buildModernEmailHtml(data, meetingTitle, meetingDate, pems) {
    const pemMap = pems.reduce((map, pem) => { map[pem.ldap] = pem.fullName; return map; }, {});
    const attendeeNames = data.selectedPems.map(ldap => `${pemMap[ldap] || ldap} (${ldap})`).join(', ');
    
    const headerTitle = `${data.workflow.toUpperCase()} Youtube T&S`;

    let workflowDisplay = data.workflow;
    if (data.vertical) {
      workflowDisplay += ` - ${data.vertical}`;
    }

    // Get the base URL for row links (custom or default)
    const baseUrlData = getSupportRowLinkBaseUrl();
    const baseUrl = baseUrlData.url;

    let keyCasesSectionHtml = '';
    if (data.keyCases && data.keyCases.length > 0) {
      const casesRowsHtml = data.keyCases.map(caseInfo => {
        // Build ID pill with type indicator
        let videoId = caseInfo.id;
        let idPill = '';
        let idTypeIndicator = '';
        
        if (videoId.includes('yurt.corp.google.com')) {
          const match = videoId.match(/entity_id=([^&]+)/);
          if (match) {
            const extractedId = decodeURIComponent(match[1]);
            idPill = `<a href="${videoId}" style="color: #1a73e8; text-decoration: none; font-weight: 500; font-family: monospace;">${extractedId}</a>`;
            idTypeIndicator = '<span style="font-size: 11px; color: #188038; margin-left: 6px;">Yurt Link</span>';
          } else {
            idPill = `<a href="${videoId}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">${videoId}</a>`;
            idTypeIndicator = '<span style="font-size: 11px; color: #188038; margin-left: 6px;">Link</span>';
          }
        } else if (videoId.includes('http')) {
          idPill = `<a href="${videoId}" style="color: #1a73e8; text-decoration: none; font-weight: 500;">${videoId}</a>`;
          idTypeIndicator = '<span style="font-size: 11px; color: #1967d2; margin-left: 6px;">External Link</span>';
        } else {
          idPill = `<span style="font-family: monospace; color: #5f6368; font-weight: 500;">${videoId}</span>`;
          idTypeIndicator = '<span style="font-size: 11px; color: #5f6368; margin-left: 6px;">Video ID</span>';
        }
        
        // Build row link pill if available
        let rowPill = '';
        if (caseInfo.row) {
          const rowLink = `${baseUrl}&range=${caseInfo.row}:${caseInfo.row}`;
          rowPill = `<div style="background-color: #fff3e0; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 12px; margin-right: 8px; margin-bottom: 10px; border: 1px solid #ffe0b2;"><span style="color: #e65100; font-weight: 500;">Row:</span> <a href="${rowLink}" style="color: #1a73e8; text-decoration: none; font-weight: bold;">${caseInfo.row}</a></div>`;
        }
        
        // Build metadata pills (Policy ID, Timestamp)
        let metadataPills = '';
        if (caseInfo.policyId) {
          metadataPills += `<div style="background-color: #e8f0fe; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 12px; margin-right: 8px; margin-bottom: 10px; border: 1px solid #d2e3fc;"><span style="color: #1967d2; font-weight: 500;">Policy ID:</span> <span style="color: #1967d2;">${caseInfo.policyId}</span></div>`;
        }
        if (caseInfo.timestamp) {
          metadataPills += `<div style="background-color: #e8f0fe; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 12px; margin-right: 8px; margin-bottom: 10px; border: 1px solid #d2e3fc;"><span style="color: #1967d2; font-weight: 500;">Timestamp:</span> <span style="color: #1967d2;">${caseInfo.timestamp}</span></div>`;
        }
        
        // Build content HTML
        let contentHtml = '';
        
        // Pills at top: Row (if exists) + ID + Policy ID + Timestamp
        if (rowPill) contentHtml += rowPill;
        contentHtml += `<div style="background-color: #e8f0fe; display: inline-block; padding: 6px 14px; border-radius: 16px; font-size: 13px; margin-right: 8px; margin-bottom: 10px; border: 1px solid #d2e3fc;">${idPill}${idTypeIndicator}</div>`;
        if (metadataPills) contentHtml += metadataPills;
        
        if (caseInfo.policyArea) {
          const style = getPolicyAreaStyle(caseInfo.policyArea);
          contentHtml += `<span style="background-color: ${style.bg}; display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; color: ${style.text}; font-weight: 500; margin-bottom: 10px;">Policy: ${caseInfo.policyArea}</span>`;
        }
        contentHtml += '<br>';
        
        if (caseInfo.question) { 
          const formattedQuestion = caseInfo.question.replace(/\n/g, '<br>').replace(/\\/g, '');
          contentHtml += `<div style="background-color: #f8f9fa; padding: 12px; border-radius: 6px; border-left: 4px solid #3498db; margin-bottom: 10px;"><strong style="color: #2c3e50; display: block; margin-bottom: 6px;">Question:</strong><span style="color: #5f6368; line-height: 1.6;">${formattedQuestion}</span></div>`; 
        }
        if (caseInfo.answer) { 
          const clarifier = caseInfo.clarifyingPem ? ` from ${pemMap[caseInfo.clarifyingPem]}` : ''; 
          const formattedAnswer = caseInfo.answer.replace(/\n/g, '<br>').replace(/\\/g, '');
          contentHtml += `<div style="background-color: #f8f9fa; padding: 12px; border-radius: 6px; border-left: 4px solid #34a853; margin-bottom: 10px;"><strong style="color: #2c3e50; display: block; margin-bottom: 6px;">Answer${clarifier}:</strong><span style="color: #5f6368; line-height: 1.6;">${formattedAnswer}</span></div>`; 
        }
        
        // Final Decision goes AFTER the answer
        if (caseInfo.decision) {
            const style = getFinalDecisionStyle(caseInfo.decision);
            contentHtml += `<p style="margin:5px 0 0 0;"><strong>Final Decision:</strong> <span style="background-color: ${style.bg}; color: ${style.text}; padding: 3px 8px; border-radius: 4px; font-weight: 600;">${caseInfo.decision}</span></p>`;
        }
        
        return `<tr style="border-bottom: 2px solid #e0e0e0;">
                  <td style="padding: 16px; vertical-align: top;">${contentHtml}</td>
                </tr>`;
      }).join('');

      keyCasesSectionHtml = `
        <h4 style="color: #202124; margin-top: 20px; font-size: 18px; border-bottom: 2px solid #34a853; padding-bottom: 8px;">Key Cases Discussed</h4>
        <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 14px; margin-top: 16px; border: 1px solid #e8eaed; border-radius: 8px; overflow: hidden;">
          <tbody>${casesRowsHtml}</tbody>
        </table>
      `;
    }

    let followUpSectionHtml = '';
    if (data.followUp === 'Yes' && data.followUpContext) {
        followUpSectionHtml = `
            <h4 style="color: #202124; margin-top: 20px;">Follow-up Required?</h4>
            <p><strong>Yes</strong></p>
            <p style="font-weight:bold; margin-bottom:5px; margin-top:15px;">Reason to follow up:</p>
            <div style="background-color: #f1f3f4; padding: 12px; border-radius: 4px; border: 1px solid #e0e0e0;">
                <p style="margin:0; white-space: pre-wrap;">${data.followUpContext}</p>
            </div>
        `;
    }

    return `
      <div style="font-family: Arial, Helvetica, sans-serif; color: #333; max-width: 1100px; margin: auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #f8f9fa; padding: 16px; text-align: center; border-bottom: 1px solid #ddd;">
          <img src="${CONFIG.URL_LOGO}" alt="Cognizant Logo" style="width: 120px; height: auto;">
          <h3 style="margin: 10px 0 0 0; color: #5f6368; font-weight: 500;">${headerTitle}</h3>
        </div>
        <div style="padding: 24px;">
          <h2 style="color: #202124; font-size: 22px;">Meeting Summary: ${workflowDisplay}</h2>
          <p>Hi Team,</p>
          <p>This email summarizes the meeting held on <strong>${meetingDate}</strong> regarding <strong>${meetingTitle}</strong>.</p>
          <h3 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 5px; margin-top: 30px;">Meeting Details</h3>
          <p><strong>Call conducted by (PEMs):</strong> ${attendeeNames}</p>
          <h3 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 5px; margin-top: 30px;">Discussion Overview</h3>
          ${data.geminiSummary ? `<h4 style="color: #202124; margin-bottom: 5px;">What was discussed?</h4><div style="background-color: #f1f3f4; padding: 12px; border-radius: 4px; border: 1px solid #e0e0e0;"><p style="margin:0; white-space: pre-wrap;">${data.geminiSummary}</p></div>` : ''}
          
          ${keyCasesSectionHtml}
          
          ${data.concerningIssues ? `<h4 style="color: #202124; margin-top: 20px;">Concerning Issues / Comments</h4><p style="white-space: pre-wrap;">${data.concerningIssues}</p>` : ''}
          
          ${followUpSectionHtml}
          
          <hr style="border: none; border-top: 1px solid #eee; margin: 30px 0;">
          <p style="margin: 0;">For any questions, please contact the QA PoCs.</p>
          <p style="margin: 5px 0 0 0;">Thank you,</p>
        </div>
        <div style="background-color: #f8f9fa; padding: 16px; text-align: center; font-size: 12px; color: #5f6368; border-top: 1px solid #ddd;">
          <p>This auto-generated summary was created by the QA Team.</p>
        </div>
      </div>
    `;
}

// --- UTILITY FUNCTIONS ---
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// Test function for backend connectivity
function testBackendFunction() {
  return {
    status: 'success',
    timestamp: new Date().toISOString(),
    spreadsheet: ss.getName(),
    sheets: ss.getSheets().map(sheet => sheet.getName())
  };
}

// Diagnostic function to check sheet contents
function diagnoseSheets() {
  console.log(' Starting sheet diagnosis...');

  const requiredSheets = ['Workflows', 'EmailTemplates', 'PEM ldaps', 'POCs', 'Managers'];
  const diagnosis = {
    spreadsheet: ss.getName(),
    sheets: {}
  };

  requiredSheets.forEach(sheetName => {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        diagnosis.sheets[sheetName] = { status: 'NOT_FOUND' };
        console.error(` Sheet '${sheetName}' not found`);
        return;
      }

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues() : [];

      diagnosis.sheets[sheetName] = {
        status: 'OK',
        rows: lastRow,
        columns: lastColumn,
        dataRows: data.length,
        hasHeaders: lastRow >= 1,
        sampleData: data.length > 0 ? data.slice(0, 2) : []
      };

      console.log(` Sheet '${sheetName}': ${lastRow} rows, ${lastColumn} columns, ${data.length} data rows`);

      if (data.length > 0) {
        console.log(` Sample data from '${sheetName}':`, data[0]);
      }

    } catch (err) {
      diagnosis.sheets[sheetName] = { status: 'ERROR', error: err.message };
      console.error(` Error diagnosing sheet '${sheetName}':`, err.message);
    }
  });

  return diagnosis;
}

// Function to populate sample data for testing
function populateSampleData() {

  try {
    // Sample Workflows data
    let sheet = ss.getSheetByName('Workflows');
    if (!sheet) {
      sheet = ss.insertSheet('Workflows');
    }
    if (sheet.getLastRow() < 2) {
      sheet.clear();
      sheet.appendRow(['Workflow Name', 'Recipient Spreadsheet ID', 'Recipient Sheet Name', 'Recipient Range']);
      sheet.appendRow(['SA', '1abc123def456', 'Recipients', 'A2:A10']);
      sheet.appendRow(['MMI', '1abc123def456', 'Recipients', 'B2:B10']);
    }

    // Sample EmailTemplates data
    sheet = ss.getSheetByName('EmailTemplates');
    if (!sheet) {
      sheet = ss.insertSheet('EmailTemplates');
    }
    if (sheet.getLastRow() < 2) {
      sheet.clear();
      sheet.appendRow(['Template Name', 'Greeting', 'Introduction', 'Closing', 'Footer']);
      sheet.appendRow(['Default', 'Dear Team,', 'Please find below the feedback summary.', 'Best regards,', 'Quality Assurance Team']);
    }

    // Sample PEM ldaps data
    sheet = ss.getSheetByName('PEM ldaps');
    if (!sheet) {
      sheet = ss.insertSheet('PEM ldaps');
    }
    if (sheet.getLastRow() < 2) {
      sheet.clear();
      sheet.appendRow(['LDAP', 'Full Name', 'Workflow']);
      sheet.appendRow(['user1', 'John Doe', 'SA,MMI']);
      sheet.appendRow(['user2', 'Jane Smith', 'GMI']);
    }

    // Sample POCs data
    sheet = ss.getSheetByName('POCs');
    if (!sheet) {
      sheet = ss.insertSheet('POCs');
    }
    if (sheet.getLastRow() < 2) {
      sheet.clear();
      sheet.appendRow(['Email', 'Workflow', 'Vertical']);
      sheet.appendRow(['poc1@company.com', 'SA', 'ALL']);
      sheet.appendRow(['poc2@company.com', 'MMI', 'SSL']);
    }

    // Sample Managers data
    sheet = ss.getSheetByName('Managers');
    if (!sheet) {
      sheet = ss.insertSheet('Managers');
    }
    if (sheet.getLastRow() < 2) {
      sheet.clear();
      sheet.appendRow(['Email']);
      sheet.appendRow(['manager1@company.com']);
      sheet.appendRow(['manager2@company.com']);
    }

    console.log(' Sample data population completed!');
    return { status: 'success', message: 'Sample data added to empty sheets' };

  } catch (err) {
    console.error(' Error populating sample data:', err);
    return { status: 'error', message: err.message };
  }
}

// --- SUPPORT TEAM ROW LINK BASE URL MANAGEMENT ---
function getSupportRowLinkBaseUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    const customUrl = props.getProperty('SUPPORT_ROW_LINK_BASE_URL');
    
    if (customUrl) {
      return { url: customUrl, isCustom: true };
    } else {
      // Return default URL
      return { url: CONFIG.SHEET_URL_BASE, isCustom: false };
    }
  } catch (err) {
    Logger.log('Error getting support row link base URL: ' + err.message);
    return { url: CONFIG.SHEET_URL_BASE, isCustom: false, error: err.message };
  }
}

function setSupportRowLinkBaseUrl(newUrl) {
  try {
    if (!newUrl || newUrl.trim() === '') {
      throw new Error('URL cannot be empty');
    }
    
    // Basic validation - check if it looks like a Google Sheets URL
    if (!newUrl.includes('docs.google.com/spreadsheets')) {
      throw new Error('Invalid URL. Must be a Google Sheets URL.');
    }
    
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SUPPORT_ROW_LINK_BASE_URL', newUrl.trim());
    
    Logger.log('Support row link base URL updated to: ' + newUrl);
    return { status: 'success', message: 'Base URL updated successfully', url: newUrl };
  } catch (err) {
    Logger.log('Error setting support row link base URL: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

function resetSupportRowLinkBaseUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('SUPPORT_ROW_LINK_BASE_URL');
    
    Logger.log('Support row link base URL reset to default');
    return { status: 'success', message: 'Reset to default URL', url: CONFIG.SHEET_URL_BASE };
  } catch (err) {
    Logger.log('Error resetting support row link base URL: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

// ==================== CLARIFICATION SHEET URL MANAGEMENT ====================

function getClarificationSheetUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    const customUrl = props.getProperty('CLARIFICATION_SHEET_URL');
    
    if (customUrl) {
      return { url: customUrl, isCustom: true };
    } else {
      // Return default URL
      const defaultUrl = 'https://docs.google.com/spreadsheets/d/1aQ_fD7BMpi7Ba7wCAzQLfdgpEfu79CA6AnpTgouOFI4/edit?gid=1763023571#gid=1763023571';
      return { url: defaultUrl, isCustom: false };
    }
  } catch (err) {
    Logger.log('Error getting clarification sheet URL: ' + err.message);
    const defaultUrl = 'https://docs.google.com/spreadsheets/d/1aQ_fD7BMpi7Ba7wCAzQLfdgpEfu79CA6AnpTgouOFI4/edit?gid=1763023571#gid=1763023571';
    return { url: defaultUrl, isCustom: false, error: err.message };
  }
}

function setClarificationSheetUrl(newUrl) {
  try {
    if (!newUrl || newUrl.trim() === '') {
      throw new Error('URL cannot be empty');
    }
    
    // Basic validation - check if it looks like a Google Sheets URL
    if (!newUrl.includes('docs.google.com/spreadsheets')) {
      throw new Error('Invalid URL. Must be a Google Sheets URL.');
    }
    
    const props = PropertiesService.getScriptProperties();
    props.setProperty('CLARIFICATION_SHEET_URL', newUrl.trim());
    
    Logger.log('Clarification sheet URL updated to: ' + newUrl);
    return { status: 'success', message: 'Clarification sheet URL updated successfully', url: newUrl };
  } catch (err) {
    Logger.log('Error setting clarification sheet URL: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

function resetClarificationSheetUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('CLARIFICATION_SHEET_URL');
    
    const defaultUrl = 'https://docs.google.com/spreadsheets/d/1aQ_fD7BMpi7Ba7wCAzQLfdgpEfu79CA6AnpTgouOFI4/edit?gid=1763023571#gid=1763023571';
    Logger.log('Clarification sheet URL reset to default');
    return { status: 'success', message: 'Reset to default URL', url: defaultUrl };
  } catch (err) {
    Logger.log('Error resetting clarification sheet URL: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

// Function to extract spreadsheet ID and gid from URL
function extractSheetInfo(url) {
  const spreadsheetIdMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  const gidMatch = url.match(/[#&]gid=([0-9]+)/);
  
  if (!spreadsheetIdMatch) {
    throw new Error('Could not extract spreadsheet ID from URL');
  }
  
  return {
    spreadsheetId: spreadsheetIdMatch[1],
    gid: gidMatch ? gidMatch[1] : '0'
  };
}

// Function to fetch data from clarification sheet by row number
function fetchClarificationData(rowNumber) {
  try {
    Logger.log('=== fetchClarificationData called with row: ' + rowNumber + ' ===');
    
    if (!rowNumber || isNaN(rowNumber) || rowNumber < 2) {
      throw new Error('Invalid row number. Must be 2 or greater.');
    }
    
    const urlData = getClarificationSheetUrl();
    const sheetInfo = extractSheetInfo(urlData.url);
    
    Logger.log('Opening spreadsheet: ' + sheetInfo.spreadsheetId + ', gid: ' + sheetInfo.gid);
    
    const spreadsheet = SpreadsheetApp.openById(sheetInfo.spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    // Find sheet by gid
    let targetSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId().toString() === sheetInfo.gid) {
        targetSheet = sheets[i];
        break;
      }
    }
    
    if (!targetSheet) {
      // If gid not found, use first sheet
      targetSheet = sheets[0];
      Logger.log('Could not find sheet with gid ' + sheetInfo.gid + ', using first sheet: ' + targetSheet.getName());
    }
    
    const lastRow = targetSheet.getLastRow();
    if (rowNumber > lastRow) {
      throw new Error('Row number ' + rowNumber + ' exceeds last row (' + lastRow + ') in sheet');
    }
    
    // Get header row (row 1)
    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    
    // Get data row
    const rowData = targetSheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    
    Logger.log('Headers: ' + JSON.stringify(headers));
    Logger.log('Row data: ' + JSON.stringify(rowData));
    Logger.log('Row data with indices:');
    for (let i = 0; i < Math.min(rowData.length, 25); i++) {
      Logger.log('  [' + i + '] = ' + JSON.stringify(rowData[i]));
    }
    
    // Create a helper function to find column index by partial header match
    function findColumnIndex(searchTerms) {
      for (let term of searchTerms) {
        for (let i = 0; i < headers.length; i++) {
          const header = String(headers[i]).toLowerCase();
          if (header.includes(term.toLowerCase())) {
            return i;
          }
        }
      }
      return -1;
    }
    
    // Try to find columns by header name, with fallback to common indices
    const entityIdIdx = findColumnIndex(['entity id', 'entity', 'video id']);
    const policyAreaIdx = findColumnIndex(['policy area', 'policy']);
    const whatPartIdx = findColumnIndex(['what part']);
    const policyIdIdx = findColumnIndex(['policy id related', 'related policy']);
    const timestampIdx = findColumnIndex(['timestamp', 'violative']);
    const detailedQuestionIdx = findColumnIndex(['detailed question', 'grey area']);
    const pemLdapIdx = findColumnIndex(['pem ldap', 'ldap']);
    const policyIdFutureIdx = findColumnIndex(['policy id\n(future', 'future addition']);
    const pemResponseIdx = findColumnIndex(['pem response']);
    const pemSessionDecisionIdx = findColumnIndex(['pem session decision', 'session decision']);
    
    Logger.log('Found column indices: entityId=' + entityIdIdx + ', policyArea=' + policyAreaIdx + ', detailedQuestion=' + detailedQuestionIdx);
    
    // Extract data using found indices, with hardcoded fallbacks based on your sheet structure
    // Column mapping (0-indexed):
    // F (5) = Entity ID (always)
    // I (8) = Policy Area
    // J (9) = Policy ID
    // K (10) = Timestamp
    // L (11) = Detailed Question
    // P (15) = PEM LDAP
    // Q (16) = Policy ID Future/Decision
    // R (17) = PEM Response
    
    const result = {
      entityId: String((entityIdIdx >= 0 ? rowData[entityIdIdx] : rowData[5]) || ''),
      policyArea: String((policyAreaIdx >= 0 ? rowData[policyAreaIdx] : rowData[8]) || ''),
      whatPart: String((whatPartIdx >= 0 ? rowData[whatPartIdx] : '') || ''),
      policyId: String((policyIdIdx >= 0 ? rowData[policyIdIdx] : rowData[9]) || ''),
      timestamp: String((timestampIdx >= 0 ? rowData[timestampIdx] : rowData[10]) || ''),
      detailedQuestion: String((detailedQuestionIdx >= 0 ? rowData[detailedQuestionIdx] : rowData[11]) || ''),
      pemLdap: String((pemLdapIdx >= 0 ? rowData[pemLdapIdx] : rowData[15]) || ''),
      policyIdFuture: String((policyIdFutureIdx >= 0 ? rowData[policyIdFutureIdx] : rowData[16]) || ''),
      pemResponse: String((pemResponseIdx >= 0 ? rowData[pemResponseIdx] : rowData[17]) || ''),
      pemSessionDecision: String((pemSessionDecisionIdx >= 0 ? rowData[pemSessionDecisionIdx] : rowData[17]) || '')
    };
    
    Logger.log('Extracted data BEFORE return: ' + JSON.stringify(result));
    
    Logger.log('Extracted data: ' + JSON.stringify(result));
    
    return { status: 'success', data: result };
    
  } catch (err) {
    Logger.log('ERROR in fetchClarificationData: ' + err.message);
    Logger.log('Stack: ' + err.stack);
    return { status: 'error', message: err.message };
  }
}

// Fuzzy string matching for policy areas
function fuzzyMatchPolicyArea(extracted, availableOptions) {
  if (!extracted || !availableOptions || availableOptions.length === 0) {
    return null;
  }
  
  const extractedLower = extracted.toLowerCase().trim();
  
  // Exact match
  for (let option of availableOptions) {
    if (option.toLowerCase() === extractedLower) {
      return option;
    }
  }
  
  // Partial match - check if extracted contains option or vice versa
  for (let option of availableOptions) {
    const optionLower = option.toLowerCase();
    if (extractedLower.includes(optionLower) || optionLower.includes(extractedLower)) {
      return option;
    }
  }
  
  // Similarity scoring using Levenshtein-like approach
  let bestMatch = null;
  let bestScore = 0;
  
  for (let option of availableOptions) {
    const score = calculateSimilarity(extractedLower, option.toLowerCase());
    if (score > bestScore && score > 0.5) { // 50% similarity threshold
      bestScore = score;
      bestMatch = option;
    }
  }
  
  return bestMatch;
}

// Simple similarity calculation (Dice coefficient)
function calculateSimilarity(str1, str2) {
  const bigrams1 = getBigrams(str1);
  const bigrams2 = getBigrams(str2);
  
  if (bigrams1.length === 0 || bigrams2.length === 0) {
    return 0;
  }
  
  let intersection = 0;
  for (let bigram of bigrams1) {
    if (bigrams2.indexOf(bigram) !== -1) {
      intersection++;
    }
  }
  
  return (2.0 * intersection) / (bigrams1.length + bigrams2.length);
}

function getBigrams(str) {
  const bigrams = [];
  for (let i = 0; i < str.length - 1; i++) {
    bigrams.push(str.substring(i, i + 2));
  }
  return bigrams;
}

// ==================== EMAIL HISTORY FUNCTIONS ====================

function getAdminEmailHistory() {
  try {
    const sheet = ss.getSheetByName('Meeting record');
    if (!sheet) {
      Logger.log('Meeting record sheet not found');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in Meeting record sheet');
      return [];
    }
    
    // Get up to 7 columns: Timestamp, Workflow, Subject, To, CC, HTML Body, Sent By
    const lastCol = Math.min(7, sheet.getLastColumn());
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Map data to objects and reverse for chronological order (newest first)
    const history = data.map((row, index) => ({
      id: index + 2,
      timestamp: row[0] ? new Date(row[0]).toLocaleString() : 'N/A',
      workflow: row[1] || '',
      subject: row[2] || '',
      to: row[3] || '',
      cc: row[4] || '',
      ldap: row[6] ? row[6] : 'Unknown' // Column 7 - Sent By
    })).reverse();
    
    Logger.log('Retrieved ' + history.length + ' admin emails from history');
    return history;
  } catch (err) {
    Logger.log('Error getting admin email history: ' + err.message);
    return [];
  }
}
