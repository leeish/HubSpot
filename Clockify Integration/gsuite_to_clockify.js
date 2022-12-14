/**
 * This is a Google Apps Script, to be deployed within a Google sheet container.
 * These sheets are assumed to exist: "email_sent", "internal_recurrences", "meetings_booked", "customers", "config", "sku_prioritization"
 * Access to Calendar API needs to be added as a service.
 */

const version_string = "1.0.0"

//////
// Globals
//////

let config_map = undefined

const all_tags = {
  'Customer Book of Business Management': '624bb7d3a5b26c4f53265350',
  'Customer Call': '624bb7efa5b26c4f53265358',
  'Customer Deliverables': '624bb7e5a5b26c4f53265356',
  'Customer Escalations': '624bb7f6a5b26c4f5326535a',
  'Customer Flex Coverage': '624bb87fa5b26c4f5326536f',
  'Customer Prep/Follow Up': '624bb7d9a5b26c4f53265352',
  'Customer Scoping Projects': '624bb82ca5b26c4f53265365',
  'Customer Training Creation': '624bb876a5b26c4f5326536b',
  'Customer Training Live': '624bb865a5b26c4f53265369',
  'Customer Troubleshooting': '624bb7fba5b26c4f5326535c',
  'EMEA SQL': '6308cc4eea4dff0f3fba82f9',
  'HubSpot Impact': '624bb85fa5b26c4f53265367',
  'Integrations Project': '630fb42aea4dff0f3fbbae7c',
  'Internal Administration': '624bb879a5b26c4f5326536d',
  'Internal Meeting & Collaboration': '62599ff6a5b26c4f53272802',
  'Mentoring': '62599ffaa5b26c4f53272807',
  'OOO': '624bb7dfa5b26c4f53265354',
  'Professional Development & Training': '6259a004a5b26c4f53272812',
  'PSO': '62ab4b86750c033c9dde9280',
  'Special Projects': '6259a00ea5b26c4f5327281a',
}

//TODO add these to config_map or somewhere else instead
let min_adjusted_meeting_length = 0.5 // fraction
let max_meeting_start_delay = 0.33 // fraction
let min_email_minutes = 5 // minutes
let max_email_overlap = 3 // minutes, should be smaller than min_email_minutes

function update_globals() {
  if (config_map == null) {
    config_map = getConfig();
    let r = jsonResponse(UrlFetchApp.fetch("https://hubspot.clockify.me/api/v1/user", { 'headers': { 'x-api-key': config_map["clockify_key"] } }));
    config_map["user_id"] = r["id"];
    config_map["workspace_id"] = r["defaultWorkspace"];
  }
}


//////
// GAS events and UI elements and wrapped functions
//////

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    //{name: 'Validate content (placeholder)', functionName: 'validateSheet_'},
    { name: 'Refresh email and calendar logs', functionName: 'refreshSheet_' },
    { name: 'Write logs to Clockify', functionName: 'logActivities_' },
    { name: 'Get project/client/task IDs', functionName: 'fetchServices_' },
    { name: 'Update internal recurring meetings', functionName: 'updateInternals_' }
  ];
  spreadsheet.addMenu('Clockifyiable_Activities', menuItems);
}

function refreshSheet_() {
  writeRecentSentEmail();
  writeRecentMeetings();
}

function logActivities_() {
  log_all_activities();
}

function updateInternals_() {
  updateInternalRecurrences(32);
}

function dailyLogging() {
  updateInternalRecurrences();
  writeRecentSentEmail();
  writeRecentMeetings();
  log_all_activities();
}

function log_all_activities() {
  update_globals()
  var logged_intervals = get_intervals(30)
  var start_intervals = logged_intervals.length
  var start_coverage = logged_intervals.reduce((a, b) => a + b[1] - b[0], 0);
  log_meetings(logged_intervals);
  var logged_meetings = logged_intervals.length - start_intervals
  if (config_map["email_max_minutes"] != 0) {
    log_emails(logged_intervals);
  }
  let data = {
    "email": config_map["sender_email"],
    "eventName": "pe3967897_clockify_gsuite_sync",
    "properties": {
      "logged_time_entries": logged_intervals.length - start_intervals,
      "ps_role": config_map["role"],
      "version_string": version_string,
      "logged_email": logged_intervals.length - start_intervals - logged_meetings,
      "logged_minutes": Math.round(logged_intervals.reduce((a, b) => a + b[1] - b[0], -start_coverage) / (1000 * 60)),
    }
  }
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data),
    'headers': { 'authorization': "Bearer " + config_map["hs_analytics_token"] }
  };
  let response = UrlFetchApp.fetch("https://api.hubapi.com/events/v3/send", options);
}

function fetchServices_() {
  enrich_customers();
}

function enrich_customers() {
  let matched_projects = getServices();
  matchCustomerProjects(matched_projects);
}

//////
// Sheet interaction functions (GAS v1)
//////

function getConfig() {
  let config_map = {}
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("config");
  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();
  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 0; i < values.length; i++) {
    config_map[values[i][0]] = values[i][1];
  }
  if (config_map.hasOwnProperty("autoresponder_subject_strings")) {
    config_map["autoresponder_subject_strings"] = config_map["autoresponder_subject_strings"].replace(";", ",").split(",").map(x => x.trim().toLowerCase())
  }
  return config_map
}

function getServices() {
  update_globals();
  var project_requests = []
  // TODO figure out how on earth project count has crept up to > 10k
  for (var i = 0; i < 10; i++) {
    project_requests.push(
      {
        'url': 'https://hubspot.clockify.me/api/v1/workspaces/' + config_map["workspace_id"] + '/projects?page-size=3000&archived=false&page=' + i.toString() + '&hydrated=true',
        'headers': { 'x-api-key': config_map['clockify_key'] }
      }
    )
  }
  var project_batches = UrlFetchApp.fetchAll(project_requests);
  var projects_by_hid = {};
  for (let element of getHIDs()) {
    projects_by_hid[element] = []
  }
  var task = {}
  for (let outer_index = 0; outer_index < project_batches.length; outer_index++) {
    let batch_of_projects = JSON.parse(project_batches[outer_index].getContentText());
    for (let index = 0; index < batch_of_projects.length; index++) {
      let project = batch_of_projects[index];
      for (let task_index = 0; task_index < project["tasks"].length; task_index++) {
        task = project["tasks"][task_index]
        if (task["name"] in projects_by_hid && !projects_by_hid[task["name"]].map(x => (x["project"] == project["id"])).some(x => x)) {
          projects_by_hid[task["name"]].push({
            "client": project["clientId"],
            "sku": project["name"],
            "project": project["id"],
            "task": task["id"]
          })
        }
      }
    }
  }
  return projects_by_hid;
}

function getPriorityMap() {
  update_globals();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("sku_prioritization");
  let range = sheet.getDataRange();
  let values = range.getValues();
  let prio_col = undefined
  switch (config_map["role"]) {
    case "TC":
      prio_col = 1; break;
    case "IC":
      prio_col = 2; break;
    case "CT":
      prio_col = 3; break;
    case "ONB":
      prio_col = 4; break;
  }
  service_priorities = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][prio_col] != "") {
      service_priorities[values[i][0]] = values[i][prio_col];
    }
  }
  return service_priorities;
}

function matchCustomerProjects(matched_projects) {
  let service_priorities = getPriorityMap();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getSheetByName("customers");
  range = sheet.getDataRange();
  values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    let hid = Math.trunc(values[i][1]).toString()
    let identified = false;
    let chosen_index = undefined
    let current_values_valid = false;
    for (var j = 0; j < matched_projects[hid].length; j++) {
      if (matched_projects[hid][j]["sku"] == values[i][2] && matched_projects[hid][j]["project"] == values[i][4] && matched_projects[hid][j]["task"] == values[i][5]) {
        current_values_valid = true;
        identified = false;
        break;
      } else if (service_priorities.hasOwnProperty(matched_projects[hid][j]["sku"])) {
        let priority = service_priorities[matched_projects[hid][j]["sku"]]
        if (typeof chosen_index == "undefined" || priority > service_priorities[matched_projects[hid][chosen_index]["sku"]]) {
          chosen_index = j
          identified = true;
        } else if (priority == service_priorities[matched_projects[hid][chosen_index]["sku"]]) {
          identified = false;
          break;
        }
      }
    }
    if (identified) {
      values[i][2] = matched_projects[hid][chosen_index]["sku"];
      values[i][3] = matched_projects[hid][chosen_index]["client"];
      values[i][4] = matched_projects[hid][chosen_index]["project"];
      values[i][5] = matched_projects[hid][chosen_index]["task"];
      values[i][6] = JSON.stringify(matched_projects[hid]);
    } else if (current_values_valid) {
      values[i][6] = JSON.stringify(matched_projects[hid]);
    } else {
      values[i][2] = "";
      values[i][3] = "";
      values[i][4] = "";
      values[i][5] = "";
      values[i][6] = JSON.stringify(matched_projects[hid]);
    }
  }
  range.setValues(values);
}

function getHIDs() {
  let hid_array = []
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customers");
  // This represents ALL the data
  let range = sheet.getDataRange();
  let values = range.getValues();
  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 1; i < values.length; i++) {
    hid_array.push(Math.trunc(values[i][1]).toString())
  }
  return hid_array;
}

function domainToIds(domains) {
  if (typeof domains === 'string' || domains instanceof String) {
    domains = [domains]
  }
  domains = domains.map(x => x.toLowerCase());
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customers");
  let range = sheet.getDataRange();
  let values = range.getValues();
  let domain_map = {}
  for (var i = 1; i < values.length; i++) {
    let all_domains = values[i][0].replace(";", ",").toLowerCase().split(",")
    for (let j = 0; j < all_domains.length; j++) {
      let domain = all_domains[j].trim();
      domain_map[domain] = {
        client_id: values[i][3],
        project_id: values[i][4],
        task_id: values[i][5],
        hid: values[i][1]
      }
    }
  }
  let matched_client = "";
  let matched_project = "";
  let matched_task = "";
  let matching_success = false;
  let matched_domains = [];
  for (let domain of domains) {
    if (domain_map[domain]) {
      if (!matching_success) {
        matched_client = domain_map[domain]["client_id"]
        matched_project = domain_map[domain]["project_id"]
        matched_task = domain_map[domain]["task_id"]
        matched_hid = domain_map[domain]["hid"]
        matching_success = true;
        matched_domains.push(domain);
      } else if (domain_map[domain]["project_id"] != matched_project) {
        matching_success = false;
        break;
      } else {
        matched_domains.push(domain);
      }
    }
  }
  if (matching_success && matched_project != "" && matched_task != "") {
    return [matched_client, matched_project, matched_task, matched_hid, matched_domains.join(";")]
  } else {
    return [null, null, null, null, null]
  }
}

function getRecurrenceTagMap() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("internal_recurrences");
  let tag_map = {}
  if (sheet.getLastRow() > 1) {
    recurrences = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2).getValues();
    for (let recurrence of recurrences) {
      if (all_tags[recurrence[1]]) {
        tag_map[recurrence[0]] = all_tags[recurrence[1]]
      }
    }
  }
  return tag_map;
}


function extractEmailAddresses(string) {
  // via https://www.weirdgeek.com/2019/10/regular-expression-in-google-apps-script/ and https://stackoverflow.com/questions/42407785/regex-extract-email-from-strings
  // cf. https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/match
  var regExp = new RegExp("([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", "gi");
  var results = string.match(regExp);
  return results;
}

function writeRecentSentEmail() {
  update_globals();
  // https://developers.google.com/apps-script/reference/gmail
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("email_sent");
  sheet.clear();
  sheet.appendRow(["send_timestamp", "subject", "recipient_domains", "client_id", "project_id", "task_id", "hid", "matched_domains"]);
  let threads = GmailApp.search("in:sent", 0, 100);
  for (var i = 0; i < threads.length; i++) {
    let messages = threads[i].getMessages();
    for (var j = messages.length - 1; j >= 0; j--) {
      if (messages[j].getFrom().includes(config_map["sender_email"])) {
        let message_date = messages[j].getDate();
        if ((Date.now() - message_date) / (1000 * 60 * 60) < 24) {
          let message_subject = messages[j].getSubject()
          //TODO make the following exclusion strings part of the config sheet
          let out_of_office = false;
          for (let substring of config_map["autoresponder_subject_strings"]) {
            if (message_subject.toLowerCase().includes(substring)) {
              out_of_office = true;
            }
          }
          if (message_subject && !out_of_office) {
            let message_recipients = messages[j].getTo();
            let message_cc = messages[j].getCc();
            if (message_cc.length > 0) {
              message_recipients = message_recipients + ", " + message_cc
            }
            let recipients = extractEmailAddresses(message_recipients);
            let recipient_domains = []
            for (var k = 0; k < recipients.length; k++) {
              let recipient_domain = recipients[k].split("@")[1];
              if (!recipient_domains.includes(recipient_domain) && recipient_domain != "hubspot.com" && recipient_domain != "gmail.com" && !recipient_domain.includes("google.com")) {
                recipient_domains.push(recipient_domain);
              }
            }
            let matchedIds = domainToIds(recipient_domains)
            if (matchedIds[1]) {
              sheet.appendRow([message_date.getTime(), sanitize(message_subject.toLowerCase()), recipient_domains.join(";"), matchedIds[0], matchedIds[1], matchedIds[2], matchedIds[3], matchedIds[4]]);
            }
          }
        }
      }
    }
  }
}

function writeRecentMeetings() {
  update_globals();
  // https://developers.google.com/apps-script/guides/services/advanced
  // https://developers.google.com/calendar/api/v3/reference/events 
  // unfortunately couldn't use https://developers.google.com/apps-script/reference/calendar/calendar-app since it doesn't return "decline" status for an event owner
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("meetings_booked");
  sheet.clear();
  sheet.appendRow(["start_timestamp", "end_timestamp", "event_summary", "recipient_domains", "client_id", "project_id", "task_id", "hid", "matched_domains", "tag_id"]);
  let calendarId = 'primary';
  let timeUpperBound = new Date(Date.now()).toISOString()
  if (config_map["log_ahead"] == "yes") {
    timeUpperBound = new Date(Date.now() + (24 * 60 * 60 * 1000)).toISOString()
  }
  let events = Calendar.Events.list(calendarId, {
    timeMin: new Date(Date.now() - (24 * 60 * 60 * 1000)).toISOString(),
    timeMax: timeUpperBound,
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    let recurrence_tags = getRecurrenceTagMap();
    for (var i = 0; i < events.items.length; i++) {
      let event = events.items[i];
      // check and log internal, then CONTINUE
      if (event.recurringEventId && recurrence_tags[event.recurringEventId.split("_")[0]]) {
        sheet.appendRow([
          Date.parse(event.start.dateTime),
          Date.parse(event.end.dateTime),
          sanitize(event.summary.toLowerCase()),
          "hubspot.com",
          "",
          "624bb17fa5b26c4f532652fd",
          "62696d5d750c033c9dd5b598",
          53,
          "hubspot.com",
          recurrence_tags[event.recurringEventId.split("_")[0]],
        ]);
        continue;
      }
      if (!event.start.date) {
        log_event = true;
        let event_domains = []
        if (event.attendees && event.attendees.length > 0) {
          for (var k = 0; k < event.attendees.length; k++) {
            let attendee = event.attendees[k]
            if (attendee.self) {
              if (attendee.responseStatus != "accepted") {
                log_event = false;
              }
            } else {
              let attendee_domain = attendee.email.split("@")[1]
              if (!event_domains.includes(attendee_domain) && attendee_domain != "hubspot.com" && attendee_domain != "gmail.com" && !attendee_domain.includes("google.com")) {
                event_domains.push(attendee_domain);
              }
            }
          }
        }
        let matchedIds = domainToIds(event_domains);
        if (matchedIds[1] && log_event) {
          let event_start = Date.parse(event.start.dateTime);
          let event_end = Date.parse(event.end.dateTime);
          sheet.appendRow([
            event_start,
            event_end,
            sanitize(event.summary.toLowerCase()),
            event_domains.join(";"),
            matchedIds[0],
            matchedIds[1],
            matchedIds[2],
            matchedIds[3],
            matchedIds[4],
            all_tags["Customer Call"]
          ]);
        }
      }
    }
  }
}

function updateInternalRecurrences(days = 1) {
  update_globals();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("internal_recurrences");
  let interval_days = 32;
  let processed_events = new Set()
  if (sheet.getLastRow() > 1) {
    processed_events = new Set(sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat());
    interval_days = days;
  }
  let events = Calendar.Events.list('primary', {
    timeMin: new Date(Date.now() - (interval_days * 24 * 60 * 60 * 1000)).toISOString(),
    timeMax: new Date().toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 500
  });
  if (events.items && events.items.length > 0) {
    own_handle = config_map["sender_email"].split("@")[0].toLowerCase()
    for (var i = 0; i < events.items.length; i++) {
      let event = events.items[i];
      //Logger.log(event.summary)
      if (!event.start.date && event.attendees && event.recurringEventId && !processed_events.has(event.recurringEventId.split("_")[0])) {
        processed_events.add(event.recurringEventId.split("_")[0])
        //this is where the categorization happens
        let candidate_event = true;
        let hs_attendees = []
        event.attendees.forEach(att => {
          let [handle, domain] = att.email.split("@")
          if (domain == "hubspot.com") {
            if (handle.toLowerCase() != own_handle) hs_attendees.push(handle)
          } else if (domain != "gmail.com" && !domain.includes("google.com")) {
            candidate_event = false;
          }
        })
        hs_attendees.sort();
        if (hs_attendees.length == 0 && event.organizer.email == config_map["sender_email"] && event.creator.email == config_map["sender_email"]) {
          candidate_event = false
        }
        if (candidate_event) {
          let attendee_string = hs_attendees.slice(0, 3).join(", ")
          if (hs_attendees.length > 4) attendee_string += " (+" + (hs_attendees.length - 3) + " more)"
          sheet.appendRow([
            event.summary,
            attendee_string,
            event.recurringEventId.split("_")[0],
          ])
        }
      }
    }
  }
}



/**
 * This is basically a GAS rewrite of the original Python code
 */

//////
// Utility functions (v2)
//////

function jsonResponse(response) {
  return JSON.parse(response.getContentText());
}

function get_intervals(minus_x_hours = 96) {
  update_globals();
  let lower_bound = Math.floor(Date.now() - minus_x_hours * 60 * 60 * 1000);
  var newDate = new Date();
  newDate.setTime(lower_bound);
  dateString = newDate.toUTCString();
  let page_size = 0;
  let completed = false;
  let intervals = [];
  while (page_size < 1000 && completed == false) {
    page_size += 50;
    let url = "https://hubspot.clockify.me/api/v1/workspaces/" + config_map["workspace_id"] + "/user/" + config_map["user_id"] + "/time-entries?page-size=" + page_size
    let r = UrlFetchApp.fetch(url, { 'headers': { 'x-api-key': config_map["clockify_key"] } });
    my_time_entries = jsonResponse(r);
    for (let time_entry of my_time_entries) {
      time_start = Date.parse(time_entry["timeInterval"]["start"]);
      time_end = Date.parse(time_entry["timeInterval"]["end"]);
      if (time_end > lower_bound) {
        intervals.push([time_start, time_end]);
      } else {
        completed = true;
        break;
      }
    }
  }
  return intervals;
}

function sanitize(description) {
  description = description.replace(/[^a-zA-Z0-9]/g, " ");
  while (description.includes("  ")) {
    description = description.replace("  ", " ");
  }
  return description.trim();
}

function log_activity(from_timestamp, to_timestamp, description, project_id, tag_list, billable, task_id) {
  update_globals();
  let from_isoZ = new Date(from_timestamp).toISOString();
  let to_isoZ = new Date(to_timestamp).toISOString();
  let data = {
    "start": from_isoZ,
    "end": to_isoZ,
    "billable": billable,
    "projectId": project_id.toString(),
    "tagIds": tag_list,
    "description": description,
    "taskId": task_id.toString()
  };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data),
    'headers': { 'x-api-key': config_map["clockify_key"] }
  };
  let response = UrlFetchApp.fetch("https://hubspot.clockify.me/api/v1/workspaces/" + config_map["workspace_id"] + "/time-entries", options);
  if (response.getResponseCode() == 201) {
    return true;
  } else {
    Logger.log("failed to create time entry (" + description + ")");
    return false;
  }
}

function effective_meeting_times(from_timestamp, to_timestamp, logged_intervals) {
  update_globals()
  //from_timestamp = parseInt(from_timestamp);
  //to_timestamp = parseInt(to_timestamp);
  var skip = false;
  var original_length = to_timestamp - from_timestamp;
  latest_start_date = from_timestamp + original_length * max_meeting_start_delay;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (latest_start_date > logged_intervals[i][1] && logged_intervals[i][1] > from_timestamp) {
      from_timestamp = logged_intervals[i][1];
    }
    if (from_timestamp < logged_intervals[i][0] && logged_intervals[i][0] < to_timestamp) {
      to_timestamp = logged_intervals[i][0];
    }
  }
  if ((to_timestamp - from_timestamp) < original_length * min_adjusted_meeting_length) {
    skip = true;
  }
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < to_timestamp && logged_intervals[i][1] > from_timestamp) {
      skip = true;
    }
  }
  if (skip) {
    return [null, null];
  } else {
    return [from_timestamp, to_timestamp];
  }
}

function effective_email_times(send_timestamp, logged_intervals) {
  update_globals()
  send_timestamp = parseInt(send_timestamp);
  skip = false;
  upper_bound = send_timestamp;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][1] > send_timestamp && logged_intervals[i][0] < upper_bound) {
      upper_bound = logged_intervals[i][0];
    }
  }
  upper_bound = Math.min(upper_bound, send_timestamp);
  lower_bound = upper_bound - config_map["email_max_minutes"] * 60 * 1000;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < upper_bound && logged_intervals[i][1] > lower_bound) {
      lower_bound = logged_intervals[i][1];
    }
  }
  if ((upper_bound - lower_bound) * 60 * 1000 < min_email_minutes) {
    skip = true;
  }
  if ((send_timestamp - upper_bound) * 60 * 1000 > max_email_overlap) {
    skip = true;
  }
  // this loop may be redundant, not sure
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < upper_bound && logged_intervals[i][1] > lower_bound) {
      skip = true;
    }
  }
  if (skip) {
    return [null, null];
  } else {
    return [lower_bound, upper_bound];
  }
}

//////
// main interface (v2)
//////

function log_meetings(logged_intervals = [], silent = false) {
  update_globals()
  var prep_time_max = config_map["call_prep_minutes"];
  var post_time_max = config_map["call_post_minutes"];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("meetings_booked");
  let range = sheet.getDataRange();
  let values = range.getValues();
  // TODO exclude meetings everybody but yourself has declined?
  var padding_intervals = [];
  for (let i = 1; i < values.length; i++) {
    var row = {
      "start_timestamp": values[i][0],
      "end_timestamp": values[i][1],
      "project": values[i][5],
      "event_summary": values[i][2],
      "hid": values[i][7],
      "task_id": values[i][6],
      "tag_id": values[i][9],
    }
    let internal = false;
    let prefix = "CALL "
    let billable = true
    if (row["hid"] == 53) {
      internal = true;
      prefix = "INTERNAL ";
      billable = false;
    }
    if (row["project"] && row["project"] != "") {
      let from_timestamp, to_timestamp;
      [from_timestamp, to_timestamp] = effective_meeting_times(row['start_timestamp'], row['end_timestamp'], logged_intervals);
      if (from_timestamp && to_timestamp && row['project']) {
        var r = log_activity(from_timestamp, to_timestamp, prefix + row['event_summary'], row['project'], [row["tag_id"]], billable, row['task_id']);
        if (r) {
          logged_intervals.push([from_timestamp, to_timestamp]);
          if (!silent) {
            Logger.log("Logged event (" + Math.round((to_timestamp - from_timestamp) / (1000 * 60)) + "min) " + "\"" + row['event_summary'] + "\" to " + row['hid'].toString());
          }
          if (!internal) {
            padding_intervals.push({
              "from": from_timestamp,
              "to": to_timestamp,
              "summary": row['event_summary'],
              "project": row['project'],
              "task": row['task_id'],
            })
          }
        } else {
          Logger.log("FAILED to log event \"" + row['event_summary'] + "\" to " + row['hid'].toString());
        }
      } else {
        Logger.log("Cannot log event \"" + row['event_summary'] + "\" to " + row['hid'].toString() + " (coincides with logged activity)");
      }
    }
  }
  for (let i = 0; i < padding_intervals.length; i++) {
    let item = padding_intervals[i];
    // prep_call_time
    let prep_from, prep_to;
    [prep_from, prep_to] = effective_meeting_times(item["from"] - prep_time_max * 1000 * 60, item["from"], logged_intervals);
    if (prep_to == item["from"] && (prep_to - prep_from) / (1000 * 60) > prep_time_max / 2) {
      r = log_activity(prep_from, prep_to, "PREP " + item["summary"], item['project'], [all_tags["Customer Prep/Follow Up"]], true, item["task"]);
      if (r) {
        logged_intervals.push([prep_from, prep_to]);
      }
    }
    // post_call_time
    let post_from, post_to;
    [post_from, post_to] = effective_meeting_times(item["to"], item["to"] + post_time_max * 1000 * 60, logged_intervals);
    if (post_from == item["to"] && (post_to - post_from) / (1000 * 60) > post_time_max / 2) {
      r = log_activity(post_from, post_to, "POST " + item["summary"], item['project'], [all_tags["Customer Prep/Follow Up"]], true, item['task']);
      if (r) {
        logged_intervals.push([post_from, post_to])
      }
    }
  }
}

function log_emails(logged_intervals = [], silent = false) {
  update_globals()
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("email_sent");
  let range = sheet.getDataRange();
  let values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    var row = {
      "send_timestamp": values[i][0],
      "project": values[i][4],
      "subject": values[i][1],
      "hid": values[i][6],
      "task_id": values[i][5]
    }
    if (row["project"] && row["project"] != "") {
      let effective_times = effective_email_times(row['send_timestamp'], logged_intervals);
      if (effective_times[0] && effective_times[1] && row['project']) {
        var r = log_activity(effective_times[0], effective_times[1], "EMAIL " + row['subject'], row['project'], [all_tags["Customer Prep/Follow Up"]], true, row['task_id']);
        if (r) {
          if (!silent) {
            Logger.log("Logged email (" + Math.round((effective_times[1] - effective_times[0]) / (1000 * 60)) + "min) " + "\"" + row['subject'] + "\" to " + row['hid'].toString());
          }
          logged_intervals.push([effective_times[0], effective_times[1]]);
        } else {
          Logger.log("FAILED to log email \"" + row['subject'] + "\" to " + row['hid'].toString());
        }
      } else {
        Logger.log("Cannot log email \"" + row['subject'] + "\" to " + row['hid'].toString() + " (coincides with logged activity)");
      }
    }
  }
}
