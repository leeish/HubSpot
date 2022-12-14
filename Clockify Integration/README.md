# Clockify G-Suite Code Documentation

The Clockify G-Suite integration code will check for emails and calendar items from the current user and log those as activites in Clockify if there are openings at those times.

## Table of Contents
- [Globals](#globals)
- [Events](#events)
    - [onOpen](#onopen)
- [Actions](#actions)
    - [refreshSheet](#refreshsheet)
    - [logActivities](#logactivities)
    - [fetchServices](#fetchservices)
    - [updateInternals](#updateinternals)
- [Functions](#functions)
    - [domainToIds](#domaintoids)
    - [effective_email_times](#effective_email_times)
    - [effective_meeting_times](#effective_meeting_times)
    - [enrich_customers](#enrich_customers)
    - [extractEmailAddresses](#extractemailaddresses)
    - [getConfig](#getconfig)
    - [getHIDs](#gethids)
    - [get_intervals](#get_intervals)
    - [getPriorityMap](#getprioritymap)
    - [getRecurrenceTagMap](#getrecurrencetagmap)
    - [getServices](#getservices)
    - [jsonResponse](#jsonresponse)
    - [log_activity](#log_activity)
    - [log_all_activities](#log_all_activities)
    - [log_emails](#log_emails)
    - [log_meetings](#log_meetings)
    - [matchCustomerProjects](#matchcustomerprojects)
    - [sanitize](#sanitize)
    - [updateInternalRecurrences](#updateinternalrecurrences)
    - [writeRecentMeetings](#writerecentmeetings)
    - [writeRecentSentEmail](#writerecentsentemail)
- [Notes](#notes)

## Globals

| Name | Type | Description |
| --- | --- | --- |
| version_string | `String` | Versionsing string |
| config_map | `Object` | Contains config mapping from `config` sheet |
| all_tags | `Object` | Tag dictionary from Clockify |
| min_adjusted_meeting_length | `float` | ... |
| max_meeting_start_delay | `float` | ... | 
| min_email_minutes | `int` | Number of minutes for... |
| max_email_overlap | `int` | Number of minutes for... |

## Events

### onOpen

Populates custom menu in the wookbook.
___
___


## Actions

### refreshSheet

Clears the `email_sent` and `meetings_booked` sheets in the workbook and populates them with the most recent data from Gmail and Google Calendar.

**Dependencies**
- [writeRecentSentEmail](#writeRecentsentemail)
- [writeRecentMeetings](#writerecentmeetings)
___

### logActivities

Log Clockify activities to Viktor's portal.

**Dependencies**
- [log_all_activities](#log_all_activities)
___

### fetchServices

**Dependencies**
- [enrich_customers](#enrich_customers)
___

### updateInternals


**Dependencies**
- [updateInternalRecurrences](#updateinternalrecurrences)
___
___


## Functions

### domainToIds 

| Argument | Type | Description |
| --- | --- | --- |
| domains | `Array` | Array of domain strings used for matching against customers |

Returns a tuple array of matched customer data from a passed in array of domains. All values will be `null` if no match was found.

```js
    [client_id, project_id, task_id, hub_id, domains]
```
___

### effective_email_times

| Argument | Type | Description |
| --- | --- | --- |
| send_timestamp | `int` | Time email was sent |
| logged_intervals | `Array` | Array of tuples that contain two elements [start_timestamp,end_timestamp] which represent existing logged interval of time. |

Returns a single activity start/end time based on a passed in email send time and existing logged activity.
___

### effective_meeting_times

| Argument | Type | Description |
| --- | --- | --- |
| from_timestamp | `int` | Start time of the logged activity |
| to_timestamp | `int` | End time of the logged activity |
| logged_intervals | `Array` | Array of tuples that contain two elements [start_timestamp,end_timestamp] which represent existing logged interval of time. |

Returns a single activity start/end time based on passed in start/end and existing logged activity.
___

### enrich_customers

Fetches the projects from Clockify and then matches them against the `customers` sheet in the workbook.

**Dependencies**
- [getServices](#getservices)
- [matchCustomerProjects](#matchcustomerprojects)
___

### extractEmailAddresses

| Argument | Type | Description |
| --- | --- | --- |
| string | `String` | String in which to search for email addresses |

Returns an array of email addresses from an input string.
___

### getConfig

Reads the `config` sheet and returns the values as a configuration map with key value pairs where the key is the first column and the value is the second value.
___

### getHIDs

Reads the `customers` sheet in the workbook and returns an array of hubIDs.
___

### get_intervals

| Argument | Type | Description |
| --- | --- | --- |
| minux_x_hours | `int` | Lookback for events in hours from the current date/time |

Fetches the most recent time intervals between now and the passed in argument. Returns an `array` of start and end times.
___


### getPriorityMap

Reads the `sku_prioritization` sheet in the workbook and returns the service priorities based on the `role` defined in the `config_map`.
___

### getRecurrenceTagMap

?
___

### getServices

Fetches the projects from Clockify and returns an dictionary of project objects.

```js
{
    project.name: {
        client: project.clientId,
        sku: project.name,
        project: project.id,
        task: task.id
    }
}
```
___

### jsonResponse

| Argument | Type | Description |
| --- | --- | --- |
| response | `Object` | Should be the result of a `UrlFetchApp.fetch` function that has a `getResponseText()` method available. |

Returns a UrlFetchApp response text as a JSON object.
___

### log_activity

| Argument | Type | Description |
| --- | --- | --- |
| from_timestamp | `int` | Start time of the logged activity |
| to_timestamp | `int` | End time of the logged activity |
| description | `String` | Description to log with activity |
| project_id | `UUID` | Project ID in Clockify |
| tag_list | `Array` | Array of Tags to associate with the acitivty (Should limit to 1?) |
| billable | `Boolean` | Indicates if the activity is billable |
| task_id | `UUID` | Task ID in Clockify |


Uses Clockify API to log an activity based on passed in arguments. Returns true on `201` response.
___

### log_all_activities


** Dependencies **
- [get_intervals](#get_intervals)
___

### log_emails

| Argument | Type | Description | Default |
| --- | --- | --- | --- |
| logged_intervals | `Array` | Array of tuples that contain two elements [start_timestamp,end_timestamp] which represent existing logged interval of time. | `[]` | 
| silent | `Boolean` | Turn on logging with a `true` value | `false` | 

Runs through passed in `logged_intervals` and logs an activity in Clockify for each one.
___

### log_meetings

| Argument | Type | Description | Default |
| --- | --- | --- | --- |
| logged_intervals | `Array` | Array of tuples that contain two elements [start_timestamp,end_timestamp] which represent existing logged interval of time. | `[]` | 
| silent | `Boolean` | Turn on logging with a `true` value | `false` | 

Runs through passed in `logged_intervals` and logs an activity in Clockify for each one.
___

### matchCustomerProjects

Compares the passed in projects argument against the `customers` sheet in the workbook and returns matched values.
___

### sanitize

Strings a string of special characters and trims white space.
___

### updateInternalRecurrences

| Argument | Type | Description |
| --- | --- | --- |
| days | `int` | ... |

?
___

### writeRecentMeetings

Fetches recent events from the users Google Calendar and writes them to the `meetings_booked` sheet in the workbook.
___

### writeRecentSentEmail

Fetches recently sent emails from the users Gmail and writes them to the `email_sent` sheet in the workbook.
___
___

## Notes

- Mix of `camelCase` and `snake_case` throughout the document.
- Some variables in `config` sheet and other variables in code.
- `hs_analytics_token` should probably be embedded in code.
- Consider abstracting API calls to single function to separate API logic and processing logic.
- Move [enrich_customers](#enrich_customers) into [fetchServies](#fetchservies) function. ([enrich_customers](#enrich_customers) isn't used anywhere else).
- Many functions rely on [update_globals`](#update_globals) is there a way to refactor to make that cleaner.
- What config values should be optional/required if any and can we set defaults. Currently `autoresponder_subject_strings` is only optional value.
- Is there a way to validate config values and provide error/handling feedback if something bad is put in.
- Could `domains` in [domainToIds](#domaintoids) ever be a string. Judging from the code it appears not.
- [domainToIds](#domaintoids) could return a descriptive object instead of a tuple array.
- Some function names seem to indicate they'll return mutliple results but actually return a single result. [effective_meeting_times](#effective_meeting_times) returns a single meeting window. See also [effective_email_times](#effective_email_times), [domainToIds](#domain-to-ids).
- Is there a possibility to merge [effective_meeting_times](#effective_meeting_times) & [effective_email_times](#effective_email_times) into a single function that essentially handles null `to_timestamp` and uses the `email_max_minutes` instead.
- Should we implement a global debug logger with a single variable across all functions instead of just in [log_meetings](#log_meetings).
- Consider making base URLs globals instead of written each time.
- Need some explinations of some of the sheet configuration items, I think Viktor has this in a doc.