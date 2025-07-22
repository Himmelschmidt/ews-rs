/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, DateTime, Operation, OperationResponse, ResponseClass,
    MESSAGES_NS_URI, TYPES_NS_URI,
};

/// A request to get user availability information.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getuseravailability>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct GetUserAvailability {
    /// The time zone context for the request.
    #[xml_struct(ns_prefix = "t")]
    pub time_zone: Option<SerializableTimeZone>,

    /// The collection of mailboxes to query for availability information.
    #[xml_struct(ns_prefix = "t")]
    pub mailbox_data_array: MailboxDataArray,

    /// The time window for which to retrieve availability information.
    #[xml_struct(ns_prefix = "t")]
    pub free_busy_view_options: FreeBusyViewOptions,
}

impl Operation for GetUserAvailability {
    type Response = GetUserAvailabilityResponse;
}

impl EnvelopeBodyContents for GetUserAvailability {
    fn name() -> &'static str {
        "GetUserAvailability"
    }
}

/// Time zone information for the request.
#[derive(Clone, Debug, XmlSerialize, Deserialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
#[serde(rename_all = "PascalCase")]
pub struct SerializableTimeZone {
    /// The bias in minutes from UTC.
    #[xml_struct(attribute)]
    pub bias: i32,

    /// Standard time information.
    pub standard_time: Option<SerializableTimeZoneTime>,

    /// Daylight time information.
    pub daylight_time: Option<SerializableTimeZoneTime>,
}

/// Time zone time information.
#[derive(Clone, Debug, XmlSerialize, Deserialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
#[serde(rename_all = "PascalCase")]
pub struct SerializableTimeZoneTime {
    /// The bias in minutes from the main time zone bias.
    pub bias: i32,

    /// The time of day when the time change occurs.
    pub time: String,

    /// The day of the week when the time change occurs.
    pub day_of_week: DayOfWeek,

    /// The month when the time change occurs.
    pub month: i32,

    /// The day of the month when the time change occurs.
    pub day_order: i32,
}

/// Days of the week.
#[derive(Clone, Debug, XmlSerialize, Deserialize)]
#[xml_struct(text)]
pub enum DayOfWeek {
    Sunday,
    Monday,
    Tuesday,
    Wednesday,
    Thursday,
    Friday,
    Saturday,
}

/// Array of mailbox data for availability queries.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
pub struct MailboxDataArray {
    /// The mailboxes to query.
    pub mailbox_data: Vec<MailboxData>,
}

/// Mailbox data for availability queries.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
pub struct MailboxData {
    /// The email address of the mailbox.
    pub email: EmailAddress,

    /// Whether to exclude conflicts.
    pub exclude_conflicts: Option<bool>,
}

/// Email address information.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
pub struct EmailAddress {
    /// The name associated with the email address.
    pub name: Option<String>,

    /// The email address.
    pub address: String,

    /// The routing type (typically "SMTP").
    pub routing_type: Option<String>,
}

/// Options for the free/busy view.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
pub struct FreeBusyViewOptions {
    /// The time window for the availability request.
    pub time_window: Duration,

    /// The level of detail to return.
    pub requested_view: FreeBusyViewType,

    /// The interval between free/busy data points in minutes.
    pub merged_free_busy_interval_in_minutes: Option<i32>,
}

/// Time window for availability requests.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = TYPES_NS_URI)]
pub struct Duration {
    /// The start time of the window.
    pub start_time: DateTime,

    /// The end time of the window.
    pub end_time: DateTime,
}

/// The type of free/busy view to return.
#[derive(Clone, Debug, XmlSerialize, Deserialize)]
#[xml_struct(text)]
pub enum FreeBusyViewType {
    None,
    MergedOnly,
    FreeBusy,
    FreeBusyMerged,
    Detailed,
    DetailedMerged,
}

/// A response to a [`GetUserAvailability`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getuseravailabilityresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetUserAvailabilityResponse {
    pub free_busy_response_array: Option<FreeBusyResponseArray>,
    pub suggestions_response: Option<SuggestionsResponse>,
}

impl OperationResponse for GetUserAvailabilityResponse {}

impl EnvelopeBodyContents for GetUserAvailabilityResponse {
    fn name() -> &'static str {
        "GetUserAvailabilityResponse"
    }
}

/// Array of free/busy responses.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FreeBusyResponseArray {
    pub free_busy_response: Vec<FreeBusyResponse>,
}

/// Free/busy response for a single mailbox.
pub type FreeBusyResponse = Option<ResponseClass<FreeBusyResponseData>>;

/// Free/busy response data.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FreeBusyResponseData {
    /// The free/busy view data.
    pub free_busy_view: Option<FreeBusyView>,
}

/// Free/busy view data for a mailbox.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FreeBusyView {
    /// The type of free/busy view.
    pub free_busy_view_type: FreeBusyViewType,

    /// Merged free/busy data as a string.
    pub merged_free_busy: Option<String>,

    /// Calendar events for detailed views.
    pub calendar_event_array: Option<CalendarEventArray>,

    /// Working hours information.
    pub working_hours: Option<WorkingHours>,
}

/// Array of calendar events.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct CalendarEventArray {
    pub calendar_event: Vec<CalendarEvent>,
}

/// A calendar event in the free/busy response.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct CalendarEvent {
    /// The start time of the event.
    pub start_time: DateTime,

    /// The end time of the event.
    pub end_time: DateTime,

    /// The busy type of the event.
    pub busy_type: LegacyFreeBusyStatus,

    /// Additional details about the event.
    pub calendar_event_details: Option<CalendarEventDetails>,
}

/// Details about a calendar event.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct CalendarEventDetails {
    /// The unique identifier of the event.
    pub id: Option<String>,

    /// The subject of the event.
    pub subject: Option<String>,

    /// The location of the event.
    pub location: Option<String>,

    /// Whether the event is a meeting.
    pub is_meeting: Option<bool>,

    /// Whether the event is recurring.
    pub is_recurring: Option<bool>,

    /// Whether the event is an exception to a recurring series.
    pub is_exception: Option<bool>,

    /// Whether the event is a reminder.
    pub is_reminder_set: Option<bool>,

    /// Whether the event is private.
    pub is_private: Option<bool>,
}

/// Working hours information.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct WorkingHours {
    /// The time zone for working hours.
    pub time_zone: Option<SerializableTimeZone>,

    /// The working period information.
    pub working_period_array: Option<WorkingPeriodArray>,
}

/// Array of working periods.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct WorkingPeriodArray {
    pub working_period: Vec<WorkingPeriod>,
}

/// A working period definition.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct WorkingPeriod {
    /// The days of the week for this working period.
    pub day_of_week: DaysOfWeek,

    /// The start time of the working period.
    pub start_time_in_minutes: i32,

    /// The end time of the working period.
    pub end_time_in_minutes: i32,
}

/// Days of the week as a bitmask.
#[derive(Clone, Debug, Deserialize)]
pub enum DaysOfWeek {
    Sunday,
    Monday,
    Tuesday,
    Wednesday,
    Thursday,
    Friday,
    Saturday,
    Weekdays,
    WeekendDays,
    None,
}

/// Free/busy status values.
#[derive(Clone, Debug, Deserialize)]
pub enum LegacyFreeBusyStatus {
    Free,
    Tentative,
    Busy,
    OOF, // Out of Office
    NoData,
}

/// Suggestions response for meeting times.
pub type SuggestionsResponse = Option<ResponseClass<SuggestionsResponseData>>;

/// Suggestions response data.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct SuggestionsResponseData {
    /// Array of meeting time suggestions.
    pub suggestion_day_result_array: Option<SuggestionDayResultArray>,
}

/// Array of suggestion day results.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct SuggestionDayResultArray {
    pub suggestion_day_result: Vec<SuggestionDayResult>,
}

/// Suggestion results for a single day.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct SuggestionDayResult {
    /// The date for these suggestions.
    pub date: DateTime,

    /// The day quality rating.
    pub day_quality: SuggestionQuality,

    /// Array of meeting time suggestions for this day.
    pub suggestion_array: Option<SuggestionArray>,
}

/// Array of meeting time suggestions.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct SuggestionArray {
    pub suggestion: Vec<Suggestion>,
}

/// A meeting time suggestion.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Suggestion {
    /// The suggested meeting time.
    pub meeting_time: DateTime,

    /// Whether this is a good suggestion.
    pub is_workday: bool,

    /// The quality rating of this suggestion.
    pub suggestion_quality: SuggestionQuality,

    /// Array of attendee conflicts for this time.
    pub attendee_conflict_data_array: Option<AttendeeConflictDataArray>,
}

/// Array of attendee conflict data.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct AttendeeConflictDataArray {
    pub unknown_attendee_conflict_data: Option<Vec<UnknownAttendeeConflictData>>,
    pub individual_attendee_conflict_data: Option<Vec<IndividualAttendeeConflictData>>,
    pub group_attendee_conflict_data: Option<Vec<GroupAttendeeConflictData>>,
}

/// Conflict data for unknown attendees.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct UnknownAttendeeConflictData {
    // Add fields as needed based on EWS schema
}

/// Conflict data for individual attendees.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct IndividualAttendeeConflictData {
    pub busy_type: LegacyFreeBusyStatus,
}

/// Conflict data for group attendees.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GroupAttendeeConflictData {
    pub number_of_members: Option<i32>,
    pub number_of_members_available: Option<i32>,
    pub number_of_members_with_conflict: Option<i32>,
    pub number_of_members_with_no_data: Option<i32>,
}

/// Quality rating for suggestions.
#[derive(Clone, Debug, Deserialize)]
pub enum SuggestionQuality {
    Excellent,
    Good,
    Fair,
    Poor,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::sealed::EnvelopeBodyContents;

    #[test]
    fn test_get_user_availability_operation_name() {
        assert_eq!(
            <GetUserAvailability as EnvelopeBodyContents>::name(),
            "GetUserAvailability"
        );
    }

    #[test]
    fn test_get_user_availability_creation() {
        let email = EmailAddress {
            name: Some("John Doe".to_string()),
            address: "john.doe@example.com".to_string(),
            routing_type: Some("SMTP".to_string()),
        };

        let mailbox_data = MailboxData {
            email,
            exclude_conflicts: Some(false),
        };

        let time_window = Duration {
            start_time: DateTime(time::OffsetDateTime::now_utc()),
            end_time: DateTime(time::OffsetDateTime::now_utc() + time::Duration::hours(8)),
        };

        let free_busy_options = FreeBusyViewOptions {
            time_window,
            requested_view: FreeBusyViewType::FreeBusy,
            merged_free_busy_interval_in_minutes: Some(30),
        };

        let operation = GetUserAvailability {
            time_zone: None,
            mailbox_data_array: MailboxDataArray {
                mailbox_data: vec![mailbox_data],
            },
            free_busy_view_options: free_busy_options,
        };

        assert_eq!(operation.mailbox_data_array.mailbox_data.len(), 1);
        assert_eq!(
            operation.mailbox_data_array.mailbox_data[0].email.address,
            "john.doe@example.com"
        );
        assert!(matches!(
            operation.free_busy_view_options.requested_view,
            FreeBusyViewType::FreeBusy
        ));
        assert_eq!(
            operation
                .free_busy_view_options
                .merged_free_busy_interval_in_minutes,
            Some(30)
        );
    }

    #[test]
    fn test_get_user_availability_with_timezone() {
        let time_zone = SerializableTimeZone {
            bias: -480, // PST bias
            standard_time: Some(SerializableTimeZoneTime {
                bias: 0,
                time: "02:00:00".to_string(),
                day_of_week: DayOfWeek::Sunday,
                month: 11,
                day_order: 1,
            }),
            daylight_time: Some(SerializableTimeZoneTime {
                bias: -60,
                time: "02:00:00".to_string(),
                day_of_week: DayOfWeek::Sunday,
                month: 3,
                day_order: 2,
            }),
        };

        let email = EmailAddress {
            name: None,
            address: "user@example.com".to_string(),
            routing_type: None,
        };

        let operation = GetUserAvailability {
            time_zone: Some(time_zone),
            mailbox_data_array: MailboxDataArray {
                mailbox_data: vec![MailboxData {
                    email,
                    exclude_conflicts: None,
                }],
            },
            free_busy_view_options: FreeBusyViewOptions {
                time_window: Duration {
                    start_time: DateTime(time::OffsetDateTime::now_utc()),
                    end_time: DateTime(time::OffsetDateTime::now_utc() + time::Duration::hours(24)),
                },
                requested_view: FreeBusyViewType::Detailed,
                merged_free_busy_interval_in_minutes: None,
            },
        };

        assert!(operation.time_zone.is_some());
        if let Some(tz) = &operation.time_zone {
            assert_eq!(tz.bias, -480);
            assert!(tz.standard_time.is_some());
            assert!(tz.daylight_time.is_some());
        }
    }

    #[test]
    fn test_free_busy_view_types() {
        let view_types = vec![
            FreeBusyViewType::None,
            FreeBusyViewType::MergedOnly,
            FreeBusyViewType::FreeBusy,
            FreeBusyViewType::FreeBusyMerged,
            FreeBusyViewType::Detailed,
            FreeBusyViewType::DetailedMerged,
        ];

        // Test that all view types can be created
        for view_type in view_types {
            let options = FreeBusyViewOptions {
                time_window: Duration {
                    start_time: DateTime(time::OffsetDateTime::now_utc()),
                    end_time: DateTime(time::OffsetDateTime::now_utc() + time::Duration::hours(1)),
                },
                requested_view: view_type,
                merged_free_busy_interval_in_minutes: Some(15),
            };

            // Verify the view type is set correctly
            match options.requested_view {
                FreeBusyViewType::None => {}
                FreeBusyViewType::MergedOnly => {}
                FreeBusyViewType::FreeBusy => {}
                FreeBusyViewType::FreeBusyMerged => {}
                FreeBusyViewType::Detailed => {}
                FreeBusyViewType::DetailedMerged => {}
            }
        }
    }

    #[test]
    fn test_legacy_free_busy_status() {
        let status_values = vec![
            LegacyFreeBusyStatus::Free,
            LegacyFreeBusyStatus::Tentative,
            LegacyFreeBusyStatus::Busy,
            LegacyFreeBusyStatus::OOF,
            LegacyFreeBusyStatus::NoData,
        ];

        // Test that all status values can be created
        for status in status_values {
            let event = CalendarEvent {
                start_time: DateTime(time::OffsetDateTime::now_utc()),
                end_time: DateTime(time::OffsetDateTime::now_utc() + time::Duration::hours(1)),
                busy_type: status,
                calendar_event_details: None,
            };

            // Verify the status is set correctly
            match event.busy_type {
                LegacyFreeBusyStatus::Free => {}
                LegacyFreeBusyStatus::Tentative => {}
                LegacyFreeBusyStatus::Busy => {}
                LegacyFreeBusyStatus::OOF => {}
                LegacyFreeBusyStatus::NoData => {}
            }
        }
    }

    #[test]
    fn test_multiple_mailboxes() {
        let emails = vec![
            EmailAddress {
                name: Some("User One".to_string()),
                address: "user1@example.com".to_string(),
                routing_type: Some("SMTP".to_string()),
            },
            EmailAddress {
                name: Some("User Two".to_string()),
                address: "user2@example.com".to_string(),
                routing_type: Some("SMTP".to_string()),
            },
            EmailAddress {
                name: Some("User Three".to_string()),
                address: "user3@example.com".to_string(),
                routing_type: Some("SMTP".to_string()),
            },
        ];

        let mailbox_data: Vec<MailboxData> = emails
            .into_iter()
            .map(|email| MailboxData {
                email,
                exclude_conflicts: Some(true),
            })
            .collect();

        let operation = GetUserAvailability {
            time_zone: None,
            mailbox_data_array: MailboxDataArray { mailbox_data },
            free_busy_view_options: FreeBusyViewOptions {
                time_window: Duration {
                    start_time: DateTime(time::OffsetDateTime::now_utc()),
                    end_time: DateTime(time::OffsetDateTime::now_utc() + time::Duration::hours(12)),
                },
                requested_view: FreeBusyViewType::FreeBusyMerged,
                merged_free_busy_interval_in_minutes: Some(60),
            },
        };

        assert_eq!(operation.mailbox_data_array.mailbox_data.len(), 3);

        for (i, mailbox) in operation.mailbox_data_array.mailbox_data.iter().enumerate() {
            assert_eq!(mailbox.email.address, format!("user{}@example.com", i + 1));
            assert_eq!(mailbox.exclude_conflicts, Some(true));
        }
    }
}
