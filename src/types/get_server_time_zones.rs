/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, Operation, OperationResponse, ResponseClass, ResponseCode,
    MESSAGES_NS_URI,
};

/// A request to retrieve time zone definitions from the Exchange server.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getservertimezones>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct GetServerTimeZones {
    /// Whether to return the complete definitions for each time zone.
    ///
    /// When `true` (default), returns complete time zone definitions.
    /// When `false`, returns only the name and identifier for each time zone.
    #[xml_struct(attribute)]
    pub return_full_time_zone_data: Option<bool>,

    /// An optional array of time zone identifiers to retrieve.
    ///
    /// If omitted, all time zone definitions available on the server are returned.
    pub ids: Option<TimeZoneIds>,
}

impl Operation for GetServerTimeZones {
    type Response = GetServerTimeZonesResponse;
}

impl EnvelopeBodyContents for GetServerTimeZones {
    fn name() -> &'static str {
        "GetServerTimeZones"
    }
}

/// Container for time zone identifiers.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct TimeZoneIds {
    /// Array of time zone identifiers.
    #[xml_struct(ns_prefix = "t")]
    pub id: Vec<String>,
}

/// A response to a [`GetServerTimeZones`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getservertimezonesresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetServerTimeZonesResponse {
    pub response_messages: GetServerTimeZonesResponseMessages,
}

impl OperationResponse for GetServerTimeZonesResponse {}

impl EnvelopeBodyContents for GetServerTimeZonesResponse {
    fn name() -> &'static str {
        "GetServerTimeZonesResponse"
    }
}

/// A collection of responses for individual entities within a request.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetServerTimeZonesResponseMessages {
    pub get_server_time_zones_response_message: Vec<GetServerTimeZonesResponseMessage>,
}

/// A response to a request for server time zones.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getservertimezonesresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetServerTimeZonesResponseMessage {
    /// The status of the corresponding request.
    #[serde(rename = "@ResponseClass")]
    pub response_class: ResponseClass,

    pub response_code: Option<ResponseCode>,

    pub message_text: Option<String>,

    /// The time zone definitions returned by the server.
    pub time_zone_definitions: Option<TimeZoneDefinitions>,
}

/// Container for time zone definitions.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneDefinitions {
    /// Array of time zone definitions.
    #[serde(rename = "TimeZoneDefinition")]
    pub time_zone_definition: Vec<TimeZoneDefinition>,
}

/// Represents a time zone definition from the Exchange server.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/timezonedefinition>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneDefinition {
    /// The unique identifier for the time zone.
    #[serde(rename = "@Id")]
    pub id: String,

    /// The name of the time zone.
    #[serde(rename = "@Name")]
    pub name: String,

    /// The periods that define the time zone.
    pub periods: Option<TimeZonePeriods>,

    /// The transitions between different time zone periods.
    pub transitions_groups: Option<TimeZoneTransitionsGroups>,

    /// The transitions for the time zone.
    pub transitions: Option<TimeZoneTransitions>,
}

/// Container for time zone periods.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZonePeriods {
    /// Array of time zone periods.
    #[serde(rename = "Period")]
    pub period: Vec<TimeZonePeriod>,
}

/// Represents a period within a time zone definition.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZonePeriod {
    /// The bias (in minutes) for this period.
    #[serde(rename = "@Bias")]
    pub bias: String,

    /// The name of this period.
    #[serde(rename = "@Name")]
    pub name: String,

    /// The identifier for this period.
    #[serde(rename = "@Id")]
    pub id: String,
}

/// Container for time zone transition groups.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneTransitionsGroups {
    /// Array of transition groups.
    #[serde(rename = "TransitionsGroup")]
    pub transitions_group: Vec<TimeZoneTransitionsGroup>,
}

/// Represents a group of transitions for a time zone.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneTransitionsGroup {
    /// The identifier for this transitions group.
    #[serde(rename = "@Id")]
    pub id: String,

    /// The transitions in this group.
    pub transition: Option<Vec<TimeZoneTransition>>,
}

/// Container for time zone transitions.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneTransitions {
    /// Array of transitions.
    #[serde(rename = "Transition")]
    pub transition: Vec<TimeZoneTransition>,
}

/// Represents a transition between time zone periods.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneTransition {
    /// Points to a transition group.
    pub to: Option<TimeZoneTransitionTo>,
}

/// Represents the target of a time zone transition.
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct TimeZoneTransitionTo {
    /// The kind of transition target.
    #[serde(rename = "@Kind")]
    pub kind: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_server_time_zones_operation_name() {
        assert_eq!(
            <GetServerTimeZones as crate::types::sealed::EnvelopeBodyContents>::name(),
            "GetServerTimeZones"
        );
    }

    #[test]
    fn test_get_server_time_zones_creation() {
        let operation = GetServerTimeZones {
            return_full_time_zone_data: Some(false),
            ids: Some(TimeZoneIds {
                id: vec!["UTC".to_string(), "Eastern Standard Time".to_string()],
            }),
        };

        assert_eq!(operation.return_full_time_zone_data, Some(false));
        assert!(operation.ids.is_some());
        if let Some(ids) = &operation.ids {
            assert_eq!(ids.id.len(), 2);
            assert_eq!(ids.id[0], "UTC");
            assert_eq!(ids.id[1], "Eastern Standard Time");
        }
    }

    #[test]
    fn test_get_server_time_zones_minimal() {
        let operation = GetServerTimeZones {
            return_full_time_zone_data: None,
            ids: None,
        };

        assert_eq!(operation.return_full_time_zone_data, None);
        assert!(operation.ids.is_none());
    }
}
