/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use ews_proc_macros::operation_response;
use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{BaseItemId, MESSAGES_NS_URI};

/// A request to mark one or more items as read or unread.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/markasread>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
#[operation_response(MarkAsReadResponseMessage)]
pub struct MarkAsRead {
    /// Whether to mark the items as read (true) or unread (false).
    #[xml_struct(attribute)]
    pub read_flag: bool,

    /// Whether to suppress read receipts for the items.
    #[xml_struct(attribute)]
    pub suppress_read_receipts: Option<bool>,

    /// The items to mark as read or unread.
    pub item_ids: Vec<BaseItemId>,
}


/// A response to a request for marking an item as read/unread.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/markasreadresponsemessage>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct MarkAsReadResponseMessage {}

#[cfg(test)]
mod test {
    use super::*;
    use crate::{test_utils::assert_serialized_content, Operation, ResponseClass};

    #[test]
    fn test_mark_as_read_operation_name() {
        let request = MarkAsRead {
            read_flag: true,
            suppress_read_receipts: None,
            item_ids: vec![],
        };
        assert_eq!(request.name(), "MarkAsRead");
    }

    #[test]
    fn test_serialize_mark_as_read() {
        let request = MarkAsRead {
            read_flag: true,
            suppress_read_receipts: Some(false),
            item_ids: vec![BaseItemId::ItemId {
                id: "test-item-id".to_string(),
                change_key: Some("test-change-key".to_string()),
            }],
        };

        let expected = r#"<MarkAsRead xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ReadFlag="true" SuppressReadReceipts="false"><ItemIds><t:ItemId Id="test-item-id" ChangeKey="test-change-key"/></ItemIds></MarkAsRead>"#;

        assert_serialized_content(&request, "MarkAsRead", expected);
    }

    #[test]
    fn test_deserialize_mark_as_read_response() {
        let xml = r#"<MarkAsReadResponse xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
            <ResponseMessages>
                <MarkAsReadResponseMessage ResponseClass="Success">
                    <ResponseCode>NoError</ResponseCode>
                </MarkAsReadResponseMessage>
            </ResponseMessages>
        </MarkAsReadResponse>"#;

        let response: MarkAsReadResponse =
            quick_xml::de::from_str(xml).expect("should deserialize successfully");
        assert!(matches!(
            response.response_messages.response_messages[0],
            ResponseClass::Success(_)
        ));
    }
}
