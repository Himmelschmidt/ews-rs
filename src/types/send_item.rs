/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::types::sealed::EnvelopeBodyContents;
use crate::{
    BaseItemId, BaseFolderId, Operation, OperationResponse, ResponseClass, ResponseCode,
    MESSAGES_NS_URI,
};

/// A request to send one or more Exchange items.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/senditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct SendItem {
    /// Whether to save a copy of the sent item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/senditem#attributes>
    #[xml_struct(attribute)]
    pub save_item_to_folder: bool,

    /// A list of items to send.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemids>
    pub item_ids: Vec<BaseItemId>,

    /// The folder to save sent items to.
    ///
    /// If not specified and `save_item_to_folder` is true, items will be saved
    /// to the default Sent Items folder.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/saveditemfolderid>
    pub saved_item_folder_id: Option<BaseFolderId>,
}

impl Operation for SendItem {
    type Response = SendItemResponse;
}

impl EnvelopeBodyContents for SendItem {
    fn name() -> &'static str {
        "SendItem"
    }
}

/// A response to a [`SendItem`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/senditemresponse>
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct SendItemResponse {
    pub response_messages: ResponseMessages,
}

impl OperationResponse for SendItemResponse {}

impl EnvelopeBodyContents for SendItemResponse {
    fn name() -> &'static str {
        "SendItemResponse"
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseMessages {
    pub send_item_response_message: Vec<SendItemResponseMessage>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct SendItemResponseMessage {
    /// The status of the corresponding request, i.e. whether it succeeded or
    /// resulted in an error.
    #[serde(rename = "@ResponseClass")]
    pub response_class: ResponseClass,

    pub response_code: Option<ResponseCode>,

    pub message_text: Option<String>,
}

#[cfg(test)]
mod test {
    use crate::{
        test_utils::assert_deserialized_content, ResponseClass, ResponseCode,
    };

    use super::{ResponseMessages, SendItemResponse, SendItemResponseMessage};

    #[test]
    fn test_deserialize_send_item_response() {
        let content = r#"<SendItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                    <m:ResponseMessages>
                        <m:SendItemResponseMessage ResponseClass="Success">
                            <m:ResponseCode>NoError</m:ResponseCode>
                        </m:SendItemResponseMessage>
                    </m:ResponseMessages>
                    </SendItemResponse>"#;

        let expected = SendItemResponse {
            response_messages: ResponseMessages {
                send_item_response_message: vec![SendItemResponseMessage {
                    response_class: ResponseClass::Success,
                    response_code: Some(ResponseCode::NoError),
                    message_text: None,
                }],
            },
        };

        assert_deserialized_content(content, expected);
    }
}