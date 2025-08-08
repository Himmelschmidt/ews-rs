/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use ews_proc_macros::operation_response;
use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{BaseFolderId, BaseItemId, MESSAGES_NS_URI};

/// A request to send one or more Exchange items.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/senditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
#[operation_response(SendItemResponseMessage)]
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


#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct SendItemResponseMessage {}

#[cfg(test)]
mod test {
    use crate::{test_utils::assert_deserialized_content, ResponseClass};

    use super::{SendItemResponse, SendItemResponseMessage};
    use crate::ResponseMessages;

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
                response_messages: vec![ResponseClass::Success(
                    SendItemResponseMessage {},
                )],
            },
        };

        assert_deserialized_content(content, expected);
    }
}
