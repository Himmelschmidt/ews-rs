/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, BaseItemId, Items, Operation,
    OperationResponse, ResponseClass, ResponseCode, MESSAGES_NS_URI,
};

/// A request to move an Exchange item to a different folder.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/moveitem>

#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct MoveItem {
    pub to_folder_id: BaseFolderId,
    pub item_ids: Vec<BaseItemId>,
    pub return_new_item_ids: bool,
}

impl Operation for MoveItem {
    type Response = MoveItemResponse;
}

impl EnvelopeBodyContents for MoveItem {
    fn name() -> &'static str {
        "MoveItem"
    }
}

/// A response to a [`MoveItem`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getitemresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MoveItemResponse {
    pub response_messages: ResponseMessages,
}

impl OperationResponse for MoveItemResponse {}

impl EnvelopeBodyContents for MoveItemResponse {
    fn name() -> &'static str {
        "MoveItemResponse"
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseMessages {
    pub move_item_response_message: Vec<MoveItemResponseMessage>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MoveItemResponseMessage {
    /// The status of the corresponding request, i.e. whether it succeeded or
    /// resulted in an error.
    #[serde(rename = "@ResponseClass")]
    pub response_class: ResponseClass,

    pub response_code: Option<ResponseCode>,

    pub message_text: Option<String>,

    pub items: Items,
}
