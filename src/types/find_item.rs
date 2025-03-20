/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, ItemId, ItemShape, Operation,
    OperationResponse, ResponseClass, ResponseCode, Restriction, SortOrder, MESSAGES_NS_URI,
};

/// The FindItem operation searches for items that are located in a user's mailbox.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct FindItem {
    /// The traversal type for the search.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
    #[xml_struct(attribute)]
    pub traversal: Traversal,

    /// A description of the information to be included in the response for each
    /// item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemshape>
    pub item_shape: ItemShape,

    /// The folders to search.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/parentfolderids>
    pub parent_folder_ids: Vec<BaseFolderId>,

    /// The restriction or query used to filter items.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/restriction>
    pub restriction: Option<Restriction>,

    /// Defines how items are sorted in the response.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/sortorder>
    pub sort_order: Option<SortOrder>,
}

impl Operation for FindItem {
    type Response = FindItemResponse;
}

impl EnvelopeBodyContents for FindItem {
    fn name() -> &'static str {
        "FindItem"
    }
}

/// The response to a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindItemResponse {
    pub response_messages: ResponseMessages,
}

impl OperationResponse for FindItemResponse {}

impl EnvelopeBodyContents for FindItemResponse {
    fn name() -> &'static str {
        "FindItemResponse"
    }
}

/// The response messages for a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsemessages>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseMessages {
    pub find_item_response_message: Vec<FindItemResponseMessage>,
}

/// A response message for a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindItemResponseMessage {
    /// The status of the corresponding request, i.e. whether it succeeded or
    /// resulted in an error.
    #[serde(rename = "@ResponseClass")]
    pub response_class: ResponseClass,

    pub response_code: Option<ResponseCode>,

    pub message_text: Option<String>,

    pub descriptive_link_key: Option<u32>,

    /// The root folder containing the items found by the search.
    pub root_folder: Option<RootFolder>,
}

/// The root folder containing the items found by a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/rootfolder-finditemresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct RootFolder {
    /// The total number of items in the view.
    #[serde(rename = "@TotalItemsInView")]
    pub total_items_in_view: u32,

    /// Whether the response includes the last item in the range.
    #[serde(rename = "@IncludesLastItemInRange")]
    pub includes_last_item_in_range: bool,

    /// The items found by the search.
    pub items: Items,
}

/// The items found by a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/items>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Items {
    /// The message items found by the search.
    #[serde(default)]
    pub message: Vec<Message>,
}

/// A message item found by a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/message-ex15websvcsotherref>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Message {
    /// The ID of the message.
    pub item_id: ItemId,
}

/// The traversal type for a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum Traversal {
    /// A shallow traversal finds items in the folder.
    Shallow,
    /// A soft-deleted traversal finds items in the dumpster.
    SoftDeleted,
    ///Returns only the identities of associated items in the folder.
    Associated,
}
