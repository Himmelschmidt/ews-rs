/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, FolderId, FolderShape, Operation,
    OperationResponse, ResponseClass, ResponseCode, Traversal, MESSAGES_NS_URI,
};

/// The FindItem operation searches for items that are located in a user's mailbox.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct FindFolder {
    /// The traversal type for the search.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
    #[xml_struct(attribute)]
    pub traversal: Traversal,

    /// A description of the information to be included in the response for each
    /// item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemshape>
    pub folder_shape: FolderShape,

    /// The folders to search.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/parentfolderids>
    pub parent_folder_ids: Vec<BaseFolderId>,
}

impl Operation for FindFolder {
    type Response = FindFolderResponse;
}

impl EnvelopeBodyContents for FindFolder {
    fn name() -> &'static str {
        "FindFolder"
    }
}

/// The response to a ['FindFolder'] operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindFolderResponse {
    pub response_messages: ResponseMessages,
}

impl OperationResponse for FindFolderResponse {}

impl EnvelopeBodyContents for FindFolderResponse {
    fn name() -> &'static str {
        "FindFolderResponse"
    }
}

/// The response messages for a FindFolder operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsemessages>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseMessages {
    pub find_folder_response_message: Vec<FindFolderResponseMessage>,
}

/// A response message for a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindFolderResponseMessage {
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

/// The root folder containing the items found by a FindFolder operation.
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
    pub folders: Folders,
}

/// The items found by a FindFolder operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/items>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Folders {
    /// The folder items found by the search.
    #[serde(default)]
    pub folders: Vec<Folder>,
}

/// A message item found by a FindItem operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/message-ex15websvcsotherref>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Folder {
    /// The ID of the message.
    pub folder_id: FolderId,

    pub display_name: String,

    pub total_count: u32,
    pub child_folder_count: u32,
    pub unread_count: u32,
}
