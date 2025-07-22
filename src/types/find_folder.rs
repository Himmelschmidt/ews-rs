/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, FolderShape, Folders, Operation,
    OperationResponse, Paging, ResponseClass, Restriction, Traversal, MESSAGES_NS_URI,
};

/// A request to find folders matching certain criteria.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/findfolder>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct FindFolder {
    /// The traversal method for the find operation.
    #[xml_struct(attribute)]
    pub traversal: Traversal,

    /// A description of the information to be included in the response for each found folder.
    pub folder_shape: FolderShape,

    /// Paging information for the response.
    #[xml_struct(flatten, ns_prefix = "m")]
    pub paging: Option<Paging>,

    /// Restriction to apply to the search.
    pub restriction: Option<Restriction>,

    /// The parent folder IDs to search in.
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

/// A response to a [`FindFolder`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/findfolderresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindFolderResponse {
    pub response_messages: FindFolderResponseMessages,
}

impl OperationResponse for FindFolderResponse {}

impl EnvelopeBodyContents for FindFolderResponse {
    fn name() -> &'static str {
        "FindFolderResponse"
    }
}

/// A collection of responses for individual entities within a request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsemessages>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindFolderResponseMessages {
    pub find_folder_response_message: Vec<ResponseClass<FindFolderResponseMessage>>,
}

/// A response to a request for finding folders.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/findfolderresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct FindFolderResponseMessage {
    /// The root folder containing the search results.
    pub root_folder: RootFolder,
}

/// The root folder element in find responses.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/rootfolder-findfolderresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct RootFolder {
    #[serde(rename = "@IndexedPagingOffset")]
    pub indexed_paging_offset: Option<i32>,

    #[serde(rename = "@TotalItemsInView")]
    pub total_items_in_view: u32,

    #[serde(rename = "@IncludesLastItemInRange")]
    pub includes_last_item_in_range: bool,

    pub folders: Folders,
}
