/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use ews_proc_macros::operation_response;
use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    BaseFolderId, FieldOrder, ItemShape, Items, Paging, Restriction, Traversal, MESSAGES_NS_URI,
};

/// A request to find items matching certain criteria.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
#[operation_response(FindItemResponseMessage)]
pub struct FindItem {
    /// The traversal method for the find operation.
    #[xml_struct(attribute)]
    pub traversal: Traversal,

    /// A description of the information to be included in the response for each found item.
    pub item_shape: ItemShape,

    /// Paging information for the response.
    #[xml_struct(flatten, ns_prefix = "m")]
    pub paging: Option<Paging>,

    /// Restriction to apply to the search.
    pub restriction: Option<Restriction>,

    /// Sort order for the results.
    pub sort_order: Option<Vec<FieldOrder>>,

    /// The parent folder IDs to search in.
    pub parent_folder_ids: Vec<BaseFolderId>,
}

/// A response to a request for finding items.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponsemessage>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct FindItemResponseMessage {
    /// The root folder containing the search results.
    pub root_folder: RootFolder,
}

/// The root folder element in find responses.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/rootfolder-finditemresponsemessage>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct RootFolder {
    #[serde(rename = "@IndexedPagingOffset")]
    pub indexed_paging_offset: Option<i32>,

    #[serde(rename = "@TotalItemsInView")]
    pub total_items_in_view: u32,

    #[serde(rename = "@IncludesLastItemInRange")]
    pub includes_last_item_in_range: bool,

    pub items: Items,
}
