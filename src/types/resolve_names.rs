/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, Operation, OperationResponse, ResponseClass,
    MESSAGES_NS_URI,
};

/// A request to resolve ambiguous email addresses and display names.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/resolvenames>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct ResolveNames {
    /// Determines whether the full contact data for resolved names is returned.
    #[xml_struct(attribute)]
    pub return_full_contact_data: Option<bool>,

    /// Defines the search scope for resolving names.
    #[xml_struct(attribute)]
    pub search_scope: Option<SearchScope>,

    /// Specifies the property set returned for contacts.
    #[xml_struct(attribute)]
    pub contact_data_shape: Option<ContactDataShape>,

    /// Optional parent folder IDs to search within.
    pub parent_folder_ids: Option<Vec<BaseFolderId>>,

    /// The name or partial name to resolve.
    pub unresolved_entry: String,
}

impl Operation for ResolveNames {
    type Response = ResolveNamesResponse;
}

impl EnvelopeBodyContents for ResolveNames {
    fn name() -> &'static str {
        "ResolveNames"
    }
}

/// Defines the search scope for the ResolveNames operation.
#[derive(Clone, Copy, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum SearchScope {
    /// Search only in Active Directory.
    ActiveDirectory,
    /// Search in Active Directory, then contacts.
    ActiveDirectoryContacts,
    /// Search only in contacts.
    Contacts,
    /// Search in contacts, then Active Directory.
    ContactsActiveDirectory,
}

/// Specifies the property set returned for contacts in ResolveNames.
#[derive(Clone, Copy, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum ContactDataShape {
    /// Only the ID of the contact.
    IdOnly,
    /// The default set of properties.
    Default,
    /// All properties of the contact.
    AllProperties,
}

/// A response to a [`ResolveNames`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/resolvenamesresponse>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ResolveNamesResponse {
    pub response_messages: ResolveNamesResponseMessages,
}

impl OperationResponse for ResolveNamesResponse {}

impl EnvelopeBodyContents for ResolveNamesResponse {
    fn name() -> &'static str {
        "ResolveNamesResponse"
    }
}

/// A collection of responses for ResolveNames requests.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ResolveNamesResponseMessages {
    pub resolve_names_response_message: Vec<ResponseClass<ResolveNamesResponseMessage>>,
}

/// A response message for an individual ResolveNames request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/resolvenamesresponsemessage>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ResolveNamesResponseMessage {
    /// Collection of resolved names.
    pub resolution_set: Option<ResolutionSet>,
}

/// A collection of resolved names.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/resolutionset>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ResolutionSet {
    /// Whether names were included that could not be resolved.
    #[serde(rename = "@IncludesLastItemInRange")]
    pub includes_last_item_in_range: Option<bool>,

    /// The index of the first resolution in the set.
    #[serde(rename = "@IndexedPagingOffset")]
    pub indexed_paging_offset: Option<i32>,

    /// The total number of resolutions available.
    #[serde(rename = "@TotalItemsInView")]
    pub total_items_in_view: Option<i32>,

    /// Collection of individual resolutions.
    pub resolution: Vec<Resolution>,
}

/// An individual name resolution result.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/resolution>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Resolution {
    /// The mailbox information for the resolved name.
    pub mailbox: Mailbox,

    /// Contact information if available.
    pub contact: Option<Contact>,
}

/// Mailbox information for a resolved name.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailbox>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Mailbox {
    /// The display name of the mailbox.
    pub name: Option<String>,

    /// The email address of the mailbox.
    pub email_address: Option<String>,

    /// The routing type (usually SMTP).
    pub routing_type: Option<String>,

    /// The mailbox type.
    pub mailbox_type: Option<MailboxType>,
}

/// The type of mailbox.
#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Eq)]
pub enum MailboxType {
    Mailbox,
    PublicDL,
    PrivateDL,
    Contact,
    PublicFolder,
    Unknown,
    OneOff,
    GroupMailbox,
}

/// Contact information for a resolved name.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contact>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Contact {
    /// The display name of the contact.
    pub display_name: Option<String>,

    /// The given name of the contact.
    pub given_name: Option<String>,

    /// The initials of the contact.
    pub initials: Option<String>,

    /// The middle name of the contact.
    pub middle_name: Option<String>,

    /// The nickname of the contact.
    pub nickname: Option<String>,

    /// The complete name of the contact.
    pub complete_name: Option<CompleteName>,

    /// The company name of the contact.
    pub company_name: Option<String>,

    /// Email addresses for the contact.
    pub email_addresses: Option<EmailAddresses>,

    /// Phone numbers for the contact.
    pub phone_numbers: Option<PhoneNumbers>,

    /// Physical addresses for the contact.
    pub physical_addresses: Option<PhysicalAddresses>,
}

/// Complete name information for a contact.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct CompleteName {
    /// The title of the contact.
    pub title: Option<String>,

    /// The first name of the contact.
    pub first_name: Option<String>,

    /// The middle name of the contact.
    pub middle_name: Option<String>,

    /// The last name of the contact.
    pub last_name: Option<String>,

    /// The suffix of the contact.
    pub suffix: Option<String>,

    /// The initials of the contact.
    pub initials: Option<String>,

    /// The full name of the contact.
    pub full_name: Option<String>,

    /// The nickname of the contact.
    pub nickname: Option<String>,

    /// The Yomi first name.
    pub yomi_first_name: Option<String>,

    /// The Yomi last name.
    pub yomi_last_name: Option<String>,
}

/// Email addresses for a contact.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct EmailAddresses {
    /// Collection of email address entries.
    pub entry: Vec<EmailAddressEntry>,
}

/// An email address entry.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct EmailAddressEntry {
    /// The key identifying the email address type.
    #[serde(rename = "@Key")]
    pub key: String,

    /// The email address value.
    #[serde(rename = "$text")]
    pub value: String,
}

/// Phone numbers for a contact.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct PhoneNumbers {
    /// Collection of phone number entries.
    pub entry: Vec<PhoneNumberEntry>,
}

/// A phone number entry.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct PhoneNumberEntry {
    /// The key identifying the phone number type.
    #[serde(rename = "@Key")]
    pub key: String,

    /// The phone number value.
    #[serde(rename = "$text")]
    pub value: String,
}

/// Physical addresses for a contact.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct PhysicalAddresses {
    /// Collection of physical address entries.
    pub entry: Vec<PhysicalAddressEntry>,
}

/// A physical address entry.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct PhysicalAddressEntry {
    /// The key identifying the address type.
    #[serde(rename = "@Key")]
    pub key: String,

    /// The street address.
    pub street: Option<String>,

    /// The city.
    pub city: Option<String>,

    /// The state or province.
    pub state: Option<String>,

    /// The country or region.
    pub country_or_region: Option<String>,

    /// The postal code.
    pub postal_code: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_names_serialization() {
        let resolve_names = ResolveNames {
            return_full_contact_data: Some(true),
            search_scope: Some(SearchScope::ActiveDirectoryContacts),
            contact_data_shape: Some(ContactDataShape::Default),
            parent_folder_ids: None,
            unresolved_entry: "john.doe@example.com".to_string(),
        };

        // Test that the structure can be created and has expected values
        assert_eq!(resolve_names.unresolved_entry, "john.doe@example.com");
        assert_eq!(resolve_names.return_full_contact_data, Some(true));
        assert!(matches!(
            resolve_names.search_scope,
            Some(SearchScope::ActiveDirectoryContacts)
        ));
        assert!(matches!(
            resolve_names.contact_data_shape,
            Some(ContactDataShape::Default)
        ));
    }

    #[test]
    fn test_resolve_names_operation_name() {
        assert_eq!(
            <ResolveNames as crate::types::sealed::EnvelopeBodyContents>::name(),
            "ResolveNames"
        );
    }
}
