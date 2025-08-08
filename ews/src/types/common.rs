/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use std::ops::{Deref, DerefMut};

use serde::{Deserialize, Deserializer};
use time::format_description::well_known::Iso8601;
use xml_struct::XmlSerialize;

pub mod response;
pub use self::response::{ResponseClass, ResponseMessages};
pub mod message_xml;
pub use self::message_xml::MessageXml;

pub(crate) const MESSAGES_NS_URI: &str =
    "http://schemas.microsoft.com/exchange/services/2006/messages";
pub(crate) const SOAP_NS_URI: &str = "http://schemas.xmlsoap.org/soap/envelope/";
pub(crate) const TYPES_NS_URI: &str = "http://schemas.microsoft.com/exchange/services/2006/types";

/// The folder properties which should be included in the response.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/foldershape>.
#[derive(Clone, Debug, Default, XmlSerialize)]
pub struct FolderShape {
    #[xml_struct(ns_prefix = "t")]
    pub base_shape: BaseShape,
}

/// The item properties which should be included in the response.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemshape>.
#[derive(Clone, Debug, XmlSerialize)]
pub struct ItemShape {
    /// The base set of properties to include, which may be extended by other
    /// fields.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/baseshape>
    #[xml_struct(ns_prefix = "t")]
    pub base_shape: BaseShape,

    /// Whether the MIME content of an item should be included.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/includemimecontent>
    #[xml_struct(ns_prefix = "t")]
    pub include_mime_content: Option<bool>,

    /// A list of properties which should be included in addition to those
    /// implied by other fields.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/additionalproperties>
    #[xml_struct(ns_prefix = "t")]
    pub additional_properties: Option<Vec<PathToElement>>,
}

impl Default for ItemShape {
    fn default() -> Self {
        Self {
            base_shape: BaseShape::IdOnly,
            include_mime_content: None,
            additional_properties: None,
        }
    }
}

/// An identifier for a property on an Exchange entity.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum PathToElement {
    /// An identifier for an extended MAPI property.
    ///
    /// The full set of constraints on which properties may or must be set
    /// together are not expressed in the structure of this variant. Please see
    /// Microsoft's documentation for further details.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/extendedfielduri>
    // TODO: We can represent in a friendlier way with an enum, probably. A
    // property is fully specified by a type and either:
    // - A property set ID plus property name/ID, or
    // - A property tag.
    // https://github.com/thunderbird/ews-rs/issues/9
    ExtendedFieldURI {
        /// A well-known identifier for a property set.
        #[xml_struct(attribute)]
        distinguished_property_set_id: Option<DistinguishedPropertySet>,

        /// A GUID representing a property set.
        // TODO: This could use a strong type for representing a GUID.
        #[xml_struct(attribute)]
        property_set_id: Option<String>,

        /// Specifies a property by integer tag.
        // TODO: This should use an integer type, but it seems a hex
        // representation is preferred, and we should restrict the possible
        // values per the docs.
        #[xml_struct(attribute)]
        property_tag: Option<String>,

        /// The name of a property within a specified property set.
        #[xml_struct(attribute)]
        property_name: Option<String>,

        /// The dispatch ID of a property within a specified property set.
        #[xml_struct(attribute)]
        property_id: Option<String>,

        /// The value type of the desired property.
        #[xml_struct(attribute)]
        property_type: PropertyType,
    },

    /// An identifier for a property given by a well-known string.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/fielduri>
    #[allow(non_snake_case)]
    FieldURI {
        /// The well-known string.
        // TODO: Adjust xml_struct to support field renaming to avoid non-snake
        // case identifiers.
        // https://github.com/thunderbird/xml-struct-rs/issues/6
        // TODO: We could use an enum for this field. It's just large and not
        // worth typing out by hand.
        #[xml_struct(attribute)]
        field_URI: String,
    },

    /// An identifier for a specific element of a dictionary-based property.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/indexedfielduri>
    #[allow(non_snake_case)]
    IndexedFieldURI {
        /// The well-known string identifier of the property.
        #[xml_struct(attribute)]
        field_URI: String,

        /// The member within the dictionary to access.
        #[xml_struct(attribute)]
        field_index: String,
    },
}

/// Response objects available for a message item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responseobjects>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseObjects {
    /// Available reply action for the message.
    #[xml_struct(ns_prefix = "t")]
    pub reply_to_item: Option<MessageResponseObject>,

    /// Available reply-all action for the message.
    #[xml_struct(ns_prefix = "t")]
    pub reply_all_to_item: Option<MessageResponseObject>,

    /// Available forward action for the message.
    #[xml_struct(ns_prefix = "t")]
    pub forward_item: Option<MessageResponseObject>,
}

/// A response object representing an available action on a message.
///
/// These are different from the operation types and represent the available
/// actions that can be performed on a message as returned in GetItem responses.
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct MessageResponseObject {
    // This struct is intentionally empty as the XML elements typically
    // only indicate availability of the action through their presence
}

/// The identifier for an extended MAPI property.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/extendedfielduri>
// N.B.: This is copied from `PathToElement::ExtendedFieldURI`,
// which follows the same structure. However, xml-struct doesn't currently
// support using a nested structure to define an element's attributes, see
// https://github.com/thunderbird/xml-struct-rs/issues/9
#[derive(Clone, Debug, Deserialize, XmlSerialize, Eq, PartialEq)]
pub struct ExtendedFieldURI {
    /// A well-known identifier for a property set.
    #[xml_struct(attribute)]
    pub distinguished_property_set_id: Option<DistinguishedPropertySet>,

    /// A GUID representing a property set.
    // TODO: This could use a strong type for representing a GUID.
    #[xml_struct(attribute)]
    pub property_set_id: Option<String>,

    /// Specifies a property by integer tag.
    #[xml_struct(attribute)]
    pub property_tag: Option<String>,

    /// The name of a property within a specified property set.
    #[xml_struct(attribute)]
    pub property_name: Option<String>,

    /// The dispatch ID of a property within a specified property set.
    #[xml_struct(attribute)]
    pub property_id: Option<String>,

    /// The value type of the desired property.
    #[xml_struct(attribute)]
    pub property_type: PropertyType,
}

/// A well-known MAPI property set identifier.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/extendedfielduri#distinguishedpropertysetid-attribute>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, Eq, PartialEq)]
#[xml_struct(text)]
pub enum DistinguishedPropertySet {
    Address,
    Appointment,
    CalendarAssistant,
    Common,
    InternetHeaders,
    Meeting,
    PublicStrings,
    Sharing,
    Task,
    UnifiedMessaging,
}

/// The action an Exchange server will take upon creating a `Message` item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/createitem#messagedisposition-attribute>
/// and <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/updateitem#messagedisposition-attribute>
#[derive(Clone, Copy, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum MessageDisposition {
    SaveOnly,
    SendOnly,
    SendAndSaveCopy,
}

/// The type of the value of a MAPI property.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/extendedfielduri#propertytype-attribute>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, Eq, PartialEq)]
#[xml_struct(text)]
pub enum PropertyType {
    ApplicationTime,
    ApplicationTimeArray,
    Binary,
    BinaryArray,
    Boolean,
    CLSID,
    CLSIDArray,
    Currency,
    CurrencyArray,
    Double,
    DoubleArray,
    Float,
    FloatArray,
    Integer,
    IntegerArray,
    Long,
    LongArray,
    Short,
    ShortArray,
    SystemTime,
    SystemTimeArray,
    String,
    StringArray,
}

/// The base set of properties to be returned in response to our request.
/// Additional properties may be specified by the parent element.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/baseshape>.
#[derive(Clone, Copy, Debug, Default, XmlSerialize)]
#[xml_struct(text)]
pub enum BaseShape {
    /// Only the IDs of any items or folders returned.
    IdOnly,

    /// The default set of properties for the relevant item or folder.
    ///
    /// The properties returned are dependent on the type of item or folder. See
    /// the EWS documentation for details.
    #[default]
    Default,

    /// All properties of an item or folder.
    AllProperties,
}

/// The traversal method for a find operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/findfolder-operation>
#[derive(Clone, Copy, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum Traversal {
    Shallow,
    SoftDeleted,
    Associated,
}

/// The manner in which paged views of data are retrieved.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/basepagingtype>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum Paging {
    /// A view of the data paged by item index.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/indexedpageitemview>
    IndexedPageItemView(IndexedPaging),
    // TODO: Implement other paging types:
    // - FractionalPageItemView
    // - CalendarView
    // - ContactsView
}

impl Default for Paging {
    fn default() -> Self {
        Paging::IndexedPageItemView(Default::default())
    }
}

/// Defines how paged views are retrieved from a list of items.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/indexedpageitemview>
#[derive(Clone, Debug, Default, XmlSerialize)]
pub struct IndexedPaging {
    /// The maximum number of results to return.
    #[xml_struct(attribute)]
    pub max_entries_returned: Option<u32>,

    /// The offset from the beginning of the list to start from.
    #[xml_struct(attribute)]
    pub offset: u32,

    /// The point from which the offset is calculated.
    #[xml_struct(attribute)]
    pub base_point: BasePoint,
}

/// The direction from which the offset is calculated.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/indexedpageitemview>
#[derive(Clone, Copy, Debug, Default, XmlSerialize)]
#[xml_struct(text)]
pub enum BasePoint {
    #[default]
    Beginning,
    End,
}

/// A restriction or filter for a search operation.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/restriction>
#[derive(Clone, Debug, XmlSerialize)]
pub struct Restriction {
    #[xml_struct(flatten)]
    pub restriction_type: RestrictionType,
}

#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum RestrictionType {
    And(AndRestriction),
    Or(OrRestriction),
    // TODO: Not
    IsEqualTo(FieldEqualTo),
    // TODO: IsNotEqualTo, IsGreaterThan, IsGreaterThanOrEqualTo, IsLessThan, IsLessThanOrEqualTo
    // TODO: Contains, Excludes
    Exists(PathToElement),
}

/// Represents a logical AND operation between multiple restrictions.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/and>
#[derive(Clone, Debug, XmlSerialize)]
pub struct AndRestriction(#[xml_struct(ns_prefix = "t")] pub Vec<Restriction>);

/// Represents a logical OR operation between multiple restrictions.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/or>
#[derive(Clone, Debug, XmlSerialize)]
pub struct OrRestriction(#[xml_struct(ns_prefix = "t")] pub Vec<Restriction>);

// TODO: Implement NOT restriction once Box<T> serialization is resolved
// /// Represents a logical NOT operation that negates another restriction.
// ///
// /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/not>
// #[derive(Clone, Debug, XmlSerialize)]
// pub struct NotRestriction {
//     #[xml_struct(flatten, ns_prefix = "t")]
//     pub restriction: Box<Restriction>,
// }

#[derive(Clone, Debug, XmlSerialize)]
#[allow(non_snake_case)]
pub struct FieldEqualTo {
    #[xml_struct(flatten, ns_prefix = "t")]
    pub path: PathToElement,
    #[xml_struct(ns_prefix = "t")]
    pub FieldURIOrConstant: FieldURIOrConstant,
}

#[derive(Clone, Debug, XmlSerialize)]
pub struct FieldURIOrConstant {
    #[xml_struct(ns_prefix = "t")]
    pub constant: Constant,
}

#[derive(Clone, Debug, XmlSerialize)]
pub struct Constant {
    #[xml_struct(attribute)]
    pub value: String,
}

impl Restriction {
    /// Creates a new AND restriction combining multiple restrictions.
    pub fn and(restrictions: Vec<Restriction>) -> Self {
        Self {
            restriction_type: RestrictionType::And(AndRestriction(restrictions)),
        }
    }

    /// Creates a new OR restriction for multiple alternative restrictions.
    pub fn or(restrictions: Vec<Restriction>) -> Self {
        Self {
            restriction_type: RestrictionType::Or(OrRestriction(restrictions)),
        }
    }

    // TODO: Implement NOT helper once Box<T> serialization is resolved
    // /// Creates a new NOT restriction that negates another restriction.
    // pub fn not(restriction: Restriction) -> Self {
    //     Self {
    //         restriction_type: RestrictionType::Not(NotRestriction {
    //             restriction: Box::new(restriction),
    //         }),
    //     }
    // }

    /// Creates a new IsEqualTo restriction for field equality.
    pub fn equal_to(path: PathToElement, value: String) -> Self {
        Self {
            restriction_type: RestrictionType::IsEqualTo(FieldEqualTo {
                path,
                FieldURIOrConstant: FieldURIOrConstant {
                    constant: Constant { value },
                },
            }),
        }
    }

    /// Creates a new Exists restriction to check field presence.
    pub fn exists(path: PathToElement) -> Self {
        Self {
            restriction_type: RestrictionType::Exists(path),
        }
    }
}

/// Represents a single field by which to sort the results of a search.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/fieldorder>
#[derive(Clone, Debug, XmlSerialize)]
pub struct FieldOrder {
    #[xml_struct(flatten)]
    pub path: PathToElement,
    #[xml_struct(attribute)]
    pub order: SortDirection,
}

/// The direction in which to sort the results of a search.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/sortdirection>
#[derive(Clone, Copy, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum SortDirection {
    Ascending,
    Descending,
}

/// The common format for item move and copy operations.
#[derive(Clone, Debug, XmlSerialize)]
pub struct CopyMoveItemData {
    /// The destination folder for the copied/moved item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/tofolderid>
    pub to_folder_id: BaseFolderId,
    /// The unique identifiers for each item to copy/move.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemids>
    pub item_ids: Vec<BaseItemId>,
    /// Whether or not to return the new item idententifers in the response.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/returnnewitemids>
    pub return_new_item_ids: Option<bool>,
}

/// The common format for folder move and copy operations.
#[derive(Clone, Debug, XmlSerialize)]
pub struct CopyMoveFolderData {
    /// The destination folder for the copied/moved folder.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/tofolderid>
    pub to_folder_id: BaseFolderId,

    /// The identifiers for each folder to copy/move.
    ///
    /// <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folderids>
    pub folder_ids: Vec<BaseFolderId>,
}

/// The common format of folder response messages.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct FolderResponseMessage {
    pub folders: Folders,
}

/// The common format of item response messages.
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ItemResponseMessage {
    pub items: Items,
}

/// An identifier for an Exchange folder.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum BaseFolderId {
    /// An identifier for an arbitrary folder.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folderid>.
    FolderId {
        #[xml_struct(attribute)]
        id: String,

        #[xml_struct(attribute)]
        change_key: Option<String>,
    },

    /// An identifier for referencing a folder by name, e.g. "inbox" or
    /// "junkemail".
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/distinguishedfolderid>.
    DistinguishedFolderId {
        #[xml_struct(attribute)]
        id: String,

        #[xml_struct(attribute)]
        change_key: Option<String>,
    },
}

/// The unique identifier of a folder.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folderid>
#[derive(Clone, Debug, Deserialize, PartialEq, XmlSerialize, Eq)]
pub struct FolderId {
    #[serde(rename = "@Id")]
    pub id: String,

    #[serde(rename = "@ChangeKey")]
    pub change_key: Option<String>,
}

/// The manner in which items or folders are deleted.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/deletetype>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum DeleteType {
    HardDelete,
    MoveToDeletedItems,
    SoftDelete,
}

/// An identifier for an Exchange item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemids>
// N.B.: Commented-out variants are not yet implemented.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum BaseItemId {
    /// An identifier for a standard Exchange item.
    ItemId {
        #[xml_struct(attribute)]
        id: String,

        #[xml_struct(attribute)]
        change_key: Option<String>,
    },
    // OccurrenceItemId { .. }
    // RecurringMasterItemId { .. }
}

/// The unique identifier of an item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemid>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct ItemId {
    #[xml_struct(attribute)]
    #[serde(rename = "@Id")]
    pub id: String,

    #[serde(rename = "@ChangeKey")]
    #[xml_struct(attribute)]
    pub change_key: Option<String>,
}

impl ItemId {
    /// Creates a new ItemId with the given ID and no change key.
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            change_key: None,
        }
    }

    /// Creates a new ItemId with the given ID and change key.
    pub fn with_change_key(id: impl Into<String>, change_key: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            change_key: Some(change_key.into()),
        }
    }
}

/// The representation of a folder in an EWS operation.
#[derive(Clone, Debug, Deserialize, XmlSerialize, Eq, PartialEq)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum Folder {
    /// A calendar folder in a mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/calendarfolder>
    #[serde(rename_all = "PascalCase")]
    CalendarFolder {
        #[xml_struct(ns_prefix = "t")]
        folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        parent_folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        folder_class: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        display_name: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        total_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        child_folder_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        extended_property: Option<Vec<ExtendedProperty>>,
    },

    /// A contacts folder in a mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contactsfolder>
    #[serde(rename_all = "PascalCase")]
    ContactsFolder {
        #[xml_struct(ns_prefix = "t")]
        folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        parent_folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        folder_class: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        display_name: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        total_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        child_folder_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        extended_property: Option<Vec<ExtendedProperty>>,
    },

    /// A folder in a mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folder>
    #[serde(rename_all = "PascalCase")]
    Folder {
        #[xml_struct(ns_prefix = "t")]
        folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        parent_folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        folder_class: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        display_name: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        total_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        child_folder_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        extended_property: Option<Vec<ExtendedProperty>>,

        #[xml_struct(ns_prefix = "t")]
        unread_count: Option<u32>,
    },

    /// A search folder in a mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/searchfolder>
    #[serde(rename_all = "PascalCase")]
    SearchFolder {
        #[xml_struct(ns_prefix = "t")]
        folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        parent_folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        folder_class: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        display_name: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        total_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        child_folder_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        extended_property: Option<Vec<ExtendedProperty>>,
    },

    /// A task folder in a mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/tasksfolder>
    #[serde(rename_all = "PascalCase")]
    TasksFolder {
        #[xml_struct(ns_prefix = "t")]
        folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        parent_folder_id: Option<FolderId>,

        #[xml_struct(ns_prefix = "t")]
        folder_class: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        display_name: Option<String>,

        #[xml_struct(ns_prefix = "t")]
        total_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        child_folder_count: Option<u32>,

        #[xml_struct(ns_prefix = "t")]
        extended_property: Option<Vec<ExtendedProperty>>,
    },
}

impl Folder {
    pub fn new_folder(display_name: impl Into<String>) -> Self {
        Self::Folder {
            folder_id: None,
            parent_folder_id: None,
            folder_class: None,
            display_name: Some(display_name.into()),
            total_count: None,
            child_folder_count: None,
            extended_property: None,
            unread_count: None,
        }
    }

    pub fn new_calendar_folder(display_name: impl Into<String>) -> Self {
        Self::CalendarFolder {
            folder_id: None,
            parent_folder_id: None,
            folder_class: None,
            display_name: Some(display_name.into()),
            total_count: None,
            child_folder_count: None,
            extended_property: None,
        }
    }

    pub fn new_contacts_folder(display_name: impl Into<String>) -> Self {
        Self::ContactsFolder {
            folder_id: None,
            parent_folder_id: None,
            folder_class: None,
            display_name: Some(display_name.into()),
            total_count: None,
            child_folder_count: None,
            extended_property: None,
        }
    }

    pub fn new_search_folder(display_name: impl Into<String>) -> Self {
        Self::SearchFolder {
            folder_id: None,
            parent_folder_id: None,
            folder_class: None,
            display_name: Some(display_name.into()),
            total_count: None,
            child_folder_count: None,
            extended_property: None,
        }
    }

    pub fn new_tasks_folder(display_name: impl Into<String>) -> Self {
        Self::TasksFolder {
            folder_id: None,
            parent_folder_id: None,
            folder_class: None,
            display_name: Some(display_name.into()),
            total_count: None,
            child_folder_count: None,
            extended_property: None,
        }
    }
}

/// An array of items.
#[derive(Clone, Debug, Default, Deserialize, PartialEq, Eq)]
pub struct Items {
    #[serde(rename = "$value", default)]
    pub inner: Vec<RealItem>,
}

/// A collection of information on Exchange folders.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folders-ex15websvcsotherref>
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
pub struct Folders {
    #[serde(rename = "$value", default)]
    pub inner: Vec<Folder>,
}

/// An item which may appear as the result of a request to read or modify an
/// Exchange item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/items>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(variant_ns_prefix = "t")]
#[non_exhaustive]
pub enum RealItem {
    Message(Message),
    CalendarItem(Message),
    MeetingMessage(Message),
    MeetingRequest(Message),
    MeetingResponse(Message),
    MeetingCancellation(Message),
}

impl RealItem {
    /// Return the [`Message`] object contained within this [`RealItem`].
    pub fn inner_message(&self) -> &Message {
        match self {
            RealItem::Message(message)
            | RealItem::CalendarItem(message)
            | RealItem::MeetingMessage(message)
            | RealItem::MeetingRequest(message)
            | RealItem::MeetingResponse(message)
            | RealItem::MeetingCancellation(message) => message,
        }
    }

    /// Take ownership of the inner [`Message`].
    pub fn into_inner_message(self) -> Message {
        match self {
            RealItem::Message(message)
            | RealItem::CalendarItem(message)
            | RealItem::MeetingMessage(message)
            | RealItem::MeetingRequest(message)
            | RealItem::MeetingResponse(message)
            | RealItem::MeetingCancellation(message) => message,
        }
    }
}

/// An item which may appear in an item-based attachment.
///
/// See [`Attachment::ItemAttachment`] for details.
// N.B.: Commented-out variants are not yet implemented.
#[non_exhaustive]
#[derive(Clone, Debug, Deserialize)]
pub enum AttachmentItem {
    // Item(Item),
    Message(Message),
    CalendarItem(Message),
    // Contact(Contact),
    // Task(Task),
    MeetingMessage(Message),
    MeetingRequest(Message),
    MeetingResponse(Message),
    MeetingCancellation(Message),
}

/// A date and time with second precision.
// `time` provides an `Option<OffsetDateTime>` deserializer, but it does not
// work with map fields which may be omitted, as in our case.
// We also need to handle Exchange Server timestamps that sometimes omit timezone info.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DateTime(pub time::OffsetDateTime);

// Helper module for flexible datetime deserialization
mod flexible_datetime {
    use serde::{Deserialize, Deserializer};

    pub fn deserialize<'de, D>(deserializer: D) -> Result<time::OffsetDateTime, D::Error>
    where
        D: Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;

        // The time crate's ISO8601 parser can handle both formats if we configure it properly
        // First try with the default ISO8601 parser which expects timezone
        if let Ok(dt) =
            time::OffsetDateTime::parse(&s, &time::format_description::well_known::Iso8601::DEFAULT)
        {
            return Ok(dt);
        }

        // If no timezone, parse as PrimitiveDateTime and assume UTC
        // This handles the Exchange Server case where timezone is omitted
        if let Ok(pdt) = time::PrimitiveDateTime::parse(
            &s,
            &time::format_description::well_known::Iso8601::DEFAULT,
        ) {
            return Ok(pdt.assume_utc());
        }

        Err(serde::de::Error::custom(format!(
            "Unable to parse datetime '{s}'. Expected ISO8601 format with or without timezone."
        )))
    }
}

impl<'de> serde::Deserialize<'de> for DateTime {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        flexible_datetime::deserialize(deserializer).map(DateTime)
    }
}

impl XmlSerialize for DateTime {
    /// Serializes a `DateTime` as an XML text content node by formatting the
    /// inner [`time::OffsetDateTime`] as an ISO 8601-compliant string.
    fn serialize_child_nodes<W>(
        &self,
        writer: &mut quick_xml::Writer<W>,
    ) -> Result<(), xml_struct::Error>
    where
        W: std::io::Write,
    {
        let time = self
            .0
            .format(&Iso8601::DEFAULT)
            .map_err(|err| xml_struct::Error::Value(err.into()))?;

        time.serialize_child_nodes(writer)
    }
}

/// An email message.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/message-ex15websvcsotherref>
#[derive(Clone, Debug, Default, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Message {
    /// The MIME content of the item.
    #[xml_struct(ns_prefix = "t")]
    pub mime_content: Option<MimeContent>,

    /// The item's Exchange identifier.
    #[xml_struct(ns_prefix = "t")]
    pub item_id: Option<ItemId>,

    /// The identifier for the containing folder.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/parentfolderid>
    #[xml_struct(ns_prefix = "t")]
    pub parent_folder_id: Option<FolderId>,

    /// The Exchange class value of the item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemclass>
    #[xml_struct(ns_prefix = "t")]
    pub item_class: Option<String>,

    /// The subject of the item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/subject>
    #[xml_struct(ns_prefix = "t")]
    pub subject: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub sensitivity: Option<Sensitivity>,

    #[xml_struct(ns_prefix = "t")]
    pub body: Option<Body>,

    #[xml_struct(ns_prefix = "t")]
    pub attachments: Option<Attachments>,

    #[xml_struct(ns_prefix = "t")]
    pub date_time_received: Option<DateTime>,

    #[xml_struct(ns_prefix = "t")]
    pub size: Option<usize>,

    /// A list of categories describing an item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/categories-ex15websvcsotherref>
    #[xml_struct(ns_prefix = "t")]
    pub categories: Option<Vec<StringElement>>,

    // Extended MAPI properties of the message.
    #[xml_struct(ns_prefix = "t")]
    pub extended_property: Option<Vec<ExtendedProperty>>,

    #[xml_struct(ns_prefix = "t")]
    pub importance: Option<Importance>,

    #[xml_struct(ns_prefix = "t")]
    pub in_reply_to: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub is_submitted: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_draft: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_from_me: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_resend: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_unmodified: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub internet_message_headers: Option<InternetMessageHeaders>,

    #[xml_struct(ns_prefix = "t")]
    pub date_time_sent: Option<DateTime>,

    #[xml_struct(ns_prefix = "t")]
    pub date_time_created: Option<DateTime>,

    #[xml_struct(ns_prefix = "t")]
    pub reminder_due_by: Option<DateTime>,

    #[xml_struct(ns_prefix = "t")]
    pub reminder_is_set: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub reminder_minutes_before_start: Option<usize>,

    #[xml_struct(ns_prefix = "t")]
    pub display_cc: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub display_to: Option<String>,

    /// Whether the item has (non-inline) attachments.
    ///
    /// **Important**: According to Microsoft's EWS specification, inline attachments
    /// (embedded images, etc.) are considered "hidden attachments" and do NOT affect
    /// this property. If an item only has inline attachments, this will be `false`.
    ///
    /// To check for any attachments including inline ones, use `has_any_attachments()`.
    /// To access all attachments including inline ones, check the `attachments` field directly.
    ///
    /// See: <https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/attachments-and-ews-in-exchange>
    #[xml_struct(ns_prefix = "t")]
    pub has_attachments: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub culture: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub sender: Option<Recipient>,

    #[xml_struct(ns_prefix = "t")]
    pub to_recipients: Option<ArrayOfRecipients>,

    #[xml_struct(ns_prefix = "t")]
    pub cc_recipients: Option<ArrayOfRecipients>,

    #[xml_struct(ns_prefix = "t")]
    pub bcc_recipients: Option<ArrayOfRecipients>,

    #[xml_struct(ns_prefix = "t")]
    pub is_read_receipt_requested: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_delivery_receipt_requested: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub conversation_index: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub conversation_topic: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub from: Option<Recipient>,

    #[xml_struct(ns_prefix = "t")]
    pub internet_message_id: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub is_read: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub is_response_requested: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub reply_to: Option<Recipient>,

    #[xml_struct(ns_prefix = "t")]
    pub received_by: Option<Recipient>,

    #[xml_struct(ns_prefix = "t")]
    pub received_representing: Option<Recipient>,

    #[xml_struct(ns_prefix = "t")]
    pub last_modified_name: Option<String>,

    #[xml_struct(ns_prefix = "t")]
    pub last_modified_time: Option<DateTime>,

    #[xml_struct(ns_prefix = "t")]
    pub is_associated: Option<bool>,

    #[xml_struct(ns_prefix = "t")]
    pub conversation_id: Option<ItemId>,

    #[xml_struct(ns_prefix = "t")]
    pub references: Option<String>,

    /// Response objects available for this message.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responseobjects>
    #[xml_struct(ns_prefix = "t")]
    pub response_objects: Option<ResponseObjects>,
}

impl Message {
    /// Create a new Message with all fields initialized to None.
    pub fn new() -> Self {
        Self {
            mime_content: None,
            item_id: None,
            parent_folder_id: None,
            item_class: None,
            subject: None,
            sensitivity: None,
            body: None,
            attachments: None,
            date_time_received: None,
            size: None,
            categories: None,
            extended_property: None,
            importance: None,
            in_reply_to: None,
            is_submitted: None,
            is_draft: None,
            is_from_me: None,
            is_resend: None,
            is_unmodified: None,
            internet_message_headers: None,
            date_time_sent: None,
            date_time_created: None,
            reminder_due_by: None,
            reminder_is_set: None,
            reminder_minutes_before_start: None,
            display_cc: None,
            display_to: None,
            has_attachments: None,
            culture: None,
            sender: None,
            to_recipients: None,
            cc_recipients: None,
            bcc_recipients: None,
            is_read_receipt_requested: None,
            is_delivery_receipt_requested: None,
            conversation_index: None,
            conversation_topic: None,
            from: None,
            internet_message_id: None,
            is_read: None,
            is_response_requested: None,
            reply_to: None,
            received_by: None,
            received_representing: None,
            last_modified_name: None,
            last_modified_time: None,
            is_associated: None,
            conversation_id: None,
            references: None,
            response_objects: None,
        }
    }

    /// Create a new Message with basic information for common use cases.
    pub fn with_basic_info(
        subject: impl Into<String>,
        body: Body,
        to_recipients: ArrayOfRecipients,
    ) -> Self {
        let mut message = Self::new();
        message.subject = Some(subject.into());
        message.body = Some(body);
        message.to_recipients = Some(to_recipients);
        message
    }

    /// Returns true if the message has any attachments, including inline attachments.
    ///
    /// This method checks the actual `attachments` collection rather than relying on the
    /// `has_attachments` property, which according to Microsoft's EWS specification does
    /// not include inline attachments (embedded images, etc.).
    ///
    /// Use this method when you need to know if there are any attachments at all,
    /// including inline/embedded ones.
    pub fn has_any_attachments(&self) -> bool {
        self.attachments
            .as_ref()
            .map(|attachments| !attachments.inner.is_empty())
            .unwrap_or(false)
    }

    /// Returns true if the message has any inline attachments (embedded images, etc.).
    ///
    /// Inline attachments are typically embedded in HTML email bodies and referenced
    /// by Content-ID (cid:) links.
    pub fn has_inline_attachments(&self) -> bool {
        self.attachments
            .as_ref()
            .map(|attachments| {
                attachments.inner.iter().any(|attachment| match attachment {
                    Attachment::FileAttachment { is_inline, .. } => is_inline.unwrap_or(false),
                    Attachment::ItemAttachment { is_inline, .. } => is_inline.unwrap_or(false),
                })
            })
            .unwrap_or(false)
    }

    /// Returns true if the message has any non-inline (regular) attachments.
    ///
    /// This typically corresponds to files that users would save or download,
    /// as opposed to embedded images or inline content.
    pub fn has_regular_attachments(&self) -> bool {
        self.attachments
            .as_ref()
            .map(|attachments| {
                attachments.inner.iter().any(|attachment| match attachment {
                    Attachment::FileAttachment { is_inline, .. } => !is_inline.unwrap_or(false),
                    Attachment::ItemAttachment { is_inline, .. } => !is_inline.unwrap_or(false),
                })
            })
            .unwrap_or(false)
    }
}

/// An extended MAPI property of an Exchange item or folder.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/extendedproperty>
#[allow(non_snake_case)]
#[derive(Clone, Debug, Deserialize, XmlSerialize, Eq, PartialEq)]
pub struct ExtendedProperty {
    #[xml_struct(ns_prefix = "t")]
    pub extended_field_URI: ExtendedFieldURI,

    #[xml_struct(ns_prefix = "t")]
    pub value: String,
}

/// A list of attachments.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachments-ex15websvcsotherref>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct Attachments {
    #[serde(rename = "$value")]
    #[xml_struct(flatten)]
    pub inner: Vec<Attachment>,
}

/// A newtype around a vector of `Recipient`s, that is deserialized using
/// `deserialize_recipients`.
#[derive(Clone, Debug, Default, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct ArrayOfRecipients(
    #[serde(deserialize_with = "deserialize_recipients")] pub Vec<Recipient>,
);

impl Deref for ArrayOfRecipients {
    type Target = Vec<Recipient>;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl DerefMut for ArrayOfRecipients {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.0
    }
}

/// A single mailbox.
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Recipient {
    #[xml_struct(ns_prefix = "t")]
    pub mailbox: Mailbox,
}

impl Recipient {
    /// Creates a new Recipient from an email address.
    pub fn new(email_address: impl Into<String>) -> Self {
        Self {
            mailbox: Mailbox {
                name: None,
                email_address: email_address.into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
        }
    }

    /// Creates a new Recipient with both email address and name.
    pub fn with_name(email_address: impl Into<String>, name: impl Into<String>) -> Self {
        Self {
            mailbox: Mailbox {
                name: Some(name.into()),
                email_address: email_address.into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
        }
    }

    /// Creates a new Recipient from a Mailbox.
    pub fn from_mailbox(mailbox: Mailbox) -> Self {
        Self { mailbox }
    }
}

/// Deserializes a list of recipients.
///
/// `quick-xml`'s `serde` implementation requires the presence of an
/// intermediate type when dealing with lists, and this is not compatible with
/// our model for serialization.
///
/// We could directly deserialize into a `Vec<Mailbox>`, which would also
/// simplify this function a bit, but this would mean using different models
/// to represent single vs. multiple recipient(s).
fn deserialize_recipients<'de, D>(deserializer: D) -> Result<Vec<Recipient>, D::Error>
where
    D: Deserializer<'de>,
{
    #[derive(Clone, Debug, Deserialize)]
    #[serde(rename_all = "PascalCase")]
    struct MailboxSequence {
        mailbox: Vec<Mailbox>,
    }

    let seq = MailboxSequence::deserialize(deserializer)?;

    Ok(seq
        .mailbox
        .into_iter()
        .map(|mailbox| Recipient { mailbox })
        .collect())
}

/// A list of Internet Message Format headers.
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct InternetMessageHeaders {
    pub internet_message_header: Vec<InternetMessageHeader>,
}

/// A reference to a user or address which can send or receive mail.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailbox>
#[derive(Clone, Debug, Default, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct Mailbox {
    /// The name of this mailbox's user.
    #[xml_struct(ns_prefix = "t")]
    pub name: Option<String>,

    /// The email address for this mailbox.
    #[xml_struct(ns_prefix = "t")]
    pub email_address: String,

    /// The protocol used in routing to this mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/routingtype-emailaddress>
    pub routing_type: Option<RoutingType>,

    /// The type of sender/recipient represented by this mailbox.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailboxtype>
    pub mailbox_type: Option<MailboxType>,

    /// An identifier for a contact or list of contacts corresponding to this
    /// mailbox.
    pub item_id: Option<ItemId>,
}

impl Mailbox {
    /// Create a new Mailbox with the given email address.
    pub fn new(email_address: impl Into<String>) -> Self {
        Self {
            name: None,
            email_address: email_address.into(),
            routing_type: None,
            mailbox_type: None,
            item_id: None,
        }
    }

    /// Create a new Mailbox with the given email address and name.
    pub fn with_name(email_address: impl Into<String>, name: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            email_address: email_address.into(),
            routing_type: None,
            mailbox_type: None,
            item_id: None,
        }
    }
}

/// A protocol used in routing mail.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/routingtype-emailaddress>
#[derive(Clone, Copy, Debug, Default, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(text)]
pub enum RoutingType {
    #[default]
    SMTP,
    EX,
}

/// The type of sender or recipient a mailbox represents.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailboxtype>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(text)]
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

/// The priority level of an item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/importance>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(text)]
pub enum Importance {
    Low,
    Normal,
    High,
}

/// A string value.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/string>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct StringElement {
    /// The string content.
    pub string: String,
}

/// The sensitivity of the contents of an item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/sensitivity>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(text)]
pub enum Sensitivity {
    Normal,
    Personal,
    Private,
    Confidential,
}

/// The body of an item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/body>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct Body {
    /// The content type of the body.
    #[serde(rename = "@BodyType")]
    #[xml_struct(attribute)]
    pub body_type: BodyType,

    /// Whether the body has been truncated.
    #[serde(rename = "@IsTruncated")]
    #[xml_struct(attribute)]
    pub is_truncated: Option<bool>,

    /// The content of the body.
    // TODO: It's not immediately obvious why this tag may be empty, but it has
    // been encountered in real world responses. Needs a closer look.
    #[serde(rename = "$text")]
    #[xml_struct(flatten)]
    pub content: Option<String>,
}

impl Body {
    /// Create a new Body with text content.
    pub fn text(content: impl Into<String>) -> Self {
        Self {
            body_type: BodyType::Text,
            content: Some(content.into()),
            is_truncated: None,
        }
    }

    /// Create a new Body with HTML content.
    pub fn html(content: impl Into<String>) -> Self {
        Self {
            body_type: BodyType::HTML,
            content: Some(content.into()),
            is_truncated: None,
        }
    }

    /// Create a new Body with the specified body type and content.
    pub fn new(body_type: BodyType, content: impl Into<String>) -> Self {
        Self {
            body_type,
            content: Some(content.into()),
            is_truncated: None,
        }
    }
}

/// The content type of an item's body.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/body>
#[derive(Clone, Copy, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[xml_struct(text)]
pub enum BodyType {
    HTML,
    Text,
}

/// An attachment to an Exchange item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachments-ex15websvcsotherref>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub enum Attachment {
    /// An attachment containing an Exchange item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemattachment>
    #[serde(rename_all = "PascalCase")]
    ItemAttachment {
        /// An identifier for the attachment.
        attachment_id: AttachmentId,

        /// The name of the attachment.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/name-attachmenttype>
        name: String,

        /// The MIME type of the attachment's content.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contenttype>
        /// Sometimes there isn't a content type on an attachment
        content_type: Option<String>,

        /// An arbitrary identifier for the attachment.
        ///
        /// This field is not set by Exchange and is intended for use by
        /// external applications.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contentid>
        content_id: Option<String>,

        /// A URI representing the location of the attachment's content.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contentlocation>
        content_location: Option<String>,

        /// The size of the attachment's content in bytes.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/size>
        size: Option<usize>,

        /// The most recent modification time for the attachment.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/lastmodifiedtime>
        last_modified_time: Option<DateTime>,

        /// Whether the attachment appears inline in the item body.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/isinline>
        is_inline: Option<bool>,
        // XXX: With this field in place, parsing will fail if there is no
        // `AttachmentItem` in the response.
        // See https://github.com/tafia/quick-xml/issues/683
        // /// The attached item.
        // #[serde(rename = "$value")]
        // content: Option<AttachmentItem>,
    },

    /// An attachment containing a file.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/fileattachment>
    #[serde(rename_all = "PascalCase")]
    FileAttachment {
        /// An identifier for the attachment.
        attachment_id: AttachmentId,

        /// The name of the attachment.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/name-attachmenttype>
        name: String,

        /// The MIME type of the attachment's content.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contenttype>
        content_type: String,

        /// An arbitrary identifier for the attachment.
        ///
        /// This field is not set by Exchange and is intended for use by
        /// external applications.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contentid>
        content_id: Option<String>,

        /// A URI representing the location of the attachment's content.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/contentlocation>
        content_location: Option<String>,

        /// The size of the attachment's content in bytes.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/size>
        size: Option<usize>,

        /// The most recent modification time for the attachment.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/lastmodifiedtime>
        last_modified_time: Option<DateTime>,

        /// Whether the attachment appears inline in the item body.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/isinline>
        is_inline: Option<bool>,

        /// Whether the attachment represents a contact photo.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/iscontactphoto>
        is_contact_photo: Option<bool>,

        /// The base64-encoded content of the attachment.
        ///
        /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/content>
        content: Option<String>,
    },
}

/// An identifier for an attachment.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachmentid>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct AttachmentId {
    /// A unique identifier for the attachment.
    #[serde(rename = "@Id")]
    #[xml_struct(attribute)]
    pub id: String,

    /// The unique identifier of the item to which it is attached.
    #[serde(rename = "@RootItemId")]
    #[xml_struct(attribute)]
    pub root_item_id: Option<String>,

    /// The change key of the item to which it is attached.
    #[serde(rename = "@RootItemChangeKey")]
    #[xml_struct(attribute)]
    pub root_item_change_key: Option<String>,
}

/// The content of an item, represented according to MIME (Multipurpose Internet
/// Mail Extensions).
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mimecontent>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
pub struct MimeContent {
    /// The character set of the MIME content if it contains [RFC 2045]-encoded
    /// text.
    ///
    /// [RFC 2045]: https://datatracker.ietf.org/doc/html/rfc2045
    #[serde(rename = "@CharacterSet")]
    #[xml_struct(attribute)]
    pub character_set: Option<String>,

    /// The item content.
    #[serde(rename = "$text")]
    #[xml_struct(flatten)]
    pub content: String,
}

/// The headers of an Exchange item's MIME content.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/internetmessageheader>
#[derive(Clone, Debug, Deserialize, XmlSerialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct InternetMessageHeader {
    /// The name of the header.
    #[serde(rename = "@HeaderName")]
    #[xml_struct(attribute)]
    pub header_name: String,

    /// The value of the header.
    #[serde(rename = "$text")]
    #[xml_struct(flatten)]
    pub value: String,
}

#[cfg(test)]
mod tests {

    use super::*;
    use crate::{test_utils::assert_serialized_content, Error};

    /// Tests that an [`ArrayOfRecipients`] correctly serializes into XML. It
    /// should serialize as multiple `<t:Mailbox>` elements, one per [`Recipient`].
    #[test]
    fn serialize_array_of_recipients() -> Result<(), Error> {
        // Define the recipients to serialize.
        let alice = Recipient {
            mailbox: Mailbox {
                name: Some("Alice Test".into()),
                email_address: "alice@test.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
        };

        let bob = Recipient {
            mailbox: Mailbox {
                name: Some("Bob Test".into()),
                email_address: "bob@test.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
        };

        let recipients = ArrayOfRecipients(vec![alice, bob]);

        // Ensure the structure of the XML document is correct.
        let expected = "<Recipients><t:Mailbox><t:Name>Alice Test</t:Name><t:EmailAddress>alice@test.com</t:EmailAddress></t:Mailbox><t:Mailbox><t:Name>Bob Test</t:Name><t:EmailAddress>bob@test.com</t:EmailAddress></t:Mailbox></Recipients>";

        assert_serialized_content(&recipients, "Recipients", expected);

        Ok(())
    }

    /// Tests that deserializing a sequence of `<t:Mailbox>` XML elements
    /// results in an [`ArrayOfRecipients`] with one [`Recipient`] per
    /// `<t:Mailbox>` element.
    #[test]
    fn deserialize_array_of_recipients() -> Result<(), Error> {
        // The raw XML to deserialize.
        let xml = "<Recipients><t:Mailbox><t:Name>Alice Test</t:Name><t:EmailAddress>alice@test.com</t:EmailAddress></t:Mailbox><t:Mailbox><t:Name>Bob Test</t:Name><t:EmailAddress>bob@test.com</t:EmailAddress></t:Mailbox></Recipients>";

        // Deserialize the raw XML, with `serde_path_to_error` to help
        // troubleshoot any issue.
        let mut de = quick_xml::de::Deserializer::from_reader(xml.as_bytes());
        let recipients: ArrayOfRecipients = serde_path_to_error::deserialize(&mut de)?;

        // Ensure we have the right number of recipients in the resulting
        // `ArrayOfRecipients`.
        assert_eq!(recipients.0.len(), 2);

        // Ensure the first recipient correctly has a name and address.
        assert_eq!(
            recipients.first().expect("no recipient at index 0"),
            &Recipient {
                mailbox: Mailbox {
                    name: Some("Alice Test".into()),
                    email_address: "alice@test.com".into(),
                    routing_type: None,
                    mailbox_type: None,
                    item_id: None,
                },
            }
        );

        // Ensure the second recipient correctly has a name and address.
        assert_eq!(
            recipients.get(1).expect("no recipient at index 1"),
            &Recipient {
                mailbox: Mailbox {
                    name: Some("Bob Test".into()),
                    email_address: "bob@test.com".into(),
                    routing_type: None,
                    mailbox_type: None,
                    item_id: None,
                },
            }
        );

        Ok(())
    }

    /// Tests the creation and serialization of AND compound restrictions.
    #[test]
    fn test_and_restriction() -> Result<(), Error> {
        let subject_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "item:Subject".to_string(),
            },
            "Test Subject".to_string(),
        );
        let sender_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "message:Sender".to_string(),
            },
            "sender@example.com".to_string(),
        );

        let and_restriction = Restriction::and(vec![subject_restriction, sender_restriction]);

        let expected = r#"<Restriction><t:And><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Test Subject"/></t:FieldURIOrConstant></t:IsEqualTo><t:IsEqualTo><t:FieldURI FieldURI="message:Sender"/><t:FieldURIOrConstant><t:Constant Value="sender@example.com"/></t:FieldURIOrConstant></t:IsEqualTo></t:And></Restriction>"#;
        assert_serialized_content(&and_restriction, "Restriction", expected);
        Ok(())
    }

    /// Tests the creation and serialization of OR compound restrictions.
    #[test]
    fn test_or_restriction() -> Result<(), Error> {
        let subject_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "item:Subject".to_string(),
            },
            "Important".to_string(),
        );
        let subject_restriction2 = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "item:Subject".to_string(),
            },
            "Urgent".to_string(),
        );

        let or_restriction = Restriction::or(vec![subject_restriction, subject_restriction2]);

        let expected = r#"<Restriction><t:Or><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Important"/></t:FieldURIOrConstant></t:IsEqualTo><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Urgent"/></t:FieldURIOrConstant></t:IsEqualTo></t:Or></Restriction>"#;
        assert_serialized_content(&or_restriction, "Restriction", expected);
        Ok(())
    }

    // TODO: Uncomment once NOT restriction is implemented
    // /// Tests the creation and serialization of NOT compound restrictions.
    // #[test]
    // fn test_not_restriction() -> Result<(), Error> {
    //     let subject_restriction = Restriction::equal_to(
    //         PathToElement::FieldURI {
    //             field_URI: "item:Subject".to_string(),
    //         },
    //         "Spam".to_string(),
    //     );

    //     let not_restriction = Restriction::not(subject_restriction);

    //     let expected = r#"<Restriction><t:Not><t:Restriction><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Spam"/></t:FieldURIOrConstant></t:IsEqualTo></t:Restriction></t:Not></Restriction>"#;
    //     assert_serialized_content(&not_restriction, "Restriction", expected);
    //     Ok(())
    // }

    /// Tests the creation and serialization of nested compound restrictions.
    #[test]
    fn test_nested_compound_restrictions() -> Result<(), Error> {
        // Create inner OR restriction (Important OR Urgent)
        let important_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "item:Subject".to_string(),
            },
            "Important".to_string(),
        );
        let urgent_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "item:Subject".to_string(),
            },
            "Urgent".to_string(),
        );
        let subject_or = Restriction::or(vec![important_restriction, urgent_restriction]);

        // Create sender restriction
        let sender_restriction = Restriction::equal_to(
            PathToElement::FieldURI {
                field_URI: "message:Sender".to_string(),
            },
            "boss@company.com".to_string(),
        );

        // Create outer AND restriction: (Important OR Urgent) AND (sender is boss)
        let complex_restriction = Restriction::and(vec![subject_or, sender_restriction]);

        let expected = r#"<Restriction><t:And><t:Or><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Important"/></t:FieldURIOrConstant></t:IsEqualTo><t:IsEqualTo><t:FieldURI FieldURI="item:Subject"/><t:FieldURIOrConstant><t:Constant Value="Urgent"/></t:FieldURIOrConstant></t:IsEqualTo></t:Or><t:IsEqualTo><t:FieldURI FieldURI="message:Sender"/><t:FieldURIOrConstant><t:Constant Value="boss@company.com"/></t:FieldURIOrConstant></t:IsEqualTo></t:And></Restriction>"#;
        assert_serialized_content(&complex_restriction, "Restriction", expected);
        Ok(())
    }

    /// Tests the Exists restriction helper method.
    #[test]
    fn test_exists_restriction() -> Result<(), Error> {
        let exists_restriction = Restriction::exists(PathToElement::FieldURI {
            field_URI: "item:HasAttachments".to_string(),
        });

        let expected = r#"<Restriction><t:Exists><t:FieldURI FieldURI="item:HasAttachments"/></t:Exists></Restriction>"#;
        assert_serialized_content(&exists_restriction, "Restriction", expected);
        Ok(())
    }

    /// Tests deserialization of a message with HasAttachments=false but inline attachments present.
    /// This is expected EWS behavior where inline attachments are "hidden attachments" and do not
    /// affect the HasAttachments property according to Microsoft's EWS specification.
    #[test]
    fn test_message_with_inline_attachments_has_attachments_false() -> Result<(), Error> {
        // XML representing a message with HasAttachments=false but inline attachments present
        let xml = r#"<Message xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:HasAttachments>false</t:HasAttachments>
            <t:Subject>Email with inline image</t:Subject>
            <t:Attachments>
                <t:FileAttachment>
                    <t:AttachmentId Id="AAMkAGE3"/>
                    <t:Name>image001.png</t:Name>
                    <t:ContentType>image/png</t:ContentType>
                    <t:ContentId>image001.png@01DB6678.C4A79FB0</t:ContentId>
                    <t:IsInline>true</t:IsInline>
                    <t:IsContactPhoto>false</t:IsContactPhoto>
                </t:FileAttachment>
            </t:Attachments>
        </Message>"#;

        // Deserialize the raw XML, with `serde_path_to_error` to help
        // troubleshoot any issue.
        let mut de = quick_xml::de::Deserializer::from_reader(xml.as_bytes());
        let message: Message = serde_path_to_error::deserialize(&mut de)?;

        // Verify that HasAttachments is false (as per EWS specification for inline attachments)
        assert_eq!(message.has_attachments, Some(false));

        // Verify that attachments are still present
        assert!(message.attachments.is_some());
        let attachments = message.attachments.as_ref().unwrap();
        assert_eq!(attachments.inner.len(), 1);

        // Verify it's an inline attachment
        if let Attachment::FileAttachment { is_inline, .. } = &attachments.inner[0] {
            assert_eq!(*is_inline, Some(true));
        } else {
            panic!("Expected FileAttachment");
        }

        Ok(())
    }

    /// Tests deserialization of a CalendarItem which is commonly found with calendar invites.
    #[test]
    fn test_calendar_item_deserialization() -> Result<(), Error> {
        // XML representing a CalendarItem (commonly seen with calendar invites)
        let xml = r#"<Items xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:CalendarItem>
                <t:Subject>Team Meeting</t:Subject>
                <t:Body BodyType="Text">Weekly team sync meeting</t:Body>
                <t:HasAttachments>false</t:HasAttachments>
            </t:CalendarItem>
        </Items>"#;

        // Deserialize the raw XML
        let mut de = quick_xml::de::Deserializer::from_reader(xml.as_bytes());
        let items: Items = serde_path_to_error::deserialize(&mut de)?;

        // Verify that we have one item and it's a CalendarItem
        assert_eq!(items.inner.len(), 1);

        if let RealItem::CalendarItem(message) = &items.inner[0] {
            assert_eq!(message.subject, Some("Team Meeting".to_string()));
            assert_eq!(message.has_attachments, Some(false));
        } else {
            panic!("Expected CalendarItem but got a different variant");
        }

        Ok(())
    }

    /// Tests deserialization of various meeting-related items.
    #[test]
    fn test_meeting_item_variants_deserialization() -> Result<(), Error> {
        // Test MeetingRequest
        let meeting_request_xml = r#"<Items xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:MeetingRequest>
                <t:Subject>Project Review Meeting</t:Subject>
                <t:Body BodyType="Text">Please join the project review meeting</t:Body>
            </t:MeetingRequest>
        </Items>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(meeting_request_xml.as_bytes());
        let items: Items = serde_path_to_error::deserialize(&mut de)?;
        assert_eq!(items.inner.len(), 1);
        assert!(matches!(&items.inner[0], RealItem::MeetingRequest(_)));

        // Test MeetingResponse
        let meeting_response_xml = r#"<Items xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:MeetingResponse>
                <t:Subject>Accepted: Project Review Meeting</t:Subject>
            </t:MeetingResponse>
        </Items>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(meeting_response_xml.as_bytes());
        let items: Items = serde_path_to_error::deserialize(&mut de)?;
        assert_eq!(items.inner.len(), 1);
        assert!(matches!(&items.inner[0], RealItem::MeetingResponse(_)));

        // Test MeetingCancellation
        let meeting_cancel_xml = r#"<Items xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:MeetingCancellation>
                <t:Subject>Cancelled: Project Review Meeting</t:Subject>
            </t:MeetingCancellation>
        </Items>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(meeting_cancel_xml.as_bytes());
        let items: Items = serde_path_to_error::deserialize(&mut de)?;
        assert_eq!(items.inner.len(), 1);
        assert!(matches!(&items.inner[0], RealItem::MeetingCancellation(_)));

        Ok(())
    }

    /// Tests the convenience methods for checking different types of attachments.
    #[test]
    fn test_message_attachment_convenience_methods() -> Result<(), Error> {
        // Create a message with inline attachment
        let inline_xml = r#"<Message xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:HasAttachments>false</t:HasAttachments>
            <t:Attachments>
                <t:FileAttachment>
                    <t:AttachmentId Id="inline-id"/>
                    <t:Name>image.png</t:Name>
                    <t:ContentType>image/png</t:ContentType>
                    <t:IsInline>true</t:IsInline>
                </t:FileAttachment>
            </t:Attachments>
        </Message>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(inline_xml.as_bytes());
        let inline_message: Message = serde_path_to_error::deserialize(&mut de)?;

        // Test convenience methods for inline attachment
        assert!(inline_message.has_any_attachments());
        assert!(inline_message.has_inline_attachments());
        assert!(!inline_message.has_regular_attachments());

        // Create a message with regular attachment
        let regular_xml = r#"<Message xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:HasAttachments>true</t:HasAttachments>
            <t:Attachments>
                <t:FileAttachment>
                    <t:AttachmentId Id="regular-id"/>
                    <t:Name>document.pdf</t:Name>
                    <t:ContentType>application/pdf</t:ContentType>
                    <t:IsInline>false</t:IsInline>
                </t:FileAttachment>
            </t:Attachments>
        </Message>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(regular_xml.as_bytes());
        let regular_message: Message = serde_path_to_error::deserialize(&mut de)?;

        // Test convenience methods for regular attachment
        assert!(regular_message.has_any_attachments());
        assert!(!regular_message.has_inline_attachments());
        assert!(regular_message.has_regular_attachments());

        // Create a message with no attachments
        let no_attachments_xml = r#"<Message xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:HasAttachments>false</t:HasAttachments>
            <t:Subject>No attachments</t:Subject>
        </Message>"#;

        let mut de = quick_xml::de::Deserializer::from_reader(no_attachments_xml.as_bytes());
        let no_attachments_message: Message = serde_path_to_error::deserialize(&mut de)?;

        // Test convenience methods for no attachments
        assert!(!no_attachments_message.has_any_attachments());
        assert!(!no_attachments_message.has_inline_attachments());
        assert!(!no_attachments_message.has_regular_attachments());

        Ok(())
    }

    /// Tests deserialization of mixed attachment types (ItemAttachment and FileAttachment)
    /// This reproduces the issue from GitHub issue #22 where ItemAttachment lacks ContentType
    #[test]
    fn test_mixed_attachments_item_without_content_type() -> Result<(), Error> {
        // XML representing a message with both ItemAttachment (missing ContentType) and FileAttachment
        let xml = r#"<Message xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <t:Subject>Test Meeting - Walkthrough Invite</t:Subject>
            <t:HasAttachments>true</t:HasAttachments>
            <t:Attachments>
                <t:ItemAttachment>
                    <t:AttachmentId Id="AAMkADExampleItemAttachmentIdForTesting123"/>
                    <t:Name>Test Meeting - Walkthrough Invite</t:Name>
                    <t:Size>14300</t:Size>
                    <t:LastModifiedTime>2025-01-01T12:00:00</t:LastModifiedTime>
                    <t:IsInline>false</t:IsInline>
                </t:ItemAttachment>
                <t:FileAttachment>
                    <t:AttachmentId Id="AAMkADExampleFileAttachmentIdForTesting456"/>
                    <t:Name>test-image.png</t:Name>
                    <t:ContentType>image/png</t:ContentType>
                    <t:ContentId>test-image.png@example.test</t:ContentId>
                    <t:Size>14726</t:Size>
                    <t:LastModifiedTime>2025-01-01T12:00:00</t:LastModifiedTime>
                    <t:IsInline>true</t:IsInline>
                    <t:IsContactPhoto>false</t:IsContactPhoto>
                </t:FileAttachment>
            </t:Attachments>
        </Message>"#;

        // This should work after our fix - ItemAttachment is missing ContentType
        let mut de = quick_xml::de::Deserializer::from_reader(xml.as_bytes());
        let message: Message = serde_path_to_error::deserialize(&mut de)?;

        // Verify that HasAttachments is true
        assert_eq!(message.has_attachments, Some(true));

        // Verify that attachments are present
        assert!(message.attachments.is_some());
        let attachments = message.attachments.as_ref().unwrap();
        assert_eq!(attachments.inner.len(), 2);

        // Verify first attachment is ItemAttachment without ContentType
        if let Attachment::ItemAttachment {
            name,
            is_inline,
            content_type,
            ..
        } = &attachments.inner[0]
        {
            assert_eq!(name, "Test Meeting - Walkthrough Invite");
            assert_eq!(*is_inline, Some(false));
            // ContentType should be None since it wasn't provided in XML (this is the key fix!)
            assert_eq!(content_type.as_ref(), None);
        } else {
            panic!("Expected ItemAttachment");
        }

        // Verify second attachment is FileAttachment with ContentType
        if let Attachment::FileAttachment {
            name,
            is_inline,
            content_type,
            content_id,
            ..
        } = &attachments.inner[1]
        {
            assert_eq!(name, "test-image.png");
            assert_eq!(*is_inline, Some(true));
            assert_eq!(content_type, "image/png");
            assert_eq!(content_id.as_ref().unwrap(), "test-image.png@example.test");
        } else {
            panic!("Expected FileAttachment");
        }

        Ok(())
    }
}
