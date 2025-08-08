/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use ews_proc_macros::operation_response;
use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{AttachmentId, BaseItemId, MESSAGES_NS_URI};

/// A request to create one or more attachments on an Exchange item.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/createattachment>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
#[operation_response(CreateAttachmentResponseMessage)]
pub struct CreateAttachment {
    /// The identifier of the parent Exchange store item to which the attachments will be added.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/parentitemid>
    pub parent_item_id: BaseItemId,

    /// The attachments to create.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachments-ex15websvcsotherref>
    pub attachments: Vec<NewAttachment>,
}

/// An attachment to be created, without an existing attachment ID.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum NewAttachment {
    /// An attachment containing an Exchange item.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/itemattachment>
    ItemAttachment {
        /// The name of the attachment.
        name: String,

        /// The MIME content type of the attachment.
        content_type: Option<String>,

        /// The content ID value.
        content_id: Option<String>,

        /// The content location.
        content_location: Option<String>,

        /// Whether the attachment appears inline within the parent item.
        is_inline: Option<bool>,

        /// The attached item content.
        #[xml_struct(flatten)]
        item: AttachmentItemContent,
    },

    /// An attachment containing a file.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/fileattachment>
    FileAttachment {
        /// The name of the attachment.
        name: String,

        /// The MIME content type of the attachment.
        content_type: Option<String>,

        /// The content ID value.
        content_id: Option<String>,

        /// The content location.
        content_location: Option<String>,

        /// Whether the attachment appears inline within the parent item.
        is_inline: Option<bool>,

        /// Whether this is a contact photo.
        is_contact_photo: Option<bool>,

        /// The binary content of the file (base64 encoded).
        content: String,
    },
}

/// Content for item attachments in create requests.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(variant_ns_prefix = "t")]
pub enum AttachmentItemContent {
    /// A message item to attach.
    Message {
        /// The subject of the message.
        subject: Option<String>,

        /// The body of the message.
        body: Option<MessageBody>,

        /// The sender of the message.
        from: Option<SingleRecipient>,

        /// The primary recipients.
        to_recipients: Option<Vec<EmailAddressType>>,

        /// The carbon copy recipients.
        cc_recipients: Option<Vec<EmailAddressType>>,

        /// The blind carbon copy recipients.
        bcc_recipients: Option<Vec<EmailAddressType>>,
    },
}

/// A message body for attachment items.
#[derive(Clone, Debug, XmlSerialize)]
pub struct MessageBody {
    /// The type of the body content.
    #[xml_struct(attribute)]
    pub body_type: BodyTypeValue,

    /// The body content.
    #[xml_struct(flatten)]
    pub content: Option<String>,
}

/// The type of body content.
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(text)]
pub enum BodyTypeValue {
    HTML,
    Text,
}

/// A single recipient for message items.
#[derive(Clone, Debug, XmlSerialize)]
pub struct SingleRecipient {
    /// The mailbox information.
    pub mailbox: EmailAddressType,
}

/// An email address.
#[derive(Clone, Debug, XmlSerialize)]
pub struct EmailAddressType {
    /// The display name.
    pub name: Option<String>,

    /// The email address.
    pub email_address: String,

    /// The routing type (typically "SMTP").
    pub routing_type: Option<String>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct CreateAttachmentResponseMessage {
    /// The created attachments.
    pub attachments: Option<CreatedAttachments>,
}

/// A collection of attachments created by CreateAttachment.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct CreatedAttachments {
    #[serde(rename = "$value", default)]
    pub inner: Vec<CreatedAttachment>,
}

/// An attachment that was created, containing only its ID.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
pub enum CreatedAttachment {
    /// A created item attachment.
    #[serde(rename_all = "PascalCase")]
    ItemAttachment {
        /// The identifier of the created attachment.
        attachment_id: AttachmentId,
    },

    /// A created file attachment.
    #[serde(rename_all = "PascalCase")]
    FileAttachment {
        /// The identifier of the created attachment.
        attachment_id: AttachmentId,
    },
}

#[cfg(test)]
mod test {
    use crate::{test_utils::assert_deserialized_content, AttachmentId, ResponseClass};

    use super::{
        CreateAttachmentResponse, CreateAttachmentResponseMessage, CreatedAttachment,
        CreatedAttachments,
    };
    use crate::ResponseMessages;

    #[test]
    fn test_deserialize_create_attachment_response() {
        let content = r#"<CreateAttachmentResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                    <m:ResponseMessages>
                        <m:CreateAttachmentResponseMessage ResponseClass="Success">
                            <m:ResponseCode>NoError</m:ResponseCode>
                            <m:Attachments>
                                <t:FileAttachment>
                                    <t:AttachmentId Id="AQMkADEyYjJhY2Y0LWVmYjEtNDNhNS05NTA4LWNkMzQ1ZGI1MmY2NgBGAAADYTcxYjJhNGQ2"/>
                                </t:FileAttachment>
                            </m:Attachments>
                        </m:CreateAttachmentResponseMessage>
                    </m:ResponseMessages>
                    </CreateAttachmentResponse>"#;

        let expected = CreateAttachmentResponse {
            response_messages: ResponseMessages {
                response_messages: vec![ResponseClass::Success(CreateAttachmentResponseMessage {
                    attachments: Some(CreatedAttachments {
                        inner: vec![CreatedAttachment::FileAttachment {
                            attachment_id: AttachmentId {
                                id: "AQMkADEyYjJhY2Y0LWVmYjEtNDNhNS05NTA4LWNkMzQ1ZGI1MmY2NgBGAAADYTcxYjJhNGQ2".to_string(),
                                root_item_id: None,
                                root_item_change_key: None,
                            },
                        }]
                    }),
                })],
            },
        };

        assert_deserialized_content(content, expected);
    }
}
