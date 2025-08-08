/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use ews_proc_macros::operation_response;
use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{Attachment, AttachmentId, MESSAGES_NS_URI};

/// A request to retrieve one or more attachments from Exchange items.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getattachment>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
#[operation_response(GetAttachmentResponseMessage)]
pub struct GetAttachment {
    /// Optional shape information for the attachment. Typically used to specify
    /// whether to include the attachment content in the response.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachmentshape>
    pub attachment_shape: Option<AttachmentShape>,

    /// The identifiers of the attachments to retrieve.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachmentids>
    pub attachment_ids: Vec<AttachmentId>,
}


/// Describes what information to include in attachment responses.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/attachmentshape>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct AttachmentShape {
    /// Whether to include the attachment content in the response.
    #[xml_struct(attribute)]
    pub include_mime_content: Option<bool>,

    /// The type of body content to include for item attachments.
    pub body_type: Option<BodyType>,

    /// Whether to filter HTML content when the body type is HTML.
    pub filter_html_content: Option<bool>,

    /// Additional properties to include in the response.
    pub additional_properties: Option<Vec<String>>,
}

/// The type of body content to include for item attachments.
#[derive(Clone, Debug, XmlSerialize, Deserialize)]
pub enum BodyType {
    #[serde(rename = "HTML")]
    Html,
    #[serde(rename = "Text")]
    Text,
    #[serde(rename = "Best")]
    Best,
}


#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct GetAttachmentResponseMessage {
    /// The retrieved attachments.
    pub attachments: Option<Attachments>,
}

/// A collection of attachments returned by GetAttachment.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct Attachments {
    #[serde(rename = "$value", default)]
    pub inner: Vec<Attachment>,
}

#[cfg(test)]
mod test {
    use crate::{test_utils::assert_deserialized_content, Attachment, AttachmentId, ResponseClass};

    use super::{Attachments, GetAttachmentResponse, GetAttachmentResponseMessage};
    use crate::ResponseMessages;

    #[test]
    fn test_deserialize_get_attachment_response() {
        let content = r#"<GetAttachmentResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                    <m:ResponseMessages>
                        <m:GetAttachmentResponseMessage ResponseClass="Success">
                            <m:ResponseCode>NoError</m:ResponseCode>
                            <m:Attachments>
                                <t:FileAttachment>
                                    <t:AttachmentId Id="AQMkADEyYjJhY2Y0LWVmYjEtNDNhNS05NTA4LWNkMzQ1ZGI1MmY2NgBGAAADYTcxYjJhNGQ2"/>
                                    <t:Name>test.txt</t:Name>
                                    <t:ContentType>text/plain</t:ContentType>
                                    <t:Size>5</t:Size>
                                    <t:Content>SGVsbG8=</t:Content>
                                </t:FileAttachment>
                            </m:Attachments>
                        </m:GetAttachmentResponseMessage>
                    </m:ResponseMessages>
                    </GetAttachmentResponse>"#;

        let expected = GetAttachmentResponse {
            response_messages: ResponseMessages {
                response_messages: vec![ResponseClass::Success(GetAttachmentResponseMessage {
                    attachments: Some(Attachments {
                        inner: vec![Attachment::FileAttachment {
                            attachment_id: AttachmentId {
                                id: "AQMkADEyYjJhY2Y0LWVmYjEtNDNhNS05NTA4LWNkMzQ1ZGI1MmY2NgBGAAADYTcxYjJhNGQ2".to_string(),
                                root_item_id: None,
                                root_item_change_key: None,
                            },
                            name: "test.txt".to_string(),
                            content_type: "text/plain".to_string(),
                            content_id: None,
                            content_location: None,
                            size: Some(5),
                            last_modified_time: None,
                            is_inline: None,
                            is_contact_photo: None,
                            content: Some("SGVsbG8=".to_string()),
                        }]
                    }),
                })],
            },
        };

        assert_deserialized_content(content, expected);
    }
}
