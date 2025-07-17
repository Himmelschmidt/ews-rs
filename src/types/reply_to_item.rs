/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, ArrayOfRecipients, Body, ItemId, ItemResponseMessage,
    MessageDisposition, Operation, OperationResponse, Recipient, MESSAGES_NS_URI,
};

/// A reply to the sender of an item in the Exchange store.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/replytoitem>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct ReplyToItem {
    /// The action the Exchange server will take upon creating this reply.
    #[xml_struct(attribute)]
    pub message_disposition: Option<MessageDisposition>,

    /// The subject of the reply message.
    #[xml_struct(ns_prefix = "t")]
    pub subject: Option<String>,

    /// The body content of the reply message.
    #[xml_struct(ns_prefix = "t")]
    pub body: Option<Body>,

    /// The recipients of the reply message.
    #[xml_struct(ns_prefix = "t")]
    pub to_recipients: Option<ArrayOfRecipients>,

    /// The CC recipients of the reply message.
    #[xml_struct(ns_prefix = "t")]
    pub cc_recipients: Option<ArrayOfRecipients>,

    /// The BCC recipients of the reply message.
    #[xml_struct(ns_prefix = "t")]
    pub bcc_recipients: Option<ArrayOfRecipients>,

    /// Whether a read receipt is requested for the reply.
    #[xml_struct(ns_prefix = "t")]
    pub is_read_receipt_requested: Option<bool>,

    /// Whether a delivery receipt is requested for the reply.
    #[xml_struct(ns_prefix = "t")]
    pub is_delivery_receipt_requested: Option<bool>,

    /// The sender of the reply message when sent by a delegate.
    #[xml_struct(ns_prefix = "t")]
    pub from: Option<Recipient>,

    /// The identifier of the item being replied to.
    #[xml_struct(ns_prefix = "t")]
    pub reference_item_id: ItemId,

    /// The new body content that will be prepended to the original message.
    #[xml_struct(ns_prefix = "t")]
    pub new_body_content: Option<Body>,

    /// The mailbox that received the original message.
    #[xml_struct(ns_prefix = "t")]
    pub received_by: Option<Recipient>,

    /// The user on whose behalf the original message was received.
    #[xml_struct(ns_prefix = "t")]
    pub received_representing: Option<Recipient>,
}

impl Operation for ReplyToItem {
    type Response = ReplyToItemResponse;
}

impl EnvelopeBodyContents for ReplyToItem {
    fn name() -> &'static str {
        "ReplyToItem"
    }
}

/// A response to a [`ReplyToItem`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/replytoitemresponse>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ReplyToItemResponse {
    pub response_messages: ReplyToItemResponseMessages,
}

impl OperationResponse for ReplyToItemResponse {}

impl EnvelopeBodyContents for ReplyToItemResponse {
    fn name() -> &'static str {
        "ReplyToItemResponse"
    }
}

/// A collection of responses for individual entities within a request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsemessages>
#[derive(Clone, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "PascalCase")]
pub struct ReplyToItemResponseMessages {
    pub reply_to_item_response_message: Vec<ItemResponseMessage>,
}

#[cfg(test)]
mod tests {
    use crate::{
        test_utils::{assert_deserialized_content, assert_serialized_content},
        ArrayOfRecipients, Body, BodyType, ItemId, ItemResponseMessage, Items, Mailbox,
        MessageDisposition, Recipient, ResponseClass, ResponseCode,
    };

    use super::{ReplyToItem, ReplyToItemResponse, ReplyToItemResponseMessages};

    #[test]
    fn test_serialize_reply_to_item() {
        let reply_to_item = ReplyToItem {
            message_disposition: Some(MessageDisposition::SendAndSaveCopy),
            subject: Some("Re: Test Subject".to_string()),
            body: Some(Body {
                body_type: BodyType::Text,
                is_truncated: None,
                content: Some("This is my reply.".to_string()),
            }),
            to_recipients: Some(ArrayOfRecipients(vec![Recipient {
                mailbox: Mailbox {
                    name: Some("John Doe".to_string()),
                    email_address: "john.doe@example.com".to_string(),
                    routing_type: None,
                    mailbox_type: None,
                    item_id: None,
                },
            }])),
            cc_recipients: None,
            bcc_recipients: None,
            is_read_receipt_requested: Some(false),
            is_delivery_receipt_requested: Some(false),
            from: None,
            reference_item_id: ItemId {
                id: "AAAtAEF/swbAAA=".to_string(),
                change_key: Some("EwAAABYA/s4b".to_string()),
            },
            new_body_content: Some(Body {
                body_type: BodyType::Text,
                is_truncated: None,
                content: Some("This is my reply.".to_string()),
            }),
            received_by: None,
            received_representing: None,
        };

        let expected = r#"<ReplyToItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" MessageDisposition="SendAndSaveCopy"><t:Subject>Re: Test Subject</t:Subject><t:Body BodyType="Text">This is my reply.</t:Body><t:ToRecipients><t:Mailbox><t:Name>John Doe</t:Name><t:EmailAddress>john.doe@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients><t:IsReadReceiptRequested>false</t:IsReadReceiptRequested><t:IsDeliveryReceiptRequested>false</t:IsDeliveryReceiptRequested><t:ReferenceItemId Id="AAAtAEF/swbAAA=" ChangeKey="EwAAABYA/s4b"/><t:NewBodyContent BodyType="Text">This is my reply.</t:NewBodyContent></ReplyToItem>"#;

        assert_serialized_content(&reply_to_item, "ReplyToItem", expected);
    }

    #[test]
    fn test_deserialize_reply_to_item_response() {
        let content = r#"<ReplyToItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                        <m:ResponseMessages>
                            <m:ReplyToItemResponseMessage ResponseClass="Success">
                                <m:ResponseCode>NoError</m:ResponseCode>
                                <m:Items />
                            </m:ReplyToItemResponseMessage>
                        </m:ResponseMessages>
                        </ReplyToItemResponse>"#;

        let expected = ReplyToItemResponse {
            response_messages: ReplyToItemResponseMessages {
                reply_to_item_response_message: vec![ItemResponseMessage {
                    response_class: ResponseClass::Success,
                    response_code: Some(ResponseCode::NoError),
                    message_text: None,
                    items: Items { inner: vec![] },
                }],
            },
        };

        assert_deserialized_content(content, expected);
    }
}