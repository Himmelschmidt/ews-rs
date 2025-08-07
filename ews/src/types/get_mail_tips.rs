/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, Mailbox, Operation, OperationResponse, ResponseClass,
    MESSAGES_NS_URI,
};

/// A request to get mail tips for specified recipients.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getmailtips>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct GetMailTips {
    /// The email address sending the message.
    pub sending_as: Mailbox,

    /// List of recipients to check for mail tips.
    pub recipients: Vec<Mailbox>,

    /// Types of mail tips to retrieve.
    pub mail_tips_requested: MailTipsRequested,
}

impl Operation for GetMailTips {
    type Response = GetMailTipsResponse;
}

impl EnvelopeBodyContents for GetMailTips {
    fn name() -> &'static str {
        "GetMailTips"
    }
}

/// Types of mail tips that can be requested.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailtipsrequested>
#[derive(Clone, Debug, Deserialize, XmlSerialize)]
#[xml_struct(text)]
pub enum MailTipsRequested {
    All,
    OutOfOfficeMessage,
    MailboxFullStatus,
    CustomMailTip,
    ExternalMemberCount,
    TotalMemberCount,
    MaxMessageSize,
    DeliveryRestriction,
    ModerationStatus,
    InvalidRecipient,
}

/// A response to a [`GetMailTips`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getmailtipsresponse>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetMailTipsResponse {
    pub response_messages: GetMailTipsResponseMessages,
}

impl OperationResponse for GetMailTipsResponse {}

impl EnvelopeBodyContents for GetMailTipsResponse {
    fn name() -> &'static str {
        "GetMailTipsResponse"
    }
}

/// A collection of responses for individual entities within a request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsemessages>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetMailTipsResponseMessages {
    pub get_mail_tips_response_message: Vec<ResponseClass<GetMailTipsResponseMessage>>,
}

/// A response to a request for getting mail tips.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/getmailtipsresponsemessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct GetMailTipsResponseMessage {
    /// Mail tips for the recipients.
    pub mail_tips: Option<Vec<MailTips>>,
}

/// Mail tips information for a recipient.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/mailtips>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MailTips {
    /// The recipient for whom the mail tips apply.
    pub recipient_address: Mailbox,

    /// Indicates if the recipient is valid.
    pub pending_mail_tips: Option<MailTipsRequested>,

    /// Out of office message if available.
    pub out_of_office: Option<OutOfOffice>,

    /// Whether the mailbox is full.
    pub mailbox_full: Option<bool>,

    /// Custom mail tip message.
    pub custom_mail_tip: Option<String>,

    /// Total number of members if recipient is a distribution list.
    pub total_member_count: Option<u32>,

    /// Number of external members if recipient is a distribution list.
    pub external_member_count: Option<u32>,

    /// Maximum message size the recipient can accept.
    pub max_message_size: Option<u32>,

    /// Delivery restriction information.
    pub delivery_restricted: Option<bool>,

    /// Whether the recipient is moderated.
    pub is_moderated: Option<bool>,

    /// Whether the recipient address is valid.
    pub invalid_recipient: Option<bool>,
}

/// Out of office information for a recipient.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/outofoffice>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct OutOfOffice {
    /// The out of office reply message.
    pub reply_message: Option<OutOfOfficeReplyMessage>,
}

/// Out of office reply message.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/replymessage>
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct OutOfOfficeReplyMessage {
    /// The message content.
    #[serde(rename = "$text")]
    pub message: Option<String>,

    /// The culture/language of the message.
    #[serde(rename = "@xml:lang")]
    pub culture: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{test_utils::assert_serialized_content, Error};

    #[test]
    fn test_get_mail_tips_operation_name() {
        let get_mail_tips = GetMailTips {
            sending_as: Mailbox {
                name: Some("Test User".into()),
                email_address: "test@example.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
            recipients: vec![Mailbox {
                name: Some("Recipient".into()),
                email_address: "recipient@example.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            }],
            mail_tips_requested: MailTipsRequested::All,
        };

        assert_eq!(get_mail_tips.name(), "GetMailTips");
    }

    #[test]
    fn test_serialize_get_mail_tips() -> Result<(), Error> {
        let get_mail_tips = GetMailTips {
            sending_as: Mailbox {
                name: Some("Test User".into()),
                email_address: "test@example.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            },
            recipients: vec![Mailbox {
                name: Some("Recipient".into()),
                email_address: "recipient@example.com".into(),
                routing_type: None,
                mailbox_type: None,
                item_id: None,
            }],
            mail_tips_requested: MailTipsRequested::OutOfOfficeMessage,
        };

        let expected = r#"<GetMailTips xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"><SendingAs><t:Name>Test User</t:Name><t:EmailAddress>test@example.com</t:EmailAddress></SendingAs><Recipients><t:Name>Recipient</t:Name><t:EmailAddress>recipient@example.com</t:EmailAddress></Recipients><MailTipsRequested>OutOfOfficeMessage</MailTipsRequested></GetMailTips>"#;

        assert_serialized_content(&get_mail_tips, "GetMailTips", expected);

        Ok(())
    }
}
