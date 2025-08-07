/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

use serde::Deserialize;
use xml_struct::XmlSerialize;

use crate::{
    types::sealed::EnvelopeBodyContents, BaseFolderId, DeleteType, Operation, OperationResponse,
    ResponseClass, MESSAGES_NS_URI,
};

/// A request to delete all items from one or more folders.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/emptyfolder>
#[derive(Clone, Debug, XmlSerialize)]
#[xml_struct(default_ns = MESSAGES_NS_URI)]
pub struct EmptyFolder {
    /// The method the EWS server will use to perform the deletion of items.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/emptyfolder#deletetype-attribute>
    #[xml_struct(attribute)]
    pub delete_type: DeleteType,

    /// Whether to delete subfolders as well.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/emptyfolder#deletesubfolders-attribute>
    #[xml_struct(attribute)]
    pub delete_subfolders: bool,

    /// A list of folders to empty.
    ///
    /// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/folderids>
    pub folder_ids: Vec<BaseFolderId>,
}

impl Operation for EmptyFolder {
    type Response = EmptyFolderResponse;
}

impl EnvelopeBodyContents for EmptyFolder {
    fn name() -> &'static str {
        "EmptyFolder"
    }
}

/// A response to an [`EmptyFolder`] request.
///
/// See <https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/emptyfolderresponse>
#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct EmptyFolderResponse {
    pub response_messages: ResponseMessages,
}

impl OperationResponse for EmptyFolderResponse {}

impl EnvelopeBodyContents for EmptyFolderResponse {
    fn name() -> &'static str {
        "EmptyFolderResponse"
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct ResponseMessages {
    pub empty_folder_response_message: Vec<ResponseClass<EmptyFolderResponseMessage>>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "PascalCase")]
pub struct EmptyFolderResponseMessage {}

#[cfg(test)]
mod test {
    use crate::{
        test_utils::assert_deserialized_content, BaseFolderId, DeleteType, Operation, ResponseClass,
    };

    use super::{EmptyFolder, EmptyFolderResponse, EmptyFolderResponseMessage, ResponseMessages};

    #[test]
    fn test_deserialize_empty_folder_response() {
        let content = r#"<EmptyFolderResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                    <m:ResponseMessages>
                        <m:EmptyFolderResponseMessage ResponseClass="Success">
                            <m:ResponseCode>NoError</m:ResponseCode>
                        </m:EmptyFolderResponseMessage>
                    </m:ResponseMessages>
                    </EmptyFolderResponse>"#;

        let expected = EmptyFolderResponse {
            response_messages: ResponseMessages {
                empty_folder_response_message: vec![ResponseClass::Success(
                    EmptyFolderResponseMessage {},
                )],
            },
        };

        assert_deserialized_content(content, expected);
    }

    #[test]
    fn test_create_empty_folder_request() {
        let operation = EmptyFolder {
            delete_type: DeleteType::MoveToDeletedItems,
            delete_subfolders: false,
            folder_ids: vec![BaseFolderId::DistinguishedFolderId {
                id: "junkemail".to_string(),
                change_key: None,
            }],
        };

        assert_eq!(operation.name(), "EmptyFolder");
        assert!(!operation.delete_subfolders);
        assert_eq!(operation.folder_ids.len(), 1);
    }
}
