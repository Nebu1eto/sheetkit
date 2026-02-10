//! Threaded comments XML schema structures.
//!
//! Represents `xl/threadedComments/threadedComment{N}.xml` and
//! `xl/persons/person.xml` in the OOXML package (Excel 2019+).

use serde::{Deserialize, Serialize};

/// Namespace for threaded comments (Excel 2018+).
pub const THREADED_COMMENTS_NS: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/// Content type for threaded comments parts.
pub const THREADED_COMMENTS_CONTENT_TYPE: &str = "application/vnd.ms-excel.threadedcomments+xml";

/// Content type for the person list part.
pub const PERSON_LIST_CONTENT_TYPE: &str = "application/vnd.ms-excel.person+xml";

/// Relationship type for threaded comments (worksheet-level).
pub const REL_TYPE_THREADED_COMMENT: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";

/// Relationship type for the person list (workbook-level).
pub const REL_TYPE_PERSON: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/person";

/// Root element for threaded comments XML.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "ThreadedComments")]
pub struct ThreadedComments {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "threadedComment", default)]
    pub comments: Vec<ThreadedComment>,
}

/// Individual threaded comment entry.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "threadedComment")]
pub struct ThreadedComment {
    #[serde(rename = "@ref")]
    pub cell_ref: String,

    #[serde(rename = "@dT")]
    pub date_time: String,

    #[serde(rename = "@personId")]
    pub person_id: String,

    #[serde(rename = "@id")]
    pub id: String,

    #[serde(rename = "@parentId", skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,

    #[serde(rename = "@done", skip_serializing_if = "Option::is_none", default)]
    pub done: Option<String>,

    pub text: String,
}

/// Root element for the person list XML (`xl/persons/person.xml`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "personList")]
pub struct PersonList {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "person", default)]
    pub persons: Vec<Person>,
}

/// Individual person entry in the person list.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Person {
    #[serde(rename = "@displayName")]
    pub display_name: String,

    #[serde(rename = "@id")]
    pub id: String,

    #[serde(rename = "@userId", skip_serializing_if = "Option::is_none")]
    pub user_id: Option<String>,

    #[serde(rename = "@providerId", skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
}

impl Default for ThreadedComments {
    fn default() -> Self {
        Self {
            xmlns: THREADED_COMMENTS_NS.to_string(),
            comments: Vec::new(),
        }
    }
}

impl Default for PersonList {
    fn default() -> Self {
        Self {
            xmlns: THREADED_COMMENTS_NS.to_string(),
            persons: Vec::new(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_threaded_comments_default() {
        let tc = ThreadedComments::default();
        assert_eq!(tc.xmlns, THREADED_COMMENTS_NS);
        assert!(tc.comments.is_empty());
    }

    #[test]
    fn test_person_list_default() {
        let pl = PersonList::default();
        assert_eq!(pl.xmlns, THREADED_COMMENTS_NS);
        assert!(pl.persons.is_empty());
    }

    #[test]
    fn test_threaded_comment_roundtrip() {
        let tc = ThreadedComments {
            xmlns: THREADED_COMMENTS_NS.to_string(),
            comments: vec![
                ThreadedComment {
                    cell_ref: "A1".to_string(),
                    date_time: "2024-01-15T10:30:00.00".to_string(),
                    person_id: "{PERSON-1}".to_string(),
                    id: "{COMMENT-1}".to_string(),
                    parent_id: None,
                    done: None,
                    text: "Initial comment".to_string(),
                },
                ThreadedComment {
                    cell_ref: "A1".to_string(),
                    date_time: "2024-01-15T11:00:00.00".to_string(),
                    person_id: "{PERSON-2}".to_string(),
                    id: "{REPLY-1}".to_string(),
                    parent_id: Some("{COMMENT-1}".to_string()),
                    done: Some("1".to_string()),
                    text: "This is a reply".to_string(),
                },
            ],
        };

        let xml = quick_xml::se::to_string(&tc).unwrap();
        assert!(xml.contains("A1"));
        assert!(xml.contains("Initial comment"));
        assert!(xml.contains("parentId"));
        assert!(xml.contains("done=\"1\""));

        let parsed: ThreadedComments = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.comments.len(), 2);
        assert_eq!(parsed.comments[0].cell_ref, "A1");
        assert_eq!(parsed.comments[0].id, "{COMMENT-1}");
        assert!(parsed.comments[0].parent_id.is_none());
        assert_eq!(
            parsed.comments[1].parent_id,
            Some("{COMMENT-1}".to_string())
        );
        assert_eq!(parsed.comments[1].done, Some("1".to_string()));
    }

    #[test]
    fn test_person_list_roundtrip() {
        let pl = PersonList {
            xmlns: THREADED_COMMENTS_NS.to_string(),
            persons: vec![Person {
                display_name: "John Doe".to_string(),
                id: "{PERSON-GUID}".to_string(),
                user_id: Some("user@example.com".to_string()),
                provider_id: Some("ADAL".to_string()),
            }],
        };

        let xml = quick_xml::se::to_string(&pl).unwrap();
        assert!(xml.contains("John Doe"));
        assert!(xml.contains("userId"));
        assert!(xml.contains("providerId"));

        let parsed: PersonList = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.persons.len(), 1);
        assert_eq!(parsed.persons[0].display_name, "John Doe");
        assert_eq!(
            parsed.persons[0].user_id,
            Some("user@example.com".to_string())
        );
    }

    #[test]
    fn test_threaded_comment_without_optional_fields() {
        let tc = ThreadedComment {
            cell_ref: "B2".to_string(),
            date_time: "2024-06-01T08:00:00.00".to_string(),
            person_id: "{P1}".to_string(),
            id: "{C1}".to_string(),
            parent_id: None,
            done: None,
            text: "Simple comment".to_string(),
        };

        let xml = quick_xml::se::to_string(&tc).unwrap();
        assert!(!xml.contains("parentId"));
        assert!(!xml.contains("done"));

        let parsed: ThreadedComment = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.parent_id.is_none());
        assert!(parsed.done.is_none());
    }

    #[test]
    fn test_person_without_optional_fields() {
        let p = Person {
            display_name: "Anonymous".to_string(),
            id: "{P-ANON}".to_string(),
            user_id: None,
            provider_id: None,
        };

        let xml = quick_xml::se::to_string(&p).unwrap();
        assert!(!xml.contains("userId"));
        assert!(!xml.contains("providerId"));

        let parsed: Person = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.user_id.is_none());
        assert!(parsed.provider_id.is_none());
    }

    #[test]
    fn test_parse_real_excel_threaded_comment_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" dT="2024-01-15T10:30:00.00" personId="{GUID1}" id="{GUID2}">
    <text>This is the initial comment</text>
  </threadedComment>
  <threadedComment ref="A1" dT="2024-01-15T11:00:00.00" personId="{GUID3}" id="{GUID4}" parentId="{GUID2}" done="1">
    <text>This is a reply</text>
  </threadedComment>
</ThreadedComments>"#;

        let parsed: ThreadedComments = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.comments.len(), 2);
        assert_eq!(parsed.comments[0].text, "This is the initial comment");
        assert_eq!(parsed.comments[1].text, "This is a reply");
        assert_eq!(parsed.comments[1].parent_id, Some("{GUID2}".to_string()));
        assert_eq!(parsed.comments[1].done, Some("1".to_string()));
    }

    #[test]
    fn test_parse_real_excel_person_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="John Doe" id="{GUID}" userId="user@example.com" providerId="ADAL"/>
</personList>"#;

        let parsed: PersonList = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.persons.len(), 1);
        assert_eq!(parsed.persons[0].display_name, "John Doe");
        assert_eq!(parsed.persons[0].id, "{GUID}");
    }
}
