//! Comments XML schema structures.
//!
//! Represents `xl/comments{N}.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Comments root element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "comments")]
pub struct Comments {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "authors")]
    pub authors: Authors,

    #[serde(rename = "commentList")]
    pub comment_list: CommentList,
}

/// Authors container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Authors {
    #[serde(rename = "author", default)]
    pub authors: Vec<String>,
}

/// Comment list container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CommentList {
    #[serde(rename = "comment", default)]
    pub comments: Vec<Comment>,
}

/// Individual comment.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Comment {
    #[serde(rename = "@ref")]
    pub r#ref: String,

    #[serde(rename = "@authorId")]
    pub author_id: u32,

    #[serde(rename = "text")]
    pub text: CommentText,
}

/// Comment text content.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CommentText {
    #[serde(rename = "r", default)]
    pub runs: Vec<CommentRun>,
}

/// A text run within a comment.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CommentRun {
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    pub rpr: Option<CommentRunProperties>,

    #[serde(rename = "t")]
    pub t: String,
}

/// Run properties for comment text formatting.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CommentRunProperties {
    #[serde(rename = "b", skip_serializing_if = "Option::is_none")]
    pub b: Option<BoldFlag>,

    #[serde(rename = "sz", skip_serializing_if = "Option::is_none")]
    pub sz: Option<FontSize>,
}

/// Bold flag element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BoldFlag;

/// Font size element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FontSize {
    #[serde(rename = "@val")]
    pub val: f64,
}

impl Default for Comments {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            authors: Authors {
                authors: Vec::new(),
            },
            comment_list: CommentList {
                comments: Vec::new(),
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_comments_default() {
        let comments = Comments::default();
        assert_eq!(comments.xmlns, namespaces::SPREADSHEET_ML);
        assert!(comments.authors.authors.is_empty());
        assert!(comments.comment_list.comments.is_empty());
    }

    #[test]
    fn test_comments_roundtrip() {
        let comments = Comments {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            authors: Authors {
                authors: vec!["Author1".to_string()],
            },
            comment_list: CommentList {
                comments: vec![Comment {
                    r#ref: "A1".to_string(),
                    author_id: 0,
                    text: CommentText {
                        runs: vec![CommentRun {
                            rpr: None,
                            t: "This is a comment".to_string(),
                        }],
                    },
                }],
            },
        };

        let xml = quick_xml::se::to_string(&comments).unwrap();
        assert!(xml.contains("A1"));
        assert!(xml.contains("This is a comment"));
        assert!(xml.contains("Author1"));

        let parsed: Comments = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.authors.authors.len(), 1);
        assert_eq!(parsed.comment_list.comments.len(), 1);
        assert_eq!(parsed.comment_list.comments[0].r#ref, "A1");
        assert_eq!(parsed.comment_list.comments[0].author_id, 0);
        assert_eq!(
            parsed.comment_list.comments[0].text.runs[0].t,
            "This is a comment"
        );
    }
}
