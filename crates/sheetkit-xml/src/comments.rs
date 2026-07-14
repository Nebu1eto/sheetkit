//! Comments XML schema structures.
//!
//! Represents `xl/comments{N}.xml` in the OOXML package.

use serde::{Deserialize, Deserializer, Serialize};

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
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct CommentText {
    #[serde(rename = "r", default)]
    pub runs: Vec<CommentRun>,
}

impl<'de> Deserialize<'de> for CommentText {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let elements = CommentTextElements::deserialize(deserializer)?;
        Ok(Self {
            runs: elements
                .content
                .into_iter()
                .map(|element| match element {
                    CommentTextElement::Run(run) => run,
                    CommentTextElement::DirectText(t) => CommentRun {
                        rpr: None,
                        t: t.value,
                    },
                })
                .collect(),
        })
    }
}

#[derive(Deserialize)]
struct CommentTextElements {
    #[serde(rename = "$value", default)]
    content: Vec<CommentTextElement>,
}

#[derive(Deserialize)]
enum CommentTextElement {
    #[serde(rename = "r")]
    Run(CommentRun),
    #[serde(rename = "t")]
    DirectText(CommentTextValue),
}

#[derive(Deserialize)]
struct CommentTextValue {
    #[serde(rename = "$value", default)]
    value: String,
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

    #[test]
    fn test_comment_text_deserializes_direct_rich_and_mixed_text_in_order() {
        let xml = r#"
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Author</author></authors>
  <commentList>
    <comment ref="A1" authorId="0"><text><t>plain text &amp; more</t></text></comment>
    <comment ref="A2" authorId="0"><text><r><rPr><b/></rPr><t>rich text</t></r></text></comment>
    <comment ref="A3" authorId="0"><text><t>first part</t><r><t>middle part</t></r><t>last part</t></text></comment>
  </commentList>
</comments>"#;

        let parsed: Comments = quick_xml::de::from_str(xml).unwrap();
        let direct = &parsed.comment_list.comments[0].text;
        assert_eq!(
            direct
                .runs
                .iter()
                .map(|run| run.t.as_str())
                .collect::<String>(),
            "plain text & more"
        );
        assert!(direct.runs[0].rpr.is_none());

        let rich = &parsed.comment_list.comments[1].text;
        assert_eq!(rich.runs[0].t, "rich text");
        assert!(rich.runs[0].rpr.is_some());

        let mixed = &parsed.comment_list.comments[2].text;
        assert_eq!(
            mixed
                .runs
                .iter()
                .map(|run| run.t.as_str())
                .collect::<String>(),
            "first partmiddle partlast part"
        );

        let serialized = quick_xml::se::to_string(&parsed).unwrap();
        assert_eq!(serialized.matches("plain text &amp; more").count(), 1);
        assert_eq!(serialized.matches("rich text").count(), 1);
        assert_eq!(serialized.matches("first part").count(), 1);
        let reparsed: Comments = quick_xml::de::from_str(&serialized).unwrap();
        assert_eq!(reparsed, parsed);
    }
}
