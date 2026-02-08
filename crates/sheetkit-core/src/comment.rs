//! Comment management utilities.
//!
//! Provides functions for adding, querying, and removing cell comments.

use sheetkit_xml::comments::{Authors, Comment, CommentList, CommentRun, CommentText, Comments};
use sheetkit_xml::namespaces;

/// Configuration for a cell comment.
#[derive(Debug, Clone, PartialEq)]
pub struct CommentConfig {
    /// The cell reference (e.g. "A1").
    pub cell: String,
    /// The author of the comment.
    pub author: String,
    /// The plain text of the comment.
    pub text: String,
}

/// Add a comment to a sheet's comments collection.
///
/// If `comments` is `None`, a new `Comments` structure is created.
pub fn add_comment(comments: &mut Option<Comments>, config: &CommentConfig) {
    let c = comments.get_or_insert_with(|| Comments {
        xmlns: namespaces::SPREADSHEET_ML.to_string(),
        authors: Authors {
            authors: Vec::new(),
        },
        comment_list: CommentList {
            comments: Vec::new(),
        },
    });

    // Find or add the author.
    let author_id = match c.authors.authors.iter().position(|a| a == &config.author) {
        Some(idx) => idx as u32,
        None => {
            c.authors.authors.push(config.author.clone());
            (c.authors.authors.len() - 1) as u32
        }
    };

    // Remove existing comment on the same cell if any.
    c.comment_list
        .comments
        .retain(|comment| comment.r#ref != config.cell);

    // Add the new comment.
    c.comment_list.comments.push(Comment {
        r#ref: config.cell.clone(),
        author_id,
        text: CommentText {
            runs: vec![CommentRun {
                rpr: None,
                t: config.text.clone(),
            }],
        },
    });
}

/// Get the comment for a specific cell.
///
/// Returns `None` if there is no comment on the cell.
pub fn get_comment(comments: &Option<Comments>, cell: &str) -> Option<CommentConfig> {
    let c = comments.as_ref()?;
    let comment = c.comment_list.comments.iter().find(|cm| cm.r#ref == cell)?;

    let author = c
        .authors
        .authors
        .get(comment.author_id as usize)
        .cloned()
        .unwrap_or_default();

    let text = comment
        .text
        .runs
        .iter()
        .map(|r| r.t.as_str())
        .collect::<Vec<_>>()
        .join("");

    Some(CommentConfig {
        cell: cell.to_string(),
        author,
        text,
    })
}

/// Remove a comment from a specific cell.
///
/// Returns `true` if a comment was found and removed.
pub fn remove_comment(comments: &mut Option<Comments>, cell: &str) -> bool {
    if let Some(ref mut c) = comments {
        let before = c.comment_list.comments.len();
        c.comment_list
            .comments
            .retain(|comment| comment.r#ref != cell);
        let removed = c.comment_list.comments.len() < before;

        // Clean up if no comments remain.
        if c.comment_list.comments.is_empty() {
            *comments = None;
        }

        removed
    } else {
        false
    }
}

/// Get all comments from a sheet's comments collection.
pub fn get_all_comments(comments: &Option<Comments>) -> Vec<CommentConfig> {
    match comments.as_ref() {
        Some(c) => c
            .comment_list
            .comments
            .iter()
            .map(|comment| {
                let author = c
                    .authors
                    .authors
                    .get(comment.author_id as usize)
                    .cloned()
                    .unwrap_or_default();
                let text = comment
                    .text
                    .runs
                    .iter()
                    .map(|r| r.t.as_str())
                    .collect::<Vec<_>>()
                    .join("");
                CommentConfig {
                    cell: comment.r#ref.clone(),
                    author,
                    text,
                }
            })
            .collect(),
        None => Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_add_comment() {
        let mut comments = None;
        let config = CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Hello comment".to_string(),
        };
        add_comment(&mut comments, &config);

        assert!(comments.is_some());
        let c = comments.as_ref().unwrap();
        assert_eq!(c.authors.authors.len(), 1);
        assert_eq!(c.authors.authors[0], "Alice");
        assert_eq!(c.comment_list.comments.len(), 1);
        assert_eq!(c.comment_list.comments[0].r#ref, "A1");
    }

    #[test]
    fn test_get_comment() {
        let mut comments = None;
        let config = CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test comment".to_string(),
        };
        add_comment(&mut comments, &config);

        let result = get_comment(&comments, "A1");
        assert!(result.is_some());
        let c = result.unwrap();
        assert_eq!(c.cell, "A1");
        assert_eq!(c.author, "Alice");
        assert_eq!(c.text, "Test comment");
    }

    #[test]
    fn test_get_comment_nonexistent() {
        let comments: Option<Comments> = None;
        assert!(get_comment(&comments, "A1").is_none());
    }

    #[test]
    fn test_get_comment_wrong_cell() {
        let mut comments = None;
        let config = CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test".to_string(),
        };
        add_comment(&mut comments, &config);
        assert!(get_comment(&comments, "B1").is_none());
    }

    #[test]
    fn test_remove_comment() {
        let mut comments = None;
        let config = CommentConfig {
            cell: "A1".to_string(),
            author: "Alice".to_string(),
            text: "Test".to_string(),
        };
        add_comment(&mut comments, &config);
        assert!(remove_comment(&mut comments, "A1"));
        assert!(comments.is_none());
    }

    #[test]
    fn test_remove_nonexistent_comment() {
        let mut comments: Option<Comments> = None;
        assert!(!remove_comment(&mut comments, "A1"));
    }

    #[test]
    fn test_multiple_comments_different_cells() {
        let mut comments = None;
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "Comment 1".to_string(),
            },
        );
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Comment 2".to_string(),
            },
        );
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "C3".to_string(),
                author: "Alice".to_string(),
                text: "Comment 3".to_string(),
            },
        );

        let all = get_all_comments(&comments);
        assert_eq!(all.len(), 3);

        // Verify authors are deduplicated
        let c = comments.as_ref().unwrap();
        assert_eq!(c.authors.authors.len(), 2); // Alice and Bob

        // Verify individual lookups
        let c1 = get_comment(&comments, "A1").unwrap();
        assert_eq!(c1.text, "Comment 1");
        assert_eq!(c1.author, "Alice");

        let c2 = get_comment(&comments, "B2").unwrap();
        assert_eq!(c2.text, "Comment 2");
        assert_eq!(c2.author, "Bob");
    }

    #[test]
    fn test_overwrite_comment_on_same_cell() {
        let mut comments = None;
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "Original".to_string(),
            },
        );
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Bob".to_string(),
                text: "Updated".to_string(),
            },
        );

        let all = get_all_comments(&comments);
        assert_eq!(all.len(), 1);
        assert_eq!(all[0].text, "Updated");
        assert_eq!(all[0].author, "Bob");
    }

    #[test]
    fn test_remove_one_of_multiple_comments() {
        let mut comments = None;
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Alice".to_string(),
                text: "First".to_string(),
            },
        );
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "B2".to_string(),
                author: "Bob".to_string(),
                text: "Second".to_string(),
            },
        );

        assert!(remove_comment(&mut comments, "A1"));
        assert!(comments.is_some()); // Still has B2

        let all = get_all_comments(&comments);
        assert_eq!(all.len(), 1);
        assert_eq!(all[0].cell, "B2");
    }

    #[test]
    fn test_get_all_comments_empty() {
        let comments: Option<Comments> = None;
        let all = get_all_comments(&comments);
        assert!(all.is_empty());
    }

    #[test]
    fn test_comments_xml_roundtrip() {
        let mut comments = None;
        add_comment(
            &mut comments,
            &CommentConfig {
                cell: "A1".to_string(),
                author: "Author".to_string(),
                text: "A test comment".to_string(),
            },
        );

        let c = comments.as_ref().unwrap();
        let xml = quick_xml::se::to_string(c).unwrap();
        let parsed: Comments = quick_xml::de::from_str(&xml).unwrap();

        assert_eq!(parsed.authors.authors.len(), 1);
        assert_eq!(parsed.comment_list.comments.len(), 1);
        assert_eq!(parsed.comment_list.comments[0].r#ref, "A1");
        assert_eq!(
            parsed.comment_list.comments[0].text.runs[0].t,
            "A test comment"
        );
    }
}
