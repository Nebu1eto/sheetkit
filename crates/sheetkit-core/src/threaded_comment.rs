//! Threaded comments management.
//!
//! Provides functions for adding, querying, resolving, and removing
//! threaded comments (Excel 2019+ feature). Threaded comments support
//! replies and resolved state, unlike legacy comments.

use sheetkit_xml::threaded_comment::{Person, PersonList, ThreadedComment, ThreadedComments};

use crate::error::{Error, Result};

/// Input configuration for adding a threaded comment.
#[derive(Debug, Clone, PartialEq)]
pub struct ThreadedCommentInput {
    /// Display name of the comment author.
    pub author: String,
    /// The comment text.
    pub text: String,
    /// Optional parent comment ID for replies.
    pub parent_id: Option<String>,
}

/// Input configuration for adding a person.
#[derive(Debug, Clone, PartialEq)]
pub struct PersonInput {
    /// Display name of the person.
    pub display_name: String,
    /// Optional user ID (e.g. email address).
    pub user_id: Option<String>,
    /// Optional provider ID (e.g. "ADAL").
    pub provider_id: Option<String>,
}

/// Output data for a threaded comment.
#[derive(Debug, Clone, PartialEq)]
pub struct ThreadedCommentData {
    /// Unique comment ID (GUID).
    pub id: String,
    /// Cell reference (e.g. "A1").
    pub cell_ref: String,
    /// Comment text.
    pub text: String,
    /// Author display name.
    pub author: String,
    /// Person ID (GUID).
    pub person_id: String,
    /// ISO 8601 timestamp.
    pub date_time: String,
    /// Parent comment ID (for replies).
    pub parent_id: Option<String>,
    /// Whether the comment thread is marked as resolved.
    pub done: bool,
}

/// Output data for a person.
#[derive(Debug, Clone, PartialEq)]
pub struct PersonData {
    /// Person ID (GUID).
    pub id: String,
    /// Display name.
    pub display_name: String,
    /// Optional user ID.
    pub user_id: Option<String>,
    /// Optional provider ID.
    pub provider_id: Option<String>,
}

/// Generate a simple GUID-like identifier wrapped in curly braces.
/// Uses a counter-based approach for deterministic IDs in a session.
fn generate_guid() -> String {
    use std::sync::atomic::{AtomicU64, Ordering};
    static COUNTER: AtomicU64 = AtomicU64::new(1);
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:04X}-{:012X}}}",
        (n >> 32) as u32,
        ((n >> 16) & 0xFFFF) as u16,
        (n & 0xFFFF) as u16,
        ((n >> 48) | 0x4000) as u16,
        n & 0xFFFF_FFFF_FFFF
    )
}

/// Get the current UTC timestamp in ISO 8601 format with two-digit
/// fractional seconds, matching the format Excel uses for threaded comments.
fn current_timestamp() -> String {
    let now = chrono::Utc::now();
    let base = now.format("%Y-%m-%dT%H:%M:%S").to_string();
    let centiseconds = now.timestamp_subsec_millis() / 10;
    format!("{base}.{centiseconds:02}")
}

/// Find or create a person in the person list, returning their ID.
pub fn find_or_create_person(
    person_list: &mut PersonList,
    display_name: &str,
    user_id: Option<&str>,
    provider_id: Option<&str>,
) -> String {
    if let Some(existing) = person_list
        .persons
        .iter()
        .find(|p| p.display_name == display_name)
    {
        return existing.id.clone();
    }

    let id = generate_guid();
    person_list.persons.push(Person {
        display_name: display_name.to_string(),
        id: id.clone(),
        user_id: user_id.map(|s| s.to_string()),
        provider_id: provider_id.map(|s| s.to_string()),
    });
    id
}

/// Add a threaded comment to a sheet's threaded comments collection.
///
/// Returns the generated comment ID.
pub fn add_threaded_comment(
    threaded_comments: &mut Option<ThreadedComments>,
    person_list: &mut PersonList,
    cell: &str,
    input: &ThreadedCommentInput,
) -> Result<String> {
    if let Some(ref parent_id) = input.parent_id {
        let tc = threaded_comments.as_ref();
        let parent_exists = tc
            .map(|t| t.comments.iter().any(|c| c.id == *parent_id))
            .unwrap_or(false);
        if !parent_exists {
            return Err(Error::Internal(format!(
                "parent comment '{}' not found",
                parent_id
            )));
        }
    }

    let person_id = find_or_create_person(person_list, &input.author, None, None);
    let comment_id = generate_guid();

    let tc = threaded_comments.get_or_insert_with(ThreadedComments::default);
    tc.comments.push(ThreadedComment {
        cell_ref: cell.to_string(),
        date_time: current_timestamp(),
        person_id,
        id: comment_id.clone(),
        parent_id: input.parent_id.clone(),
        done: None,
        text: input.text.clone(),
    });

    Ok(comment_id)
}

/// Get all threaded comments for a sheet.
pub fn get_threaded_comments(
    threaded_comments: &Option<ThreadedComments>,
    person_list: &PersonList,
) -> Vec<ThreadedCommentData> {
    let Some(tc) = threaded_comments.as_ref() else {
        return Vec::new();
    };

    tc.comments
        .iter()
        .map(|c| {
            let author = person_list
                .persons
                .iter()
                .find(|p| p.id == c.person_id)
                .map(|p| p.display_name.clone())
                .unwrap_or_default();
            ThreadedCommentData {
                id: c.id.clone(),
                cell_ref: c.cell_ref.clone(),
                text: c.text.clone(),
                author,
                person_id: c.person_id.clone(),
                date_time: c.date_time.clone(),
                parent_id: c.parent_id.clone(),
                done: c.done.as_deref() == Some("1"),
            }
        })
        .collect()
}

/// Get threaded comments for a specific cell.
pub fn get_threaded_comments_by_cell(
    threaded_comments: &Option<ThreadedComments>,
    person_list: &PersonList,
    cell: &str,
) -> Vec<ThreadedCommentData> {
    get_threaded_comments(threaded_comments, person_list)
        .into_iter()
        .filter(|c| c.cell_ref == cell)
        .collect()
}

/// Delete a threaded comment by its ID.
///
/// Returns `true` if the comment was found and removed.
pub fn delete_threaded_comment(
    threaded_comments: &mut Option<ThreadedComments>,
    comment_id: &str,
) -> bool {
    if let Some(ref mut tc) = threaded_comments {
        let before = tc.comments.len();
        tc.comments.retain(|c| c.id != comment_id);
        let removed = tc.comments.len() < before;

        if tc.comments.is_empty() {
            *threaded_comments = None;
        }

        removed
    } else {
        false
    }
}

/// Set the resolved (done) state of a threaded comment.
///
/// Returns `true` if the comment was found and updated.
pub fn resolve_threaded_comment(
    threaded_comments: &mut Option<ThreadedComments>,
    comment_id: &str,
    done: bool,
) -> bool {
    if let Some(ref mut tc) = threaded_comments {
        if let Some(comment) = tc.comments.iter_mut().find(|c| c.id == comment_id) {
            comment.done = if done { Some("1".to_string()) } else { None };
            return true;
        }
    }
    false
}

/// Add a person to the person list. Returns the person ID.
pub fn add_person(person_list: &mut PersonList, input: &PersonInput) -> String {
    find_or_create_person(
        person_list,
        &input.display_name,
        input.user_id.as_deref(),
        input.provider_id.as_deref(),
    )
}

/// Get all persons from the person list.
pub fn get_persons(person_list: &PersonList) -> Vec<PersonData> {
    person_list
        .persons
        .iter()
        .map(|p| PersonData {
            id: p.id.clone(),
            display_name: p.display_name.clone(),
            user_id: p.user_id.clone(),
            provider_id: p.provider_id.clone(),
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_add_threaded_comment() {
        let mut tc = None;
        let mut pl = PersonList::default();
        let input = ThreadedCommentInput {
            author: "Alice".to_string(),
            text: "Hello thread".to_string(),
            parent_id: None,
        };
        let id = add_threaded_comment(&mut tc, &mut pl, "A1", &input).unwrap();
        assert!(!id.is_empty());
        assert!(tc.is_some());

        let comments = get_threaded_comments(&tc, &pl);
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].cell_ref, "A1");
        assert_eq!(comments[0].text, "Hello thread");
        assert_eq!(comments[0].author, "Alice");
        assert!(!comments[0].done);
        assert!(comments[0].parent_id.is_none());
    }

    #[test]
    fn test_add_reply() {
        let mut tc = None;
        let mut pl = PersonList::default();

        let parent_id = add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Initial comment".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let reply_id = add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Bob".to_string(),
                text: "This is a reply".to_string(),
                parent_id: Some(parent_id.clone()),
            },
        )
        .unwrap();

        let comments = get_threaded_comments(&tc, &pl);
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[1].parent_id, Some(parent_id));
        assert_ne!(reply_id, comments[0].id);
    }

    #[test]
    fn test_reply_to_nonexistent_parent() {
        let mut tc = None;
        let mut pl = PersonList::default();
        let input = ThreadedCommentInput {
            author: "Alice".to_string(),
            text: "Bad reply".to_string(),
            parent_id: Some("{NONEXISTENT}".to_string()),
        };
        let result = add_threaded_comment(&mut tc, &mut pl, "A1", &input);
        assert!(result.is_err());
    }

    #[test]
    fn test_get_by_cell() {
        let mut tc = None;
        let mut pl = PersonList::default();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "On A1".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "B2",
            &ThreadedCommentInput {
                author: "Bob".to_string(),
                text: "On B2".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let a1_comments = get_threaded_comments_by_cell(&tc, &pl, "A1");
        assert_eq!(a1_comments.len(), 1);
        assert_eq!(a1_comments[0].text, "On A1");

        let b2_comments = get_threaded_comments_by_cell(&tc, &pl, "B2");
        assert_eq!(b2_comments.len(), 1);
        assert_eq!(b2_comments[0].text, "On B2");

        let empty = get_threaded_comments_by_cell(&tc, &pl, "C3");
        assert!(empty.is_empty());
    }

    #[test]
    fn test_delete_threaded_comment() {
        let mut tc = None;
        let mut pl = PersonList::default();

        let id = add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Delete me".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        assert!(delete_threaded_comment(&mut tc, &id));
        assert!(tc.is_none());
    }

    #[test]
    fn test_delete_nonexistent_comment() {
        let mut tc: Option<ThreadedComments> = None;
        assert!(!delete_threaded_comment(&mut tc, "{NONEXISTENT}"));
    }

    #[test]
    fn test_delete_one_of_multiple() {
        let mut tc = None;
        let mut pl = PersonList::default();

        let id1 = add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "First".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "B2",
            &ThreadedCommentInput {
                author: "Bob".to_string(),
                text: "Second".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        assert!(delete_threaded_comment(&mut tc, &id1));
        assert!(tc.is_some());
        let remaining = get_threaded_comments(&tc, &pl);
        assert_eq!(remaining.len(), 1);
        assert_eq!(remaining[0].text, "Second");
    }

    #[test]
    fn test_resolve_threaded_comment() {
        let mut tc = None;
        let mut pl = PersonList::default();

        let id = add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Resolve me".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        assert!(resolve_threaded_comment(&mut tc, &id, true));
        let comments = get_threaded_comments(&tc, &pl);
        assert!(comments[0].done);

        assert!(resolve_threaded_comment(&mut tc, &id, false));
        let comments = get_threaded_comments(&tc, &pl);
        assert!(!comments[0].done);
    }

    #[test]
    fn test_resolve_nonexistent() {
        let mut tc: Option<ThreadedComments> = None;
        assert!(!resolve_threaded_comment(&mut tc, "{NONEXISTENT}", true));
    }

    #[test]
    fn test_add_person() {
        let mut pl = PersonList::default();
        let id = add_person(
            &mut pl,
            &PersonInput {
                display_name: "Alice".to_string(),
                user_id: Some("alice@example.com".to_string()),
                provider_id: Some("ADAL".to_string()),
            },
        );

        assert!(!id.is_empty());
        let persons = get_persons(&pl);
        assert_eq!(persons.len(), 1);
        assert_eq!(persons[0].display_name, "Alice");
        assert_eq!(persons[0].user_id, Some("alice@example.com".to_string()));
    }

    #[test]
    fn test_add_duplicate_person() {
        let mut pl = PersonList::default();
        let id1 = add_person(
            &mut pl,
            &PersonInput {
                display_name: "Alice".to_string(),
                user_id: None,
                provider_id: None,
            },
        );
        let id2 = add_person(
            &mut pl,
            &PersonInput {
                display_name: "Alice".to_string(),
                user_id: None,
                provider_id: None,
            },
        );

        assert_eq!(id1, id2);
        assert_eq!(get_persons(&pl).len(), 1);
    }

    #[test]
    fn test_get_persons_empty() {
        let pl = PersonList::default();
        assert!(get_persons(&pl).is_empty());
    }

    #[test]
    fn test_get_threaded_comments_empty() {
        let tc: Option<ThreadedComments> = None;
        let pl = PersonList::default();
        assert!(get_threaded_comments(&tc, &pl).is_empty());
    }

    #[test]
    fn test_person_auto_created_on_add_comment() {
        let mut tc = None;
        let mut pl = PersonList::default();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Auto person".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let persons = get_persons(&pl);
        assert_eq!(persons.len(), 1);
        assert_eq!(persons[0].display_name, "Alice");
    }

    #[test]
    fn test_comment_has_timestamp() {
        let mut tc = None;
        let mut pl = PersonList::default();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Timestamp check".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let comments = get_threaded_comments(&tc, &pl);
        assert!(!comments[0].date_time.is_empty());
        assert!(comments[0].date_time.contains('T'));
    }

    #[test]
    fn test_multiple_authors_multiple_cells() {
        let mut tc = None;
        let mut pl = PersonList::default();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "A1",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Alice on A1".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "B2",
            &ThreadedCommentInput {
                author: "Bob".to_string(),
                text: "Bob on B2".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        add_threaded_comment(
            &mut tc,
            &mut pl,
            "C3",
            &ThreadedCommentInput {
                author: "Alice".to_string(),
                text: "Alice on C3".to_string(),
                parent_id: None,
            },
        )
        .unwrap();

        let persons = get_persons(&pl);
        assert_eq!(persons.len(), 2);

        let all = get_threaded_comments(&tc, &pl);
        assert_eq!(all.len(), 3);
    }
}
