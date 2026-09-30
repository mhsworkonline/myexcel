//! Test channel (`--features test-build`): the UI is read from `ui\` next to the exe instead of
//! being compiled in, so a UI fix only needs the files swapped and Help > Reload. The origin,
//! CSP and capabilities are the same as the release build.
use sha2::{Digest, Sha256};
use std::borrow::Cow;
use std::collections::HashMap;
use std::path::{Component, Path, PathBuf};
use std::sync::Mutex;
use tauri::utils::assets::{AssetKey, AssetsIter, CspHash};

pub struct DiskAssets {
    root: PathBuf,
    /// Inline-script hashes per HTML content; leaked so `csp_hashes` can hand out borrows.
    /// Bounded by the number of distinct UI builds loaded in one session.
    hashes: Mutex<HashMap<String, &'static [String]>>,
}

impl DiskAssets {
    pub fn new() -> Self {
        let root = std::env::current_exe()
            .ok()
            .and_then(|p| p.parent().map(|d| d.join("ui")))
            .unwrap_or_else(|| PathBuf::from("ui"));
        Self { root, hashes: Mutex::new(HashMap::new()) }
    }

    fn path(&self, key: &AssetKey) -> Option<PathBuf> {
        let rel = Path::new(key.as_ref().trim_start_matches('/'));
        if rel.components().any(|c| !matches!(c, Component::Normal(_))) {
            return None;
        }
        Some(self.root.join(rel))
    }
}

/// Hashes of `<script>` elements without `src`, as tauri-codegen computes them for embedded HTML.
fn inline_script_hashes(html: &str) -> Vec<String> {
    use base64::Engine;
    let mut out = Vec::new();
    let mut rest = html;
    while let Some(i) = rest.find("<script") {
        rest = &rest[i + 7..];
        let Some(tag_end) = rest.find('>') else { break };
        let attrs = &rest[..tag_end];
        rest = &rest[tag_end + 1..];
        let Some(close) = rest.find("</script>") else { break };
        let body = &rest[..close];
        rest = &rest[close..];
        if attrs.contains("src=") || body.is_empty() {
            continue;
        }
        let normalized = body.replace("\r\n", "\n").replace('\r', "\n");
        let digest = Sha256::digest(normalized.as_bytes());
        out.push(format!("'sha256-{}'", base64::engine::general_purpose::STANDARD.encode(digest)));
    }
    out
}

impl<R: tauri::Runtime> tauri::Assets<R> for DiskAssets {
    fn get(&self, key: &AssetKey) -> Option<Cow<'_, [u8]>> {
        std::fs::read(self.path(key)?).ok().map(Cow::Owned)
    }

    fn iter(&self) -> Box<AssetsIter<'_>> {
        Box::new(std::iter::empty())
    }

    fn csp_hashes(&self, html_path: &AssetKey) -> Box<dyn Iterator<Item = CspHash<'_>> + '_> {
        let Some(html) = self.path(html_path).and_then(|p| std::fs::read_to_string(p).ok()) else {
            return Box::new(std::iter::empty());
        };
        let mut cache = self.hashes.lock().unwrap();
        let list: &'static [String] = cache
            .entry(html.clone())
            .or_insert_with(|| Box::leak(inline_script_hashes(&html).into_boxed_slice()));
        Box::new(list.iter().map(|h| CspHash::Script(h.as_str())))
    }
}
