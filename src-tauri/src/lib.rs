use std::sync::Mutex;
use tauri::menu::{Menu, MenuItem, PredefinedMenuItem, Submenu};
use tauri::{Emitter, Manager};
use tauri_plugin_fs::FsExt;

/// File passed on the command line (file association / "Open with"), handed to the UI once.
struct StartupFile(Mutex<Option<String>>);

const SUPPORTED: [&str; 7] = [".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt", ".ods"];

#[tauri::command]
fn startup_file(state: tauri::State<StartupFile>) -> Option<String> {
    state.0.lock().unwrap().take()
}

fn file_arg(args: &[String]) -> Option<String> {
    args.iter()
        .skip(1)
        .find(|a| {
            let l = a.to_lowercase();
            SUPPORTED.iter().any(|e| l.ends_with(e)) && std::path::Path::new(a).is_file()
        })
        .cloned()
}

#[derive(serde::Serialize, Default)]
struct ClipboardData {
    html: Option<String>,
    text: Option<String>,
}

/// Read the system clipboard: "HTML Format" (what Excel puts there) plus Unicode text.
#[tauri::command]
fn clipboard_read() -> Result<ClipboardData, String> {
    #[cfg(windows)]
    {
        use clipboard_win::{formats, get_clipboard};
        let html = formats::Html::new().and_then(|f| get_clipboard::<String, _>(f).ok());
        let text = get_clipboard::<String, _>(formats::Unicode).ok();
        Ok(ClipboardData { html, text })
    }
    #[cfg(not(windows))]
    {
        Err("native clipboard is only implemented on Windows".into())
    }
}

/// Put HTML (CF_HTML) and plain text on the clipboard in one operation so Excel can paste formatting.
#[tauri::command]
fn clipboard_write(html: String, text: String) -> Result<(), String> {
    #[cfg(windows)]
    {
        use clipboard_win::{formats, options::NoClear, raw, Clipboard};
        let _guard = Clipboard::new_attempts(10).map_err(|e| e.to_string())?;
        raw::empty().map_err(|e| e.to_string())?;
        if let Some(f) = formats::Html::new() {
            raw::set_html(f.code(), &html).map_err(|e| e.to_string())?;
        }
        raw::set_string_with(&text, NoClear).map_err(|e| e.to_string())?;
        Ok(())
    }
    #[cfg(not(windows))]
    {
        let _ = (html, text);
        Err("native clipboard is only implemented on Windows".into())
    }
}

/// Menu labels carry their shortcut after a tab (displayed right-aligned by Windows) instead of a
/// real accelerator: the web UI already handles every shortcut, so binding them natively would
/// fire commands twice or steal keys from text fields.
fn build_menu(app: &tauri::AppHandle) -> tauri::Result<Menu<tauri::Wry>> {
    let item = |id: &str, label: &str| MenuItem::with_id(app, id, label, true, None::<&str>);
    let sep = || PredefinedMenuItem::separator(app);
    let file = Submenu::with_items(
        app,
        "&File",
        true,
        &[
            &item("new", "&New\tCtrl+N")?,
            &item("open", "&Open...\tCtrl+O")?,
            &item("recent", "Open &Recent...")?,
            &sep()?,
            &item("save", "&Save\tCtrl+S")?,
            &item("saveAs", "Save &As...\tF12")?,
            &item("exportPdf", "&Export as PDF...")?,
            &sep()?,
            &item("pageSetup", "Page Set&up...")?,
            &item("print", "&Print...\tCtrl+P")?,
            &sep()?,
            &item("options", "Op&tions...")?,
            &sep()?,
            &item("exit", "E&xit\tAlt+F4")?,
        ],
    )?;
    let edit = Submenu::with_items(
        app,
        "&Edit",
        true,
        &[
            &item("undo", "&Undo\tCtrl+Z")?,
            &item("redo", "&Redo\tCtrl+Y")?,
            &sep()?,
            &item("cut", "Cu&t\tCtrl+X")?,
            &item("copy", "&Copy\tCtrl+C")?,
            &item("paste", "&Paste\tCtrl+V")?,
            &item("pasteSpecial", "Paste &Special...\tCtrl+Alt+V")?,
            &sep()?,
            &item("find", "&Find...\tCtrl+F")?,
            &item("replace", "R&eplace...\tCtrl+H")?,
            &item("goto", "&Go To...\tCtrl+G")?,
            &sep()?,
            &item("selectAll", "Select &All\tCtrl+A")?,
        ],
    )?;
    let view = Submenu::with_items(
        app,
        "&View",
        true,
        &[
            &item("toggleTheme", "&Dark Mode")?,
            &item("toggleFormulaBar", "&Formula Bar")?,
            &item("toggleGridlines", "&Gridlines")?,
            &sep()?,
            &item("zoomIn", "Zoom &In\tCtrl+Wheel Up")?,
            &item("zoomOut", "Zoom &Out\tCtrl+Wheel Down")?,
            &item("zoomReset", "&Zoom 100%")?,
            &sep()?,
            &item("freeze", "Freeze &Panes")?,
        ],
    )?;
    let help = Submenu::with_items(
        app,
        "&Help",
        true,
        &[&item("shortcuts", "&Keyboard Shortcuts")?, &sep()?, &item("about", "&About MyExcel")?],
    )?;
    Menu::with_items(app, &[&file, &edit, &view, &help])
}

/// Turn off WebView2's own browser shortcuts (Ctrl+F find bar, Ctrl+P print, F5/Ctrl+R reload,
/// Ctrl+Shift+I devtools, ...) so spreadsheet shortcuts reach the page. Editing keys such as
/// copy/paste/undo in text fields are unaffected.
fn disable_browser_accelerators(window: &tauri::WebviewWindow) {
    #[cfg(windows)]
    let _ = window.with_webview(|wv| unsafe {
        use webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Settings3;
        use windows::core::Interface;
        if let Ok(core) = wv.controller().CoreWebView2() {
            if let Ok(settings) = core.Settings() {
                if let Ok(s3) = settings.cast::<ICoreWebView2Settings3>() {
                    let _ = s3.SetAreBrowserAcceleratorKeysEnabled(false);
                }
                let _ = settings.SetAreDefaultContextMenusEnabled(false);
                let _ = settings.SetIsStatusBarEnabled(false);
            }
        }
    });
    #[cfg(not(windows))]
    let _ = window;
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let args: Vec<String> = std::env::args().collect();
    let startup = file_arg(&args);
    tauri::Builder::default()
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_persisted_scope::init())
        .plugin(tauri_plugin_dialog::init())
        .manage(StartupFile(Mutex::new(startup.clone())))
        .invoke_handler(tauri::generate_handler![startup_file, clipboard_read, clipboard_write])
        .setup(move |app| {
            // A file opened via association is outside any user-granted scope: allow exactly that file.
            if let Some(p) = &startup {
                let _ = app.fs_scope().allow_file(p);
            }
            let menu = build_menu(app.handle())?;
            app.set_menu(menu)?;
            app.on_menu_event(|app, event| {
                // Forward menu commands to the web UI, which owns all command logic.
                let _ = app.emit("menu", event.id().0.clone());
            });
            if let Some(w) = app.get_webview_window("main") {
                disable_browser_accelerators(&w);
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running MyExcel");
}
