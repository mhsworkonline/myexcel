use std::sync::Mutex;
use tauri::menu::{Menu, MenuItem, PredefinedMenuItem, Submenu};
use tauri::{Emitter, Manager};

/// File passed on the command line (file association / "Open with").
struct StartupFile(Mutex<Option<String>>);

#[tauri::command]
fn startup_file(state: tauri::State<StartupFile>) -> Option<String> {
    state.0.lock().unwrap().take()
}

fn file_arg(args: &[String]) -> Option<String> {
    args.iter()
        .skip(1)
        .find(|a| {
            let l = a.to_lowercase();
            [".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".ods"].iter().any(|e| l.ends_with(e))
        })
        .cloned()
}

fn build_menu(app: &tauri::AppHandle) -> tauri::Result<Menu<tauri::Wry>> {
    let item = |id: &str, label: &str, accel: Option<&str>| MenuItem::with_id(app, id, label, true, accel);
    let file = Submenu::with_items(
        app,
        "File",
        true,
        &[
            &item("new", "New", Some("CmdOrCtrl+N"))?,
            &item("open", "Open...", Some("CmdOrCtrl+O"))?,
            &PredefinedMenuItem::separator(app)?,
            &item("save", "Save", Some("CmdOrCtrl+S"))?,
            &item("saveAs", "Save As...", Some("F12"))?,
            &item("exportPdf", "Export as PDF...", None)?,
            &PredefinedMenuItem::separator(app)?,
            &item("print", "Print...", Some("CmdOrCtrl+P"))?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::quit(app, Some("Exit"))?,
        ],
    )?;
    let edit = Submenu::with_items(
        app,
        "Edit",
        true,
        &[
            &item("undo", "Undo", None)?,
            &item("redo", "Redo", None)?,
            &PredefinedMenuItem::separator(app)?,
            &PredefinedMenuItem::cut(app, None)?,
            &PredefinedMenuItem::copy(app, None)?,
            &PredefinedMenuItem::paste(app, None)?,
            &PredefinedMenuItem::separator(app)?,
            &item("find", "Find...", None)?,
            &item("replace", "Replace...", None)?,
            &item("goto", "Go To...", None)?,
        ],
    )?;
    let view = Submenu::with_items(
        app,
        "View",
        true,
        &[
            &item("toggleTheme", "Toggle Dark Mode", None)?,
            &item("zoomIn", "Zoom In", None)?,
            &item("zoomOut", "Zoom Out", None)?,
            &item("zoomReset", "Zoom 100%", None)?,
        ],
    )?;
    let help = Submenu::with_items(app, "Help", true, &[&item("about", "About MyExcel", None)?])?;
    Menu::with_items(app, &[&file, &edit, &view, &help])
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let args: Vec<String> = std::env::args().collect();
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .manage(StartupFile(Mutex::new(file_arg(&args))))
        .invoke_handler(tauri::generate_handler![startup_file])
        .setup(|app| {
            let menu = build_menu(app.handle())?;
            app.set_menu(menu)?;
            app.on_menu_event(|app, event| {
                // Forward menu commands to the web UI.
                let _ = app.emit("menu", event.id().0.clone());
            });
            if let Some(w) = app.get_webview_window("main") {
                let _ = w.set_title("MyExcel");
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running MyExcel");
}
