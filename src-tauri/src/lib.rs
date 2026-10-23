//! Enveloppe native de Grillon : des fenêtres qui affichent l'appli web (dossier `dist/`).
//! L'ouverture / l'enregistrement des fichiers passent par les plugins dialog et fs (côté JS).
//! Ici : plusieurs fenêtres (un classeur par fenêtre) et l'ouverture d'un fichier par double-clic
//! dans le Finder.

use std::collections::HashMap;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Mutex;
use tauri::menu::{AboutMetadata, Menu, PredefinedMenuItem, Submenu};
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};

#[derive(Default)]
struct Opens {
    /// Vrai une fois que la première fenêtre a récupéré les fichiers reçus au démarrage.
    ready: Mutex<bool>,
    /// Fichiers demandés par le système avant que l'interface soit prête.
    pending: Mutex<Vec<String>>,
    /// Fichier à ouvrir dans une fenêtre créée par `spawn_window` (étiquette → chemin).
    windows: Mutex<HashMap<String, String>>,
    counter: AtomicUsize,
}

/// Crée une fenêtre de plus, avec éventuellement un fichier à y ouvrir.
fn spawn_window(app: &AppHandle, path: Option<String>) -> Result<(), String> {
    let state = app.state::<Opens>();
    let n = state.counter.fetch_add(1, Ordering::SeqCst) + 1;
    let label = format!("doc-{n}");
    if let Some(p) = path {
        state.windows.lock().unwrap().insert(label.clone(), p);
    }
    WebviewWindowBuilder::new(app, &label, WebviewUrl::App(format!("index.html?w={label}").into()))
        .title("Grillon")
        .inner_size(1360.0, 860.0)
        .min_inner_size(720.0, 480.0)
        .disable_drag_drop_handler()
        .build()
        .map(|_| ())
        .map_err(|e| e.to_string())
}

/// Appelée par la première fenêtre au démarrage : renvoie (et vide) la liste des fichiers à ouvrir.
#[tauri::command]
fn take_pending_opens(state: tauri::State<Opens>) -> Vec<String> {
    *state.ready.lock().unwrap() = true;
    std::mem::take(&mut *state.pending.lock().unwrap())
}

/// Appelée par une fenêtre créée par `new_window` : le fichier qu'elle doit ouvrir, s'il y en a un.
#[tauri::command]
fn take_window_open(label: String, state: tauri::State<Opens>) -> Option<String> {
    state.windows.lock().unwrap().remove(&label)
}

/// Fichier > Nouveau / Ouvrir : une fenêtre de plus (asynchrone : la création d'une fenêtre
/// ne doit pas bloquer le fil des commandes).
#[tauri::command]
async fn new_window(app: AppHandle, path: Option<String>) -> Result<(), String> {
    spawn_window(&app, path)
}

/// Barre des menus du Mac en français (celle de Tauri par défaut est en anglais). Les commandes d'édition
/// restent les éléments du système : ⌘Z, ⌘C, ⌘V… arrivent ainsi à la page comme avant.
fn french_menu<R: tauri::Runtime>(app: &AppHandle<R>) -> tauri::Result<Menu<R>> {
    let about = AboutMetadata { name: Some("Grillon".into()), ..Default::default() };
    let sep = || PredefinedMenuItem::separator(app);
    Menu::with_items(
        app,
        &[
            &Submenu::with_items(
                app,
                "Grillon",
                true,
                &[
                    &PredefinedMenuItem::about(app, Some("À propos de Grillon"), Some(about))?,
                    &sep()?,
                    &PredefinedMenuItem::services(app, Some("Services"))?,
                    &sep()?,
                    &PredefinedMenuItem::hide(app, Some("Masquer Grillon"))?,
                    &PredefinedMenuItem::hide_others(app, Some("Masquer les autres"))?,
                    &PredefinedMenuItem::show_all(app, Some("Tout afficher"))?,
                    &sep()?,
                    &PredefinedMenuItem::quit(app, Some("Quitter Grillon"))?,
                ],
            )?,
            &Submenu::with_items(
                app,
                "Fichier",
                true,
                &[&PredefinedMenuItem::close_window(app, Some("Fermer la fenêtre"))?],
            )?,
            &Submenu::with_items(
                app,
                "Édition",
                true,
                &[
                    &PredefinedMenuItem::undo(app, Some("Annuler"))?,
                    &PredefinedMenuItem::redo(app, Some("Rétablir"))?,
                    &sep()?,
                    &PredefinedMenuItem::cut(app, Some("Couper"))?,
                    &PredefinedMenuItem::copy(app, Some("Copier"))?,
                    &PredefinedMenuItem::paste(app, Some("Coller"))?,
                    &PredefinedMenuItem::select_all(app, Some("Tout sélectionner"))?,
                ],
            )?,
            &Submenu::with_items(
                app,
                "Présentation",
                true,
                &[&PredefinedMenuItem::fullscreen(app, Some("Plein écran"))?],
            )?,
            &Submenu::with_items(
                app,
                "Fenêtre",
                true,
                &[
                    &PredefinedMenuItem::minimize(app, Some("Réduire"))?,
                    &PredefinedMenuItem::maximize(app, Some("Agrandir"))?,
                ],
            )?,
        ],
    )
}

pub fn run() {
    tauri::Builder::default()
        .menu(|app| french_menu(app))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        // Mises à jour automatiques (Releases GitHub) : vérification et installation depuis l'interface.
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .manage(Opens::default())
        .invoke_handler(tauri::generate_handler![take_pending_opens, take_window_open, new_window])
        .build(tauri::generate_context!())
        .expect("erreur au démarrage de Grillon")
        .run(|app, event| {
            // macOS : double-clic sur un .xlsx / .csv associé à Grillon.
            #[cfg(target_os = "macos")]
            if let tauri::RunEvent::Opened { urls } = &event {
                let paths: Vec<String> = urls
                    .iter()
                    .filter_map(|u| u.to_file_path().ok())
                    .map(|p| p.to_string_lossy().into_owned())
                    .collect();
                let state = app.state::<Opens>();
                let ready = *state.ready.lock().unwrap();
                if ready {
                    // L'appli tourne déjà : chaque fichier s'ouvre dans sa propre fenêtre.
                    for p in paths {
                        let _ = spawn_window(app, Some(p));
                    }
                } else if !paths.is_empty() {
                    state.pending.lock().unwrap().extend(paths);
                }
            }
            #[cfg(not(target_os = "macos"))]
            let _ = (app, event);
        });
}
