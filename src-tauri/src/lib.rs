//! Enveloppe native de Grillon : une fenêtre qui affiche l'appli web (dossier `dist/`).
//! L'ouverture / l'enregistrement des fichiers passent par les plugins dialog et fs (côté JS).
//! Ici : seulement l'ouverture d'un fichier par double-clic dans le Finder.

use std::sync::Mutex;
use tauri::{Emitter, Manager};

/// Fichiers demandés par le système avant que l'interface soit prête.
#[derive(Default)]
struct PendingOpens(Mutex<Vec<String>>);

/// Appelée par l'interface au démarrage : renvoie (et vide) la liste des fichiers à ouvrir.
#[tauri::command]
fn take_pending_opens(state: tauri::State<PendingOpens>) -> Vec<String> {
    std::mem::take(&mut *state.0.lock().unwrap())
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .manage(PendingOpens::default())
        .invoke_handler(tauri::generate_handler![take_pending_opens])
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
                if !paths.is_empty() {
                    app.state::<PendingOpens>().0.lock().unwrap().extend(paths.clone());
                    let _ = app.emit("grillon://open", paths);
                }
            }
            #[cfg(not(target_os = "macos"))]
            let _ = (app, event);
        });
}
