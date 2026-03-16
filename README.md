## Okoskert Cloud Functions

Ez a projekt az Okoskert rendszer Firebase Cloud Functions kódját tartalmazza, elsősorban **projekt riport / Excel export** készítésére Firestore adatokból, valamint a fejlesztői környezet (emulátor) indításához szükséges segédscriptet.

---

### Fő elemek

- **`run-functions.sh`**  
  Segédscript a Firebase Functions emulátor indításához lokálisan (Mac / Linux, `zsh`/`bash` környezetben).

- **`firebase.json`**  
  Firebase projekt konfiguráció (emulátorok, functions beállítások stb.).

- **`functions/main.py`**  
  A Firebase Functions belépési pontja.  
  Jelenleg a legfontosabb HTTP trigger:

  - **`projectExport`** (`@https_fn.on_request()`):
    - `GET`/`POST` kérés `projectId` query paraméterrel.
    - Beolvassa a Firestore-ból:
      - `projects/{projectId}`
      - `worklogs` (collection group, `assignedProjectId == projectId`)
      - `materials` (collection group, `projectId == projectId`)
      - `users` (`teamId == project.teamId`)
      - `machines` (`teamId == project.teamId`)
      - `projects/{projectId}/machineWorklog`
    - Ezekből Excel fájlt generál (`export_excel.build_export_xlsx`).
    - Feltölti a Firebase Storage-be: `exports/{projectId}/projekt_jelentes_YYYYMMDD_HHMMSS.xlsx`
    - Visszaad egy JSON választ:
      - `fileName`
      - `storagePath`
      - opcionálisan `downloadUrl` (aláírt URL, ha elérhető a service account kulcs).

- **`functions/export_excel.py`**  
  A tényleges Excel generálást végzi `pandas` + `openpyxl` segítségével.
  - Projekt metaadatok lap: `Projekt`
  - Anyagfelhasználás: `Alapanyagok`
  - Munkaidő és bér: `Munkadíjak`
  - Munkagépek üzemóra: `Munkagépek`
  - Magyar nyelvű oszlopnevek, formázás, összesítések.

- **`functions/requirements.txt`**  
  A Cloud Function Python függőségei:
  - `firebase_functions`
  - `firebase-admin`
  - `pandas`
  - `openpyxl`

---

## Fejlesztői környezet beállítása

### Előfeltételek

- **Node.js** (Firebase CLI-hez)
- **Python 3.11+** (a Cloud Functions Python runtime-hoz)
- **Firebase CLI** (`firebase-tools`)  
  Telepítés:
  ```bash
  npm install -g firebase-tools