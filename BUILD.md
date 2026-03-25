# 🔧 Build Diagnostic PC

Ce document décrit les étapes pour générer l’exécutable et l’installateur du projet Diagnostic PC.

---

## 📦 Prérequis

- PowerShell 7 (pwsh)
- Module ps2exe
- Inno Setup

---

## 🛠️ 1. Build de l’exécutable

Depuis le dossier du projet :

pwsh -ExecutionPolicy Bypass -File build.ps1

Résultat :
- Génère debug_pc.exe dans le dossier courant

---

## 📦 2. Build de l’installateur

1. Ouvrir Inno Setup
2. Charger le fichier installer.iss
3. Cliquer sur Compile

Résultat :
- Génère Setup_Diagnostic_PC.exe (souvent dans Output/)

---

## 🔢 3. Gestion de version

Avant chaque build, mettre à jour :

Dans le script PowerShell :
$AppVersion = "1.0.0"

Dans installer.iss :
AppVersion=1.0.0
OutputBaseFilename=Setup_Diagnostic_PC_1.0.0

Format recommandé :
1.0.0
1.0.1
1.1.0

---

## 🚀 4. Publication sur GitHub

1. Aller dans Releases
2. Cliquer sur Create a new release
3. Tag : v1.0.0
4. Titre : Diagnostic PC v1.0.0
5. Ajouter le fichier Setup_Diagnostic_PC_1.0.0.exe
6. Cliquer sur Publish release

---

## 🔍 5. Vérification

Tester l’API GitHub :

Invoke-RestMethod https://api.github.com/repos/Pinkywhisky/diagnostic-pc/releases/latest

Doit retourner un JSON avec tag_name, html_url, etc.

---

Tester dans l’application :

- Lancer l’outil
- Cliquer sur Mise à jour

Résultat attendu :
- "Vous êtes à jour"
- ou "Nouvelle version disponible"

---

## ⚠️ Notes importantes

- Le repository doit être public pour que la vérification fonctionne
- Une release en draft ne sera pas détectée
- Toujours incrémenter la version avant build

---

## 💡 Bonnes pratiques

- Ne pas versionner les .exe dans Git
- Utiliser les Releases GitHub pour la distribution
- Garder une cohérence entre :
  - $AppVersion
  - le tag GitHub (vX.X.X)
  - le nom du setup

---

## 📁 Structure recommandée

Projet_debug_pc/
├── debug_pc.ps1
├── build.ps1
├── build.bat
├── installer.iss
├── README.md
├── BUILD.md
└── .gitignore

---

## 🧠 Résumé

1. Mettre à jour la version  
2. Build .exe  
3. Build setup  
4. Publier sur GitHub  
5. Tester