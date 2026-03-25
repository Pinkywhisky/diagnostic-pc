# 🔧 Build Diagnostic PC

## Prérequis

- PowerShell 7 (pwsh)
- Module ps2exe
- Inno Setup

---

## 1. Build de l'exe

```powershell
pwsh -ExecutionPolicy Bypass -File build.ps1

---

## 2. Build du setup

1. Ouvrir Inno Setup
2. Charger installer.iss
3. Cliquer sur Compile

## 3. Version

Mettre à jour :

* $AppVersion dans debug_pc.ps1
* AppVersion dans installer.iss

## 4. Release GitHub

Aller dans Releases
Create new release
Tag : vX.X
Upload Setup_Diagnostic_PC.exe