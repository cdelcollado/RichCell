# Instal·lació del complement "Editor HTML Productes"

## Requisits previs

- Windows 10 / 11
- Microsoft Excel 365
- [Node.js LTS](https://nodejs.org/en/download) instal·lat

---

## Instal·lació (un sol pas)

Obre PowerShell com a **Administrador** i executa:

```powershell
powershell -ExecutionPolicy Bypass -File "\\NOM-SERVIDOR\HTML_in_Excel\deploy\install.ps1"
```

> Substitueix `NOM-SERVIDOR` pel nom de l'ordinador que té el servidor original.

L'script fa automàticament:
- Copia els fitxers a `C:\HTML_in_Excel`
- Instal·la el certificat SSL de confiança
- Configura el servidor per arrencar automàticament a l'inici de sessió
- Comparteix la carpeta i configura Excel

---

## Després de la instal·lació
0. **Anar al directori on esta instal·lat el programa c:\HTML_in_Excel i executar node server.js **
1. **Tanca Excel** si estava obert
2. **Torna a obrir Excel**
3. **Inserir** → **Complements** → pestanya **CARPETA COMPARTIDA**
4. Selecciona **Editor HTML Productes** → **Afegeix**

---

## Desinstal·lació

```powershell
# Atura i elimina la tasca programada del servidor
Unregister-ScheduledTask -TaskName "HTMLProductEditor_Server" -Confirm:$false

# Elimina la compartició de xarxa
Remove-SmbShare -Name "HTML_in_Excel" -Force

# Elimina els fitxers (opcional)
Remove-Item "C:\HTML_in_Excel" -Recurse -Force
```
