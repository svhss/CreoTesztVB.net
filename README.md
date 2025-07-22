# Creo API Setup for VB.NET

## 1. Környezeti változó beállítása

Nyomd meg: `Windows + R`
Írd be:

```
rundll32 sysdm.cpl,EditEnvironmentVariable
```

Új rendszer változó hozzáadása:

* **Név**: `pro_comm_msg_Exe`
* **Érték**: például

  ```
  "C:\Program Files\PTC\Creo 9.0\Common Files\x86e_win64\obj\pro_comm_msg.exe"
  ```

## 2. Creo API regisztrálása

Futtasd ezt a batch fájlt rendszergazdaként:

```
"C:\Program Files\PTC\Creo 9.0\Parametric\bin\vb_api_register.bat"
```

Ez regisztrálja a COM-típus könyvtárakat a VB.NET használatához.

## 3. PFCLS hivatkozás hozzáadása VB.NET projektben

Visual Studio-ban:

* Projekt → Jobb klikk → Add → Reference
* Válaszd a **COM** fület
* Keresd meg és válaszd ki:
  **Creo VB API Type Library for Creo Parametric**
  (ProgID: `PFCLS.pfcGlobal`)

Ezután a következő import érvényes:

```vbnet
Imports pfcls
```
