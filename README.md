# 🚛 Warehouse Management

An integrated **Excel + VBA automation system** for container inventory and GPS tracking.  
This tool helps logistics teams manage real-time warehouse inspection ("Nghiệm Kho") data, vehicle GPS logs, and audit events — all directly within a single Excel workbook.

---

## 📦 Overview

**Warehouse Management** is part of the broader *VanTaiBT* logistics ecosystem.  
It provides a centralized dashboard for recording, auditing, and visualizing container operations through VBA macros, dynamic shapes, and GPS integration.

Key features:
- **Warehouse inspection (“Nghiệm kho”)** tracking for each container/vehicle.
- **GPS log parsing** to map movements and confirm arrival/departure events.
- **Audit Log System** integrated with unique `Machine_ID` and timestamp.
- **Excel VBA automation** for UI navigation, data validation, and report generation.
- **Offline-first architecture**: works fully in Excel, with optional Google Sheet sync.

---

## ⚙️ Technical Architecture

| Layer | Description |
|-------|--------------|
| **Frontend (Excel UI)** | Custom shapes, buttons, icons (e.g., `Shp_Btn_Style01`, `Animate_Shape_Box`) designed for modern dashboard control |
| **Logic Layer (VBA)** | Class-based modules (`CArrayHelper`, `App_State`) for performance and array handling |
| **Data Layer (Sheets)** | Sheets like `Audit_Log`, `NhatKy_Audit`, and `Machine_Registry` storing structured transaction data |
| **External Integration** | Calls `Google Apps Script WebApp` endpoints (`exists_machine`, `Audit_Log`) for online syncing |
| **Database** |Using `Microsoft Assets` |

---

## 📋 Main Modules

| Module | Function |
|---------|-----------|
| `Audit_Log` | Records local events with timestamps and Machine IDs |
| `GPS_Import` | Parses and validates GPS data from external CSV or API logs |
| `Nghiem_Kho_Form` | Provides user interface for warehouse inspection entry |
| `Machine_ID_Check` | Verifies machine registration with Google Apps Script |
| `Optimize_Performance` | Toggles Excel performance features (ScreenUpdating, Events) |
| `CArrayHelper` | Manages dynamic arrays with Option Base 1 for range alignment |

---

## 🔐 Data Security & Traceability

- Each user action is logged into `NhatKy_Audit` with:
  - **UserID / MachineID**
  - **Action type (Insert, Update, Delete)**
  - **Timestamp**
- Optional cloud sync via Apps Script ensures audit consistency across devices.

---

## 🧭 Usage Guide

1. **Enable Macros** when opening the file.  
2. Navigate using the **Home Dashboard**.  
3. Use the **GPS Import** button to load GPS data (`.csv` or API).  
4. Review and confirm “Nghiệm Kho” results.  
5. Logs are automatically written into `Audit_Log`.  
6. Optionally trigger **Sync to Google Sheet** (if connected).


## 🧩 Dependencies

- **Excel 2019 / 365** (with VBA enabled)
- **Windows 10+**
- **Internet connection** (for cloud sync)
- Optional: Google Apps Script endpoints  
  - `/exists_machine`  
  - `/Audit_Log`
