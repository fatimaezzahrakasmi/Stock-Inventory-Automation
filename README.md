# 📦 Stock Management & Automated Invoicing (VBA)

A professional automation solution built with Excel and VBA to manage the full lifecycle of retail operations, including inventory tracking, client databases, and automated PDF invoicing.

---

## 🚀 Main Features

### 🔄 Dynamic Stock Synchronization
* [cite_start]**Real-time Updates**: Automatically calculates current stock using: `Initial Stock + Entries - Exits`. [cite: 17]
* **Visual Status Indicators**: Uses VBA to color-code inventory health:
  * [cite_start]**Red**: Rupture de stock (Stock = 0). [cite: 17]
  * [cite_start]**Yellow**: Faible stock (Stock ≤ 10). [cite: 17]
  * [cite_start]**Green**: En stock (Stock > 10). [cite: 17]
* [cite_start]**Automatic Upsert**: If a product reference is missing from the master list, the system automatically appends it. [cite: 17]

### 📄 Automated Billing Pipeline
* [cite_start]**Unique ID Generation**: Creates standardized invoice numbers (e.g., `F-2024-09-0001`). [cite: 17]
* [cite_start]**PDF Automation**: Exports the `facture` sheet to PDF and saves it to a specific client folder. [cite: 17]
* [cite_start]**OS Integration**: Automatically creates local directories (`MkDir`) for new clients on the computer. [cite: 17]
* [cite_start]**Smart Hyperlinking**: Adds a direct link to the client's invoice folder inside the Excel database. [cite: 17]

### 👥 Integrated CRM
* [cite_start]**Client Registration**: Captures names, addresses, and contact info via a custom UI form. [cite: 17]
* [cite_start]**Auto-Indexing**: Assigns unique IDs to every new customer automatically. [cite: 17]

---

## 🏗️ Technical Architecture

### **Worksheet Structure**
| Sheet | Purpose |
| :--- | :--- |
| **vente** | [cite_start]Transaction input and sales log. [cite: 17] |
| **stock** | [cite_start]Master inventory with visual health alerts. [cite: 17] |
| **client** | [cite_start]Customer database and document links. [cite: 17] |
| **details** | [cite_start]Transaction history for stock movements. [cite: 17] |
| **facture** | [cite_start]Automated invoice template. [cite: 17] |

### **VBA Module Breakdown**
* [cite_start]**Module 1**: `AddToTable` - Appends sales data to the structured table. [cite: 17]
* [cite_start]**Module 2**: Navigation - Handles switching between the 5 main sheets. [cite: 17]
* [cite_start]**Module 3**: `transfertData` - The core engine for stock math and coloring. [cite: 17]
* [cite_start]**Module 4**: `AjouteClient` - Manages the CRM and ID generation. [cite: 17]
* [cite_start]**Module 5**: `GenererSauvgarderFacture` - Manages PDF export and file system I/O. [cite: 17]

---

## 🛠️ Setup Instructions

1. **Enable Macros**: Click "Enable Content" when opening the `.xlsm` file.
2. **Directory Config**: Ensure the path `C:\Users\pc\Desktop\Factures\` exists or update the path in **Module 5**.
3. **Usage Flow**:
   * Add a client in the **Client** sheet.
   * Record a sale in the **Vente** sheet.
   * Run the **Update Stock** script.
   * Generate the PDF from the **Facture** sheet.

---
