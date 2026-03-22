# 📦 Stock Management & Automated Invoicing (VBA)

A professional automation solution built with Excel and VBA to manage the full lifecycle of retail operations, including inventory tracking, client databases, and automated PDF invoicing.

---

## 🚀 Main Features

### 🔄 Dynamic Stock Synchronization
* **Real-time Updates**: Automatically calculates current stock using: `Initial Stock + Entries - Exits`. 
* **Visual Status Indicators**: Uses VBA to color-code inventory health:
* **Red**: Rupture de stock (Stock = 0).
* **Yellow**: Faible stock (Stock ≤ 10). 
* **Green**: En stock (Stock > 10). 
* **Automatic Upsert**: If a product reference is missing from the master list, the system automatically appends it.

### 📄 Automated Billing Pipeline
* **Unique ID Generation**: Creates standardized invoice numbers (e.g., `F-2024-09-0001`). 
* **PDF Automation**: Exports the `facture` sheet to PDF and saves it to a specific client folder. 
* **OS Integration**: Automatically creates local directories (`MkDir`) for new clients on the computer. 
* **Smart Hyperlinking**: Adds a direct link to the client's invoice folder inside the Excel database. 

### 👥 Integrated CRM
* **Client Registration**: Captures names, addresses, and contact info via a custom UI form. 
* **Auto-Indexing**: Assigns unique IDs to every new customer automatically. 

---

## 🏗️ Technical Architecture

### **Worksheet Structure**
| Sheet | Purpose |
| :--- | :--- |
| **vente** | Transaction input and sales log.  |
| **stock** | Master inventory with visual health alerts.  |
| **client** | Customer database and document links.  |
| **details** | Transaction history for stock movements.  |
| **facture** | Automated invoice template.  |

### **VBA Module Breakdown**
* **Module 1**: `AddToTable` - Appends sales data to the structured table. 
* **Module 2**: Navigation - Handles switching between the 5 main sheets. 
* **Module 3**: `transfertData` - The core engine for stock math and coloring. 
* **Module 4**: `AjouteClient` - Manages the CRM and ID generation. 
* **Module 5**: `GenererSauvgarderFacture` - Manages PDF export and file system I/O. 

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
