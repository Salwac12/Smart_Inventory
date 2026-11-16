# Smart Inventory & BOM Management System

### 🌐 Live Website

🚀 **[CLICK HERE TO OPEN SMART INVENTORY SYSTEM](https://smart-inventory-salwac12.vercel.app/)**

A web-based inventory and Bill of Materials (BOM) management system developed using Python Flask, PostgreSQL, HTML, CSS, and JavaScript.


# Smart Inventory & BOM Management System

A web-based Inventory and Bill of Materials (BOM) Management System developed using Python, Flask, SQLite, HTML, CSS, and JavaScript.

The application is designed to simplify electronic component inventory management and production material planning for electronics manufacturing businesses.

## Features

### 1. Inventory Management
- Register electronic components.
- Maintain component part numbers and descriptions.
- Track stock quantities.
- Store supplier and pricing information.
- Monitor minimum stock levels.
- Identify low-stock and out-of-stock components.

### 2. Stock Management
- Record Stock IN and Stock OUT transactions.
- Automatically update inventory quantities.
- Prevent stock issues exceeding available quantities.
- Maintain transaction history.

### 3. BOM Management
- Upload Excel Bills of Materials.
- Preview uploaded components.
- Save BOMs with product names and revisions.
- View previously saved BOMs.

### 4. BOM Inventory Matching
- Match customer BOM components with internal inventory records.
- Identify unmatched components.
- Add missing components to inventory.

### 5. Material Requirement Planning
- Enter a production quantity.
- Calculate component requirements.
- Compare requirements against available inventory.
- Identify material shortages.
- Estimate purchase requirement costs.

### 6. Excel Reports
Generate downloadable purchase requirement reports with:
- To Order
- Needs Review
- Report Summary

### 7. Dashboard
- Inventory summary.
- Stock alerts.
- Stock movement visualization.
- Recent transaction history.
- Quick navigation between modules.

## Technologies Used

**Backend**
- Python
- Flask

**Database**
- SQLite

**Frontend**
- HTML
- CSS
- JavaScript

**Excel Processing**
- openpyxl

**Development Tools**
- Visual Studio Code
- Git
- GitHub

## Installation

Clone the repository:

```bash
git clone https://github.com/Salwac12/AI_Smart_Inventory.git
```

Navigate to the project directory:

```bash
cd AI_Smart_Inventory
```

Create a virtual environment:

```bash
python -m venv venv
```

Activate the virtual environment on Windows:

```powershell
.\venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the application:

```bash
python app.py
```

Open the dashboard:

```text
http://127.0.0.1:5000/
```

The application initializes its local SQLite database when started.

## Application Workflow

1. Register electronic components in inventory.
2. Record incoming or outgoing stock.
3. Upload and save a customer's BOM.
4. Match BOM components with inventory.
5. Enter the planned production quantity.
6. Calculate material shortages.
7. Download the purchase requirement Excel report.

## Important Notes

- The application is currently designed for local development and demonstration.
- Inventory data is stored in a local SQLite database.
- The database file is excluded from version control.
- Production deployment requires additional configuration for persistent storage and security.

## Future Improvements

Potential enhancements include user authentication, role-based access control, and support for a production database.

## Project Purpose

This project demonstrates practical applications of Python web development, database management, electronic component inventory tracking, and production material planning.
