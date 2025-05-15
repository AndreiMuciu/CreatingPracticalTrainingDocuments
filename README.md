# 📄 Internship Agreement Manager — Automated Document Generation in C#

This project streamlines the management and automated generation of **internship agreements** for students. Built in **C#** using **Visual Studio**, it leverages the power of `DocX` for Word file manipulation and `CSVHelper` for structured data handling via CSV files.

---

## 🚀 Features

### 🧠 Intelligent .docx Analysis & Completion

- Supports both **Romanian** and **English** Word templates
- Automatically identifies fields marked with `...` and maps them to appropriate types (`string`, `int`, etc.)
- Handles structured input for:
  - Student details
  - Company information
  - Academic supervisor
  - Company-assigned tutor

### 🔁 Full CRUD on Agreements

- **Load** agreements from a CSV file  
- **Create** new internship agreements via console input  
- **Update** existing entries by ID  
- **Delete** unwanted entries from the collection

### 📋 Display Agreements

- View a list of all agreements showing key information (student name, company name, etc.)

### 📤 Export Agreements to CSV

- Save all data fields into a **structured CSV file** for external storage or sharing

### 📄 Generate Word Documents

- Create `.docx` files for individual agreements (based on ID)
- Batch-generate `.docx` files for all agreements with the format:  
  `"FirstName LastName - CompanyName - Agreement.docx"`

---

## 🛠️ Technologies Used

- **C#** — Core programming language  
- **[DocX](https://github.com/xceedsoftware/DocX)** — Word document generation and editing  
- **[CSVHelper](https://joshclose.github.io/CsvHelper/)** — Easy CSV read/write operations  
- **Visual Studio** — Development environment

---

## 📦 Project Structure

This project follows object-oriented design, with the following main classes:

- `Student` — Contains student-specific data (name, ID, etc.)
- `Company` — Represents the internship provider
- `Tutor` — Stores the company tutor’s information
- `AcademicSupervisor` — University contact person
- `Agreement` — Aggregates all above into a full contract
- `AgreementCollection` — Manages multiple agreements (CRUD support)

---

## ▶️ How to Use

### 1. Setup

- Clone the repository
- Open the solution in **Visual Studio**
- Install dependencies using **NuGet Package Manager**:
  - `DocX`
  - `CSVHelper`

### 2. Run the Application

- Use the **console interface** to:
  - Load agreements from CSV
  - Create new agreements
  - Edit or delete existing ones
  - Export data as CSV
  - Generate Word documents

### 3. Generate .docx Files

- Follow the on-screen prompts to generate:
  - Individual agreements by ID
  - All agreements in batch mode

---

## 📝 Example File Naming

```
John Doe - Softvision - Agreement.docx
Ana Popescu - Continental - Agreement.docx
```

---

## 📂 Example Templates

Ensure your `.docx` templates include fields marked with `...` to be correctly identified and replaced.

---

## 🧠 Why This Project?

Internship documentation often involves repetitive tasks. This tool removes manual work, reduces errors, and ensures all agreement fields are consistently filled using structured data and automation.

---

## 📜 License

This project is intended for academic and educational use. Modify and reuse freely.

---

Looking to extend this project with a GUI or database backend? Feel free to fork it and experiment!
