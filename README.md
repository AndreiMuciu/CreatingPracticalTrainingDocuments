# CreatingPracticalTrainingDocuments

This project facilitates the management and automated generation of practice agreements for students. It is developed using C#, leveraging the DocX and CSVHelper libraries within Visual Studio. The main objective is to streamline the process of filling, storing, and generating documents related to practice agreements.

Features
Core Functionalities:
1. Analysis and Completion of .docx Agreements

-Supports both Romanian and English templates.
-Automatically identifies fields marked with "..." and maps them to appropriate variable types (string, int, etc.).
-Ensures precise modeling of all required fields, such as student details, company information, academic supervisor, and tutor details.

2. CRUD Operations on Agreements

-Load: Import a list of agreements from a CSV file into memory.
-Create: Add new agreements to the collection with structured data input.
-Update: Edit agreements based on their unique ID.
-Delete: Remove agreements from the collection if needed.

3. Display Agreements

-List all agreements with basic details like the student's name and associated company.

4. Export to CSV

-Save all agreement fields to a structured CSV file for easy data management and sharing.

5. Generate .docx Documents

-Create individual .docx files for specific agreements based on their ID.
-Batch-generate .docx files for all agreements in the format:
"FirstName LastName - CompanyName - Agreement.docx".

Technologies and Libraries Used:
-C#: Core programming language for the project.
-DocX: For generating and modifying Word documents.
-CSVHelper: For handling CSV file imports and exports.
-Visual Studio: Development environment.


Project Structure:
The project uses object-oriented design principles to model real-world entities. Key classes include:

-Student: Stores student-specific data (name, student ID, etc.).
-Company: Represents the company providing the internship.
-Tutor: Stores details of the company’s assigned tutor.
-AcademicSupervisor: Stores details of the academic supervisor.
-Agreement: Aggregates all the above details into a single entity.
-AgreementCollection: Manages a collection of agreements, supporting CRUD operations.

How to Use:
1. Setup:
-Clone the repository and open the solution in Visual Studio.
-Ensure the DocX and CSVHelper packages are installed (via NuGet Package Manager).

2. Run the Application:
-Use the console-based interface to load agreements, create new ones, edit existing entries, and export data.

3. Generate Documents:
-Follow the prompts to generate .docx files for individual or all agreements.
