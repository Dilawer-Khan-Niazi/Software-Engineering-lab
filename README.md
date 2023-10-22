# Software-Engineering-lab
Imports and Excel File Handling:

The code starts by importing necessary modules, including os, openpyxl for working with Excel files, and datetime for handling dates.
It defines a function create_excel_file() to create an Excel file named "library_data.xlsx" if it doesn't exist. This file serves as the data store for the library system.
User Registration and Login:

The program allows users to register with a username, password, and user type (e.g., Librarian or Member).
User information is stored in the "Users" sheet of the Excel file.
A login function allows users to log in by checking the provided username and password against the stored user data.
Password Reset:

Users can reset their passwords by providing their username and a new password.
Book Catalog Management:

The program enables the addition and deletion of books in the library's catalog.
Book data is stored in the "Books" sheet of the Excel file.
Issuing and Returning Books:

Members can issue and return books. When a book is issued, its status is updated to "Not Available," and the issuance details are recorded in the "Issuance Tracking" sheet.
When a book is returned, the return date is updated in the issuance record, and the book's status is changed to "Available."
Searching for Books:

Users can search for books by title or author. The program prompts the user to enter a search query and displays matching books from the catalog.
Displaying the Book Catalog:

Users can view the entire catalog of books, including their titles and authors.
Main Program Loop:

The code runs in a loop that presents a menu of options to the user.
Users can choose from a range of options, including registration, login, password reset, catalog management, issuing and returning books, searching, and displaying the catalog.
Exit:

Users have the option to quit the program.
