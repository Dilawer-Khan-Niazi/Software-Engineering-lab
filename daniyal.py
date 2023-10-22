def search_books():
    create_excel_file()
    workbook = openpyxl.load_workbook('library_data.xlsx')
    books_sheet = workbook['Books']

    while True:
        search_query = input("Enter a book title or author (or type 'q' to quit): ").strip()

        if search_query.lower() == 'q':
            break

        found_books = []

        for row in books_sheet.iter_rows(min_row=2, values_only=True):
            title, _, author, _ = row  # Adjust this line to match the row structure
            if search_query.lower() in title.lower() or search_query.lower() in author.lower():
                found_books.append((title, author))

        if found_books:
            print("Matching Books:")
            for title, author in found_books:
                print(f"Title: {title}, Author: {author}")
        else:
            print("No matching books found.")

    workbook.close()

def display_catalog():
    create_excel_file()
    workbook = openpyxl.load_workbook('library_data.xlsx')
    books_sheet = workbook['Books']

    print("Catalog of Books:")
    for row in books_sheet.iter_rows(min_row=2, values_only=True):
        title, _ , author, _  = row  
        print(f"Title: {title}, Author: {author}")

    workbook.close()