# MSList-to-Excel

This project addresses a limitation when exporting data from Microsoft Lists. By default, Microsoft Lists provides an export in CSV format, but when the the file surpasses the 5000 registers, a direct conversion to Excel (.xlsx) is not possible due to formatting inconsistencies. The goal of this project is to build a lightweight and reliable tool that processes the exported CSV file and transforms it into a clean, usable Excel file.

The solution will:

-  Parse CSV exports from Microsoft Lists, handling irregular delimiters, quotes, and UTF-8 encoding issues.

-  Normalize headers and values, ensuring compatibility with Excel’s stricter schema.

-  Convert CSV to Excel (.xlsx) using libraries such as Pandas or openpyxl, maintaining data integrity and allowing further enhancements such as styling, column widths, or summary sheets.

-  Provide a basic graphic interface where the user can drop a CSV file and receive a ready-to-use Excel file in return.

This tool ensures that organizations relying on Microsoft Lists can continue analyzing, sharing, and integrating data seamlessly in Excel, even when direct conversions fail.

