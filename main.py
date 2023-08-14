# generate_excel.py
import openpyxl
import unittest
import sys
import importlib


# Create an Excel file and write test details
def write_tests_to_excel(test_classes, excel_file):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Test Details"

    sheet["A1"] = "Test Class"
    sheet["B1"] = "Test Method"

    for row, test_class in enumerate(test_classes, start=2):
        for method_name in dir(test_class):
            if method_name.startswith("test_"):
                sheet.cell(row=row, column=1, value=test_class.__name__)
                sheet.cell(row=row, column=2, value=method_name)

                row += 1

    wb.save(excel_file)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python generate_excel.py <test_file.py> <output_file.xlsx>")
        sys.exit(1)

    test_file = sys.argv[1]
    output_file = sys.argv[2]

    # Import the test classes from the provided test file
    try:
        module = importlib.import_module(test_file.replace(".py", ""))
    except ImportError:
        print(f"Error: Failed to import test file '{test_file}'")
        sys.exit(1)

    test_classes = [
        cls
        for cls in module.__dict__.values()
        if isinstance(cls, type) and issubclass(cls, unittest.TestCase)
    ]

    if not test_classes:
        print(f"No test classes found in '{test_file}'")
        sys.exit(1)

    write_tests_to_excel(test_classes, output_file)

    print(f"Test details extracted and written to {output_file}")
