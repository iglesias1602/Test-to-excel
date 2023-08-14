import openpyxl
import unittest


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
    test_classes = [
        TestProduit,
        TestInventory,
        TestDistributeur,
        TestClient,
        TestTransaction,
    ]
    excel_output_file = "test_details.xlsx"

    write_tests_to_excel(test_classes, excel_output_file)

    print(f"Test details extracted and written to {excel_output_file}")
