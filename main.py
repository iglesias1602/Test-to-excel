import openpyxl
import unittest
import sys
import importlib
import io
import contextlib


# Create an Excel file and write test results
def write_test_results_to_excel(test_results, excel_file):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Test Results"

    sheet["A1"] = "Test Class"
    sheet["B1"] = "Test Method"
    sheet["C1"] = "Input"
    sheet["D1"] = "Result"

    for row, (test_class, test_method, input_data, result) in enumerate(
        test_results, start=2
    ):
        sheet.cell(row=row, column=1, value=test_class)
        sheet.cell(row=row, column=2, value=test_method)
        sheet.cell(row=row, column=3, value=input_data)
        sheet.cell(row=row, column=4, value=result)

    wb.save(excel_file)


# Custom TestResult class to capture test results
class CustomTestResult(unittest.TextTestResult):
    def __init__(self, stream=sys.stdout, descriptions=1, verbosity=1):
        super().__init__(stream, descriptions, verbosity)
        self.test_results = []

    def addSuccess(self, test):
        super().addSuccess(test)
        test_class = test.__class__.__name__
        test_method = test._testMethodName
        input_data = getattr(test, "input_data", "")
        self.test_results.append((test_class, test_method, input_data, "OK"))

    def addFailure(self, test, err):
        super().addFailure(test, err)
        test_class = test.__class__.__name__
        test_method = test._testMethodName
        input_data = getattr(test, "input_data", "")
        self.test_results.append((test_class, test_method, input_data, "NOT OK"))


# Run tests and capture results
def run_tests(test_module):
    test_suite = unittest.defaultTestLoader.loadTestsFromModule(test_module)
    result = CustomTestResult()
    test_suite.run(result)
    return result.test_results


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(
            "Usage: python run_tests_and_generate_excel.py <test_file.py> <output_file.xlsx>"
        )
        sys.exit(1)

    test_file = sys.argv[1]
    output_file = sys.argv[2]

    try:
        module = importlib.import_module(test_file.replace(".py", ""))
    except ImportError:
        print(f"Error: Failed to import test file '{test_file}'")
        sys.exit(1)

    test_results = run_tests(module)
    write_test_results_to_excel(test_results, output_file)

    print(f"Test results generated and written to {output_file}")
