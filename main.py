import openpyxl
import re


# Read the Python test file and extract test cases
def extract_test_cases(file_path):
    test_cases = []

    with open(file_path, "r") as f:
        content = f.read()
        # Modify this regex pattern based on your test file's structure
        test_matches = re.findall(r'def test_(.*?):\s*?r""".*?"""', content, re.DOTALL)

        for test_match in test_matches:
            test_name = test_match[0]
            test_description = re.search(
                r'"""(.*?)"""', test_match[1], re.DOTALL
            ).group(1)
            test_cases.append((test_name, test_description))

    return test_cases


# Create an Excel file and write test cases
def write_to_excel(test_cases, excel_file):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Test Cases"

    sheet["A1"] = "Test Name"
    sheet["B1"] = "Description"

    for row, (test_name, description) in enumerate(test_cases, start=2):
        sheet.cell(row=row, column=1, value=test_name)
        sheet.cell(row=row, column=2, value=description)

    wb.save(excel_file)


if __name__ == "__main__":
    python_test_file = "path/to/your/test_file.py"
    excel_output_file = "test_cases.xlsx"

    test_cases = extract_test_cases(python_test_file)
    write_to_excel(test_cases, excel_output_file)

    print(f"Test cases extracted and written to {excel_output_file}")
