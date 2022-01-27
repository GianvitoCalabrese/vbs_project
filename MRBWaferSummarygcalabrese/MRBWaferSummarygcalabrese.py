import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet.range("A1").value == "Hello xlwings!":
        sheet.range("A1").value = "Bye xlwings!"
    else:
        sheet.range("A1").value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("MRBWaferSummarygcalabrese.xlsm").set_mock_caller()
    main()
