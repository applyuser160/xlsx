import sys
from xlsx import load_workbook

def test_read_and_write_cell_value():
    # 'data/sample.xlsx' を読み込む
    wb = load_workbook("data/sample.xlsx")

    # シート名の一覧を確認
    print(f"Sheet names: {wb.sheetnames}")

    # "シート1" シートを取得
    ws = wb["シート1"]
    print(f"Successfully got worksheet: {ws.name}")

    # セルの値を取得して確認
    val_a1 = ws["A1"].value
    print(f"Value of A1: {val_a1}")
    assert val_a1 == "1.0"

    val_b2 = ws["B2"].value
    print(f"Value of B2: {val_b2}")
    assert val_b2 == "4.0"

    # 存在しないセルの値がNoneであることを確認
    val_z99 = ws["Z99"].value
    print(f"Value of Z99: {val_z99}")
    assert val_z99 is None

    # A1セルの値を変更
    ws["A1"].value = "99.9"
    print(f"Changed A1 value to: 99.9")

    # 変更を新しいファイルに保存
    output_path = "data/test_output.xlsx"
    wb.copy(output_path)
    print(f"Saved changes to {output_path}")

    # 保存したファイルを読み直して確認
    wb_new = load_workbook(output_path)
    ws_new = wb_new["シート1"]
    val_a1_new = ws_new["A1"].value
    print(f"Value of A1 from new file: {val_a1_new}")
    assert val_a1_new == "99.9"

    # 元のB2の値が変わっていないことも確認
    val_b2_new = ws_new["B2"].value
    print(f"Value of B2 from new file: {val_b2_new}")
    assert val_b2_new == "4.0"

    # cell()メソッドのテスト
    b1_cell = ws.cell(row=1, column=2)
    print(f"Value of B1 (via cell method): {b1_cell.value}")
    assert b1_cell.value == "3.0"

    # cell()メソッド経由での書き込み
    b1_cell.value = "3.14"
    print(f"Changed B1 value to: 3.14")
    assert ws["B1"].value == "3.14"


    print("All tests passed!")


def test_read_shared_strings():
    wb = load_workbook("data/shared_strings_test.xlsx")
    print(f"Sheet names in shared_strings_test.xlsx: {wb.sheetnames}")
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    print(f"Testing sheet: {sheet_name}")
    # 文字列のテスト
    val_a1 = ws["A1"].value
    print(f"A1 value: {val_a1}")
    assert val_a1 == "Hello"
    assert ws["A2"].value == "World"
    assert ws["A3"].value == "Hello"
    print("String values are correct.")

    # 数値のテスト
    assert ws["B1"].value == "123"
    assert ws["B2"].value == "45.67"
    print("Numeric values are correct.")

    print("Shared strings test passed!")


if __name__ == "__main__":
    test_read_and_write_cell_value()
    test_read_shared_strings()
