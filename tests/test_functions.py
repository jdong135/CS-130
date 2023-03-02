"""
Test for implementation of Lark Module formula Evaluator
"""

import decimal
import unittest
from context import sheets


class FunctionTests(unittest.TestCase):
    """
    Tests for Functions in lark
    """
    def test_bad_func_name(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AD(1,1)")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_NAME)

    def test_and1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(1, 1, 1)")
        wb.set_cell_contents("sheet1", "A2", "=AND(1, 0, 1)")
        wb.set_cell_contents("sheet1", "A3", "=AND(0, 1, 0)")
        wb.set_cell_contents("sheet1", "A4", "=AND(0, 0, 0)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), False)

    def test_and2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(1, 1 + 1, 1)")
        wb.set_cell_contents("sheet1", "A2", "=AND(1, 1 - 1, 1)")
        wb.set_cell_contents("sheet1", "A3", "=AND(0, 1, A2)")
        wb.set_cell_contents("sheet1", "A4", "=AND(A3, 0, A2)")
        wb.set_cell_contents(
            "sheet1", "A5", "=AND(1, 1 + 1, AND(1, AND(1, 5 * 4)))")
        wb.set_cell_contents("sheet1", "A6", "=AND(AND(1))")
        wb.set_cell_contents("sheet1", "A7", "=AND(AND(0))")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A5"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A6"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A7"), False)

    def test_and3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(TruE, tRUE, TRUE)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)

    def test_and_error1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(A1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1").get_type(),
                         sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_move_AND(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(A1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), False)
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), True)

    def test_move_AND_abs_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(A$1)")
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), False)

    def test_move_AND_diff_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "=true")
        wb.set_cell_contents("sheet1", "B1", "=AND(sheet1!A1)")
        wb.move_cells('sheet1', 'B1', 'B1', 'B2', 'sheet2')
        self.assertEqual(wb.get_cell_value('sheet2', 'B2'), True)

    def test_move_AND_ref_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=true")
        wb.set_cell_contents("sheet1", "B2", "=AND(A1)")
        wb.copy_cells('sheet1', 'B2', 'B2', 'A2', 'sheet1')
        value = wb.get_cell_value('sheet1', 'A2')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_break_AND_circref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(C1)")
        wb.set_cell_contents("sheet1", "B1", "=AND(A1)")
        wb.set_cell_contents("sheet1", "C1", "=AND(B1)")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value = wb.get_cell_value('sheet1', 'B1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value = wb.get_cell_value('sheet1', 'C1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.move_cells('sheet1', 'A1', 'A1', 'A2', 'sheet1')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), False)

    def test_AND_on_str(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(\"a\")")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_AND_invalid_num_args(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND()")
        value = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value, sheets.cell_error.CellError))
        self.assertTrue(value.get_type() ==
                        sheets.cell_error.CellErrorType.TYPE_ERROR)

    def test_AND_str_concat(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'tru")
        wb.set_cell_contents("sheet1", "B1", "=AND(A1 & \"e\")")
        self.assertEqual(wb.get_cell_value('sheet1', 'b1'), True)

    def test_xor(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=XOR(1, 0)")
        wb.set_cell_contents("sheet1", "A2", "=XOR (1, 0,  1)")
        wb.set_cell_contents("sheet1", "A3", "=XOR(1, 0, 1, AND(1 , 0))")
        wb.set_cell_contents(
            "sheet1", "A4", "=XOR (1, 0,  1, AND(1, XOR(1, 0, 1 , 1, 0), 1))")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A4"), True)

    def test_or(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=OR(1, 0, 0)")
        wb.set_cell_contents("sheet1", "A2", "=OR(0, 0, 0)")
        wb.set_cell_contents(
            "sheet1", "A3", "=OR(AND(True, true, 1, 5, -3), 0, FALSE)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), True)

    def test_not(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=NOT(0)")
        wb.set_cell_contents("sheet1", "A2", "=NOT(1)")
        wb.set_cell_contents(
            "sheet2", "A3", "=OR(XOR(1, 0, True), 0, NOT(True))")
        wb.set_cell_contents("sheet1", "A3", "=NOT(sheet2!A3)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), True)

    def test_iferror_ref_literal_and_non_literal_quotes(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=IFERROR("#REF!", 1)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "#REF!")
        wb.set_cell_contents("sheet1", "A2", "=IFERROR(#REF!, 1)")
        self.assertEqual(wb.get_cell_value("sheet1", "a2"), decimal.Decimal(1))

    def test_iferror_ref_literal1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=IFERROR(#REF!, 1)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal(1))

    def test_iferror_ref_literal2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", '=IFERROR("#REF!", 1)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "#REF!")

    def test_indirect_concat(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A5", '5')
        wb.set_cell_contents("sheet1", "B5", '62')
        wb.set_cell_contents("sheet1", "B1", '=indirect("B" & A5)')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'b1'), decimal.Decimal(62))
        wb.set_cell_contents("sheet1", "B1", '=indirect("A" & A5)')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'b1'), decimal.Decimal(5))

    def test_if_lazy_eval(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "B1", '=1/0')
        wb.set_cell_contents("sheet1", "a1", '=if(1, C1, B1)')
        self.assertEqual(wb.get_cell_value(
            'sheet1', 'a1'), decimal.Decimal(0))

    def test_update_circ_ref_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=ISERROR(B1)')
        wb.set_cell_contents("sheet1", "b1", '=ISERROR(C1)')
        wb.set_cell_contents("sheet1", "b1", '=ISERROR(A1)')
        value_a = wb.get_cell_value('sheet1', 'A1')
        self.assertTrue(isinstance(value_a, sheets.cell_error.CellError))
        self.assertTrue(value_a.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value_b = wb.get_cell_value('sheet1', 'B1')
        self.assertTrue(isinstance(value_b, sheets.cell_error.CellError))
        self.assertTrue(value_b.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_indirect_circ_ref1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=INDIRECT(\"A1\")")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_indirect_circ_ref2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=INDIRECT(A1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_lazy_if_ignore_circ_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=IF(A2, B1, C1)")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        wb.set_cell_contents("sheet1", "A2", "FALSE")

    def test_indirect_cell_ref_err(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=indirect(123)')
        self.assertTrue(wb.get_cell_value('sheet1', 'a1').get_type()
                        == sheets.cell_error.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents("sheet1", "a2", '=indirect(A20)')
        self.assertTrue(wb.get_cell_value('sheet1', 'a2').get_type()
                        == sheets.cell_error.CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents("sheet1", "a3", '=100')
        wb.set_cell_contents("sheet1", "a4", '=indirect(a3)')
        self.assertTrue(wb.get_cell_value('sheet1', 'a4').get_type()
                        == sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_invalid_commas_in_func(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=AND(,)')
        wb.set_cell_contents("sheet1", "a2", '=AND(,1)')
        wb.set_cell_contents("sheet1", "a3", '=AND(1,)')
        wb.set_cell_contents("sheet1", "a4", '=AND(,1,)')
        wb.set_cell_contents("sheet1", "a5", '=AND(1, , 1)')
        self.assertTrue(wb.get_cell_value('sheet1', 'a1').get_type()
                        == sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertTrue(wb.get_cell_value('sheet1', 'a2').get_type()
                        == sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertTrue(wb.get_cell_value('sheet1', 'a3').get_type()
                        == sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertTrue(wb.get_cell_value('sheet1', 'a4').get_type()
                        == sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertTrue(wb.get_cell_value('sheet1', 'a5').get_type()
                        == sheets.cell_error.CellErrorType.PARSE_ERROR)

    def test_boolean_function_equality(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=False == AND(TRUE, false)')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_ignore_whitespace_after_func(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=AND   (1, 1)')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), True)

    def test_ref_to_bad_name(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=FUNCTION(1, 1)')
        self.assertTrue(wb.get_cell_value('sheet1', 'a1').get_type()
                        == sheets.cell_error.CellErrorType.BAD_NAME)
        wb.set_cell_contents('sheet1', 'a2', '=A1')
        self.assertTrue(wb.get_cell_value('sheet1', 'a2').get_type()
                        == sheets.cell_error.CellErrorType.BAD_NAME)

    def test_AND_of_ref_err(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "a1", '=sheet5!a1')
        wb.set_cell_contents('sheet1', 'b1', '=AND(a1)')
        self.assertTrue(wb.get_cell_value('sheet1', 'b1').get_type()
                        == sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_empty_cell_arg_if(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "B1", "=IF(TRUE, A1)")
        wb.set_cell_contents("sheet1", "C1", "=IF(FALSE, A1)")
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(0))
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), False)

    def test_empty_cell_arg_exact(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "B1", "'")
        wb.set_cell_contents("sheet1", "C1", "=EXACT(A1, B1)")
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)

    def test_SCC_in_iserror1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        wb.set_cell_contents("sheet1", "C1", "=ISERROR(B1)")
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)

    def test_SCC_in_iserror2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=ISERROR(B1)")
        wb.set_cell_contents("sheet1", "B1", "=ISERROR(A1)")
        wb.set_cell_contents("sheet1", "C1", "=ISERROR(B1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)

    def test_if_update_to_circ_ref1(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "C1", "5")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        wb.set_cell_contents("sheet1", "A1", "=IF(A2, B1, C1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(5))
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(5))
        wb.set_cell_contents("sheet1", "A2", "TRUE")
        value_a = wb.get_cell_value("sheet1", "A1")
        value_b = wb.get_cell_value("sheet1", "B1")
        self.assertTrue(isinstance(value_a, sheets.cell_error.CellError))
        self.assertTrue(value_a.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        value_b = wb.get_cell_value('sheet1', 'B1')
        self.assertTrue(isinstance(value_b, sheets.cell_error.CellError))
        self.assertTrue(value_b.get_type() ==
                        sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "A2", "FALSE")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(5))
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(5))

    def test_iferror_circ_ref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=ISERROR(A1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)

    def test_nested_indirect_exact(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "D")
        wb.set_cell_contents("Sheet1", "B1", "'1")
        wb.set_cell_contents(
            "Sheet1", "C1", "=EXACT(\"HELLO\", \"HELL\" & INDIRECT(A1 & B1))")
        wb.set_cell_contents("Sheet1", "D1", "O")
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)

    def test_iserror_circref_multiple_types(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=ISERROR(A1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "=ISERROR(A1)")
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), True)

    def test_iserror_non_participating_circref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        # Order 1: A -> B -> C
        wb.set_cell_contents("sheet1", "A1", "=B1")
        wb.set_cell_contents("sheet1", "B1", "=A1")
        wb.set_cell_contents("sheet1", "C1", "=ISERROR(B1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)
        # Order 2: A -> C -> B
        wb.set_cell_contents("sheet1", "A2", "=B2")
        wb.set_cell_contents("sheet1", "C2", "=ISERROR(B2)")
        wb.set_cell_contents("sheet1", "B2", "=A2")
        self.assertEqual(wb.get_cell_value("Sheet1", "A2").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B2").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C2"), True)
        # Order 3: B -> A -> C
        wb.set_cell_contents("sheet1", "B3", "=A3")
        wb.set_cell_contents("sheet1", "A3", "=B3")
        wb.set_cell_contents("sheet1", "C3", "=ISERROR(B3)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A3").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B3").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C3"), True)
        # Order 4: B -> C -> A
        wb.set_cell_contents("sheet1", "B4", "=A4")
        wb.set_cell_contents("sheet1", "C4", "=ISERROR(B4)")
        wb.set_cell_contents("sheet1", "A4", "=B4")
        self.assertEqual(wb.get_cell_value("Sheet1", "A4").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B4").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C4"), True)
        # Order 5: C -> A -> B
        wb.set_cell_contents("sheet1", "C5", "=ISERROR(B5)")
        wb.set_cell_contents("sheet1", "A5", "=B5")
        wb.set_cell_contents("sheet1", "B5", "=A5")
        self.assertEqual(wb.get_cell_value("Sheet1", "A5").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B5").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C5"), True)
        # Order 6: C -> B -> A
        wb.set_cell_contents("sheet1", "C6", "=ISERROR(B6)")
        wb.set_cell_contents("sheet1", "B6", "=A6")
        wb.set_cell_contents("sheet1", "A6", "=B6")
        self.assertEqual(wb.get_cell_value("Sheet1", "A6").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B6").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C6"), True)

    def test_three_iserror_two_circref(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        # Order 1: A -> B -> C
        wb.set_cell_contents("sheet1", "A1", "=ISERROR(B1)")
        wb.set_cell_contents("sheet1", "B1", "=ISERROR(A1)")
        wb.set_cell_contents("sheet1", "C1", "=ISERROR(B1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C1"), True)
        # Order 2: A -> C -> B
        wb.set_cell_contents("sheet1", "A2", "=ISERROR(B2)")
        wb.set_cell_contents("sheet1", "C2", "=ISERROR(B2)")
        wb.set_cell_contents("sheet1", "B2", "=ISERROR(A2)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A2").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B2").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C2"), True)
        # Order 3: B -> A -> C
        wb.set_cell_contents("sheet1", "B3", "=ISERROR(A3)")
        wb.set_cell_contents("sheet1", "A3", "=ISERROR(B3)")
        wb.set_cell_contents("sheet1", "C3", "=ISERROR(B3)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A3").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B3").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C3"), True)
        # Order 4: B -> C -> A
        wb.set_cell_contents("sheet1", "B4", "=ISERROR(A4)")
        wb.set_cell_contents("sheet1", "C4", "=ISERROR(B4)")
        wb.set_cell_contents("sheet1", "A4", "=ISERROR(B4)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A4").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B4").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C4"), True)
        # Order 5: C -> A -> B
        wb.set_cell_contents("sheet1", "C5", "=ISERROR(B5)")
        wb.set_cell_contents("sheet1", "A5", "=ISERROR(B5)")
        wb.set_cell_contents("sheet1", "B5", "=ISERROR(A5)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A5").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B5").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C5"), True)
        # Order 6: C -> B -> A
        wb.set_cell_contents("sheet1", "C6", "=ISERROR(B6)")
        wb.set_cell_contents("sheet1", "A6", "=ISERROR(B6)")
        wb.set_cell_contents("sheet1", "B6", "=ISERROR(A6)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A6").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B6").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("sheet1", "C6"), True)

    def test_bad_name_propagation(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=BADFUNCNAME(A2)")
        wb.set_cell_contents("sheet1", "A2", "=A1")
        self.assertTrue(wb.get_cell_value('sheet1', 'a1').get_type()
                        == sheets.cell_error.CellErrorType.BAD_NAME)
        self.assertTrue(wb.get_cell_value('sheet1', 'a2').get_type()
                        == sheets.cell_error.CellErrorType.BAD_NAME)

    def test_move_indirect_reference(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=INDIRECT(B$1)")
        wb.set_cell_contents("sheet1", "B1", "C1")
        wb.set_cell_contents("sheet1", "C1", "=5")
        wb.move_cells("sheet1", "A1", "A1", "A2")
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), decimal.Decimal(5))

    def test_logic_error_propagation(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=NOT(AND(1, 2, XOR(1, 2, (OR(1, 1/0)), 1)))")
        self.assertEqual(wb.get_cell_value('sheet1', 'a1').get_type(
        ), sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)

    def test_exact_error_propagation(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=EXACT(#REF!, #REF!)")
        wb.set_cell_contents("sheet1", "A2", "=EXACT(#REF!, #ERROR!)")
        wb.set_cell_contents("sheet1", "A3", "=EXACT(#ERROR!, #REF!)")
        self.assertEqual(wb.get_cell_value('sheet1', 'a1').get_type(
        ), sheets.cell_error.CellErrorType.BAD_REFERENCE)
        self.assertEqual(wb.get_cell_value('sheet1', 'a2').get_type(
        ), sheets.cell_error.CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value('sheet1', 'a3').get_type(
        ), sheets.cell_error.CellErrorType.PARSE_ERROR)

    def test_isblank_ignore_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=ISBLANK(#REF!)")
        wb.set_cell_contents("sheet1", "A2", "=ISBLANK(1 + 1 + 1 / (1 - 1))")
        wb.set_cell_contents("sheet1", "B3", "=sheet2!a1")
        wb.set_cell_contents("sheet1", "A3", "=ISBLANK(B3)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A2"), False)
        self.assertEqual(wb.get_cell_value("sheet1", "A3"), False)

    def test_indirect_multiple_cases(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A2", "A3")
        wb.set_cell_contents("sheet1", "A3", "=5")
        wb.set_cell_contents("sheet1", "A1", "=INDIRECT(A2)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(5))
        wb.set_cell_contents("sheet1", "A1", "=INDIRECT(\"A2\")")     
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), "A3")
        wb.set_cell_contents("sheet1", "A1", "=INDIRECT(A4)")
        self.assertTrue(wb.get_cell_value('sheet1', 'A1').get_type()
                        == sheets.cell_error.CellErrorType.BAD_REFERENCE)

    def test_multi_sheet_function_calls(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "'Hell")
        wb.set_cell_contents("sheet2", "A1", "'Hel")
        wb.set_cell_contents("sheet2", "A2", "'o")
        wb.set_cell_contents("sheet1", "A2", "'lo")
        wb.set_cell_contents(
            "sheet3", "A1", "=EXACT(sheet1!A1 & sheet2!A2, sheet2!A1 & sheet1!A2)")
        self.assertEqual(wb.get_cell_value("sheet3", "A1"), True)

    def test_copy_indirect_function_block_detect_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=5")
        wb.set_cell_contents("sheet1", "A2", "=2 / A1")
        wb.set_cell_contents("sheet1", "A3", "A2")
        wb.set_cell_contents("sheet1", "A4", "=INDIRECT(A3) + 1")
        self.assertEqual(
            wb.get_cell_value("sheet1", "A4"), decimal.Decimal("1.4"))
        wb.copy_cells("sheet1", "A2", "A4", "B2")
        self.assertEqual(wb.get_cell_value('sheet1', 'B2').get_type(
        ), sheets.cell_error.CellErrorType.DIVIDE_BY_ZERO)
        self.assertEqual(
            wb.get_cell_value("sheet1", "B4"), decimal.Decimal("1.4"))

    def test_break_circ_ref_update_choose(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=CHOOSE(1, B1, 0)")
        wb.set_cell_contents("sheet1", "B1", "=CHOOSE(1, A1, 0)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "A1", "=CHOOSE(1, 1, B1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(1))
        wb.set_cell_contents("sheet1", "A1", "=CHOOSE(1, B1, B1)")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "B1").get_type(
        ), sheets.cell_error.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "B1", "=CHOOSE(1, 1, B1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), decimal.Decimal(1))
        self.assertEqual(wb.get_cell_value("sheet1", "B1"), decimal.Decimal(1))

    def test_isblank_cell_refs(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=ISBLANK(B1)")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)
        wb.set_cell_contents("sheet1", "B1", "=0")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), False)
        wb.set_cell_contents("sheet1", "B1", "FALSE")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), False)
        wb.set_cell_contents("sheet1", "B1", "=C1")
        self.assertEqual(wb.get_cell_value("sheet1", "A1"), True)

if __name__ == "__main__":
    unittest.main()
