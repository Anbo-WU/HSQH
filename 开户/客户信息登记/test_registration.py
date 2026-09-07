"""规则与错误取值回归测试；不需要启动 OCR。"""
import unittest

from 登记 import (beneficiary_names, capital_wan, credit_class, dates, expiry,
                financial_value, question18, suitability)
from ocr_support import Line


def line(text, x, y, width=.07):
    return Line(text, .99, x, y-.006, x+width, y+.006)


class RegistrationTests(unittest.TestCase):
    def test_capital_units_and_chinese_numbers(self):
        for text, expected in [('壹仟万元整', 1000), ('壹亿贰仟万元整', 12000),
                               ('10000000元', 1000), ('1.5亿元', 15000),
                               ('人民币贰佰伍拾万元整', 250), ('拾万元整', 10),
                               ('壹仟贰佰叁拾肆万伍仟陆佰柒拾捌元整', 1234.5678)]:
            with self.subTest(text=text):
                self.assertEqual(capital_wan(text), expected)
        for text in ['1000', '壹仟万美元', '不清楚']:
            with self.assertRaises(ValueError):
                capital_wan(text)

    def test_expiry_and_invalid_dates(self):
        self.assertEqual(expiry('2021.08.11-2041.08.11'), '2041.08.11')
        self.assertEqual(expiry('2021-08-11至2041-08-11'), '2041.08.11')
        self.assertEqual(expiry('2021年8月11日至长期'), '长期')
        self.assertEqual(dates('226年8月31日 2026年2月30日'), [])
        with self.assertRaises(ValueError):
            expiry('看不清')
        with self.assertRaises(ValueError):
            expiry('2021.08.11-2041.99.11')

    def test_credit_boundaries(self):
        self.assertEqual(credit_class(80), '正常')
        self.assertEqual(credit_class(100), '正常')
        self.assertEqual(credit_class(77), '关注')
        self.assertEqual(credit_class(60.5), '关注')
        for value in [60, 59, 0, -1]:
            with self.assertRaises(ValueError):
                credit_class(value)

    def test_all_suitability_suffixes(self):
        for suffix in ['C4', 'C5', 'C4低买高', 'B类专业交易者']:
            self.assertEqual(suitability('396、某公司-'+suffix), suffix)
        self.assertEqual(suitability('396、某公司'), '专业交易者')

    def test_every_beneficiary_across_pages_blank_slots_and_duplicates(self):
        def person(name, y):
            return [line('*股东/控制人/高管', .06, y, .14),
                    line('*姓名', .21, y), line(name, .33, y),
                    line('*身份证件号码', .55, y, .12)]
        first = person('张三', .27)+person('李四', .45)
        second = person('王五', .27)+person('张三', .45)
        second += [line('*股东/控制人/高管', .06, .65, .14),
                   line('*姓名', .21, .65), line('*身份证件号码', .55, .65)]
        value, _ = beneficiary_names([first, second])
        self.assertEqual(value, '张三、李四、王五')

    def test_incomplete_beneficiary_is_reported(self):
        entries = [line('*股东/控制人/高管', .06, .27, .14), line('*姓名', .21, .27),
                   line('*身份证件号码', .55, .27), line('123456789012345678', .70, .27)]
        with self.assertRaises(ValueError):
            beneficiary_names([entries])

    def test_question_answer_only_at_stem_end(self):
        self.assertEqual(question18([line('18.贵单位认为自己能承受的最大投资损失是多少？B', .1, .88, .7)]), 'B')
        self.assertEqual(question18([line('18.最大投资损失是多少B', .1, .88, .7)]), 'B')
        self.assertEqual(question18([line('18.贵单位认为自己能承受的最大投资', .1, .85, .7),
                                     line('损失是多少？', .1, .88, .4), line('D', .8, .88)]), 'D')
        with self.assertRaises(ValueError):
            question18([line('18.最大投资损失是多少？', .1, .8, .5), line('B.√10%-30%', .1, .85)])

    def test_balance_columns_exclude_row_number_and_opening_balance(self):
        lines = [line('单位：元', .82, .1),
                 line('行次', .247, .14, .031), line('期末余额', .278, .14),
                 line('年初余额', .388, .14, .054),
                 line('行次', .654, .14, .032), line('期末余额', .710, .14, .054),
                 line('年初余额', .813, .14, .053),
                 line('资产合计', .111, .68, .053), line('30', .248, .68, .020),
                 line('8,146,053.66', .290, .68, .075),
                 line('2,563,901.29负债和所有者权益', .383, .68, .235),
                 line('53', .656, .68, .020), line('8,146,053.66', .713, .68, .075)]
        value, _ = financial_value(lines, '资产合计', '期末余额')
        self.assertEqual(value, 8146053.66)

    def test_profit_annual_column_negative_and_units(self):
        lines = [line('单位：万元', .80, .1), line('行次', .51, .18, .025),
                 line('本期金额', .60, .18, .06), line('本年累计金额', .75, .18, .10),
                 line('四、净利润（净亏损以负号填列）', .10, .89, .38),
                 line('31', .51, .89, .025), line('276.00', .61, .89, .07),
                 line('-3,020.15', .75, .89, .10)]
        value, _ = financial_value(lines, '净利润', '本年累计金额')
        self.assertEqual(value, -30201500)


if __name__ == '__main__':
    unittest.main()
