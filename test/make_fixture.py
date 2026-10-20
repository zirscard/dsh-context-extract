#!/usr/bin/env python3
"""Generate sample inputs for dsh-context-extract self-test.

Run:  python3 test/make_fixture.py
Produces test/fixtures/sample-financial.xlsx, sample-ledger.csv, sample-contract.txt.
"""
from __future__ import annotations

import csv
import os

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")
os.makedirs(OUT, exist_ok=True)


def make_financial():
    wb = Workbook()
    ws = wb.active
    ws.title = "资产负债表"
    ws.append(["项目", "2024-12-31", "2023-12-31", "单位"])
    ws.append(["货币资金", 120000000, 98000000, "人民币元"])
    ws.append(["应收账款", 45000000, 40000000, "人民币元"])
    ws.append(["存货", 60000000, 55000000, "人民币元"])
    ws.append(["资产总计", 500000000, 450000000, "人民币元"])
    ws.append(["短期借款", 80000000, 70000000, "人民币元"])
    ws.append(["应付账款", 60000000, 55000000, "人民币元"])
    ws.append(["负债合计", 300000000, 260000000, "人民币元"])
    ws.append(["未分配利润", 100000000, 90000000, "人民币元"])
    ws.append(["所有者权益合计", 200000000, 190000000, "人民币元"])

    ws2 = wb.create_sheet("利润表")
    ws2.append(["项目", "2024年度", "2023年度"])
    ws2.append(["营业收入", 800000000, 700000000])
    ws2.append(["营业总成本", 700000000, 610000000])
    ws2.append(["净利润", 60000000, 55000000])

    ws3 = wb.create_sheet("现金流量表")
    ws3.append(["项目", "2024年度", "2023年度"])
    ws3.append(["经营活动产生的现金流量净额", 50000000, 40000000])
    ws3.append(["投资活动产生的现金流量净额", -20000000, -15000000])
    ws3.append(["筹资活动产生的现金流量净额", -10000000, -5000000])
    ws3.append(["期初现金及现金等价物余额", 98000000, 90000000])
    ws3.append(["期末现金及现金等价物余额", 118000000, 110000000])
    wb.save(os.path.join(OUT, "sample-financial.xlsx"))


def make_ledger():
    with open(os.path.join(OUT, "sample-ledger.csv"), "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["日期", "客户", "合同号", "金额(万元)", "回款"])
        w.writerow(["2026-01-05", "北京某某科技有限公司", "HT-001", 320, "已回"])
        w.writerow(["2026-02-11", "上海某某集团有限公司", "HT-002", 150, "未回"])
        w.writerow(["2026-03-20", "深圳某某合伙企业", "HT-003", 480, "已回"])


def make_contract():
    text = (
        "技术服务合同\n"
        "甲方：北京某某科技有限公司\n"
        "乙方：上海某某集团有限公司\n\n"
        "第一条 服务内容\n乙方为甲方提供数据平台建设与运维服务，服务期为二十四个月。\n\n"
        "第二条 服务费用\n合同总金额为人民币伍佰万元整，分四期支付。\n\n"
        "第三条 违约责任\n任何一方逾期付款超过三十日的，应按日万分之五支付违约金。\n"
    )
    with open(os.path.join(OUT, "sample-contract.txt"), "w", encoding="utf-8") as f:
        f.write(text)


if __name__ == "__main__":
    make_financial()
    make_ledger()
    make_contract()
    print("fixtures written to", OUT)
