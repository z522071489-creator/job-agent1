# -*- coding: utf-8 -*-
"""
本地调试工具 - 模拟 SCF 返回数据测试 Excel 生成和微信推送
不依赖真实采集，用于快速验证后两个步骤是否正常工作

使用方法：
  SERVERCHAN_KEY=SCTxxxxxxx python scripts/test_local.py
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.run import make_excel, push_wechat
from datetime import date

# 模拟数据
MOCK_JOBS = [
    {
        "job_title": "综合管理专员",
        "company": "山西煤炭运销集团有限公司",
        "city": "太原",
        "education": "本科",
        "job_type": "社会招聘",
        "salary": "6000-8000元/月",
        "deadline": "2026-06-30",
        "summary": "负责公司日常行政管理、文件归档及协调沟通工作",
        "score": 90,
        "article_url": "https://example.com/job1",
        "pub_date": str(date.today()),
        "source": "微信公众号·山西国企直聘",
        "exclude_reason": None,
    },
    {
        "job_title": "财务会计",
        "company": "太原重型机械集团有限公司",
        "city": "太原",
        "education": "本科",
        "job_type": "社会招聘",
        "salary": None,
        "deadline": "2026-06-15",
        "summary": "负责企业日常账务处理、报表编制及税务申报",
        "score": 80,
        "article_url": "https://example.com/job2",
        "pub_date": str(date.today()),
        "source": "国聘网",
        "exclude_reason": None,
    },
    {
        "job_title": "运营专员",
        "company": "山西焦煤集团有限责任公司",
        "city": "太原",
        "education": "大专",
        "job_type": "社会招聘",
        "salary": "5000-7000元/月",
        "deadline": None,
        "summary": "协助开展线上运营活动、数据分析及内容管理",
        "score": 70,
        "article_url": "https://example.com/job3",
        "pub_date": str(date.today()),
        "source": "微信公众号·三晋国企",
        "exclude_reason": None,
    },
    {
        "job_title": "客服专员",
        "company": "吕梁市某国有企业",
        "city": "吕梁",
        "education": "高中",
        "job_type": "社会招聘",
        "salary": None,
        "deadline": None,
        "summary": "负责客户咨询、投诉处理及满意度维护工作",
        "score": 60,
        "article_url": "https://example.com/job4",
        "pub_date": str(date.today()),
        "source": "山西省国资委",
        "exclude_reason": None,
    },
]

MOCK_STATS = {
    "total_articles": 28,
    "wx_count": 18,
    "website_count": 10,
    "job_count": 4,
    "priority_count": 2,
    "errors": ["山西省人事考试网: 连接超时"],
}


if __name__ == "__main__":
    today = str(date.today())
    print("🧪 本地测试模式\n")

    # 测试 Excel 生成
    excel_path = f"/tmp/招聘汇总_测试_{today}.xlsx"
    make_excel(MOCK_JOBS, excel_path)
    print(f"✅ Excel 测试文件已生成: {excel_path}")

    # 测试微信推送（需要设置环境变量）
    key = os.environ.get("SERVERCHAN_KEY", "")
    if key:
        print("\n📱 测试微信推送...")
        push_wechat(MOCK_STATS, MOCK_JOBS, "https://example.com/download", today)
    else:
        print("\n⚠️ 未设置 SERVERCHAN_KEY，跳过推送测试")
        print("   运行方式：SERVERCHAN_KEY=SCTxxxxxx python scripts/test_local.py")
