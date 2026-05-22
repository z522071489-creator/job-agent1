#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
春招岗位自动采集智能体
每天 09:00 自动运行，采集公众号招聘文章，AI筛选打分，生成Excel，推送微信
"""

import os
import re
import json
import time
import requests
import tempfile
from datetime import datetime, date
from io import BytesIO
from urllib.parse import quote

from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image

# ============================================================
# 配置区 —— 通过 GitHub Secrets 注入
# ============================================================
KIMI_API_KEY   = os.environ.get("KIMI_API_KEY", "")
SERVERCHAN_KEY = os.environ.get("SERVERCHAN_KEY", "")

# 监控的公众号列表
ACCOUNTS = [
    "国企指南",
    "三晋国企",
    "山西国企直聘",
    "国聘行动",
    "优聘计划",
    "晋招聘",
    "三晋招聘",
    "山西招聘汇总",
    "晋师招聘",
    "山西热门招聘",
    "太原人才大市场公司招聘找工作",
]

# 关键词配置
TITLE_INCLUDE_KW = ["招聘", "春招", "校招", "岗位", "录用", "招录", "公告", "招募", "用人"]
TITLE_EXCLUDE_KW = ["实习", "兼职", "劳务", "外包"]
SCORE_THRESHOLD  = 50   # 低于此分不导出到Excel

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://weixin.sogou.com/",
}

# ============================================================
# 1. 搜狗微信搜索 —— 获取公众号文章列表
# ============================================================
def fetch_articles_sogou(account: str) -> list[dict]:
    """通过搜狗微信搜索，获取指定公众号最新招聘文章（最多15篇）"""
    articles = []
    seen_urls = set()

    # 用账号名 + 多个招聘关键词组合搜索，覆盖更多文章
    queries = [
        f"{account} 招聘",
        f"{account} 岗位",
        f"{account} 公告",
    ]

    for q in queries:
        url = (
            "https://weixin.sogou.com/weixin"
            f"?type=2&query={quote(q)}&ie=utf8"
        )
        try:
            resp = requests.get(url, headers=HEADERS, timeout=12)
            soup = BeautifulSoup(resp.text, "lxml")
            items = soup.select(".news-list li")

            for item in items:
                title_el = item.select_one("h3 a")
                date_el  = item.select_one(".s2")
                src_el   = item.select_one(".account")

                if not title_el:
                    continue

                source = src_el.get_text(strip=True) if src_el else ""
                # 只保留确实来自目标公众号的文章
                if account not in source and source:
                    continue

                art_url = title_el.get("href", "")
                if art_url in seen_urls:
                    continue
                seen_urls.add(art_url)

                articles.append({
                    "title":    title_el.get_text(strip=True),
                    "url":      art_url,
                    "pub_date": date_el.get_text(strip=True) if date_el else "",
                    "source":   account,
                })

                if len(articles) >= 15:
                    break

            time.sleep(1.5)

        except Exception as e:
            print(f"  [搜狗] {account} 搜索失败: {e}")

        if len(articles) >= 15:
            break

    return articles[:15]


# ============================================================
# 2. 标题关键词过滤
# ============================================================
def filter_by_title(articles: list[dict]) -> list[dict]:
    result = []
    for art in articles:
        t = art["title"]
        if any(kw in t for kw in TITLE_INCLUDE_KW):
            if not any(kw in t for kw in TITLE_EXCLUDE_KW):
                result.append(art)
    return result


# ============================================================
# 3. 抓取文章正文 + 图片列表
# ============================================================
def fetch_article_detail(url: str) -> tuple[str, list[str]]:
    """返回 (正文文本, 文章内图片URL列表)"""
    if not url.startswith("http"):
        return "", []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        soup = BeautifulSoup(resp.text, "lxml")

        body = soup.select_one("#js_content, .rich_media_content")
        if not body:
            return soup.get_text()[:4000], []

        text = body.get_text(separator="\n", strip=True)[:5000]

        imgs = []
        for img in body.select("img"):
            src = img.get("data-src") or img.get("src", "")
            if src and src.startswith("http") and "mmbiz" in src:
                imgs.append(src)

        return text, imgs[:12]

    except Exception as e:
        print(f"    抓取正文失败: {e}")
        return "", []


# ============================================================
# 4. 二维码识别
# ============================================================
def decode_qr_from_images(img_urls: list[str]) -> list[str]:
    """识别文章图片中的二维码，返回解码后的链接"""
    links = []
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
    except ImportError:
        return links

    for url in img_urls:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=8)
            img = Image.open(BytesIO(resp.content)).convert("RGB")
            codes = pyzbar_decode(img)
            for code in codes:
                data = code.data.decode("utf-8", errors="ignore")
                if data.startswith("http"):
                    links.append(data)
        except Exception:
            pass

    return links


# ============================================================
# 5. Kimi API —— 语义筛选 + 打分
# ============================================================
SYSTEM_PROMPT = """你是一个专业的招聘信息分析助手。

## 筛选规则
✅ 保留条件：
- 工作地点：太原、吕梁、山西全省任意城市
- 岗位类型：销售、交付专员、行政、国企、事业单位、正式岗、全职
- 学历：本科及以上（未注明学历也保留）

❌ 直接排除（score自动-100）：
- 实习、兼职、劳务派遣、外包
- 销售保险、销售房产、销售贷款

## 打分规则（总分从0开始叠加）
- 工作地点匹配（太原/吕梁/山西）：+30
- 岗位类型匹配（销售/行政/交付等）：+20
- 国企/央企/事业单位：+20
- 有明确薪资信息：+10
- 有明确截止日期：+10
- 含排除关键词：-100

## 输出格式（仅输出纯JSON，不含任何说明文字）
{
  "jobs": [
    {
      "job_name": "岗位名称（简洁）",
      "city": "工作城市",
      "score": 75,
      "salary_info": "5000-8000元/月 或 空字符串",
      "deadline": "2025-03-31 或 空字符串",
      "summary": "50字以内岗位摘要",
      "apply_link": "投递链接 或 空字符串",
      "excluded": false,
      "exclude_reason": "排除原因 或 空字符串"
    }
  ]
}"""


def analyze_with_kimi(article: dict, qr_links: list[str]) -> list[dict]:
    """调用Kimi API分析文章内容，提取并打分岗位"""
    if not KIMI_API_KEY:
        print("  ⚠️  KIMI_API_KEY 未配置，跳过AI分析")
        return []

    qr_info = ""
    if qr_links:
        qr_info = "\n\n投递二维码解码链接：\n" + "\n".join(qr_links)

    user_msg = (
        f"来源公众号：{article['source']}\n"
        f"文章标题：{article['title']}\n"
        f"发布日期：{article['pub_date']}\n"
        f"原文链接：{article['url']}\n\n"
        f"文章正文：\n{article.get('content', '（未能获取正文）')}"
        f"{qr_info}\n\n"
        f"请分析以上内容，提取所有岗位并打分。"
    )

    try:
        resp = requests.post(
            "https://api.moonshot.cn/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {KIMI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "moonshot-v1-32k",
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
                "temperature": 0.1,
                "max_tokens":  2048,
            },
            timeout=40,
        )
        raw = resp.json()["choices"][0]["message"]["content"]

        # 容错解析JSON
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            print(f"    Kimi 返回内容无法解析: {raw[:200]}")
            return []

        data = json.loads(m.group())
        jobs = data.get("jobs", [])

        for job in jobs:
            job["source"]        = article["source"]
            job["article_url"]   = article["url"]
            job["pub_date"]      = article["pub_date"]
            job["article_title"] = article["title"]
            job["qr_links"]      = qr_links

        return jobs

    except Exception as e:
        print(f"    Kimi API 调用失败: {e}")
        return []


# ============================================================
# 6. 生成 Excel
# ============================================================
def generate_excel(all_jobs: list[dict]) -> str:
    """生成标准化岗位清单，返回文件路径"""

    valid = [j for j in all_jobs if not j.get("excluded") and j.get("score", 0) >= SCORE_THRESHOLD]
    valid.sort(key=lambda x: x.get("score", 0), reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"岗位清单_{date.today().strftime('%m%d')}"

    # ---- 表头 ----
    headers = [
        "匹配度", "投递优先级", "岗位名称", "工作城市",
        "发布日期", "投递链接", "薪资信息", "岗位摘要",
        "原文链接", "来源公众号",
    ]

    h_fill = PatternFill("solid", fgColor="2E4057")
    h_font = Font(color="FFFFFF", bold=True, size=11, name="微软雅黑")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = h_fill
        c.font = h_font
        c.alignment = center
    ws.row_dimensions[1].height = 32

    # ---- 数据行 ----
    fills = {
        "✅ 优先投递": PatternFill("solid", fgColor="D4EDDA"),
        "⚠️ 待复核":  PatternFill("solid", fgColor="FFF3CD"),
    }

    for ri, job in enumerate(valid, 2):
        score = job.get("score", 0)
        priority = "✅ 优先投递" if score >= 80 else "⚠️ 待复核"

        apply_link = job.get("apply_link", "")
        if not apply_link and job.get("qr_links"):
            apply_link = job["qr_links"][0]
        if not apply_link:
            apply_link = job.get("article_url", "")

        row = [
            score,
            priority,
            job.get("job_name", ""),
            job.get("city", ""),
            job.get("pub_date", ""),
            apply_link,
            job.get("salary_info", ""),
            job.get("summary", ""),
            job.get("article_url", ""),
            job.get("source", ""),
        ]

        rf = fills.get(priority)
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            if rf:
                c.fill = rf
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if ci == 1:
                c.font = Font(bold=True, size=12)
                c.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[ri].height = 40

    # ---- 列宽 ----
    widths = [8, 13, 24, 10, 10, 42, 16, 40, 42, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    path = f"/tmp/招聘岗位_{date.today().strftime('%Y%m%d')}.xlsx"
    wb.save(path)
    return path


# ============================================================
# 7. 上传Excel到临时文件托管，获取下载链接
# ============================================================
def upload_excel(path: str) -> str:
    """上传到 file.io，返回7天有效下载链接"""
    try:
        with open(path, "rb") as f:
            resp = requests.post(
                "https://file.io",
                files={"file": (os.path.basename(path), f)},
                data={"expires": "7d"},
                timeout=20,
            )
        data = resp.json()
        if data.get("success"):
            return data.get("link", "")
    except Exception as e:
        print(f"  文件上传失败: {e}")
    return ""


# ============================================================
# 8. Server酱推送微信
# ============================================================
def push_wechat(title: str, body: str):
    if not SERVERCHAN_KEY:
        print("  ⚠️  SERVERCHAN_KEY 未配置，跳过推送")
        return
    try:
        resp = requests.post(
            f"https://sctapi.ftqq.com/{SERVERCHAN_KEY}.send",
            data={"title": title, "desp": body},
            timeout=12,
        )
        r = resp.json()
        if r.get("code") == 0:
            print("  ✅ 微信推送成功")
        else:
            print(f"  ⚠️  推送返回: {r}")
    except Exception as e:
        print(f"  ❌ 推送失败: {e}")


# ============================================================
# 主流程
# ============================================================
def main():
    start = datetime.now()
    print(f"\n{'='*50}")
    print(f"🚀 春招岗位采集启动 {start.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}\n")

    # ---- Step 1: 采集 ----
    all_articles = []
    for account in ACCOUNTS:
        print(f"📡 采集: {account}")
        arts = fetch_articles_sogou(account)
        print(f"   找到 {len(arts)} 篇文章")
        all_articles.extend(arts)

    # ---- Step 2: 标题过滤 ----
    filtered = filter_by_title(all_articles)
    print(f"\n🔍 标题过滤: {len(all_articles)} → {len(filtered)} 篇招聘文章\n")

    # ---- Step 3-5: 正文+二维码+AI ----
    all_jobs = []
    for i, art in enumerate(filtered, 1):
        print(f"[{i}/{len(filtered)}] {art['title'][:35]}...")
        content, imgs = fetch_article_detail(art["url"])
        art["content"] = content

        qr_links = decode_qr_from_images(imgs)
        if qr_links:
            print(f"   🔳 识别到二维码: {len(qr_links)} 个")

        jobs = analyze_with_kimi(art, qr_links)
        valid_in_art = [j for j in jobs if not j.get("excluded")]
        print(f"   💼 提取岗位: {len(jobs)} 个 | 有效: {len(valid_in_art)} 个")
        all_jobs.extend(jobs)
        time.sleep(2)

    # ---- Step 6: 统计 ----
    valid_jobs    = [j for j in all_jobs if not j.get("excluded") and j.get("score", 0) >= SCORE_THRESHOLD]
    priority_jobs = [j for j in valid_jobs if j.get("score", 0) >= 80]

    print(f"\n📊 汇总: 招聘文章{len(filtered)}篇 | 有效岗位{len(valid_jobs)}个 | 优先投递{len(priority_jobs)}个")

    # ---- Step 7: 生成Excel ----
    excel_path = ""
    dl_link = ""
    if valid_jobs:
        excel_path = generate_excel(all_jobs)
        print(f"📁 Excel已生成: {excel_path}")
        dl_link = upload_excel(excel_path)
        if dl_link:
            print(f"🔗 下载链接: {dl_link}")

    # ---- Step 8: 推送微信 ----
    elapsed = (datetime.now() - start).seconds
    summary_lines = [
        f"今日共采集 **{len(filtered)}** 篇招聘文章，筛选出 **{len(valid_jobs)}** 个符合要求的岗位，其中优先投递 **{len(priority_jobs)}** 个",
        "",
        f"⏱ 运行耗时：{elapsed}秒",
        "",
    ]

    if priority_jobs:
        summary_lines += ["### ✅ 优先投递岗位（前5条）", ""]
        for job in priority_jobs[:5]:
            name   = job.get("job_name", "")
            city   = job.get("city", "")
            score  = job.get("score", 0)
            salary = job.get("salary_info", "")
            src    = job.get("source", "")
            salary_str = f" | {salary}" if salary else ""
            summary_lines.append(f"- **{name}** · {city}{salary_str} · {score}分 · {src}")

        if len(priority_jobs) > 5:
            summary_lines.append(f"\n...共 {len(priority_jobs)} 个优先岗位，详见Excel")

    if dl_link:
        summary_lines += ["", f"📥 [点击下载Excel岗位清单]({dl_link})（7天有效）"]
    elif valid_jobs:
        summary_lines += ["", "📥 Excel已上传至 GitHub Actions Artifacts，请登录GitHub下载"]

    if not valid_jobs:
        summary_lines += ["", "💬 今日未发现符合条件的岗位，明日继续监控。"]

    push_wechat(
        title=f"🎯 招聘速报 {date.today().strftime('%m/%d')} · {len(valid_jobs)}个岗位 · {len(priority_jobs)}个优先",
        body="\n".join(summary_lines),
    )

    print(f"\n✅ 全部完成！共耗时 {elapsed} 秒\n")


if __name__ == "__main__":
    main()
