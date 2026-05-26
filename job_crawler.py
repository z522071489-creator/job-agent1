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
from datetime import datetime, date, timedelta
from io import BytesIO
from urllib.parse import quote, urlparse, parse_qs
from pathlib import Path

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

ACCOUNTS = [
    "国企指南", "三晋国企", "山西国企直聘", "国聘行动", "优聘计划",
    "晋招聘", "三晋招聘", "山西招聘汇总", "晋师招聘",
    "山西热门招聘", "太原人才大市场公司招聘找工作",
]

TITLE_INCLUDE_KW = ["招聘", "春招", "校招", "岗位", "录用", "招录", "公告", "招募", "用人"]
TITLE_EXCLUDE_KW = ["实习", "兼职", "劳务", "外包"]
SCORE_THRESHOLD  = 50

# ============================================================
# 个人求职条件 —— 通过 GitHub Secrets 注入
# ============================================================
# 学历：本科及以下（排除：硕士、博士、研究生）
MAX_EDUCATION = "本科"
EXCLUDE_EDUCATION = ["硕士", "研究生", "博士", "硕士研究生", "博士研究生"]

# 性别：男性（排除：限女性的岗位）
PREFERRED_GENDER = "不限"
EXCLUDE_GENDER = ["女性", "女"]


HEADERS_PC = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# ============================================================
# 1. Bing 搜索获取微信文章列表（核心改动）
# ============================================================
def fetch_articles_bing(account: str) -> list[dict]:
    """通过 Bing 搜索 site:mp.weixin.qq.com 获取公众号文章"""
    articles = []
    seen_urls = set()
 
    # 多个搜索词组合，覆盖更多文章
    queries = [
        f'site:mp.weixin.qq.com "{account}" 招聘',
        f'site:mp.weixin.qq.com "{account}" 岗位',
        f'site:mp.weixin.qq.com "{account}" 公告',
    ]
 
    for q in queries:
        url = f"https://www.bing.com/search?q={quote(q)}&count=10&setlang=zh-CN"
        try:
            resp = requests.get(url, headers=HEADERS_PC, timeout=15)
            soup = BeautifulSoup(resp.text, "lxml")
 
            # Bing 搜索结果条目
            for item in soup.select("li.b_algo"):
                title_el = item.select_one("h2 a")
                desc_el  = item.select_one(".b_caption p, .b_algoSlug")
                date_el  = item.select_one(".b_attribution cite, .news_dt")
 
                if not title_el:
                    continue
 
                art_url = title_el.get("href", "")
                # 只要微信文章链接
                if "mp.weixin.qq.com" not in art_url:
                    continue
                if art_url in seen_urls:
                    continue
                seen_urls.add(art_url)
 
                title   = title_el.get_text(strip=True)
                pub_date = date_el.get_text(strip=True) if date_el else ""
 
                articles.append({
                    "title":    title,
                    "url":      art_url,
                    "pub_date": pub_date,
                    "source":   account,
                })
 
                if len(articles) >= 15:
                    break
 
            time.sleep(2)  # Bing 反爬间隔
 
        except Exception as e:
            print(f"  [Bing] {account} 搜索失败: {e}")
 
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
# 3. 抓取微信文章正文
# ============================================================
def fetch_article_detail(url: str) -> tuple[str, list[str]]:
    """抓取微信文章正文和图片"""
    if not url.startswith("http"):
        return "", []
    try:
        resp = requests.get(url, headers=HEADERS_PC, timeout=15)
        soup = BeautifulSoup(resp.text, "lxml")
 
        body = soup.select_one("#js_content, .rich_media_content")
        if not body:
            text = soup.get_text()[:5000]
            print(f"    未找到正文容器，使用全文: {len(text)} 字")
            return text, []
 
        text = body.get_text(separator="\n", strip=True)[:6000]
        imgs = []
        for img in body.select("img"):
            src = img.get("data-src") or img.get("src", "")
            if src and src.startswith("http") and "mmbiz" in src:
                imgs.append(src)
 
        print(f"    正文: {len(text)} 字，图片: {len(imgs)} 张")
        return text, imgs[:10]
 
    except Exception as e:
        print(f"    抓取正文失败: {e}")
        return "", []
 
 
# ============================================================
# 4. 二维码识别
# ============================================================
def decode_qr(img_urls: list[str]) -> list[str]:
    links = []
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
    except ImportError:
        return links
    for url in img_urls:
        try:
            resp = requests.get(url, headers=HEADERS_PC, timeout=8)
            img = Image.open(BytesIO(resp.content)).convert("RGB")
            for code in pyzbar_decode(img):
                data = code.data.decode("utf-8", errors="ignore")
                if data.startswith("http"):
                    links.append(data)
        except Exception:
            pass
    return links
 
 
# ============================================================
# 5. Kimi API 语义筛选 + 打分
# ============================================================
SYSTEM_PROMPT = """你是一个专业的招聘信息分析助手。
 
## 前置筛选条件（必须同时满足才保留）
1. 招聘类型：必须是【社会招聘】（含校招+社招混合批次均可）
   ❌ 排除：仅限应届生/仅限校招/实习/兼职/志愿者
2. 学历要求：【高中及以上】均保留
   ✅ 保留：高中、中专、大专、本科、研究生、学历不限
   ❌ 排除：仅限硕士/博士（不同时招本科的）
3. 工作性质：正式用工
   ❌ 排除：实习、兼职、劳务派遣、外包、临时工
 
## 打分规则（从0分开始累加）
+ 工作地点为太原/吕梁/山西：+30分
+ 岗位类型匹配（销售/行政/交付/运营/管理/技术）：+20分
+ 国企/央企/事业单位：+20分
+ 有明确薪资信息：+10分
+ 有明确截止日期：+10分
- 含排除关键词（实习/外包/派遣/仅限应届）：-100分
 
## 严格要求
- 正文不足100字时，直接返回空jobs列表，绝对不能编造
- 岗位名称、城市、薪资、截止日期必须来自原文
- 不确定的字段填空字符串，不要猜测
- 一篇文章可以提取多个不同岗位
 
## 输出（纯JSON，不含任何说明文字）
{
  "jobs": [
    {
      "job_name": "岗位名称（来自原文）",
      "employer": "招聘单位",
      "city": "工作城市（来自原文，不确定填空）",
      "recruitment_type": "社会招聘/校招/混合",
      "education_req": "学历要求（来自原文，没有填不限）",
      "score": 75,
      "salary_info": "薪资（来自原文，没有填空）",
      "deadline": "截止日期（来自原文，没有填空）",
      "summary": "80字内摘要（基于原文）",
      "apply_link": "投递链接（来自原文，没有填空）",
      "excluded": false,
      "exclude_reason": ""
    }
  ]
}"""
 
 
def analyze_with_kimi(article: dict, qr_links: list[str]) -> list[dict]:
    if not KIMI_API_KEY:
        print("  ⚠️  KIMI_API_KEY 未配置")
        return []
 
    content = article.get("content", "")
    if len(content) < 100:
        print(f"  ⚠️  正文太短({len(content)}字)，跳过")
        return []
 
    qr_info = ("\n\n二维码解码链接：\n" + "\n".join(qr_links)) if qr_links else ""
    user_msg = (
        f"来源公众号：{article['source']}\n"
        f"文章标题：{article['title']}\n"
        f"发布日期：{article['pub_date']}\n"
        f"原文链接：{article['url']}\n\n"
        f"文章正文：\n{content}"
        f"{qr_info}\n\n"
        f"请分析以上招聘文章，提取所有符合前置条件（社会招聘+高中及以上）的岗位并打分。"
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
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            print(f"    Kimi返回无法解析")
            return []
        data  = json.loads(m.group())
        jobs  = data.get("jobs", [])
        for job in jobs:
            job["source"]      = article["source"]
            job["article_url"] = article["url"]
            job["pub_date"]    = article["pub_date"]
            job["qr_links"]    = qr_links
        return jobs
    except Exception as e:
        print(f"    Kimi调用失败: {e}")
        return []
 
 
# ============================================================
# 6. 生成 Excel
# ============================================================
def generate_excel(all_jobs: list[dict]) -> str:
    valid = [
        j for j in all_jobs
        if not j.get("excluded") and j.get("score", 0) >= SCORE_THRESHOLD
    ]
    valid.sort(key=lambda x: x.get("score", 0), reverse=True)
 
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"岗位清单_{date.today().strftime('%m%d')}"
 
    headers = [
        "匹配度", "投递优先级", "岗位名称", "招聘单位",
        "工作城市", "学历要求", "招聘类型", "发布日期",
        "投递链接", "薪资信息", "截止日期", "岗位摘要",
        "原文链接", "来源公众号",
    ]
 
    h_fill = PatternFill("solid", fgColor="2E4057")
    h_font = Font(color="FFFFFF", bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
 
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = h_fill
        c.font = h_font
        c.alignment = center
    ws.row_dimensions[1].height = 32
 
    fills = {
        "✅ 优先投递": PatternFill("solid", fgColor="D4EDDA"),
        "⚠️ 待复核":  PatternFill("solid", fgColor="FFF3CD"),
    }
 
    for ri, job in enumerate(valid, 2):
        score    = job.get("score", 0)
        priority = "✅ 优先投递" if score >= 80 else "⚠️ 待复核"
        link     = (job.get("apply_link") or
                    (job.get("qr_links") or [""])[0] or
                    job.get("article_url", ""))
        row = [
            score, priority,
            job.get("job_name", ""),
            job.get("employer", ""),
            job.get("city", ""),
            job.get("education_req", ""),
            job.get("recruitment_type", ""),
            job.get("pub_date", ""),
            link,
            job.get("salary_info", ""),
            job.get("deadline", ""),
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
        ws.row_dimensions[ri].height = 45
 
    widths = [8, 13, 22, 20, 10, 10, 10, 10, 42, 15, 12, 36, 42, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
 
    # 同时打印岗位明细到日志（方便在GitHub Actions里直接查看）
    print("\n" + "="*60)
    print(f"📋 岗位明细（共{len(valid)}个）")
    print("="*60)
    for job in valid:
        print(f"\n【{job.get('job_name','')}】{job.get('employer','')} | {job.get('city','')} | {job.get('score',0)}分")
        print(f"  学历：{job.get('education_req','')} | 类型：{job.get('recruitment_type','')} | 薪资：{job.get('salary_info','未知')}")
        print(f"  截止：{job.get('deadline','未知')} | 来源：{job.get('source','')}")
        print(f"  摘要：{job.get('summary','')}")
        print(f"  链接：{job.get('article_url','')}")
    print("="*60 + "\n")
 
    path = f"/tmp/招聘岗位_{date.today().strftime('%Y%m%d')}.xlsx"
    wb.save(path)
    return path, valid
 
 
# ============================================================
# 7. 上传 Excel
# ============================================================
def upload_excel(path: str) -> str:
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
# 8. Server酱推送
# ============================================================
def push_wechat(title: str, body: str):
    if not SERVERCHAN_KEY:
        print("  ⚠️  SERVERCHAN_KEY未配置")
        return
    try:
        resp = requests.post(
            f"https://sctapi.ftqq.com/{SERVERCHAN_KEY}.send",
            data={"title": title, "desp": body},
            timeout=12,
        )
        r = resp.json()
        print("  ✅ 微信推送成功" if r.get("code") == 0 else f"  ⚠️ 推送: {r}")
    except Exception as e:
        print(f"  ❌ 推送失败: {e}")
 
 
# ============================================================
# 主流程
# ============================================================
def main():
    start = datetime.now()
    print(f"\n{'='*55}")
    print(f"🚀 春招岗位采集 v3 启动 {start.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📌 数据源：Bing搜索（稳定）")
    print(f"📌 筛选：社会招聘 | 高中及以上 | 太原/吕梁/山西优先")
    print(f"{'='*55}\n")
 
    # Step 1: 采集
    all_articles = []
    for account in ACCOUNTS:
        print(f"📡 采集: {account}")
        arts = fetch_articles_bing(account)
        print(f"   找到 {len(arts)} 篇文章")
        all_articles.extend(arts)
 
    if not all_articles:
        print("\n⚠️  所有公众号均未采集到文章，请检查网络或搜索词配置")
        push_wechat(
            title=f"⚠️ 招聘采集异常 {date.today().strftime('%m/%d')}",
            body="今日采集到0篇文章，可能是搜索引擎暂时不可用，明日自动重试。"
        )
        return
 
    # Step 2: 标题过滤
    filtered = filter_by_title(all_articles)
    print(f"\n🔍 标题过滤: {len(all_articles)} → {len(filtered)} 篇\n")
 
    # Step 3-5: 正文+二维码+AI
    all_jobs = []
    for i, art in enumerate(filtered, 1):
        print(f"[{i}/{len(filtered)}] {art['title'][:40]}...")
        content, imgs = fetch_article_detail(art["url"])
        art["content"] = content
        qr_links = decode_qr(imgs)
        if qr_links:
            print(f"   🔳 二维码: {len(qr_links)} 个")
        jobs = analyze_with_kimi(art, qr_links)
        valid_n = len([j for j in jobs if not j.get("excluded")])
        print(f"   💼 提取: {len(jobs)} 个，有效: {valid_n} 个")
        all_jobs.extend(jobs)
        time.sleep(2)
 
    # Step 6: Excel
    valid_jobs, priority_jobs = [], []
    excel_path, dl_link = "", ""
    if all_jobs:
        excel_path, valid_jobs_list = generate_excel(all_jobs)
        valid_jobs    = valid_jobs_list
        priority_jobs = [j for j in valid_jobs if j.get("score", 0) >= 80]
        dl_link = upload_excel(excel_path)
 
    elapsed = (datetime.now() - start).seconds
    print(f"\n📊 汇总: 文章{len(filtered)}篇 | 有效{len(valid_jobs)}个 | 优先{len(priority_jobs)}个 | 耗时{elapsed}秒")
 
    # Step 7: 推送
    lines = [
        f"今日共采集 **{len(filtered)}** 篇招聘文章，"
        f"筛选出 **{len(valid_jobs)}** 个符合要求的岗位，"
        f"其中优先投递 **{len(priority_jobs)}** 个",
        f"\n> 筛选条件：社会招聘 | 高中及以上 | 太原/吕梁/山西优先",
        f"\n⏱ 耗时：{elapsed}秒  \n",
    ]
 
    if priority_jobs:
        lines += ["### ✅ 优先投递岗位", ""]
        for job in priority_jobs[:8]:
            sal  = f" | {job['salary_info']}"  if job.get("salary_info")  else ""
            edu  = f" | {job['education_req']}" if job.get("education_req") else ""
            dead = f" | 截止{job['deadline']}"  if job.get("deadline")     else ""
            lines.append(
                f"**{job.get('job_name','')}** · {job.get('employer','')} · "
                f"{job.get('city','')}{sal}{edu}{dead} · {job.get('score',0)}分  "
            )
        if len(priority_jobs) > 8:
            lines.append(f"\n...共{len(priority_jobs)}个优先岗位")
 
    if valid_jobs and len(valid_jobs) > len(priority_jobs):
        lines += ["", "### ⚠️ 待复核岗位", ""]
        review = [j for j in valid_jobs if j.get("score", 0) < 80]
        for job in review[:5]:
            lines.append(
                f"**{job.get('job_name','')}** · {job.get('employer','')} · "
                f"{job.get('city','')} · {job.get('score',0)}分  "
            )
 
    if dl_link:
        lines += ["", f"📥 [点击下载Excel岗位清单]({dl_link})（7天有效）"]
    elif valid_jobs:
        lines += ["", "📥 Excel已上传至 GitHub Actions Artifacts，登录GitHub下载"]
 
    if not valid_jobs:
        lines += ["", "💬 今日未发现符合条件的岗位。"]
 
    push_wechat(
        title=f"🎯 招聘速报 {date.today().strftime('%m/%d')} · {len(valid_jobs)}个岗位 · {len(priority_jobs)}个优先",
        body="\n".join(lines),
    )
    print("✅ 完成！\n")
 
 
if __name__ == "__main__":
    main()
