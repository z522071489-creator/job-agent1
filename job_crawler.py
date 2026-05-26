#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# v4 - Baidu search as primary, debug output added

import os
import re
import json
import time
import requests
from datetime import datetime, date
from io import BytesIO
from urllib.parse import quote

from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image

# ============================================================
# Config
# ============================================================
KIMI_API_KEY   = os.environ.get("KIMI_API_KEY", "")
SERVERCHAN_KEY = os.environ.get("SERVERCHAN_KEY", "")

ACCOUNTS = [
    "国企指南", "三晋国企", "山西国企直聘", "国聘行动", "优聘计划",
    "晋招聘", "三晋招聘", "山西招聘汇总", "晋师招聘",
    "山西热门招聘", "太原人才大市场公司招聘找工作",
]

TITLE_INCLUDE_KW = ["招聘", "春招", "校招", "岗位", "录用", "招录", "公告", "招募"]
TITLE_EXCLUDE_KW = ["实习", "兼职", "劳务", "外包"]
SCORE_THRESHOLD  = 50

HEADERS_BAIDU = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Referer": "https://www.baidu.com/",
}

HEADERS_WX = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9",
}


# ============================================================
# 1. Baidu search for WeChat articles
# ============================================================
def fetch_articles_baidu(account):
    articles = []
    seen_urls = set()

    queries = [
        'site:mp.weixin.qq.com {} 招聘'.format(account),
        'site:mp.weixin.qq.com {} 岗位'.format(account),
    ]

    for q in queries:
        url = "https://www.baidu.com/s?wd={}&rn=20&ie=utf-8".format(quote(q))
        try:
            resp = requests.get(url, headers=HEADERS_BAIDU, timeout=15)
            print("  Baidu status={} len={}".format(resp.status_code, len(resp.text)))

            soup = BeautifulSoup(resp.text, "lxml")

            # Baidu result selectors - try multiple
            items = (soup.select("div.result") or
                     soup.select("div.c-container") or
                     soup.select("[tpl='se_com_default']"))

            print("  Baidu items found: {}".format(len(items)))

            for item in items:
                # find link
                link_el = (item.select_one("h3.t a") or
                           item.select_one("h3 a") or
                           item.select_one("a[href*='mp.weixin']"))
                if not link_el:
                    continue

                href = link_el.get("href", "")
                title = link_el.get_text(strip=True)

                # Baidu wraps real URLs - follow redirect
                if "baidu.com/link" in href or not href.startswith("http"):
                    try:
                        r2 = requests.get(href, headers=HEADERS_BAIDU,
                                          timeout=8, allow_redirects=True)
                        final_url = r2.url
                    except Exception:
                        final_url = href
                else:
                    final_url = href

                if "mp.weixin.qq.com" not in final_url:
                    continue
                if final_url in seen_urls:
                    continue
                seen_urls.add(final_url)

                date_el = item.select_one(".c-color-gray2, .newTimeFactor_before_abs, .c-abstract")
                pub_date = ""
                if date_el:
                    txt = date_el.get_text(strip=True)
                    # extract date pattern
                    m = re.search(r"\d{4}[-年]\d{1,2}[-月]\d{1,2}", txt)
                    if m:
                        pub_date = m.group()

                articles.append({
                    "title":    title,
                    "url":      final_url,
                    "pub_date": pub_date,
                    "source":   account,
                })
                print("  + {}".format(title[:30]))

                if len(articles) >= 15:
                    break

            time.sleep(3)

        except Exception as e:
            print("  [Baidu] {} failed: {}".format(account, e))

        if len(articles) >= 15:
            break

    return articles[:15]


# ============================================================
# 2. Title filter
# ============================================================
def filter_by_title(articles):
    result = []
    for art in articles:
        t = art["title"]
        if any(kw in t for kw in TITLE_INCLUDE_KW):
            if not any(kw in t for kw in TITLE_EXCLUDE_KW):
                result.append(art)
    return result


# ============================================================
# 3. Fetch WeChat article content
# ============================================================
def fetch_article_detail(url):
    if not url.startswith("http"):
        return "", []
    try:
        resp = requests.get(url, headers=HEADERS_WX, timeout=15)
        soup = BeautifulSoup(resp.text, "lxml")
        body = soup.select_one("#js_content, .rich_media_content")
        if not body:
            text = soup.get_text()[:5000]
            print("    no body container, text={}".format(len(text)))
            return text, []
        text = body.get_text(separator="\n", strip=True)[:6000]
        imgs = []
        for img in body.select("img"):
            src = img.get("data-src") or img.get("src", "")
            if src and src.startswith("http") and "mmbiz" in src:
                imgs.append(src)
        print("    text={} chars, imgs={}".format(len(text), len(imgs)))
        return text, imgs[:10]
    except Exception as e:
        print("    fetch failed: {}".format(e))
        return "", []


# ============================================================
# 4. QR decode
# ============================================================
def decode_qr(img_urls):
    links = []
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
    except ImportError:
        return links
    for url in img_urls:
        try:
            resp = requests.get(url, headers=HEADERS_WX, timeout=8)
            img = Image.open(BytesIO(resp.content)).convert("RGB")
            for code in pyzbar_decode(img):
                data = code.data.decode("utf-8", errors="ignore")
                if data.startswith("http"):
                    links.append(data)
        except Exception:
            pass
    return links


# ============================================================
# 5. Kimi API
# ============================================================
SYSTEM_PROMPT = """你是一个专业的招聘信息分析助手。

## 前置筛选条件（必须同时满足才保留）
1. 招聘类型：必须是社会招聘（含校招+社招混合批次均可）
   排除：仅限应届生/仅限校招/实习/兼职/志愿者
2. 学历要求：高中及以上均保留
   保留：高中、中专、大专、本科、研究生、学历不限
   排除：仅限硕士/博士（不同时招本科的）
3. 工作性质：正式用工
   排除：实习、兼职、劳务派遣、外包、临时工

## 打分规则（从0分开始）
工作地点为太原/吕梁/山西 +30
岗位类型匹配（销售/行政/交付/运营/管理/技术）+20
国企/央企/事业单位 +20
有明确薪资信息 +10
有明确截止日期 +10
含排除关键词（实习/外包/派遣/仅限应届）-100

## 严格要求
- 正文不足100字时直接返回空jobs列表，不能编造
- 岗位名称、城市、薪资、截止日期必须来自原文
- 不确定的字段填空字符串，不要猜测

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


def analyze_with_kimi(article, qr_links):
    if not KIMI_API_KEY:
        print("  KIMI_API_KEY not set")
        return []
    content = article.get("content", "")
    if len(content) < 100:
        print("  content too short ({}), skip".format(len(content)))
        return []
    qr_info = ("\nQR links:\n" + "\n".join(qr_links)) if qr_links else ""
    user_msg = (
        "来源公众号：{}\n文章标题：{}\n发布日期：{}\n原文链接：{}\n\n"
        "文章正文：\n{}{}\n\n"
        "请分析以上招聘文章，提取所有符合前置条件（社会招聘+高中及以上）的岗位并打分。"
    ).format(
        article["source"], article["title"],
        article["pub_date"], article["url"],
        content, qr_info
    )
    try:
        resp = requests.post(
            "https://api.moonshot.cn/v1/chat/completions",
            headers={"Authorization": "Bearer {}".format(KIMI_API_KEY),
                     "Content-Type": "application/json"},
            json={
                "model": "moonshot-v1-32k",
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
                "temperature": 0.1,
                "max_tokens": 2048,
            },
            timeout=40,
        )
        raw = resp.json()["choices"][0]["message"]["content"]
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            return []
        jobs = json.loads(m.group()).get("jobs", [])
        for job in jobs:
            job["source"]      = article["source"]
            job["article_url"] = article["url"]
            job["pub_date"]    = article["pub_date"]
            job["qr_links"]    = qr_links
        return jobs
    except Exception as e:
        print("    Kimi failed: {}".format(e))
        return []


# ============================================================
# 6. Generate Excel + print log
# ============================================================
def generate_excel(all_jobs):
    valid = [j for j in all_jobs
             if not j.get("excluded") and j.get("score", 0) >= SCORE_THRESHOLD]
    valid.sort(key=lambda x: x.get("score", 0), reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "岗位清单_{}".format(date.today().strftime("%m%d"))

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
        "优先投递": PatternFill("solid", fgColor="D4EDDA"),
        "待复核":   PatternFill("solid", fgColor="FFF3CD"),
    }
    for ri, job in enumerate(valid, 2):
        score    = job.get("score", 0)
        priority = "优先投递" if score >= 80 else "待复核"
        link = (job.get("apply_link") or
                (job.get("qr_links") or [""])[0] or
                job.get("article_url", ""))
        row = [
            score, priority,
            job.get("job_name", ""), job.get("employer", ""),
            job.get("city", ""), job.get("education_req", ""),
            job.get("recruitment_type", ""), job.get("pub_date", ""),
            link, job.get("salary_info", ""), job.get("deadline", ""),
            job.get("summary", ""), job.get("article_url", ""),
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

    widths = [8, 10, 22, 20, 10, 10, 10, 10, 42, 15, 12, 36, 42, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # Print to log so visible in GitHub Actions
    print("\n" + "="*60)
    print("RESULT: {} jobs found".format(len(valid)))
    print("="*60)
    for job in valid:
        print("\n[{}] {} | {} | score={}".format(
            job.get("job_name",""), job.get("employer",""),
            job.get("city",""), job.get("score",0)))
        print("  edu={} | salary={} | deadline={}".format(
            job.get("education_req",""), job.get("salary_info","N/A"),
            job.get("deadline","N/A")))
        print("  {}".format(job.get("summary","")))
        print("  URL: {}".format(job.get("article_url","")))
    print("="*60 + "\n")

    path = "/tmp/招聘岗位_{}.xlsx".format(date.today().strftime("%Y%m%d"))
    wb.save(path)
    return path, valid


# ============================================================
# 7. Upload
# ============================================================
def upload_excel(path):
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
        print("upload failed: {}".format(e))
    return ""


# ============================================================
# 8. WeChat push
# ============================================================
def push_wechat(title, body):
    if not SERVERCHAN_KEY:
        print("SERVERCHAN_KEY not set")
        return
    try:
        resp = requests.post(
            "https://sctapi.ftqq.com/{}.send".format(SERVERCHAN_KEY),
            data={"title": title, "desp": body},
            timeout=12,
        )
        r = resp.json()
        print("push ok" if r.get("code") == 0 else "push: {}".format(r))
    except Exception as e:
        print("push failed: {}".format(e))


# ============================================================
# Main
# ============================================================
def main():
    start = datetime.now()
    print("\n" + "="*55)
    print("v4 start - {}".format(start.strftime("%Y-%m-%d %H:%M:%S")))
    print("source=Baidu | filter=social+highschool+ | priority=Shanxi")
    print("="*55 + "\n")

    all_articles = []
    for account in ACCOUNTS:
        print("fetch: {}".format(account))
        arts = fetch_articles_baidu(account)
        print("  got {} articles".format(len(arts)))
        all_articles.extend(arts)

    print("\ntotal articles: {}".format(len(all_articles)))

    if not all_articles:
        msg = "今日采集0篇文章，搜索引擎可能暂时不可用，明日自动重试。"
        print(msg)
        push_wechat("招聘采集异常 {}".format(date.today().strftime("%m/%d")), msg)
        return

    filtered = filter_by_title(all_articles)
    print("after title filter: {} articles\n".format(len(filtered)))

    all_jobs = []
    for i, art in enumerate(filtered, 1):
        print("[{}/{}] {}".format(i, len(filtered), art["title"][:40]))
        content, imgs = fetch_article_detail(art["url"])
        art["content"] = content
        qr_links = decode_qr(imgs)
        jobs = analyze_with_kimi(art, qr_links)
        valid_n = len([j for j in jobs if not j.get("excluded")])
        print("  jobs={} valid={}".format(len(jobs), valid_n))
        all_jobs.extend(jobs)
        time.sleep(2)

    valid_jobs, priority_jobs = [], []
    excel_path = dl_link = ""
    if all_jobs:
        excel_path, valid_jobs = generate_excel(all_jobs)
        priority_jobs = [j for j in valid_jobs if j.get("score", 0) >= 80]
        dl_link = upload_excel(excel_path)

    elapsed = (datetime.now() - start).seconds
    print("done: articles={} valid={} priority={} elapsed={}s".format(
        len(filtered), len(valid_jobs), len(priority_jobs), elapsed))

    lines = [
        "今日共采集 **{}** 篇招聘文章，筛选出 **{}** 个符合要求的岗位，其中优先投递 **{}** 个".format(
            len(filtered), len(valid_jobs), len(priority_jobs)),
        "\n> 筛选条件：社会招聘 | 高中及以上 | 太原/吕梁/山西优先",
        "\n耗时：{}秒\n".format(elapsed),
    ]
    if priority_jobs:
        lines += ["### 优先投递", ""]
        for job in priority_jobs[:8]:
            sal  = " | {}".format(job["salary_info"]) if job.get("salary_info") else ""
            edu  = " | {}".format(job["education_req"]) if job.get("education_req") else ""
            dead = " | 截止{}".format(job["deadline"]) if job.get("deadline") else ""
            lines.append("**{}** · {} · {}{}{}{}  · {}分  ".format(
                job.get("job_name",""), job.get("employer",""),
                job.get("city",""), sal, edu, dead, job.get("score",0)))
    review = [j for j in valid_jobs if j.get("score", 0) < 80]
    if review:
        lines += ["", "### 待复核", ""]
        for job in review[:5]:
            lines.append("**{}** · {} · {} · {}分  ".format(
                job.get("job_name",""), job.get("employer",""),
                job.get("city",""), job.get("score",0)))
    if dl_link:
        lines += ["", "[点击下载Excel]({})（7天有效）".format(dl_link)]
    elif valid_jobs:
        lines += ["", "Excel已上传至GitHub Actions Artifacts"]
    if not valid_jobs:
        lines += ["", "今日未发现符合条件的岗位。"]

    push_wechat(
        "招聘速报 {} · {}个岗位 · {}个优先".format(
            date.today().strftime("%m/%d"), len(valid_jobs), len(priority_jobs)),
        "\n".join(lines),
    )
    print("all done\n")


if __name__ == "__main__":
    main()
