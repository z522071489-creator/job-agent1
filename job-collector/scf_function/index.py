# -*- coding: utf-8 -*-
"""
腾讯云函数 SCF 入口文件
负责：采集所有数据源 + 调用 Kimi AI 分析 + 返回结果

部署方式：将整个 scf_function/ 目录打包为 zip 上传到 SCF
运行时：Python 3.9+
超时：600秒（在 SCF 控制台设置）
内存：512 MB
"""

import json
import time
import random
import re
import os
import logging
from datetime import datetime, date
from typing import Optional

import requests
from bs4 import BeautifulSoup
from openai import OpenAI

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────
# 配置区（通过 SCF 环境变量注入，避免硬编码）
# ─────────────────────────────────────────
KIMI_API_KEY = os.environ.get("KIMI_API_KEY", "")
KIMI_MODEL = "moonshot-v1-32k"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# 要监控的微信公众号名称列表
WX_ACCOUNTS = [
    "国企指南", "三晋国企", "山西国企直聘", "国聘行动", "优聘计划",
    "晋招聘", "三晋招聘", "山西招聘汇总", "晋师招聘",
    "山西热门招聘", "太原人才大市场公司招聘找工作",
]

# ─────────────────────────────────────────
# 采集模块 1：搜狗微信搜索
# ─────────────────────────────────────────

def fetch_sogou_wx_account(account_name: str, max_articles: int = 3) -> list[dict]:
    """通过搜狗微信搜索采集指定公众号最新文章"""
    results = []
    try:
        # Step 1: 搜索公众号主页
        search_url = "https://weixin.sogou.com/weixin"
        params = {"type": "1", "query": account_name, "ie": "utf8"}
        resp = requests.get(search_url, params=params, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # 找到账号对应的链接（type=1 返回公众号列表）
        account_link = None
        items = soup.select(".news-box .news-list li")
        for item in items:
            title_el = item.select_one("a.account_name")
            if title_el and account_name in title_el.get_text(strip=True):
                account_link = title_el.get("href")
                break

        if not account_link:
            logger.warning(f"未找到公众号: {account_name}")
            return results

        if not account_link.startswith("http"):
            account_link = "https://weixin.sogou.com" + account_link

        time.sleep(random.uniform(2, 4))

        # Step 2: 进入公众号主页获取文章列表
        resp2 = requests.get(account_link, headers=HEADERS, timeout=15)
        resp2.raise_for_status()
        soup2 = BeautifulSoup(resp2.text, "html.parser")

        articles = soup2.select(".news-box .news-list li")[:max_articles]
        for art in articles:
            link_el = art.select_one("h3 a")
            date_el = art.select_one(".s-p")
            if not link_el:
                continue
            title = link_el.get_text(strip=True)
            url = link_el.get("href", "")
            if not url.startswith("http"):
                url = "https://weixin.sogou.com" + url
            pub_date = date_el.get_text(strip=True) if date_el else ""

            # Step 3: 抓取正文
            content = fetch_wx_article_content(url)
            results.append({
                "title": title,
                "url": url,
                "pub_date": pub_date,
                "content": content,
                "source": f"微信公众号·{account_name}",
            })
            time.sleep(random.uniform(3, 6))

    except Exception as e:
        logger.error(f"采集公众号 {account_name} 失败: {e}")
    return results


def fetch_wx_article_content(url: str) -> str:
    """抓取微信文章正文（搜狗中转链接 or mp.weixin.qq.com）"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20, allow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        # 微信文章正文容器
        content_div = soup.select_one("#js_content") or soup.select_one(".rich_media_content")
        if content_div:
            return content_div.get_text(separator="\n", strip=True)[:3000]
        return soup.get_text(separator="\n", strip=True)[:2000]
    except Exception as e:
        logger.error(f"抓取文章内容失败 {url}: {e}")
        return ""


# ─────────────────────────────────────────
# 采集模块 2：太原市公共就业服务中心
# ─────────────────────────────────────────

def fetch_taiyuan_rsj() -> list[dict]:
    """采集 rsj.taiyuan.gov.cn 最新招聘公告"""
    results = []
    base_url = "http://rsj.taiyuan.gov.cn"
    try:
        # 尝试招聘公告页面（路径根据实际页面调整）
        page_url = f"{base_url}/zsjy/zpxx/index.htm"
        resp = requests.get(page_url, headers=HEADERS, timeout=20)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")

        items = soup.select("ul.list li a, .news-list a, table td a")[:15]
        for item in items:
            title = item.get_text(strip=True)
            if len(title) < 5:
                continue
            href = item.get("href", "")
            if not href.startswith("http"):
                href = base_url + "/" + href.lstrip("/")
            results.append({
                "title": title,
                "url": href,
                "pub_date": str(date.today()),
                "content": fetch_page_content(href),
                "source": "太原市公共就业服务中心",
            })
            time.sleep(1)
    except Exception as e:
        logger.error(f"采集太原就业中心失败: {e}")
    return results


# ─────────────────────────────────────────
# 采集模块 3：国聘网
# ─────────────────────────────────────────

def fetch_iguopin() -> list[dict]:
    """采集国聘网山西/太原岗位"""
    results = []
    try:
        # 国聘网岗位搜索 API（非官方，以实际接口为准）
        api_url = "https://www.iguopin.com/api/position/list"
        params = {
            "cityName": "太原",
            "pageNum": 1,
            "pageSize": 20,
            "orderBy": "publishTime",
        }
        resp = requests.get(api_url, params=params, headers=HEADERS, timeout=15)
        data = resp.json()
        items = data.get("data", {}).get("list", []) or data.get("result", [])
        for item in items:
            results.append({
                "title": item.get("positionName", ""),
                "url": f"https://www.iguopin.com/position/{item.get('id', '')}",
                "pub_date": item.get("publishTime", str(date.today()))[:10],
                "content": json.dumps(item, ensure_ascii=False),
                "source": "国聘网",
            })
    except Exception as e:
        logger.error(f"采集国聘网失败，尝试备用方式: {e}")
        # 备用：直接抓页面
        try:
            page_url = "https://www.iguopin.com/position/list?cityName=太原"
            resp = requests.get(page_url, headers=HEADERS, timeout=20)
            soup = BeautifulSoup(resp.text, "html.parser")
            for card in soup.select(".position-item, .job-item")[:20]:
                title_el = card.select_one("h3, .position-name, .job-name")
                link_el = card.select_one("a")
                if title_el:
                    href = link_el.get("href", "") if link_el else ""
                    if href and not href.startswith("http"):
                        href = "https://www.iguopin.com" + href
                    results.append({
                        "title": title_el.get_text(strip=True),
                        "url": href,
                        "pub_date": str(date.today()),
                        "content": card.get_text(separator="\n", strip=True),
                        "source": "国聘网",
                    })
        except Exception as e2:
            logger.error(f"国聘网备用方式也失败: {e2}")
    return results


# ─────────────────────────────────────────
# 采集模块 4：山西省国资委
# ─────────────────────────────────────────

def fetch_shanxi_gzw() -> list[dict]:
    """采集山西省国资委招聘信息"""
    results = []
    base_url = "http://gzw.shanxi.gov.cn"
    try:
        # 信息公开/招聘公告
        page_url = f"{base_url}/zwgk/rsxx/zpgg/index.htm"
        resp = requests.get(page_url, headers=HEADERS, timeout=20)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for link in soup.select(".news-list a, ul.list li a, .xxgk-list a")[:15]:
            title = link.get_text(strip=True)
            if len(title) < 5:
                continue
            href = link.get("href", "")
            if not href.startswith("http"):
                href = base_url + "/" + href.lstrip("/")
            results.append({
                "title": title,
                "url": href,
                "pub_date": str(date.today()),
                "content": fetch_page_content(href),
                "source": "山西省国资委",
            })
            time.sleep(1)
    except Exception as e:
        logger.error(f"采集山西国资委失败: {e}")
    return results


# ─────────────────────────────────────────
# 采集模块 5：山西省人事考试网
# ─────────────────────────────────────────

def fetch_shanxi_rst() -> list[dict]:
    """采集山西省人事考试网招聘公告"""
    results = []
    base_url = "http://rst.shanxi.gov.cn"
    try:
        page_url = f"{base_url}/rsw/other/web/info/listSxRs.do?categoryId=190"
        resp = requests.get(page_url, headers=HEADERS, timeout=20)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for item in soup.select("table tr td a, .list-content a, ul.newsList li a")[:15]:
            title = item.get_text(strip=True)
            if len(title) < 5:
                continue
            href = item.get("href", "")
            if not href.startswith("http"):
                href = base_url + href if href.startswith("/") else base_url + "/" + href
            results.append({
                "title": title,
                "url": href,
                "pub_date": str(date.today()),
                "content": fetch_page_content(href),
                "source": "山西省人事考试网",
            })
            time.sleep(1)
    except Exception as e:
        logger.error(f"采集山西人事考试网失败: {e}")
    return results


# ─────────────────────────────────────────
# 辅助：通用页面正文抓取
# ─────────────────────────────────────────

def fetch_page_content(url: str) -> str:
    """抓取普通网页正文"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        main = soup.select_one(".content, .article, #content, #article, .TRS_Editor, .view") or soup.body
        if main:
            return main.get_text(separator="\n", strip=True)[:3000]
    except Exception as e:
        logger.error(f"抓取页面 {url} 失败: {e}")
    return ""


# ─────────────────────────────────────────
# AI 分析模块（Kimi）
# ─────────────────────────────────────────

SYSTEM_PROMPT = """你是招聘信息筛选助手。我会给你一篇文章内容（标题+正文），你需要：

1. 判断文章中是否包含招聘信息（若不包含，返回空列表）
2. 提取所有招聘岗位，每个岗位输出一个JSON对象
3. 严格按照筛选规则处理
4. 所有字段必须来自原文，不得编造

筛选规则：
✅ 保留：社会招聘（含校招+社招混合）、学历高中及以上、正式用工
❌ 排除：实习/兼职/劳务派遣/外包/纯应届生校招/销售保险/房产/贷款

打分规则（100分制）：
- 工作地点太原/吕梁/山西：+30
- 岗位类型匹配（销售/行政/交付/运营/管理/技术/客服/财务/人事）：+20
- 国企/央企/事业单位：+20
- 有明确薪资：+10
- 有明确截止日期：+10
- 含实习/兼职/劳务派遣/外包/保险销售/房产销售/贷款：-100（直接输出空）

输出格式（JSON数组，没有岗位则返回 []）：
[
  {
    "job_title": "岗位名称",
    "company": "招聘单位",
    "city": "工作城市",
    "education": "学历要求",
    "job_type": "招聘类型（社会招聘/校招+社招/其他）",
    "salary": "薪资（原文没有写null）",
    "deadline": "截止日期（原文没有写null）",
    "summary": "岗位简介（50字内）",
    "score": 打分数字,
    "exclude_reason": "排除原因（被排除才填，否则填null）"
  }
]"""


def analyze_with_kimi(article: dict) -> list[dict]:
    """调用 Kimi API 分析单篇文章，返回岗位列表"""
    if not KIMI_API_KEY:
        logger.error("KIMI_API_KEY 未配置")
        return []
    
    text = f"标题：{article['title']}\n\n来源：{article['source']}\n\n正文：\n{article['content']}"
    if len(text) > 28000:
        text = text[:28000] + "\n...(内容过长已截断)"

    try:
        client = OpenAI(api_key=KIMI_API_KEY, base_url="https://api.moonshot.cn/v1")
        response = client.chat.completions.create(
            model=KIMI_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": text},
            ],
            temperature=0.1,
            max_tokens=3000,
        )
        raw = response.choices[0].message.content.strip()
        # 提取 JSON 数组
        match = re.search(r"\[.*\]", raw, re.DOTALL)
        if match:
            jobs = json.loads(match.group())
            # 附加元信息
            for job in jobs:
                job["article_url"] = article["url"]
                job["pub_date"] = article.get("pub_date", "")
                job["source"] = article["source"]
            return jobs
    except json.JSONDecodeError as e:
        logger.error(f"Kimi 返回 JSON 解析失败: {e}\n原始输出: {raw[:300]}")
    except Exception as e:
        logger.error(f"Kimi API 调用失败: {e}")
    return []


# ─────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────

def run_collection() -> dict:
    """执行完整采集流程，返回结果字典"""
    all_articles = []
    stats = {
        "wx_count": 0,
        "website_count": 0,
        "job_count": 0,
        "priority_count": 0,
        "errors": [],
    }

    # 1. 采集微信公众号（搜狗，限速：每账号间隔5-10秒）
    logger.info("开始采集微信公众号...")
    for account in WX_ACCOUNTS:
        articles = fetch_sogou_wx_account(account, max_articles=3)
        all_articles.extend(articles)
        stats["wx_count"] += len(articles)
        logger.info(f"  {account}: {len(articles)} 篇")
        time.sleep(random.uniform(5, 10))

    # 2. 采集招聘网站
    logger.info("开始采集招聘网站...")
    for fetcher, name in [
        (fetch_taiyuan_rsj, "太原就业中心"),
        (fetch_iguopin, "国聘网"),
        (fetch_shanxi_gzw, "山西国资委"),
        (fetch_shanxi_rst, "山西人事考试网"),
    ]:
        try:
            arts = fetcher()
            all_articles.extend(arts)
            stats["website_count"] += len(arts)
            logger.info(f"  {name}: {len(arts)} 条")
        except Exception as e:
            stats["errors"].append(f"{name}: {str(e)}")
            logger.error(f"  {name} 失败: {e}")
        time.sleep(2)

    logger.info(f"采集完成：共 {len(all_articles)} 篇文章，开始 AI 分析...")

    # 3. AI 分析
    all_jobs = []
    for i, article in enumerate(all_articles):
        logger.info(f"  分析第 {i+1}/{len(all_articles)} 篇: {article['title'][:30]}")
        jobs = analyze_with_kimi(article)
        all_jobs.extend(jobs)
        time.sleep(0.5)  # Kimi 限速

    # 4. 过滤：排除被标记的岗位，保留分数>=50
    valid_jobs = []
    for job in all_jobs:
        if job.get("exclude_reason"):
            continue
        if job.get("score", 0) < 50:
            continue
        valid_jobs.append(job)

    # 5. 排序：分数从高到低
    valid_jobs.sort(key=lambda x: x.get("score", 0), reverse=True)

    stats["job_count"] = len(valid_jobs)
    stats["priority_count"] = sum(1 for j in valid_jobs if j.get("score", 0) >= 80)
    stats["total_articles"] = len(all_articles)

    logger.info(f"分析完成：{len(valid_jobs)} 个有效岗位，{stats['priority_count']} 个优先岗位")

    return {
        "jobs": valid_jobs,
        "stats": stats,
        "date": str(date.today()),
    }


# ─────────────────────────────────────────
# SCF 入口函数
# ─────────────────────────────────────────

def main_handler(event, context):
    """腾讯云函数入口"""
    logger.info("SCF 函数启动")
    try:
        result = run_collection()
        return {
            "statusCode": 200,
            "body": json.dumps(result, ensure_ascii=False),
        }
    except Exception as e:
        logger.error(f"执行失败: {e}")
        return {
            "statusCode": 500,
            "body": json.dumps({"error": str(e)}, ensure_ascii=False),
        }


# 本地调试入口
if __name__ == "__main__":
    result = run_collection()
    print(json.dumps(result["stats"], ensure_ascii=False, indent=2))
    print(f"共 {len(result['jobs'])} 个岗位")
