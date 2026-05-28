# 招聘信息采集系统 - 完整部署指南

## 系统架构

```
GitHub Actions（定时09:00）
    ↓ HTTP POST
腾讯云函数 SCF（国内IP）
    ├── 搜狗微信搜索（11个公众号）
    ├── 太原市公共就业服务中心
    ├── 国聘网
    ├── 山西省国资委
    └── 山西省人事考试网
    ↓ Kimi AI 分析+打分
GitHub Actions（生成Excel）
    ├── 上传到 GitHub Release（可下载）
    └── Server酱 → 微信推送
```

---

## 第一步：创建腾讯云函数 SCF

> 核心：让代码跑在国内服务器上，解决403问题

### 1.1 注册/登录腾讯云
访问 https://console.cloud.tencent.com/

### 1.2 开通云函数服务
控制台 → 搜索「云函数」→ 开通服务（免费，每月100万次）

### 1.3 创建函数
1. 进入「云函数 → 函数服务 → 新建」
2. 配置如下：
   - **创建方式**：自定义创建
   - **函数名称**：`job-collector`
   - **地域**：华北地区（北京）（离山西最近）
   - **运行环境**：Python 3.9
   - **函数代码**：选择「本地上传zip」（见下方打包说明）
   - **执行方法**：`index.main_handler`

3. **高级配置**：
   - 内存：512 MB
   - 超时：**600 秒**（必须设为最大值）
   - 环境变量：添加 `KIMI_API_KEY = 你的Key`

### 1.4 打包并上传代码

```bash
# 在本地执行：
cd scf_function/
pip install -r requirements.txt -t .
zip -r ../job-collector-scf.zip .
```

上传 `job-collector-scf.zip` 到 SCF 控制台。

### 1.5 配置 HTTP 触发器
1. SCF 函数详情 → 「触发管理 → 创建触发器」
2. 触发方式：**API网关**
3. 认证方式：**API密钥认证**（重要！防止被别人调用）
4. 创建后，记录「访问路径」，格式类似：
   ```
   https://service-xxxxxxxx.gz.apigw.tencentcs.com/release/
   ```

---

## 第二步：配置 GitHub Secrets

在你的 GitHub 仓库 → Settings → Secrets and variables → Actions → New repository secret

| Secret 名称 | 值 | 说明 |
|-------------|-----|------|
| `SCF_URL` | SCF触发器URL | 上一步复制的API网关地址 |
| `SERVERCHAN_KEY` | SCTxxxxxxxx... | Server酱 SendKey |

> `GITHUB_TOKEN` 不需要手动设置，Actions 自动注入。

---

## 第三步：推送代码到 GitHub

```bash
# 在项目根目录：
git init
git add .
git commit -m "feat: 招聘采集系统初始化"
git remote add origin https://github.com/你的用户名/你的仓库名.git
git push -u origin main
```

---

## 第四步：测试运行

1. GitHub 仓库 → Actions → 「每日招聘采集」
2. 点击「Run workflow」手动触发
3. 查看日志，确认：
   - SCF 调用成功
   - 返回岗位数据
   - Excel 生成成功
   - 微信收到推送

---

## 常见问题

### Q: SCF 超时了怎么办？
A: 减少采集的公众号数量，或者把微信公众号采集和网站采集拆成两个函数。

### Q: 搜狗返回验证码/429了？
A: 在 `index.py` 中增加 `time.sleep` 时间，或减少单次采集的公众号数量。
   搜狗限速较严，建议每个账号间隔 8-15 秒。

### Q: Kimi 余额不够了？
A: 可以修改 `KIMI_MODEL = "moonshot-v1-8k"` 使用更便宜的模型，
   或者将 AI 分析移到 GitHub Actions 端（非国内IP没问题）。

### Q: 某个网站访问失败？
A: 政府网站改版频繁，需要手动更新 `index.py` 中对应函数的 CSS 选择器。

### Q: 推送了但 Excel 下载不了？
A: GitHub Release 可能需要仓库设为 Public，或者在 Actions 里下载 Artifact。

---

## 项目结构

```
job-collector/
├── .github/
│   └── workflows/
│       └── daily_job.yml        # GitHub Actions 定时任务
├── scf_function/
│   ├── index.py                 # SCF 云函数主体（采集+AI分析）
│   └── requirements.txt         # SCF 依赖
├── scripts/
│   └── run.py                   # GitHub Actions 本地脚本（生成Excel+推送）
└── README.md                    # 本文档
```

---

## 扩展：微信公众号采集说明

系统使用**搜狗微信搜索** (`weixin.sogou.com`) 抓取公众号文章。

限制说明：
- 搜狗限速：建议每个账号间隔 5-10 秒
- 每天每个账号最多采集 3 篇最新文章
- 如遇大量 403 / 验证码，说明 IP 被临时封禁，次日自动恢复
- 腾讯云北京节点 IP 通常不在搜狗黑名单内

备用方案：
- 使用 RSSHub 公共实例 (rsshub.app) 订阅部分公众号
  例如：`https://rsshub.app/wechat/mp/[biz参数]`
