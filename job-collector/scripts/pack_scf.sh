#!/bin/bash
# 打包 SCF 函数代码的脚本
# 在项目根目录执行：bash scripts/pack_scf.sh

set -e
echo "📦 打包 SCF 函数代码..."

cd scf_function/
pip install -r requirements.txt -t . --quiet
zip -r ../job-collector-scf.zip . -x "*.pyc" -x "__pycache__/*"
cd ..

echo "✅ 已生成 job-collector-scf.zip"
echo "   文件大小：$(du -sh job-collector-scf.zip | cut -f1)"
echo ""
echo "下一步：上传此 zip 文件到腾讯云 SCF 控制台"
