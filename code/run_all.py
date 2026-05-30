#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_all.py — 一键复现完整分析流程（全部使用真实数据）
=====================================================
Usage:
  python run_all.py                # 使用已有真实数据运行完整分析 + 出图
  python run_all.py --fetch        # 重新从 API 拉取数据 (DHS + GHO + World Bank)
  python run_all.py --skip-figs    # 仅计算指标, 不出图

Data sources (all real, all API-retrieved):
  DHS Program Indicator Data API  — 烟草使用分层数据 (财富/教育/城乡/性别)
  WHO GHO OData API               — 国家层面无烟烟草 & 吸烟患病率
  World Bank API                  — GDP, Gini, 城镇化率等协变量

Output:
  output/figures/   — 6 张出版级图表 (PNG 300dpi + PDF 矢量)
  output/tables/    — 不平等指标结果表 (CSV)
  data/             — 原始拉取数据

Pipeline:
  1. 数据拉取 (fetch_all_real_data.py) — 仅 --fetch 模式
  2. 不平等指标计算 (analysis_real.py)
  3. 出版级图表生成 (figures_real.py)
"""

import sys, os, argparse

SRC = os.path.join(os.path.dirname(__file__), "src")
sys.path.insert(0, SRC)
ROOT = os.path.dirname(__file__)


def main():
    ap = argparse.ArgumentParser(description="口腔健康不平等 — 真实数据分析")
    ap.add_argument("--fetch", action="store_true",
                    help="重新从 API 拉取数据 (需联网, 约3-5分钟)")
    ap.add_argument("--skip-figs", action="store_true",
                    help="仅计算指标, 不出图")
    args = ap.parse_args()

    print("=" * 65)
    print("  口腔健康不平等 — 社会经济梯度分析")
    print("  ALL REAL DATA from DHS + WHO GHO + World Bank APIs")
    print("=" * 65)

    # ------- Step 1: Fetch (optional) -------
    if args.fetch:
        print("\n[Step 1] 从 API 拉取数据 ...\n")
        ret = os.system(f"{sys.executable} "
                        f"{os.path.join(SRC, 'fetch_all_real_data.py')}")
        if ret != 0:
            print("数据拉取出错, 尝试使用已有数据.")
    else:
        master_path = os.path.join(ROOT, "data", "analysis_master.csv")
        if os.path.exists(master_path):
            print(f"\n[Step 1] 使用已有数据: {master_path}")
        else:
            print("\n[Step 1] 未找到数据, 自动拉取 ...\n")
            os.system(f"{sys.executable} "
                      f"{os.path.join(SRC, 'fetch_all_real_data.py')}")

    # ------- Step 2: Analysis -------
    print("\n[Step 2] 计算不平等指标 (真实数据) ...\n")
    ret = os.system(f"{sys.executable} "
                    f"{os.path.join(SRC, 'analysis_real.py')}")
    if ret != 0:
        print("分析出错, 退出.")
        sys.exit(1)

    # ------- Step 3: Figures -------
    if not args.skip_figs:
        print("\n[Step 3] 生成出版级图表 (真实数据) ...\n")
        ret = os.system(f"{sys.executable} "
                        f"{os.path.join(SRC, 'figures_real.py')}")
        if ret != 0:
            print("图表生成出错.")

    print("\n" + "=" * 65)
    print("  完成.")
    print(f"  图表: output/figures/  (6 张, PNG + PDF)")
    print(f"  表格: output/tables/inequality_results.csv")
    print(f"  数据: data/analysis_master.csv")
    print("=" * 65)


if __name__ == "__main__":
    main()
