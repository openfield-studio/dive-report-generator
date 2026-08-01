# -*- coding: utf-8 -*-
"""
潜水作業計画・記録（xls）の記入例を自動生成するツール。

減圧計算は実際の運用ファイル「潜水作業計画A4版.xlsm」の標準空気減圧表・
繰り返し潜水表シート（セル値を直接抽出、写真からの手動転記ではない）と、
同ファイルの実際の数式（入力データ(計画)シート）を再現している。
適用減圧深度は「その日の最大深度」ではなく、各回ごとに「深度が12m未満なら
12m、それ以外は3の倍数に切り上げ」という式で個別に計算する（実際のExcel
数式と同一）。表1データは深度12～42mの範囲（3の倍数の深度）を完全にカバー。
45m以上の深度、および繰返潜水グループ記号の「調整」ルール（前回との比較に
よるランクアップ等）は簡易近似で代用している。

【重要】上記の通り一部簡易近似を含む。実際の潜水作業計画・安全書類として
使用しないでください。あくまで「記入例」「練習用サンプル」の自動生成が目的です。

使い方:
    python dive_record_generator.py
でGUIが起動します。
"""

import datetime
import hashlib
import json
import math
import os
import random
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from tkinter import ttk, filedialog, messagebox

import jpholiday
import win32com.client as win32

import license_core

APP_VERSION = "1.0.0"
GITHUB_UPDATE_REPO = "openfield-studio/dive-report-generator"

TEMPLATE_NAME = "571d8290bfc37d2165393aa2.xls"

if getattr(sys, "frozen", False):
    # PyInstallerでexe化されている場合:
    # 設定・出力はexeと同じフォルダに、テンプレート等の同梱リソースはexe内部の展開先から読む。
    SCRIPT_DIR = os.path.dirname(sys.executable)
    RESOURCE_DIR = getattr(sys, "_MEIPASS", SCRIPT_DIR)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    RESOURCE_DIR = SCRIPT_DIR

TEMPLATE_PATH = os.path.join(RESOURCE_DIR, TEMPLATE_NAME)
SETTINGS_DIR = os.path.join(SCRIPT_DIR, "settings")
LICENSE_FILE = os.path.join(SCRIPT_DIR, "license.key")
TRIAL_FILE = os.path.join(SCRIPT_DIR, "trial.dat")


def _sanitize_filename(name: str) -> str:
    return "".join(c for c in name if c not in '\\/:*?"<>|').strip()


def settings_path_for(project_name: str) -> str:
    return os.path.join(SETTINGS_DIR, _sanitize_filename(project_name) + ".json")


_RESERVED_SETTINGS_NAMES = {"field_history"}


def list_saved_projects():
    if not os.path.isdir(SETTINGS_DIR):
        return []
    names = [
        os.path.splitext(f)[0]
        for f in os.listdir(SETTINGS_DIR)
        if f.lower().endswith(".json") and os.path.splitext(f)[0] not in _RESERVED_SETTINGS_NAMES
    ]
    return sorted(names)


def save_settings(project_name: str, data: dict):
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    with open(settings_path_for(project_name), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_settings(project_name: str) -> dict:
    with open(settings_path_for(project_name), "r", encoding="utf-8") as f:
        return json.load(f)


HISTORY_PATH = os.path.join(SETTINGS_DIR, "field_history.json")
HISTORY_MAX_ITEMS = 30


def load_history() -> dict:
    if not os.path.exists(HISTORY_PATH):
        return {}
    with open(HISTORY_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_history(data: dict):
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_to_history(history: dict, field_key: str, value: str):
    value = (value or "").strip()
    if not value:
        return
    items = history.setdefault(field_key, [])
    if value in items:
        items.remove(value)
    items.insert(0, value)
    del items[HISTORY_MAX_ITEMS:]


# Excel定数（win32com早期バインディング無しでも使えるよう直接数値指定）
XL_TYPE_PDF = 0  # xlTypePDF


def seireki_to_reiwa(year: int) -> int:
    return year - 2018


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


# ---------------------------------------------------------------------------
# 公式減圧表データ（実際の運用xlsm「潜水作業計画A4版.xlsm」の
#   標準空気減圧表／繰り返し潜水表 シートから直接抽出。写真からの手動転記ではなく
#   セル値をそのまま読み取っているため、深度12～42mの範囲は完全に正確）。
# 表1: 標準空気潜水減圧表（深度 12/15/18/21/24/27/30/33/36/39/42m の実データ）
#      各行 = (滞底時間(分), 繰返潜水グループ記号 or None, {停止深度(m): 停止時間(分), ...})
# 45m以上、および3m/6m/9m単独潜水（下記compute_applied_depth参照）は表に無いため
# 簡易近似にフォールバックする。
# ---------------------------------------------------------------------------
TABLE1 = {
    12: [
        (20, "A", {}), (30, "B", {}), (60, "D", {}), (90, "G", {}),
        (120, "H", {}), (150, "J", {}), (180, "M", {3: 5}),
        (200, None, {3: 10}), (210, None, {3: 15}), (220, None, {3: 19}), (240, None, {3: 26}),
        (270, None, {3: 35}), (300, None, {3: 44}), (330, None, {3: 53}), (360, None, {3: 62}),
    ],
    15: [
        (10, "A", {}), (20, "B", {}), (30, "C", {}), (40, "D", {}),
        (50, "E", {}), (60, "F", {}), (75, "G", {}), (100, "I", {3: 4}),
        (120, "K", {3: 9}), (125, "K", {3: 12}), (130, "L", {3: 15}), (140, "M", {3: 20}),
        (150, None, {3: 25}), (160, None, {3: 30}), (170, None, {3: 34}), (180, None, {3: 39}),
        (200, None, {3: 49}), (220, None, {3: 58}), (240, None, {3: 69}), (260, None, {3: 80}),
        (280, None, {3: 90}),
    ],
    18: [
        (10, "A", {}), (20, "B", {}), (30, "D", {}), (40, "E", {}),
        (50, "F", {}), (60, "G", {3: 4}), (80, "I", {3: 9}), (90, "J", {3: 15}),
        (100, "K", {3: 23}), (110, "L", {3: 29}), (120, "M", {3: 35}),
        (130, None, {6: 1, 3: 40}), (140, None, {6: 1, 3: 46}), (150, None, {6: 2, 3: 52}),
        (160, None, {6: 3, 3: 59}), (170, None, {6: 5, 3: 65}), (180, None, {6: 7, 3: 73}),
        (190, None, {6: 9, 3: 80}), (200, None, {6: 11, 3: 87}), (210, None, {6: 14, 3: 91}),
        (220, None, {6: 17, 3: 97}), (230, None, {6: 20, 3: 103}), (240, None, {6: 23, 3: 109}),
    ],
    21: [
        (10, "A", {}), (20, "C", {}), (25, "D", {}), (30, "D", {}), (35, "E", {}),
        (40, "F", {3: 4}), (50, "G", {3: 9}), (60, "H", {3: 11}), (70, "J", {6: 2, 3: 17}),
        (80, "K", {6: 3, 3: 25}), (90, "M", {6: 4, 3: 32}), (100, "N", {6: 5, 3: 39}),
        (110, None, {6: 7, 3: 46}), (120, None, {6: 9, 3: 54}), (130, None, {6: 13, 3: 62}),
        (140, None, {6: 16, 3: 71}), (150, None, {6: 19, 3: 77}), (160, None, {6: 22, 3: 85}),
        (170, None, {6: 27, 3: 93}), (180, None, {6: 31, 3: 101}), (190, None, {6: 35, 3: 109}),
        (200, None, {6: 38, 3: 117}),
    ],
    24: [
        (10, "A", {}), (15, "C", {}), (20, "D", {}), (25, "E", {}), (30, "F", {3: 3}),
        (40, "G", {3: 9}), (50, "H", {6: 3, 3: 11}), (55, "I", {6: 4, 3: 15}),
        (60, "J", {6: 5, 3: 21}), (65, "J", {6: 6, 3: 25}), (70, "K", {6: 6, 3: 30}),
        (75, "L", {6: 7, 3: 34}), (80, "M", {6: 8, 3: 37}),
        (85, None, {6: 10, 3: 42}), (90, None, {6: 12, 3: 46}), (95, None, {6: 14, 3: 50}),
        (100, None, {6: 15, 3: 55}), (110, None, {9: 1, 6: 20, 3: 64}), (120, None, {9: 2, 6: 23, 3: 72}),
        (130, None, {9: 3, 6: 27, 3: 82}), (140, None, {9: 4, 6: 31, 3: 93}),
        (150, None, {9: 6, 6: 36, 3: 104}), (160, None, {9: 8, 6: 39, 3: 114}),
    ],
    27: [
        (5, "A", {}), (10, "B", {}), (15, "C", {}), (20, "D", {}), (25, "E", {3: 5}),
        (30, "F", {3: 7}), (40, "H", {6: 4, 3: 10}), (45, "I", {6: 5, 3: 14}),
        (50, "J", {6: 6, 3: 20}), (55, "K", {6: 7, 3: 26}), (60, "L", {9: 1, 6: 8, 3: 31}),
        (65, None, {9: 2, 6: 8, 3: 36}), (70, None, {9: 2, 6: 11, 3: 40}), (75, None, {9: 3, 6: 13, 3: 46}),
        (80, None, {9: 3, 6: 15, 3: 51}), (85, None, {9: 4, 6: 16, 3: 56}), (90, None, {9: 4, 6: 19, 3: 60}),
        (95, None, {9: 5, 6: 22, 3: 64}), (100, None, {9: 6, 6: 24, 3: 70}), (110, None, {9: 9, 6: 27, 3: 82}),
        (120, None, {9: 11, 6: 32, 3: 95}),
    ],
    30: [
        (5, "A", {}), (10, "B", {}), (15, "D", {}), (20, "E", {3: 6}), (25, "F", {6: 1, 3: 9}),
        (30, "G", {6: 3, 3: 10}), (35, "H", {6: 5, 3: 11}), (40, "I", {6: 7, 3: 16}),
        (45, "J", {9: 1, 6: 8, 3: 23}), (50, "K", {9: 2, 6: 8, 3: 29}), (55, "L", {9: 3, 6: 9, 3: 34}),
        (60, None, {9: 4, 6: 10, 3: 40}), (65, None, {9: 4, 6: 14, 3: 46}), (70, None, {9: 5, 6: 16, 3: 52}),
        (75, None, {9: 6, 6: 18, 3: 56}), (80, None, {9: 7, 6: 22, 3: 61}), (85, None, {9: 9, 6: 23, 3: 67}),
        (90, None, {12: 1, 9: 11, 6: 25, 3: 75}), (95, None, {12: 2, 9: 12, 6: 27, 3: 82}),
        (100, None, {12: 2, 9: 14, 6: 31, 3: 90}), (105, None, {12: 2, 9: 15, 6: 34, 3: 98}),
        (110, None, {12: 3, 9: 17, 6: 38, 3: 105}),
    ],
    33: [
        (5, "A", {}), (10, "B", {}), (12, "C", {}), (15, "D", {3: 2}), (20, "F", {6: 1, 3: 9}),
        (25, "G", {6: 4, 3: 10}), (30, "H", {6: 7, 3: 10}), (35, "I", {9: 1, 6: 8, 3: 16}),
        (40, "J", {9: 3, 6: 8, 3: 24}), (45, "K", {9: 4, 6: 9, 3: 31}), (50, "M", {9: 5, 6: 9, 3: 38}),
        (55, "N", {9: 6, 6: 13, 3: 44}),
        (60, None, {9: 7, 6: 15, 3: 51}), (65, None, {12: 1, 9: 7, 6: 18, 3: 55}),
        (70, None, {12: 2, 9: 8, 6: 21, 3: 62}), (75, None, {12: 2, 9: 11, 6: 23, 3: 68}),
        (80, None, {12: 3, 9: 13, 6: 26, 3: 77}), (85, None, {12: 3, 9: 15, 6: 30, 3: 86}),
        (90, None, {12: 4, 9: 16, 6: 34, 3: 95}), (95, None, {12: 4, 9: 18, 6: 38, 3: 105}),
        (100, None, {12: 6, 9: 20, 6: 42, 3: 114}), (105, None, {12: 8, 9: 21, 6: 45, 3: 123}),
        (110, None, {12: 9, 9: 23, 6: 48, 3: 130}),
    ],
    36: [
        (5, "A", {}), (10, "C", {}), (15, "E", {3: 7}), (20, "F", {6: 2, 3: 10}),
        (25, "G", {6: 6, 3: 10}), (30, "I", {9: 2, 6: 8, 3: 14}), (35, "J", {9: 4, 6: 8, 3: 24}),
        (40, "K", {9: 6, 6: 8, 3: 32}), (45, "M", {12: 1, 9: 6, 6: 10, 3: 38}),
        (50, "N", {12: 2, 9: 7, 6: 13, 3: 46}),
        (55, None, {12: 3, 9: 7, 6: 16, 3: 53}), (60, None, {12: 4, 9: 7, 6: 19, 3: 59}),
        (65, None, {12: 4, 9: 10, 6: 22, 3: 66}), (70, None, {12: 5, 9: 13, 6: 27, 3: 75}),
        (75, None, {12: 6, 9: 15, 6: 31, 3: 86}), (80, None, {12: 7, 9: 17, 6: 35, 3: 97}),
        (85, None, {15: 1, 12: 8, 9: 18, 6: 40, 3: 107}), (90, None, {15: 1, 12: 10, 9: 20, 6: 42, 3: 118}),
        (95, None, {15: 2, 12: 11, 9: 22, 6: 46, 3: 128}), (100, None, {15: 2, 12: 13, 9: 24, 6: 50, 3: 136}),
    ],
    39: [
        (5, "A", {}), (8, "B", {}), (10, "C", {3: 2}), (15, "E", {6: 1, 3: 8}),
        (20, "G", {6: 5, 3: 10}), (25, "H", {9: 2, 6: 7, 3: 11}), (30, "J", {9: 4, 6: 8, 3: 22}),
        (35, "K", {12: 1, 9: 6, 6: 9, 3: 30}), (40, "M", {12: 2, 9: 7, 6: 9, 3: 39}),
        (45, "N", {12: 4, 9: 7, 6: 13, 3: 47}),
        (50, None, {12: 5, 9: 7, 6: 17, 3: 53}), (55, None, {12: 6, 9: 8, 6: 20, 3: 61}),
        (60, None, {15: 1, 12: 6, 9: 11, 6: 25, 3: 70}), (65, None, {15: 2, 12: 6, 9: 14, 6: 30, 3: 82}),
        (70, None, {15: 2, 12: 7, 9: 17, 6: 34, 3: 94}), (75, None, {15: 3, 12: 8, 9: 18, 6: 39, 3: 106}),
        (80, None, {15: 3, 12: 11, 9: 20, 6: 42, 3: 118}), (85, None, {15: 4, 12: 13, 9: 22, 6: 47, 3: 129}),
        (90, None, {15: 4, 12: 15, 9: 24, 6: 52, 3: 138}),
    ],
    42: [
        (7, "B", {}), (10, "D", {3: 4}), (15, "F", {6: 3, 3: 9}), (20, "G", {9: 1, 6: 7, 3: 10}),
        (25, "I", {9: 4, 6: 8, 3: 17}), (30, "K", {12: 1, 9: 6, 6: 8, 3: 28}),
        (35, "L", {12: 2, 9: 7, 6: 9, 3: 37}), (40, "N", {12: 4, 9: 7, 6: 12, 3: 46}),
        (45, "O", {15: 1, 12: 5, 9: 8, 6: 16, 3: 53}),
        (50, None, {15: 2, 12: 6, 9: 8, 6: 21, 3: 62}), (55, None, {15: 3, 12: 6, 9: 12, 6: 27, 3: 73}),
        (60, None, {15: 4, 12: 6, 9: 15, 6: 32, 3: 86}), (65, None, {15: 4, 12: 8, 9: 17, 6: 37, 3: 99}),
        (70, None, {15: 5, 12: 11, 9: 18, 6: 40, 3: 114}),
        (75, None, {18: 1, 15: 5, 12: 13, 9: 21, 6: 45, 3: 126}),
        (80, None, {18: 1, 15: 6, 12: 15, 9: 23, 6: 51, 3: 137}),
        (85, None, {18: 2, 15: 7, 12: 16, 9: 25, 6: 57, 3: 146}),
        (90, None, {18: 2, 15: 9, 12: 16, 9: 28, 6: 65, 3: 152}),
    ],
}
TABLE1_PRECISE_MAX_DEPTH = 42  # これ以上の深度(45m~)は簡易近似にフォールバック


def compute_applied_depth(depth_m: float) -> int:
    """適用減圧深度を求める。実際の運用ファイルの数式をそのまま再現:
    深度が12m未満なら12m。12m以上なら3の倍数ちょうどならそのまま、
    それ以外は3の倍数に切り上げる。（各回ごとに個別に計算する。日の最大深度ではない）"""
    if depth_m < 12:
        return 12
    if depth_m % 3 == 0:
        return int(depth_m)
    return int(3 * (depth_m // 3 + 1))


# 表4A: 繰り返し潜水ファクター(RF)表。行=繰返潜水グループ記号(RG)、列=水面待機時間(SI)帯。
# 各帯の下限（分）。最後の帯は 840~1339分、それ以降(22時間超)もRF=1.0として扱う。
TABLE4A_BANDS_MIN = [15, 30, 60, 90, 120, 180, 240, 360, 540, 720, 840]
TABLE4A = {
    "A": [1.4, 1.2, 1.1, 1.1, 1.1, 1.1, 1.1, 1.1, 1.1, 1.1, 1.0],
    "B": [1.5, 1.3, 1.2, 1.2, 1.2, 1.1, 1.1, 1.1, 1.1, 1.1, 1.0],
    "C": [1.6, 1.4, 1.3, 1.2, 1.2, 1.2, 1.1, 1.1, 1.1, 1.1, 1.0],
    "D": [1.8, 1.5, 1.4, 1.3, 1.3, 1.2, 1.2, 1.1, 1.1, 1.1, 1.0],
    "E": [1.9, 1.6, 1.5, 1.4, 1.3, 1.3, 1.2, 1.2, 1.1, 1.1, 1.0],
    "F": [2.0, 1.7, 1.6, 1.5, 1.4, 1.3, 1.3, 1.2, 1.1, 1.1, 1.0],
    "G": [None, 1.9, 1.7, 1.6, 1.5, 1.4, 1.3, 1.2, 1.1, 1.1, 1.0],
    "H": [None, None, 1.9, 1.7, 1.6, 1.5, 1.4, 1.3, 1.1, 1.1, 1.0],
    "I": [None, None, 2.0, 1.8, 1.7, 1.5, 1.4, 1.3, 1.1, 1.1, 1.0],
    "J": [None, None, None, 1.9, 1.8, 1.6, 1.5, 1.3, 1.2, 1.1, 1.0],
    "K": [None, None, None, 2.0, 1.9, 1.7, 1.5, 1.3, 1.2, 1.1, 1.0],
    "L": [None, None, None, None, 2.0, 1.7, 1.6, 1.4, 1.2, 1.1, 1.0],
    "M": [None, None, None, None, None, 1.8, 1.6, 1.4, 1.2, 1.1, 1.0],
    "N": [None, None, None, None, None, 1.9, 1.7, 1.4, 1.2, 1.1, 1.0],
    "O": [None, None, None, None, None, 2.0, 1.7, 1.4, 1.2, 1.1, 1.0],
}


def table1_lookup(applied_depth: int, ebt_min: float):
    """表1から (適用潜水時間, 繰返潜水グループ記号, 停止時間dict{深度:分}) を引く。
    applied_depth は compute_applied_depth() で計算済みの値をそのまま渡すこと。
    表の範囲(42m)を超える深度の場合は None を返す（呼び出し側で簡易近似にフォールバック）。"""
    rows = TABLE1.get(applied_depth)
    if not rows:
        return None
    for bt, rg, stops in rows:
        if ebt_min <= bt:
            return bt, rg, stops
    return rows[-1]


def table4a_lookup(rg: str, si_min: float) -> float:
    """表4Aから繰り返し潜水ファクター(RF)を引く。表にない(-)組み合わせの場合は
    より長い待機時間側の値（より安全側＝小さいRF）にフォールバックする。"""
    if si_min > 1339:  # 840~1339分の帯を超えたら待機十分としてRF=1.0
        return 1.0
    row = TABLE4A.get(rg)
    if not row:
        return 1.0
    si_min = max(si_min, TABLE4A_BANDS_MIN[0])
    idx = 0
    for i, lower in enumerate(TABLE4A_BANDS_MIN):
        if si_min >= lower:
            idx = i
    for j in range(idx, len(row)):
        if row[j] is not None:
            return row[j]
    return 1.0


def compute_one_entry(no, carried_rg, si, depth, bt, start_dt, no_deco_shallow=False):
    """1回分の潜水を、表1・表4Aに基づいて計算する。
    適用深度(表引きに使う深度)はこの回自身の深度から compute_applied_depth() で個別に求める
    （日の最大深度ではない。前回の結果には依存しない）。
    no_deco_shallow=True の場合、深度が3/6/9m（12m未満）の回は表1を使わず、
    常に無減圧（浮上停止なし、RGは前回のまま）として扱う。"""
    if no == 1:
        rf = 1.0  # 初期潜水（14時間以上の待機後）はファクター1.0
    else:
        rf = table4a_lookup(carried_rg, si)

    ebt = round(rf * bt)

    shallow_no_deco = no_deco_shallow and depth < 12
    if shallow_no_deco:
        applied_depth = depth
        applied_time = ebt
        stops = {}
        rg_out = carried_rg
    else:
        applied_depth = compute_applied_depth(depth)
        looked_up = table1_lookup(applied_depth, ebt)
        if looked_up is not None:
            applied_time, rg_out, stops = looked_up
        else:
            # 表1のデータ範囲(42m)を超える深度は簡易近似にフォールバック
            applied_time = int(math.ceil(ebt / 10.0) * 10)
            stop3 = 0 if ebt <= 95 else (5 if ebt <= 110 else 6)
            stops = {3: stop3} if stop3 else {}
            rg_out = None

    if rg_out is None:
        # 限界潜水範囲（表にRG記載なし）の場合は直前のRGを維持するフォールバック
        rg_out = carried_rg

    ascent_time = int(math.ceil(depth / 10.0))
    total_ascent = ascent_time + sum(stops.values())
    end_dt = start_dt + datetime.timedelta(minutes=bt)
    complete_dt = end_dt + datetime.timedelta(minutes=total_ascent)
    service_total = bt + total_ascent

    return {
        "no": no,
        "rg": carried_rg,
        "rf": rf,
        "si": si,
        "depth": depth,
        "start_dt": start_dt,
        "bt": bt,
        "end_dt": end_dt,
        "ebt": ebt,
        "applied_depth": applied_depth,
        "applied_time": applied_time,
        "rg_out": rg_out,
        "shallow_no_deco": shallow_no_deco,
        "stops": stops,
        "ascent_time": ascent_time,
        "total_ascent": total_ascent,
        "complete_dt": complete_dt,
        "service_total": service_total,
    }


def compute_dive_chain(start_dts, depths, bts, no_deco_shallow=False):
    """1～4回分の潜水を順番に計算する。繰返潜水グループ記号(RG)は前回の結果を次回に引き継ぐ。
    待機時間(SI)は「前回の浮上完了時刻」と「今回の潜降開始時刻」の実際の差から自動的に求める
    （潜降開始時刻の間隔＝待機時間、というのが実際の意味と一致するようにするため）。"""
    entries = []
    carried_rg = "A"
    prev_complete_dt = None
    for i in range(len(start_dts)):
        si = 0 if i == 0 else max(0, round((start_dts[i] - prev_complete_dt).total_seconds() / 60))
        e = compute_one_entry(i + 1, carried_rg, si, depths[i], bts[i], start_dts[i], no_deco_shallow)
        entries.append(e)
        carried_rg = e["rg_out"]
        prev_complete_dt = e["complete_dt"]
    return entries


def _make_day_seed(plan_date: datetime.date, params: dict, salt: str) -> int:
    """日付＋潜水スケジュールに関わるパラメータだけから決定的にシードを作る。
    これにより「計画のみ」と「実施込み」を別々に生成しても、同じ日付・同じ設定であれば
    計画側の数値が完全に一致する（実施用の乱数は別系統のsaltで完全に分離）。"""
    key_parts = [
        plan_date.isoformat(),
        str(params.get("base_times")),
        str(params.get("time_jitter")),
        str(params.get("base_depths")),
        str(params.get("depth_jitter")),
        str(params.get("base_bt")),
        str(params.get("bt_jitter")),
        str(params.get("seed", "")),
        salt,
    ]
    key = "|".join(key_parts)
    h = hashlib.sha256(key.encode("utf-8")).hexdigest()
    return int(h[:16], 16)


class DiveDay:
    """1日分の潜水記録データ"""

    def __init__(self, plan_date, exec_date, entries):
        self.plan_date = plan_date
        self.exec_date = exec_date
        self.entries = entries  # list of dict (1～4件)


def build_day(plan_date: datetime.date, params: dict, rng: random.Random) -> DiveDay:
    base_times = params["base_times"]  # ["08:30","10:30","13:00","15:30"]
    base_depths = params["base_depths"]  # [12,9,6,3]
    base_bt = params["base_bt"]
    bt_jitter = params["bt_jitter"]
    depth_jitter = params["depth_jitter"]
    time_jitter = params["time_jitter"]

    start_dts = []
    depths = []
    bts = []

    for i in range(len(base_times)):
        hh, mm = [int(x) for x in base_times[i].split(":")]
        jitter_min = rng.randint(-time_jitter, time_jitter) if time_jitter else 0
        start_dts.append(datetime.datetime(2000, 1, 1, hh, mm) + datetime.timedelta(minutes=jitter_min))

        depth = base_depths[i] + (rng.randint(-depth_jitter, depth_jitter) if depth_jitter else 0)
        depths.append(clamp(depth, 1, 40))

        bt = base_bt + (rng.randint(-bt_jitter, bt_jitter) if bt_jitter else 0)
        bts.append(clamp(bt, 5, 240))

    # 待機時間(SI)は潜降開始時刻の間隔から自動的に決まる（別入力は不要）。
    entries = compute_dive_chain(start_dts, depths, bts, params.get("no_deco_shallow", False))

    # 天候・風向・風速・波高・視界・透明度はすべて空欄のままにする（自動記入しない）。

    exec_date = plan_date + datetime.timedelta(days=1)
    return DiveDay(plan_date, exec_date, entries)


PLAN_ROW_BASE = 33  # rows 33-36 (計画)
JISSHI_ROW_BASE = 41  # rows 41-44 (実施)

# ○深度ごとの浮上停止時間 列（AF~AK＝18/15/12/9/6/3m）
STOP_DEPTH_COLS = {18: 32, 15: 33, 12: 34, 9: 35, 6: 36, 3: 37}


def fill_time_cell(ws, row, col_h, col_m, dt):
    ws.Cells(row, col_h).Value = dt.hour
    ws.Cells(row, col_m).Value = dt.minute


def fill_table_section(
    ws, row_base, day: DiveDay, jitter_actual: bool, rng: random.Random,
    deviation_pct: float = 5.0, depth_deviation_m: float = 1.0, bt_deviation_min: float = 3.0,
    no_deco_shallow: bool = False,
):
    # 列: C=回数 D=RG(通常/前回引継ぎ) F=SI H=RF J=深度 L,M=潜降開始(時,分)
    # P=BT R,S=浮上開始(時,分) T=EBT X=適用深度 Z=適用時間
    # AB=指定(RG結果) AC=調整 AD=浮上時間 AJ=停止(6m) AK=停止(3m) AL=合計 AN,AO=浮上完了(時,分)
    # AP=潜水業務時間合計 AR=累計
    if jitter_actual:
        # 実施は計画に対して、潜降開始時刻は±deviation_pct％、潜水時間(BT)は±bt_deviation_min分、
        # 深度は±depth_deviation_mメートルの幅でそれぞれずれる。
        # 表1・表4Aによる計算は実施側の値で独立に再計算する（待機時間SIも実施側の実際の間隔を使う）。
        start_dts, depths, bts = [], [], []
        for e in day.entries:
            pct = deviation_pct / 100.0
            time_frac = rng.uniform(-pct, pct)
            bt_dev = rng.uniform(-bt_deviation_min, bt_deviation_min)
            depth_dev = rng.uniform(-depth_deviation_m, depth_deviation_m)
            start_dts.append(e["start_dt"] + datetime.timedelta(minutes=round(e["bt"] * time_frac)))
            bts.append(clamp(round(e["bt"] + bt_dev), 5, 240))
            depths.append(clamp(round(e["depth"] + depth_dev), 1, 40))

        entries = compute_dive_chain(start_dts, depths, bts, no_deco_shallow)
    else:
        entries = day.entries

    cum = 0
    graph_entries = []
    for e in entries:
        r = row_base + (e["no"] - 1)

        rg_out = e["rg_out"]
        rg_adjust = rg_out if rng.random() > 0.15 else chr(clamp(ord(rg_out) + rng.choice([-1, 1]), ord("A"), ord("O")))
        cum += e["service_total"]

        ws.Cells(r, 3).Value = e["no"]
        # 深度3・6・9m常時無減圧の回は前回からの引継ぎ記号も未記入にする
        if not e["shallow_no_deco"]:
            ws.Cells(r, 4).Value = e["rg"]
        ws.Cells(r, 6).Value = e["si"]
        if not e["shallow_no_deco"]:
            ws.Cells(r, 8).Value = e["rf"]
        ws.Cells(r, 10).Value = e["depth"]
        fill_time_cell(ws, r, 12, 13, e["start_dt"])
        ws.Cells(r, 16).Value = e["bt"]
        fill_time_cell(ws, r, 18, 19, e["end_dt"])
        ws.Cells(r, 20).Value = e["ebt"]
        if e["no"] == 1:
            # テンプレートの1回目の行だけEBTセルの左半分(T列)にグレー塗りが
            # 残っており2～4回目には無い。不要な塗りなので消す。
            ws.Cells(r, 20).Interior.Pattern = -4142  # xlNone
        ws.Cells(r, 24).Value = e["applied_depth"]
        ws.Cells(r, 26).Value = e["applied_time"]
        # 深度3・6・9m常時無減圧の回は繰返潜水グループ記号を未記入にする
        if not e["shallow_no_deco"]:
            ws.Cells(r, 28).Value = rg_out
            ws.Cells(r, 29).Value = rg_adjust
        ws.Cells(r, 30).Value = e["ascent_time"]
        for stop_depth, col in STOP_DEPTH_COLS.items():  # AF~AK列（18/15/12/9/6/3m停止）
            v = e["stops"].get(stop_depth, 0)
            if v:
                ws.Cells(r, col).Value = v
        ws.Cells(r, 38).Value = e["total_ascent"]  # AL列 合計
        fill_time_cell(ws, r, 40, 41, e["complete_dt"])  # AN, AO
        ws.Cells(r, 42).Value = e["service_total"]  # AP列
        ws.Cells(r, 44).Value = cum  # AR列 累計

        graph_entries.append({
            "start": e["start_dt"], "end": e["end_dt"], "depth": e["depth"],
            "ascent_time": e["ascent_time"], "stops": e["stops"],
        })

    return graph_entries, entries[-1]


# 潜水深度グラフの軸校正: 固定ピクセル値ではなく、実際のワークシートの
# セル位置（列幅・行の高さ）から都度計算する。これにより、xls/xlsm保存時の
# 列幅の丸め方の違いなどで数字とグラフ線がズレることがなくなる。
#
# 【重要】このテンプレートの時刻軸は「1時間=1列」ではなく「1時間=3列
# （細かい点線区切り3本で1時間分、太い実線が1時間ごとの境界）」という構造。
# 以前は1時間=1列という誤った前提でC列(3)～Q列(17)の15列にしか数字を
# 書いていなかったため、実際のグリッド（C列～AS列付近まで）の左1/3に
# 数字とグラフ線が圧縮され、「時刻が左に偏る」不具合の原因になっていた。
# 5時はラベル("時刻")直後の単独1列（C列=3）、6時以降は3列ブロックの
# 中央列に数字を中央揃えで置く（実機の罫線パターンから実測して確認済み）。
_HOUR_ROW = 12        # 時刻軸の行
_SURFACE_ROW = 13     # 船上(0m)の行
_DEPTH_COL = 2         # 深度目盛りが入っている列(B)
_DEPTH_FIRST_ROW = 14  # 深度目盛りの最初の行
_DEPTH_LAST_ROW = 27   # 深度目盛りの最後の行


def _hour_label_col(hour: int) -> int:
    """時刻の数字を書き込む（＝グラフのx座標基準にする）列番号を返す。"""
    if hour == 5:
        return 3
    block_start = 4 + 3 * (hour - 6)
    return block_start + 1  # 3列ブロックの中央列


def _build_axis(ws):
    """このワークブックの実際のセル位置から、時刻→x座標・深度→y座標の変換関数を作る。"""
    x_by_hour = {}
    for hour in range(5, 20):
        col = _hour_label_col(hour)
        cell = ws.Cells(_HOUR_ROW, col)
        x_by_hour[hour] = cell.Left + cell.Width / 2.0

    y_by_depth = {0.0: ws.Cells(_SURFACE_ROW, _DEPTH_COL).Top}
    for r in range(_DEPTH_FIRST_ROW, _DEPTH_LAST_ROW + 1):
        v = ws.Cells(r, _DEPTH_COL).Value
        if v not in (None, ""):
            y_by_depth[float(v)] = ws.Cells(r, _DEPTH_COL).Top

    depths_sorted = sorted(y_by_depth.keys())
    hours_sorted = sorted(x_by_hour.keys())

    def time_to_x(dt: datetime.datetime) -> float:
        hour_frac = clamp(dt.hour + dt.minute / 60.0, hours_sorted[0], hours_sorted[-1])
        h0 = int(hour_frac)
        h1 = min(h0 + 1, hours_sorted[-1])
        if h0 not in x_by_hour:
            h0 = hours_sorted[0]
        if h1 not in x_by_hour:
            h1 = hours_sorted[-1]
        frac = hour_frac - h0
        x0, x1 = x_by_hour[h0], x_by_hour[h1]
        return x0 + (x1 - x0) * frac

    def depth_to_y(depth: float) -> float:
        depth = clamp(depth, depths_sorted[0], depths_sorted[-1])
        d0, d1 = depths_sorted[0], depths_sorted[-1]
        for i in range(len(depths_sorted) - 1):
            if depths_sorted[i] <= depth <= depths_sorted[i + 1]:
                d0, d1 = depths_sorted[i], depths_sorted[i + 1]
                break
        y0, y1 = y_by_depth[d0], y_by_depth[d1]
        frac = (depth - d0) / (d1 - d0) if d1 != d0 else 0.0
        return y0 + (y1 - y0) * frac

    return time_to_x, depth_to_y


def _rgb(r, g, b):
    return r + g * 256 + b * 65536


def draw_profile(ws, graph_entries, rgb_color, dashed=False):
    """潜降・浮上は瞬間ではなく、10m/分の速度で斜めに時間がかかる形で描画する
    （潜降時間＝浮上時間＝ascent_time分、という前提で往復とも同じ傾き）。
    減圧停止がある場合は、浮上の斜め線の途中に停止深度での水平区間（停止時間分）
    を挟む。各区間の所要時間はascent_time（表の「浮上時間」＝停止を除いた
    純粋な移動時間）を深度差の比で配分し、合計が浮上完了時刻（表のAN,AO列）と
    ずれないようにしている。"""
    time_to_x, depth_to_y = _build_axis(ws)
    y_surface = depth_to_y(0)
    for ge in graph_entries:
        slope_min = ge["ascent_time"]
        depth = ge["depth"]
        x_start = time_to_x(ge["start"])
        x_bottom_start = time_to_x(ge["start"] + datetime.timedelta(minutes=slope_min))
        x_bottom_end = time_to_x(ge["end"])
        y_depth = depth_to_y(depth)

        segments = [
            (x_start, y_surface, x_bottom_start, y_depth),
            (x_bottom_start, y_depth, x_bottom_end, y_depth),
        ]

        # 浮上区間: 深い停止深度から順に「斜めに上昇→その深度で水平停止」を繰り返し、
        # 最後に残りを水面まで斜めに上昇する。
        prev_depth = depth
        prev_time = ge["end"]
        for stop_depth in sorted(ge["stops"].keys(), reverse=True):
            if stop_depth >= prev_depth:
                continue
            travel_min = slope_min * (prev_depth - stop_depth) / depth if depth else 0
            arrive_time = prev_time + datetime.timedelta(minutes=travel_min)
            x1, y1 = time_to_x(prev_time), depth_to_y(prev_depth)
            x2, y2 = time_to_x(arrive_time), depth_to_y(stop_depth)
            segments.append((x1, y1, x2, y2))

            stop_min = ge["stops"][stop_depth]
            leave_time = arrive_time + datetime.timedelta(minutes=stop_min)
            x3 = time_to_x(leave_time)
            segments.append((x2, y2, x3, y2))

            prev_depth = stop_depth
            prev_time = leave_time

        travel_min = slope_min * prev_depth / depth if depth else 0
        surface_time = prev_time + datetime.timedelta(minutes=travel_min)
        x1, y1 = time_to_x(prev_time), depth_to_y(prev_depth)
        x2 = time_to_x(surface_time)
        segments.append((x1, y1, x2, y_surface))

        for x1, y1, x2, y2 in segments:
            shp = ws.Shapes.AddLine(x1, y1, x2, y2)
            shp.Line.ForeColor.RGB = rgb_color
            shp.Line.Weight = 2.25  # 実際の運用ファイルのグラフ線と同じ太さ
            if dashed:
                shp.Line.DashStyle = 4  # msoLineDash


def fill_hour_axis(ws):
    """潜水深度グラフ上部の時刻軸（5～19時）にサンプル同様の数字を入れる。
    テンプレート自体にはこの軸の数字が入っていないため、ここで補う。
    列は_hour_label_col()で決定し、_build_axis()と同じ列・同じ「セル中央」を
    基準にすることで、数字とグラフ線が必ず一致するようにしている。
    また、テンプレートの「9時」セルだけフォントサイズが他と異なる（8pt）など
    セルごとにバラつきがあるため、ここで全セル明示的に統一する。"""
    for hour in range(5, 20):
        col = _hour_label_col(hour)
        cell = ws.Cells(_HOUR_ROW, col)
        cell.Value = hour
        cell.HorizontalAlignment = -4108  # xlCenter
        cell.Font.Size = 9


def fix_depth_axis_alignment(ws):
    """深度目盛り（3,6,9…列B）はテンプレートで下揃え(xlVAlignBottom)になっており、
    数字が実際の行の上端（＝_build_axisがそのままy座標に使う位置）より
    行の高さ分（約23pt）下にずれて表示される。これがグラフの水深と数字が
    ズレて見える原因なので、上揃えに直す。"""
    for r in range(_DEPTH_FIRST_ROW, _DEPTH_LAST_ROW + 1):
        cell = ws.Cells(r, _DEPTH_COL)
        if cell.Value not in (None, ""):
            cell.VerticalAlignment = -4160  # xlVAlignTop


# 「設備等の点検表」の各項目名と、○/／を書き込むセル位置。
# 罫線を実測したところ、項目名の表示領域(はみ出し込みで元々1つの枠として
# 罫線で囲まれている範囲)と、実際にチェックを書き込む「□」の枠(項目名の
# 枠とは別に単独の罫線で囲まれたセル)は別物で、□は項目名の枠のすぐ右
# 隣にある。項目名セルは複数列結合して確実に収まる幅を確保し、□の枠に
# マークを書く。
# (項目名, 行, 項目名の結合開始列, 項目名の結合終了列, □の枠の列)
EQUIPMENT_ITEMS = [
    ("潜水器", 36, 47, 50, 51),
    ("送気管", 37, 47, 50, 51),
    ("さがり綱", 38, 47, 50, 51),
    ("圧力調整器", 39, 47, 50, 51),
    ("空気圧縮機", 40, 47, 50, 51),
    ("空気清浄装置", 41, 47, 50, 51),
    ("水中時計", 42, 47, 50, 51),
    ("流量計", 43, 47, 50, 51),
    ("通話装置の感度", 44, 47, 50, 51),
    ("携行物", 45, 47, 50, 51),
    ("小型船舶操縦士免許", 36, 52, 56, 57),
    ("潜水士免許", 37, 52, 56, 57),
    ("送気員教育修了証", 38, 52, 56, 57),
    ("巻上げ機教育修了証", 39, 52, 56, 57),
    ("国際信号旗A", 40, 52, 56, 57),
]


def fill_equipment_checklist(ws, checks: dict):
    """設備等の点検表にチェック状態を反映する。チェックあり＝○、なし＝／。"""
    for name, row, name_col_start, name_col_end, mark_col in EQUIPMENT_ITEMS:
        name_range = ws.Range(ws.Cells(row, name_col_start), ws.Cells(row, name_col_end))
        name_range.Merge()
        mark_cell = ws.Cells(row, mark_col)
        mark_cell.HorizontalAlignment = -4108  # xlCenter
        mark_cell.Value = "○" if checks.get(name, True) else "／"


def fill_header(ws, day: DiveDay, const_fields: dict, prev_day_info: dict = None):
    plan = day.plan_date
    exe = day.exec_date

    ws.Cells(2, 2).Value = "令和"
    ws.Cells(2, 3).Value = seireki_to_reiwa(plan.year)
    ws.Cells(2, 5).Value = plan.month
    ws.Cells(2, 7).Value = plan.day

    ws.Cells(2, 10).Value = "令和"
    ws.Cells(2, 11).Value = seireki_to_reiwa(exe.year)
    ws.Cells(2, 13).Value = exe.month
    ws.Cells(2, 15).Value = exe.day

    # テンプレートに残っている「（平成」表記を「（令和」に変更する（計画変更日欄）
    ws.Cells(3, 1).Value = "（令和"

    # 会社名・現場名など（任意・全日共通の固定値。空欄なら未入力のまま）
    if const_fields.get("jigyousha"):
        ws.Cells(4, 4).Value = const_fields["jigyousha"]
    if const_fields.get("kouji"):
        ws.Cells(4, 15).Value = const_fields["kouji"]
    if const_fields.get("motouke"):
        ws.Cells(5, 15).Value = const_fields["motouke"]
    if const_fields.get("basho"):
        # このテンプレートは「作業港・場所」の値セルだけ他の行(工事件名等、O~AA列/342pt)より
        # 狭く(O~T列/157.5pt)なっている。本物のxlsm(潜水作業計画A3版)ではこの行も他行と
        # 同じO~AA列幅なので、それに合わせて結合幅を広げてから中央揃えする
        # （狭いままだと中央揃えでも他行と開始位置がズレて見える）。
        basho_range = ws.Range(ws.Cells(6, 15), ws.Cells(6, 27))
        basho_range.Merge()
        basho_range.Value = const_fields["basho"]
        basho_range.HorizontalAlignment = -4108  # xlCenter
    if const_fields.get("funamei"):
        ws.Cells(7, 15).Value = const_fields["funamei"]
        ws.Cells(7, 15).HorizontalAlignment = -4108  # xlCenter
    if const_fields.get("sensuishi"):
        ws.Cells(5, 4).Value = const_fields["sensuishi"]
    # 職務・潜水通信方式は、実際の運用ファイルでは常に同じ定型文が入っている
    # （未入力ならこの既定文言を使う。ユーザー指定があればそちらを優先）
    ws.Cells(6, 4).Value = const_fields.get("shokumu") or "潜水作業：管理者・指揮者・潜水士"
    # テンプレートのこのセルだけフォントサイズが10ptで他の欄(11pt)より小さいため統一する
    ws.Cells(6, 4).Font.Size = 11
    if const_fields.get("sokiin"):
        ws.Cells(7, 4).Value = const_fields["sokiin"]
    if const_fields.get("renrakuin"):
        ws.Cells(8, 4).Value = const_fields["renrakuin"]
    # 「潜水・通信方式」も同様に他行より狭い(O~T列)ため、O~AA列に結合幅を広げる
    tsushin_range = ws.Range(ws.Cells(8, 15), ws.Cells(8, 27))
    tsushin_range.Merge()
    tsushin_range.Value = const_fields.get("tsushin") or "ヘルメット・フーカー・スクーバー，有線・無線・未使用"
    tsushin_range.HorizontalAlignment = -4108  # xlCenter
    # 責任者名・記入者名のセルは1列(26pt)しか幅が無く、中央揃えにすると
    # 左右均等にはみ出そうとして左側の隣接セル(ラベル)にぶつかり文字の
    # 先頭が欠けてしまう。右側(52,53列)は空欄なので3列分を結合して
    # 十分な幅を確保してから中央揃えする。
    if const_fields.get("sekininsha"):
        rng = ws.Range(ws.Cells(6, 51), ws.Cells(6, 53))
        rng.Merge()
        rng.Value = const_fields["sekininsha"]
        rng.HorizontalAlignment = -4108  # xlCenter
    if const_fields.get("kinyuusha"):
        rng = ws.Range(ws.Cells(7, 51), ws.Cells(7, 53))
        rng.Merge()
        rng.Value = const_fields["kinyuusha"]
        rng.HorizontalAlignment = -4108  # xlCenter

    # 天候・風向・風速・波高・視界・透明度はすべて空欄のままにする（自動記入しない）。

    # 前日の最終浮上時刻・最終繰返潜水記号RG（1日目など前日データが無い場合は空欄のまま）
    if prev_day_info:
        last_complete_dt = prev_day_info.get("last_complete_dt")
        last_rg = prev_day_info.get("last_rg")
        if last_complete_dt is not None:
            ws.Cells(5, 35).Value = last_complete_dt.hour
            ws.Cells(5, 37).Value = last_complete_dt.minute
        if last_rg:
            ws.Cells(6, 35).Value = last_rg

    gyoumu_start = const_fields.get("gyoumu_start", "").strip()
    gyoumu_end = const_fields.get("gyoumu_end", "").strip()

    if gyoumu_start:
        hh, mm = [int(x) for x in gyoumu_start.split(":")]
        ws.Cells(7, 35).Value = hh
        ws.Cells(7, 37).Value = mm
    else:
        first_start = day.entries[0]["start_dt"]
        work_start = first_start - datetime.timedelta(minutes=90)
        ws.Cells(7, 35).Value = work_start.hour
        ws.Cells(7, 37).Value = work_start.minute

    if gyoumu_end:
        hh, mm = [int(x) for x in gyoumu_end.split(":")]
        ws.Cells(8, 35).Value = hh
        ws.Cells(8, 37).Value = mm
    else:
        last_complete = day.entries[-1]["complete_dt"]
        work_end = last_complete + datetime.timedelta(minutes=30)
        ws.Cells(8, 35).Value = work_end.hour
        ws.Cells(8, 37).Value = work_end.minute


def _com_retry(fn, attempts=5, delay_sec=1.5):
    """Excel COM呼び出しは起動直後やファイルロック解放待ちで一時的に失敗することがあるため、
    数回リトライしてから諦める。"""
    last_err = None
    for _ in range(attempts):
        try:
            return fn()
        except Exception as ex:
            last_err = ex
            time.sleep(delay_sec)
    raise last_err


def _open_workbook_with_retry(excel, path):
    return _com_retry(lambda: excel.Workbooks.Open(path))


def _generate_one_file(excel, out_dir, plan_date, day, plan_rng, params, include_actual_this_file, prev_day_info=None):
    """1ファイル分（計画のみ、または計画+実施）を生成して保存する。
    (base_name, この日の最終潜水の情報{last_complete_dt, last_rg}) を返す。"""
    wb = _open_workbook_with_retry(excel, TEMPLATE_PATH)
    try:
        ws = wb.Worksheets(1)

        fill_hour_axis(ws)
        fix_depth_axis_alignment(ws)
        # 「計画変更時」欄（未使用）の1回目もEBTセルに同じ不要なグレー塗りが
        # 残っているので、他の行と同様に消しておく。
        ws.Cells(PLAN_ROW_BASE + 4, 20).Interior.Pattern = -4142  # xlNone
        fill_header(ws, day, params.get("const_fields", {}), prev_day_info)
        plan_graph, plan_last_entry = fill_table_section(ws, PLAN_ROW_BASE, day, jitter_actual=False, rng=plan_rng)
        last_entry_for_carry = plan_last_entry

        jisshi_graph = None
        if include_actual_this_file:
            actual_rng = random.Random(_make_day_seed(plan_date, params, "jisshi"))
            deviation_pct = params.get("jisshi_deviation_pct", 5.0)
            depth_deviation_m = params.get("jisshi_depth_deviation_m", 1.0)
            bt_deviation_min = params.get("jisshi_bt_deviation_min", 3.0)
            jisshi_graph, jisshi_last_entry = fill_table_section(
                ws, JISSHI_ROW_BASE, day, jitter_actual=True, rng=actual_rng,
                deviation_pct=deviation_pct, depth_deviation_m=depth_deviation_m,
                bt_deviation_min=bt_deviation_min,
                no_deco_shallow=params.get("no_deco_shallow", False),
            )
            last_entry_for_carry = jisshi_last_entry  # 実施ありなら実施の方が実際の結果に近い
            fill_equipment_checklist(ws, params.get("equipment_checks", {}))

        if params.get("draw_graph", True):
            draw_profile(ws, plan_graph, _rgb(0, 0, 0))  # 計画・・・黒色
            if include_actual_this_file:
                draw_profile(ws, jisshi_graph, _rgb(255, 0, 0), dashed=True)  # 実施・・・赤色

        kouji = params.get("const_fields", {}).get("kouji", "").strip()
        safe_kouji = _sanitize_filename(kouji)
        name_part = f"{safe_kouji}_" if safe_kouji else ""
        jisshi_part = "実施_" if include_actual_this_file else ""
        base_name = f"潜水作業計画記録_{jisshi_part}{name_part}{plan_date.strftime('%Y%m%d')}"
        xls_path = os.path.join(out_dir, base_name + ".xls")
        xlsm_path = os.path.join(out_dir, base_name + ".xlsm")
        pdf_path = os.path.join(out_dir, base_name + ".pdf")

        export_xls = params.get("export_xls", True)
        export_xlsm = params.get("export_xlsm", False)
        export_pdf = params.get("export_pdf", True)

        for p in (xls_path, xlsm_path, pdf_path):
            if os.path.exists(p):
                os.remove(p)

        if export_xls:
            _com_retry(lambda: wb.SaveAs(xls_path, FileFormat=56))  # 56 = xlExcel8 (.xls)
        if export_xlsm:
            _com_retry(lambda: wb.SaveAs(xlsm_path, FileFormat=52))  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        if export_pdf:
            _com_retry(lambda: ws.ExportAsFixedFormat(XL_TYPE_PDF, pdf_path))
    finally:
        wb.Close(False)

    return base_name, last_entry_for_carry


def generate(params: dict, progress_cb=None, cancel_check=None):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"テンプレートが見つかりません: {TEMPLATE_PATH}")

    out_dir = params["out_dir"]
    os.makedirs(out_dir, exist_ok=True)

    include_actual = params.get("include_actual", False)

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    include_saturday = params.get("include_saturday", False)
    include_sunday = params.get("include_sunday", False)
    include_holiday = params.get("include_holiday", False)
    num_days = params["num_days"]

    try:
        generated = 0
        offset = 0
        max_offset = num_days * 3 + 60  # 無限ループ防止の安全キャップ
        prev_day_info = None  # 初日は前日の記入なし。以後は「実際に生成した直近の日」の最終潜水情報を引き継ぐ（土日祝をまたいでも維持）。
        while generated < num_days and offset <= max_offset:
            if cancel_check and cancel_check():
                break
            plan_date = params["start_date"] + datetime.timedelta(days=offset)
            offset += 1
            if plan_date.weekday() == 5 and not include_saturday:  # 土曜日
                continue
            if plan_date.weekday() == 6 and not include_sunday:  # 日曜日
                continue
            if jpholiday.is_holiday(plan_date) and not include_holiday:  # 祭日
                continue
            generated += 1
            i = generated - 1

            # 計画側は日付＋スケジュール設定のみから決定的に生成する。
            # 「計画のみ」と「実施込み」を別々に生成しても、同じ設定なら計画の数値は一致する。
            plan_rng = random.Random(_make_day_seed(plan_date, params, "plan"))
            day = build_day(plan_date, params, plan_rng)

            # 実施を含める場合は「計画のみ」と「計画+実施」の2ファイルを両方出力する。
            # 両ファイルとも同じ「前日」情報を使う。
            base_name, last_entry = _generate_one_file(excel, out_dir, plan_date, day, plan_rng, params, False, prev_day_info)
            if include_actual:
                base_name, last_entry = _generate_one_file(excel, out_dir, plan_date, day, plan_rng, params, True, prev_day_info)

            # 深度3・6・9m常時無減圧の回は表内のRG表示も空欄にしているため、
            # その回が最終回だった場合は翌日の「最終繰返潜水記号RG」も空欄にする。
            prev_day_info = {
                "last_complete_dt": last_entry["complete_dt"],
                "last_rg": None if last_entry["shallow_no_deco"] else last_entry["rg_out"],
            }

            if progress_cb:
                progress_cb(i + 1, params["num_days"], base_name)
    finally:
        excel.Quit()


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root, trial_remaining=None):
        self.root = root
        self.trial_remaining = trial_remaining
        root.title("潜水作業計画・記録 記入例ジェネレーター")
        root.geometry("1180x760")

        pad = {"padx": 8, "pady": 4}

        # ボタン・ステータス・注意書きは下部に固定（ウィンドウが縦に伸びても常に見える）
        bottom_bar = ttk.Frame(root)
        bottom_bar.pack(side="bottom", fill="x")

        # 上部は左右2カラムのスクロール不要レイアウトにして縦の長さを抑える
        columns = ttk.Frame(root)
        columns.pack(side="top", fill="both", expand=True)
        left = ttk.Frame(columns)
        right = ttk.Frame(columns)
        left.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=8, anchor="n")
        right.pack(side="left", fill="both", expand=True, padx=(4, 8), pady=8, anchor="n")

        self.history = load_history()
        self.common_field_widgets = {}
        self.common_field_labels = {}

        def make_label(parent, text, row):
            ttk.Label(parent, text=text).grid(row=row, column=0, sticky="w", **pad)

        def make_common_field(parent, row, field_key, label_text, width=26, default_values=None):
            make_label(parent, label_text, row)
            var = tk.StringVar(value="")
            defaults = list(default_values or [])
            values = defaults + [v for v in self.history.get(field_key, []) if v not in defaults]
            combo = ttk.Combobox(parent, textvariable=var, values=values, width=width)
            combo.grid(row=row, column=1, sticky="w", **pad)
            self.common_field_widgets[field_key] = combo
            self.common_field_labels[field_key] = label_text
            return var

        # ------------------------------------------------------------------
        # 左カラム: スケジュール設定
        # ------------------------------------------------------------------
        lr = 0

        make_label(left, "保存済みの工事設定", lr)
        f0 = ttk.Frame(left)
        f0.grid(row=lr, column=1, sticky="w", **pad)
        self.project_choice = tk.StringVar(value="")
        self.project_combo = ttk.Combobox(
            f0, textvariable=self.project_choice, values=list_saved_projects(), width=20, state="readonly"
        )
        self.project_combo.pack(side="left")
        ttk.Button(f0, text="読み込む", command=self.on_load_settings).pack(side="left", padx=4)
        lr += 1

        today = datetime.date.today()

        make_label(left, "開始日（計画日・西暦）", lr)
        self.start_year = tk.StringVar(value=str(today.year))
        self.start_month = tk.StringVar(value=str(today.month))
        self.start_day = tk.StringVar(value=str(today.day))
        f = ttk.Frame(left)
        f.grid(row=lr, column=1, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.start_year, width=5).pack(side="left")
        ttk.Label(f, text="年").pack(side="left")
        ttk.Entry(f, textvariable=self.start_month, width=3).pack(side="left")
        ttk.Label(f, text="月").pack(side="left")
        ttk.Entry(f, textvariable=self.start_day, width=3).pack(side="left")
        ttk.Label(f, text="日").pack(side="left")
        lr += 1

        make_label(left, "生成日数", lr)
        self.num_days = tk.StringVar(value="5")
        ttk.Entry(left, textvariable=self.num_days, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "生成対象に含める日（チェックなしは除外）", lr)
        f_days = ttk.Frame(left)
        f_days.grid(row=lr, column=1, sticky="w", **pad)
        self.include_saturday = tk.BooleanVar(value=False)
        self.include_sunday = tk.BooleanVar(value=False)
        self.include_holiday = tk.BooleanVar(value=False)
        ttk.Checkbutton(f_days, text="土曜日", variable=self.include_saturday).pack(side="left")
        ttk.Checkbutton(f_days, text="日曜日", variable=self.include_sunday).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(f_days, text="祭日", variable=self.include_holiday).pack(side="left", padx=(8, 0))
        lr += 1

        self.include_actual = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            left,
            text="実施を含めて出力する（チェック時は「計画のみ」「計画+実施」の2ファイルを出力）",
            variable=self.include_actual,
        ).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        self.no_deco_shallow = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            left,
            text="深度3・6・9mは常に無減圧（浮上停止なし）として出力する",
            variable=self.no_deco_shallow,
        ).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "各回の潜降開始時刻（1～4回, カンマ区切り）", lr)
        self.base_times = tk.StringVar(value="08:30,10:30,13:00,15:30")
        ttk.Entry(left, textvariable=self.base_times, width=26).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "潜降開始時刻のばらつき（±分）", lr)
        self.time_jitter = tk.StringVar(value="5")
        ttk.Entry(left, textvariable=self.time_jitter, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "各回の潜水深度（潜降開始時刻と同じ数, m, カンマ区切り）", lr)
        self.base_depths = tk.StringVar(value="12,9,6,3")
        ttk.Entry(left, textvariable=self.base_depths, width=26).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "深度のばらつき（±m）", lr)
        self.depth_jitter = tk.StringVar(value="1")
        ttk.Entry(left, textvariable=self.depth_jitter, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "潜水時間 BT（分）", lr)
        self.base_bt = tk.StringVar(value="60")
        ttk.Entry(left, textvariable=self.base_bt, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "潜水時間のばらつき（±分）", lr)
        self.bt_jitter = tk.StringVar(value="5")
        ttk.Entry(left, textvariable=self.bt_jitter, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "実施の潜降開始時刻のばらつき（計画に対して±％）", lr)
        self.jisshi_deviation_pct = tk.StringVar(value="5")
        ttk.Entry(left, textvariable=self.jisshi_deviation_pct, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "実施の潜水時間のばらつき（計画に対して±分）", lr)
        self.jisshi_bt_deviation_min = tk.StringVar(value="3")
        ttk.Entry(left, textvariable=self.jisshi_bt_deviation_min, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        make_label(left, "実施の水深のばらつき（計画に対して±ｍ）", lr)
        self.jisshi_depth_deviation_m = tk.StringVar(value="1")
        ttk.Entry(left, textvariable=self.jisshi_depth_deviation_m, width=6).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        ttk.Label(
            left,
            text="※ 待機時間(SI)は各回の潜降開始時刻の間隔から自動計算されます（入力不要）",
            foreground="gray30",
        ).grid(row=lr, column=1, sticky="w", **pad)
        lr += 1

        ttk.Separator(left, orient="horizontal").grid(row=lr, column=0, columnspan=2, sticky="ew", pady=8)
        lr += 1

        make_label(left, "出力フォルダ", lr)
        f2 = ttk.Frame(left)
        f2.grid(row=lr, column=1, sticky="w", **pad)
        default_out_dir = os.path.join(os.path.expanduser("~"), "Documents", "潜水作業計画記録")
        self.out_dir = tk.StringVar(value=default_out_dir)
        ttk.Entry(f2, textvariable=self.out_dir, width=26).pack(side="left")
        ttk.Button(f2, text="参照...", command=self.browse_dir).pack(side="left", padx=4)
        lr += 1

        make_label(left, "出力形式", lr)
        f3 = ttk.Frame(left)
        f3.grid(row=lr, column=1, sticky="w", **pad)
        self.export_xls = tk.BooleanVar(value=False)
        self.export_xlsm = tk.BooleanVar(value=False)
        self.export_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(f3, text="XLS", variable=self.export_xls).pack(side="left")
        ttk.Checkbutton(f3, text="XLSM", variable=self.export_xlsm).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(f3, text="PDF", variable=self.export_pdf).pack(side="left", padx=(12, 0))
        lr += 1

        self.draw_graph = tk.BooleanVar(value=True)
        ttk.Checkbutton(left, text="潜水プロファイルの折れ線グラフを描画する（計画=黒, 実施=赤破線）", variable=self.draw_graph).grid(
            row=lr, column=1, sticky="w", **pad
        )
        lr += 1

        # ------------------------------------------------------------------
        # 右カラム: 全日共通項目
        # ------------------------------------------------------------------
        rr = 0
        ttk.Label(right, text="全日共通項目（任意）", font=("", 10, "bold")).grid(
            row=rr, column=0, columnspan=2, sticky="w", **pad
        )
        rr += 1

        self.jigyousha = make_common_field(right, rr, "jigyousha", "事業者名"); rr += 1
        self.kouji = make_common_field(right, rr, "kouji", "工事件名"); rr += 1
        self.motouke = make_common_field(right, rr, "motouke", "元請会社名"); rr += 1
        self.basho = make_common_field(right, rr, "basho", "作業港・場所"); rr += 1
        self.funamei = make_common_field(right, rr, "funamei", "船名"); rr += 1
        self.sensuishi = make_common_field(right, rr, "sensuishi", "潜水士名"); rr += 1
        self.shokumu = make_common_field(
            right, rr, "shokumu", "職務",
            default_values=["潜水作業：管理者", "潜水作業：指揮者", "潜水作業：潜水士"],
        ); rr += 1
        self.sokiin = make_common_field(right, rr, "sokiin", "送気員"); rr += 1
        self.renrakuin = make_common_field(right, rr, "renrakuin", "連絡員"); rr += 1

        make_label(right, "潜水・通信方式", rr)
        f_hoshiki = ttk.Frame(right)
        f_hoshiki.grid(row=rr, column=1, sticky="w", **pad)
        sensui_defaults = ["ヘルメット", "フーカー", "スクーバ"]
        self.sensui_hoshiki = tk.StringVar(value="")
        sensui_values = sensui_defaults + [v for v in self.history.get("sensui_hoshiki", []) if v not in sensui_defaults]
        sensui_combo = ttk.Combobox(f_hoshiki, textvariable=self.sensui_hoshiki, values=sensui_values, width=10)
        sensui_combo.pack(side="left")
        self.common_field_widgets["sensui_hoshiki"] = sensui_combo
        self.common_field_labels["sensui_hoshiki"] = "潜水方式"

        ttk.Label(f_hoshiki, text="通信方式").pack(side="left", padx=(10, 2))
        tsushin_defaults = ["有線", "無線", "未使用"]
        self.tsushin_hoshiki = tk.StringVar(value="")
        tsushin_values = tsushin_defaults + [v for v in self.history.get("tsushin_hoshiki", []) if v not in tsushin_defaults]
        tsushin_combo = ttk.Combobox(f_hoshiki, textvariable=self.tsushin_hoshiki, values=tsushin_values, width=10)
        tsushin_combo.pack(side="left")
        self.common_field_widgets["tsushin_hoshiki"] = tsushin_combo
        self.common_field_labels["tsushin_hoshiki"] = "通信方式"
        rr += 1
        self.gyoumu_start = make_common_field(right, rr, "gyoumu_start", "業務開始時刻（HH:MM, 空欄で自動）", width=10); rr += 1
        self.gyoumu_end = make_common_field(right, rr, "gyoumu_end", "業務完了時刻（HH:MM, 空欄で自動）", width=10); rr += 1
        self.kinyuusha = make_common_field(right, rr, "kinyuusha", "記入者名"); rr += 1
        self.sekininsha = make_common_field(right, rr, "sekininsha", "責任者名"); rr += 1

        f_common_btns = ttk.Frame(right)
        f_common_btns.grid(row=rr, column=1, sticky="w", **pad)
        ttk.Button(f_common_btns, text="共通項目を保存（履歴に追加）", command=self.on_save_common_fields).pack(
            side="left"
        )
        ttk.Button(f_common_btns, text="履歴を編集・削除...", command=self.on_manage_history).pack(
            side="left", padx=(6, 0)
        )
        rr += 1

        ttk.Separator(right, orient="horizontal").grid(row=rr, column=0, columnspan=2, sticky="ew", pady=8)
        rr += 1
        ttk.Label(right, text="設備等の点検表（実施を含める場合のみ反映。チェック=○ 未チェック=／）", font=("", 10, "bold")).grid(
            row=rr, column=0, columnspan=2, sticky="w", **pad
        )
        rr += 1

        eq_frame = ttk.Frame(right)
        eq_frame.grid(row=rr, column=0, columnspan=2, sticky="w", **pad)
        rr += 1
        self.equipment_vars = {}
        for i, item in enumerate(EQUIPMENT_ITEMS):
            name = item[0]
            var = tk.BooleanVar(value=True)
            self.equipment_vars[name] = var
            ttk.Checkbutton(eq_frame, text=name, variable=var).grid(
                row=i % 8, column=i // 8, sticky="w", padx=4, pady=1
            )

        # ------------------------------------------------------------------
        # 下部バー: 生成ボタン・ステータス・注意書き（常に画面下に固定表示）
        # ------------------------------------------------------------------
        ttk.Separator(bottom_bar, orient="horizontal").pack(fill="x")

        btn_row = ttk.Frame(bottom_bar)
        btn_row.pack(fill="x", pady=8)
        self.gen_button = ttk.Button(btn_row, text="生成", command=self.on_generate)
        self.gen_button.pack(side="left", padx=8)
        self.cancel_requested = False
        self.stop_button = ttk.Button(btn_row, text="停止", command=self.on_stop_generate, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        ttk.Button(btn_row, text="設定を保存（工事件名で保存）", command=self.on_save_settings).pack(side="left", padx=8)
        ttk.Button(btn_row, text="このPCのID / ライセンス認証", command=self.on_show_license_info).pack(side="left", padx=8)

        self.trial_label_var = tk.StringVar(value="")
        self.trial_label = ttk.Label(bottom_bar, textvariable=self.trial_label_var, foreground="#b36b00")
        self.trial_label.pack(anchor="w", padx=8, pady=(4, 0))
        self._update_trial_label()

        self.status = tk.StringVar(value="")
        ttk.Label(bottom_bar, textvariable=self.status, foreground="blue").pack(anchor="w", padx=8)

        note = "注意: 実際の潜水作業の安全書類としては使用しないでください（記入例・サンプル用途専用）。"
        ttk.Label(bottom_bar, text=note, foreground="gray30", justify="left").pack(anchor="w", padx=8, pady=(0, 8))

        # 保存・読み込み対象のフィールド一覧（工事件名(kouji)がファイル名のキーになる）
        self.field_vars = {
            "num_days": self.num_days,
            "include_saturday": self.include_saturday,
            "include_sunday": self.include_sunday,
            "include_holiday": self.include_holiday,
            "include_actual": self.include_actual,
            "no_deco_shallow": self.no_deco_shallow,
            "jisshi_deviation_pct": self.jisshi_deviation_pct,
            "jisshi_bt_deviation_min": self.jisshi_bt_deviation_min,
            "jisshi_depth_deviation_m": self.jisshi_depth_deviation_m,
            "base_times": self.base_times,
            "time_jitter": self.time_jitter,
            "base_depths": self.base_depths,
            "depth_jitter": self.depth_jitter,
            "base_bt": self.base_bt,
            "bt_jitter": self.bt_jitter,
            "jigyousha": self.jigyousha,
            "kouji": self.kouji,
            "motouke": self.motouke,
            "basho": self.basho,
            "funamei": self.funamei,
            "sensuishi": self.sensuishi,
            "shokumu": self.shokumu,
            "sokiin": self.sokiin,
            "renrakuin": self.renrakuin,
            "sensui_hoshiki": self.sensui_hoshiki,
            "tsushin_hoshiki": self.tsushin_hoshiki,
            "gyoumu_start": self.gyoumu_start,
            "gyoumu_end": self.gyoumu_end,
            "kinyuusha": self.kinyuusha,
            "sekininsha": self.sekininsha,
            "out_dir": self.out_dir,
            "export_xls": self.export_xls,
            "export_xlsm": self.export_xlsm,
            "export_pdf": self.export_pdf,
            "draw_graph": self.draw_graph,
        }

        # 起動時に、画面からはみ出さない範囲で全項目が一目で見えるサイズに自動調整する。
        # update_idletasks()だけだとgridレイアウトの実寸がまだ確定していないことが
        # あるため、update()で確実に確定させてから測る。余白も少し足しておく。
        root.update()
        req_w = root.winfo_reqwidth() + 20
        req_h = root.winfo_reqheight() + 20
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        win_w = min(req_w, screen_w - 40)
        win_h = min(req_h, screen_h - 80)
        root.geometry(f"{win_w}x{win_h}")

        self._check_for_update_async()

    def _check_for_update_async(self):
        def worker():
            info = fetch_latest_release_info()
            self.root.after(0, self._on_update_check_result, info)

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_check_result(self, info):
        if not info or not info.get("tag_name"):
            return
        latest = info["tag_name"]
        if _version_tuple(latest) <= _version_tuple(APP_VERSION):
            return
        if messagebox.askyesno(
            "アップデートのお知らせ",
            f"新しいバージョン {latest} が公開されています（現在: v{APP_VERSION}）。\n"
            "ダウンロードページを開きますか？",
        ):
            webbrowser.open(info["html_url"])

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def _record_common_field_history(self):
        """全日共通項目に今入っている値を、それぞれのフィールドの履歴に追加してドロップダウンへ反映する。"""
        for field_key, combo in self.common_field_widgets.items():
            value = combo.get().strip()
            if value:
                add_to_history(self.history, field_key, value)
        save_history(self.history)
        for field_key, combo in self.common_field_widgets.items():
            combo["values"] = self.history.get(field_key, [])

    def on_save_common_fields(self):
        self._record_common_field_history()
        self.status.set("全日共通項目をドロップダウンの履歴に保存しました。次回起動時も選択できます。")

    def on_manage_history(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("共通項目の履歴を編集・削除")
        dlg.geometry("420x380")

        ttk.Label(dlg, text="項目を選択:").pack(anchor="w", padx=8, pady=(8, 2))
        field_choice = tk.StringVar()
        field_combo = ttk.Combobox(
            dlg, textvariable=field_choice, state="readonly",
            values=[self.common_field_labels[k] for k in self.common_field_widgets],
        )
        field_combo.pack(fill="x", padx=8)

        listbox = tk.Listbox(dlg)
        listbox.pack(fill="both", expand=True, padx=8, pady=8)

        label_to_key = {v: k for k, v in self.common_field_labels.items()}

        def refresh_listbox(*_args):
            listbox.delete(0, tk.END)
            key = label_to_key.get(field_choice.get())
            if key:
                for item in self.history.get(key, []):
                    listbox.insert(tk.END, item)

        field_combo.bind("<<ComboboxSelected>>", refresh_listbox)
        if self.common_field_labels:
            field_choice.set(next(iter(self.common_field_labels.values())))
            refresh_listbox()

        def delete_selected():
            key = label_to_key.get(field_choice.get())
            sel = listbox.curselection()
            if not key or not sel:
                return
            value = listbox.get(sel[0])
            items = self.history.get(key, [])
            if value in items:
                items.remove(value)
            save_history(self.history)
            if key in self.common_field_widgets:
                self.common_field_widgets[key]["values"] = self.history.get(key, [])
            refresh_listbox()

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btn_frame, text="選択した項目を削除", command=delete_selected).pack(side="left")
        ttk.Button(btn_frame, text="閉じる", command=dlg.destroy).pack(side="right")

    def _update_trial_label(self):
        if _check_saved_license():
            self.trial_label_var.set("ライセンス認証済み")
        elif self.trial_remaining is not None:
            self.trial_label_var.set(
                f"試用期間中：あと{self.trial_remaining}日でライセンス認証が必要になります（コピー直後から{license_core.TRIAL_DAYS}日間）"
            )
        else:
            self.trial_label_var.set("")

    def on_show_license_info(self):
        if _show_activation_dialog(self.root):
            messagebox.showinfo("ライセンス認証", "認証が完了しました。")
        self._update_trial_label()

    def on_save_settings(self):
        kouji = self.kouji.get().strip()
        if not kouji:
            messagebox.showerror("保存エラー", "「工事件名」を入力してから保存してください（工事件名が保存名になります）。")
            return
        data = {k: v.get() for k, v in self.field_vars.items()}
        data["equipment_checks"] = {name: var.get() for name, var in self.equipment_vars.items()}
        try:
            save_settings(kouji, data)
        except Exception as ex:
            messagebox.showerror("保存エラー", str(ex))
            return
        self._record_common_field_history()
        self.project_combo["values"] = list_saved_projects()
        self.project_choice.set(kouji)
        messagebox.showinfo("保存しました", f"「{kouji}」の設定として保存しました。\n"
                             "工事件名を変更してから保存すると、別の設定として新規保存されます。")

    def on_load_settings(self):
        name = self.project_choice.get().strip()
        if not name:
            messagebox.showerror("読み込みエラー", "保存済みの工事設定を選択してください。")
            return
        try:
            data = load_settings(name)
        except Exception as ex:
            messagebox.showerror("読み込みエラー", str(ex))
            return
        for k, v in data.items():
            if k in self.field_vars:
                self.field_vars[k].set(v)
        for eq_name, checked in data.get("equipment_checks", {}).items():
            if eq_name in self.equipment_vars:
                self.equipment_vars[eq_name].set(checked)
        self.status.set(f"「{name}」の設定を読み込みました。")

    def on_generate(self):
        try:
            params = self.collect_params()
        except Exception as ex:
            messagebox.showerror("入力エラー", str(ex))
            return

        self._record_common_field_history()

        self.cancel_requested = False
        self.gen_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status.set("生成中...")
        self.root.update_idletasks()

        def progress_cb(done, total, name):
            self.status.set(f"生成中... {done}/{total} ({name})")
            self.root.update()  # 停止ボタンのクリックを生成中でも受け付けるため

        try:
            generate(params, progress_cb=progress_cb, cancel_check=lambda: self.cancel_requested)
            if self.cancel_requested:
                self.status.set(f"停止しました。ここまでの出力先: {params['out_dir']}")
            else:
                self.status.set(f"完了しました。出力先: {params['out_dir']}")
                if messagebox.askyesno("完了", "生成が完了しました。出力フォルダを開きますか？"):
                    os.startfile(params["out_dir"])
        except Exception as ex:
            messagebox.showerror("生成エラー", str(ex))
            self.status.set("エラーが発生しました。")
        finally:
            self.gen_button.config(state="normal")
            self.stop_button.config(state="disabled")

    def on_stop_generate(self):
        self.cancel_requested = True
        self.status.set("停止中...（キリのよいところで停止します）")

    def collect_params(self):
        start_date = datetime.date(
            int(self.start_year.get()), int(self.start_month.get()), int(self.start_day.get())
        )
        base_times = [t.strip() for t in self.base_times.get().split(",") if t.strip() != ""]
        if not (1 <= len(base_times) <= 4):
            raise ValueError("潜降開始時刻は1～4個をカンマ区切りで入力してください（テンプレートの表が4回分までのため）")

        depth_parts = [p.strip() for p in self.base_depths.get().split(",") if p.strip() != ""]
        if len(depth_parts) != len(base_times):
            raise ValueError(f"潜水深度パターンは潜降開始時刻と同じ{len(base_times)}個を入力してください")
        base_depths = [int(p) for p in depth_parts]

        if not self.export_xls.get() and not self.export_xlsm.get() and not self.export_pdf.get():
            raise ValueError("出力形式は XLS・XLSM・PDF のうち少なくとも1つを選択してください")

        return {
            "start_date": start_date,
            "num_days": int(self.num_days.get()),
            "base_times": base_times,
            "time_jitter": int(self.time_jitter.get()),
            "base_depths": base_depths,
            "depth_jitter": int(self.depth_jitter.get()),
            "base_bt": int(self.base_bt.get()),
            "bt_jitter": int(self.bt_jitter.get()),
            "out_dir": self.out_dir.get(),
            "include_saturday": self.include_saturday.get(),
            "include_sunday": self.include_sunday.get(),
            "include_holiday": self.include_holiday.get(),
            "include_actual": self.include_actual.get(),
            "no_deco_shallow": self.no_deco_shallow.get(),
            "jisshi_deviation_pct": float(self.jisshi_deviation_pct.get()),
            "jisshi_bt_deviation_min": float(self.jisshi_bt_deviation_min.get()),
            "jisshi_depth_deviation_m": float(self.jisshi_depth_deviation_m.get()),
            "export_xls": self.export_xls.get(),
            "export_xlsm": self.export_xlsm.get(),
            "export_pdf": self.export_pdf.get(),
            "draw_graph": self.draw_graph.get(),
            "equipment_checks": {name: var.get() for name, var in self.equipment_vars.items()},
            "const_fields": {
                "jigyousha": self.jigyousha.get().strip(),
                "kouji": self.kouji.get().strip(),
                "motouke": self.motouke.get().strip(),
                "basho": self.basho.get().strip(),
                "funamei": self.funamei.get().strip(),
                "sensuishi": self.sensuishi.get().strip(),
                "shokumu": self.shokumu.get().strip(),
                "sokiin": self.sokiin.get().strip(),
                "renrakuin": self.renrakuin.get().strip(),
                "tsushin": "，".join(
                    v for v in (self.sensui_hoshiki.get().strip(), self.tsushin_hoshiki.get().strip()) if v
                ),
                "gyoumu_start": self.gyoumu_start.get().strip(),
                "gyoumu_end": self.gyoumu_end.get().strip(),
                "kinyuusha": self.kinyuusha.get().strip(),
                "sekininsha": self.sekininsha.get().strip(),
            },
        }


def _check_saved_license() -> bool:
    if not os.path.exists(LICENSE_FILE):
        return False
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            saved_key = f.read().strip()
    except OSError:
        return False
    machine_id = license_core.get_machine_id()
    return license_core.verify_license_key(machine_id, saved_key)


def _show_activation_dialog(root) -> bool:
    """このPC用のライセンスキー入力画面を表示する。認証成功でTrueを返す。"""
    machine_id = license_core.get_machine_id()
    result = {"ok": False}

    dialog = tk.Toplevel(root)
    dialog.title("ライセンス認証")
    dialog.resizable(False, False)
    dialog.grab_set()
    dialog.protocol("WM_DELETE_WINDOW", lambda: on_cancel())

    pad = {"padx": 12, "pady": 6}

    tk.Label(dialog, text="このPCは未認証です。下記の「このPCのID」を開発元に伝えて\n発行された「ライセンスキー」を入力してください。",
              justify="left").grid(row=0, column=0, columnspan=2, sticky="w", **pad)

    tk.Label(dialog, text="このPCのID").grid(row=1, column=0, sticky="w", **pad)
    id_entry = tk.Entry(dialog, width=48)
    id_entry.insert(0, machine_id)
    id_entry.config(state="readonly")
    id_entry.grid(row=1, column=1, **pad)

    def copy_id():
        root.clipboard_clear()
        root.clipboard_append(machine_id)

    tk.Button(dialog, text="コピー", command=copy_id).grid(row=1, column=2, **pad)

    tk.Label(dialog, text="ライセンスキー").grid(row=2, column=0, sticky="w", **pad)
    key_entry = tk.Entry(dialog, width=48)
    key_entry.grid(row=2, column=1, **pad)
    key_entry.focus_set()

    status_label = tk.Label(dialog, text="", fg="red")
    status_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=12)

    def on_activate():
        entered = key_entry.get().strip()
        if license_core.verify_license_key(machine_id, entered):
            try:
                with open(LICENSE_FILE, "w", encoding="utf-8") as f:
                    f.write(entered)
            except OSError as e:
                status_label.config(text=f"保存に失敗しました: {e}")
                return
            result["ok"] = True
            dialog.destroy()
        else:
            status_label.config(text="ライセンスキーが正しくありません。")

    def on_cancel():
        result["ok"] = False
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.grid(row=4, column=0, columnspan=3, pady=10)
    tk.Button(btn_frame, text="認証", width=12, command=on_activate).pack(side="left", padx=6)
    tk.Button(btn_frame, text="終了", width=12, command=on_cancel).pack(side="left", padx=6)

    dialog.bind("<Return>", lambda e: on_activate())
    dialog.update_idletasks()
    dialog.wait_window(dialog)
    return result["ok"]


def _version_tuple(v: str):
    v = v.strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def fetch_latest_release_info():
    """GitHubの最新リリース情報を取得する。
    ネットワークエラー・リリース未作成等の場合は None を返す（起動をブロックしない）。"""
    url = f"https://api.github.com/repos/{GITHUB_UPDATE_REPO}/releases/latest"
    req = urllib.request.Request(url, headers={"User-Agent": "DiveReportGenerator"})
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return {"tag_name": data.get("tag_name", ""), "html_url": data.get("html_url", "")}
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, ValueError, OSError):
        return None


def ensure_licensed(root):
    """(起動してよいか, 試用期間の残り日数 or None) を返す。
    残り日数がNoneなのは正式ライセンス認証済みの場合。"""
    if _check_saved_license():
        return True, None
    remaining = license_core.get_trial_remaining_days(TRIAL_FILE)
    if remaining > 0:
        return True, remaining
    ok = _show_activation_dialog(root)
    return ok, None


def main():
    root = tk.Tk()
    root.withdraw()
    ok, trial_remaining = ensure_licensed(root)
    if not ok:
        root.destroy()
        return
    root.deiconify()
    App(root, trial_remaining=trial_remaining)
    root.mainloop()


if __name__ == "__main__":
    main()
