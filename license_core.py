# -*- coding: utf-8 -*-
"""
配布先PCを限定するためのライセンスキーのロジック（公開鍵署名方式）。

仕組み:
  1. 配布先PCで「マシンID」を取得する（Windowsのインストール固有GUID）。
  2. 開発者側の key_generator.py が、秘密鍵(private_key.hex)でそのマシンID
     に対する電子署名を計算し、それをライセンスキーとして発行する。
  3. 配布先PCのアプリ（exeに埋め込まれた公開鍵 _PUBLIC_KEY_HEX）は、
     入力されたライセンスキーが自分のマシンIDへの正しい署名かどうかを検証する。

  公開鍵からは秘密鍵を計算で導き出すことができないため、exeを解析されても
  （このソースコードごと読まれても）秘密鍵(private_key.hex)そのものが
  別途漏洩しない限り、第三者が新しいライセンスキーを偽造することはできない。
  private_key.hex は配布物（exe）には一切含めず、開発者の手元にのみ置くこと。
"""

import base64
import datetime
import hashlib
import os
import subprocess

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey, Ed25519PublicKey

# 検証用の公開鍵（配布して問題ない）。対応する秘密鍵は private_key.hex にのみ存在する。
_PUBLIC_KEY_HEX = "482387780bca5d873f186986131d556b6557b132c231b0eefa2d20bea6c086d5"

TRIAL_DAYS = 30


def _read_registry_machine_guid():
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
            0,
            winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
        ) as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
            if value:
                return value.strip()
    except Exception:
        return None
    return None


def _read_wmic_uuid():
    try:
        out = subprocess.check_output(
            ["wmic", "csproduct", "get", "uuid"],
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        lines = [l.strip() for l in out.decode("utf-8", "ignore").splitlines() if l.strip()]
        if len(lines) >= 2 and lines[1].upper() != "UUID":
            return lines[1]
    except Exception:
        return None
    return None


def get_machine_id() -> str:
    """このPCを識別するID文字列を返す（表示・入力用に整形済み）。"""
    raw = _read_registry_machine_guid() or _read_wmic_uuid()
    if not raw:
        # 最終手段: ホスト名などから簡易的に生成（複数PCで重複しうるが動作は継続する）
        import platform
        raw = "FALLBACK-" + platform.node()
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()[:20]
    return "-".join(digest[i:i + 4] for i in range(0, 20, 4))


def _normalize_machine_id(machine_id: str) -> bytes:
    return machine_id.strip().upper().replace(" ", "").encode("utf-8")


def compute_license_key(machine_id: str, private_key_path: str) -> str:
    """開発者側専用。private_key.hex を使ってこのマシンID用のライセンスキーを発行する。"""
    with open(private_key_path, "r", encoding="ascii") as f:
        priv_bytes = bytes.fromhex(f.read().strip())
    priv = Ed25519PrivateKey.from_private_bytes(priv_bytes)
    signature = priv.sign(_normalize_machine_id(machine_id))
    return base64.urlsafe_b64encode(signature).decode("ascii").rstrip("=")


def verify_license_key(machine_id: str, entered_key: str) -> bool:
    entered = (entered_key or "").strip()
    if not entered:
        return False
    try:
        padded = entered + "=" * (-len(entered) % 4)
        signature = base64.urlsafe_b64decode(padded)
        pub = Ed25519PublicKey.from_public_bytes(bytes.fromhex(_PUBLIC_KEY_HEX))
        pub.verify(signature, _normalize_machine_id(machine_id))
        return True
    except (InvalidSignature, ValueError):
        return False


_TRIAL_REG_PATH = r"Software\DiveReportGenerator"
_TRIAL_REG_VALUE = "TrialStart"


def _read_registry_trial_start():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _TRIAL_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, _TRIAL_REG_VALUE)
            return datetime.date.fromisoformat(value)
    except Exception:
        return None


def _write_registry_trial_start(date_) -> bool:
    try:
        import winreg
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _TRIAL_REG_PATH) as key:
            winreg.SetValueEx(key, _TRIAL_REG_VALUE, 0, winreg.REG_SZ, date_.isoformat())
        return True
    except Exception:
        return False


def get_trial_remaining_days(trial_file: str) -> int:
    """試用期間の残り日数を返す（0以下ならライセンス認証が必要）。

    試用開始日はこのPCのレジストリ(HKCU)に記録する。フォルダ内のファイルを
    削除してコピーし直しても、同じPC・同じWindowsユーザーである限り
    レジストリの記録は残るため、試用期間を使い回すことはできない
    （trial_fileは、このレジストリ方式導入前の旧バージョンからの移行用の
    フォールバックとしてのみ使う）。"""
    today = datetime.date.today()
    first_run = _read_registry_trial_start()
    if first_run is None:
        # 旧バージョン(ファイルのみで管理)からの移行: ファイルに記録があればそれを引き継ぐ
        if os.path.exists(trial_file):
            try:
                with open(trial_file, "r", encoding="utf-8") as f:
                    first_run = datetime.date.fromisoformat(f.read().strip())
            except (OSError, ValueError):
                first_run = None
        if first_run is None:
            first_run = today
        if not _write_registry_trial_start(first_run):
            # レジストリに書き込めない場合のみファイルにフォールバック
            try:
                with open(trial_file, "w", encoding="utf-8") as f:
                    f.write(first_run.isoformat())
            except OSError:
                pass
    elapsed = (today - first_run).days
    if elapsed < 0:
        elapsed = 0
    return TRIAL_DAYS - elapsed
