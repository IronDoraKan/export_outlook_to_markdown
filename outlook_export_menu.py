# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import argparse
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime, time
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, unquote

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError as exc:
    print("缺少 Outlook 讀取套件。請先安裝：py -m pip install pywin32")
    raise SystemExit(1) from exc

try:
    from markdownify import markdownify as html_to_markdown
except ImportError as exc:
    print("缺少 HTML 轉 Markdown 套件。請先安裝：py -m pip install markdownify")
    raise SystemExit(1) from exc


MAIL_ITEM_CLASS = 43
DEFAULT_OUTPUT_DIR = Path("outlook_markdown")
SCAN_SET_FILE = Path("scan_set.ini")
MANIFEST_FILE = ".export_manifest.json"
ATTACHMENT_DIR_NAME = "附件"
ATTACHMENT_EXPORT_VERSION = 2

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
PR_ATTACH_LONG_FILENAME = "http://schemas.microsoft.com/mapi/proptag/0x3707001F"


@dataclass
class ExportStats:
    scanned: int = 0
    exported: int = 0
    skipped_existing: int = 0
    skipped_outside_date: int = 0
    skipped_unlisted_folder: int = 0
    errors: int = 0


def pause() -> None:
    input("\n按 Enter 回到選單...")


def parse_date(prompt: str, default: datetime | None = None, end_of_day: bool = False) -> datetime:
    while True:
        default_text = f" [{default:%Y-%m-%d}]" if default else ""
        raw = input(f"{prompt}{default_text}: ").strip()
        if not raw and default:
            base = default
            return datetime.combine(base.date(), time.max if end_of_day else time.min)

        try:
            base = datetime.strptime(raw, "%Y-%m-%d")
            return datetime.combine(base.date(), time.max if end_of_day else time.min)
        except ValueError:
            print("日期格式請使用 YYYY-MM-DD，例如 2026-05-14。")


def safe_filename(value: str, fallback: str = "untitled", max_length: int = 90) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value or "")
    value = re.sub(r"\s+", " ", value).strip(" .")
    if not value:
        value = fallback
    if len(value) > max_length:
        value = value[:max_length].rstrip(" .")
    return value or fallback


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 2
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def load_manifest(output_dir: Path) -> dict:
    manifest_path = output_dir / MANIFEST_FILE
    if not manifest_path.exists():
        return {"messages": {}}

    try:
        return json.loads(manifest_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        backup_path = manifest_path.with_suffix(".json.bak")
        manifest_path.replace(backup_path)
        print(f"匯出紀錄檔格式有誤，已備份為 {backup_path.name}，並建立新的紀錄。")
        return {"messages": {}}


def save_manifest(output_dir: Path, manifest: dict) -> None:
    manifest_path = output_dir / MANIFEST_FILE
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def create_scan_set_template(scan_set_file: Path = SCAN_SET_FILE) -> None:
    if scan_set_file.exists():
        return

    scan_set_file.write_text(
        "\n".join(
            [
                "# 每行填一個要掃描的 Outlook 資料夾名稱。",
                "# 可填資料夾名稱，例如：Inbox",
                "# 也可填完整路徑，例如：\\\\your.name@example.com\\Inbox\\客戶郵件",
                "# 空白行與 # 開頭的行會略過。",
                "",
                "Inbox",
                "",
            ]
        ),
        encoding="utf-8",
    )


def load_scan_targets(scan_set_file: Path = SCAN_SET_FILE) -> set[str]:
    create_scan_set_template(scan_set_file)
    targets: set[str] = set()

    for line in scan_set_file.read_text(encoding="utf-8-sig").splitlines():
        value = line.strip()
        if not value or value.startswith("#"):
            continue
        targets.add(normalize_folder_token(value))

    return targets


def normalize_folder_token(value: str) -> str:
    normalized = value.strip().replace("/", "\\")
    normalized = re.sub(r"\\+", r"\\", normalized)
    return normalized.strip("\\").casefold()


def folder_matches_scan_targets(folder, targets: set[str]) -> bool:
    folder_name = normalize_folder_token(get_com_text(folder, "Name"))
    folder_path = normalize_folder_token(get_com_text(folder, "FolderPath"))
    return folder_name in targets or folder_path in targets


def get_com_text(com_object, attr_name: str) -> str:
    try:
        value = getattr(com_object, attr_name)
    except Exception:
        return ""
    return str(value or "").strip()


def get_property(com_object, property_name: str) -> str:
    try:
        return str(com_object.PropertyAccessor.GetProperty(property_name) or "").strip()
    except Exception:
        return ""


def normalize_com_datetime(value) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return datetime(
            value.year,
            value.month,
            value.day,
            getattr(value, "hour", 0),
            getattr(value, "minute", 0),
            getattr(value, "second", 0),
        )
    return None


def get_mail_time(mail) -> datetime | None:
    for attr_name in ("ReceivedTime", "SentOn", "CreationTime"):
        try:
            value = getattr(mail, attr_name)
        except Exception:
            continue

        normalized = normalize_com_datetime(value)
        if normalized:
            return normalized

    return None


def get_received_time(mail) -> datetime | None:
    try:
        received = mail.ReceivedTime
    except Exception:
        return None
    return normalize_com_datetime(received)


def get_message_key(mail) -> str:
    internet_id = get_property(mail, PR_INTERNET_MESSAGE_ID)
    if internet_id:
        return f"internet:{internet_id}"

    entry_id = get_com_text(mail, "EntryID")
    store_id = get_com_text(mail.Parent, "StoreID")
    return f"entry:{store_id}:{entry_id}"


def iter_folders(folder) -> Iterable:
    yield folder
    try:
        subfolders = folder.Folders
    except Exception:
        return

    for index in range(1, subfolders.Count + 1):
        try:
            yield from iter_folders(subfolders.Item(index))
        except Exception:
            continue


def iter_all_mail_folders(namespace) -> Iterable:
    stores = namespace.Folders
    for store_index in range(1, stores.Count + 1):
        store = stores.Item(store_index)
        yield from iter_folders(store)


def get_attachment_content_id(attachment) -> str:
    content_id = get_property(attachment, PR_ATTACH_CONTENT_ID)
    return content_id.strip("<>")


def save_attachments(mail, attachment_dir: Path, mail_stem: str) -> tuple[list[str], dict[str, str], list[str]]:
    attachment_dir.mkdir(parents=True, exist_ok=True)
    links: list[str] = []
    cid_map: dict[str, str] = {}
    saved_files: list[str] = []

    try:
        attachments = mail.Attachments
        count = attachments.Count
    except Exception:
        return links, cid_map, saved_files

    for index in range(1, count + 1):
        try:
            attachment = attachments.Item(index)
            raw_name = (
                get_property(attachment, PR_ATTACH_LONG_FILENAME)
                or get_com_text(attachment, "FileName")
                or get_com_text(attachment, "DisplayName")
                or f"attachment_{index}"
            )
            safe_name = safe_filename(raw_name, fallback=f"attachment_{index}", max_length=120)
            target = unique_path(attachment_dir / f"{mail_stem}_{index:02d}_{safe_name}")
            attachment.SaveAsFile(str(target.resolve()))

            relative = quote(f"{ATTACHMENT_DIR_NAME}/{target.name}", safe="/")
            display_name = target.name
            content_id = get_attachment_content_id(attachment)
            mime_tag = get_property(attachment, PR_ATTACH_MIME_TAG)

            if content_id:
                cid_map[content_id] = relative
                cid_map[content_id.casefold()] = relative

            if mime_tag.lower().startswith("image/"):
                links.append(f"![{display_name}]({relative})")
            else:
                links.append(f"[{display_name}]({relative})")
            saved_files.append(str(target))
        except Exception as exc:
            links.append(f"_附件匯出失敗：{exc}_")

    return links, cid_map, saved_files


def replace_cid_links(html: str, cid_map: dict[str, str]) -> str:
    def replace_match(match: re.Match) -> str:
        raw_cid = unquote(match.group(1)).strip("<>")
        return cid_map.get(raw_cid) or cid_map.get(raw_cid.casefold()) or match.group(0)

    return re.sub(r"cid:([^\"'>\s)]+)", replace_match, html or "", flags=re.IGNORECASE)


def make_markdown_body(mail, cid_map: dict[str, str]) -> str:
    html_body = get_com_text(mail, "HTMLBody")
    if html_body:
        html_body = replace_cid_links(html_body, cid_map)
        markdown = html_to_markdown(
            html_body,
            heading_style="ATX",
            bullets="-",
            strip=["style", "script"],
        )
        return markdown.strip()

    plain_body = get_com_text(mail, "Body")
    return plain_body.strip() or "_No body content_"


def build_mail_markdown(mail, folder_path: str, attachment_links: list[str], cid_map: dict[str, str]) -> str:
    subject = get_com_text(mail, "Subject") or "(No subject)"
    sender_name = get_com_text(mail, "SenderName") or "(Unknown sender)"
    sender_email = get_com_text(mail, "SenderEmailAddress")
    to_line = get_com_text(mail, "To")
    cc_line = get_com_text(mail, "CC")
    received = get_mail_time(mail)
    received_text = received.strftime("%Y-%m-%d %H:%M:%S") if received else ""
    markdown_body = make_markdown_body(mail, cid_map)

    sender = sender_name
    if sender_email:
        sender = f"{sender_name} <{sender_email}>"

    lines = [
        f"# {subject}",
        "",
        "## Metadata",
        "",
        f"- Date: {received_text}",
        f"- From: {sender}",
        f"- To: {to_line}",
        f"- Cc: {cc_line}",
        f"- Folder: {folder_path}",
        "",
        "## Body",
        "",
        markdown_body,
        "",
    ]

    if attachment_links:
        lines.extend(["## Attachments", ""])
        lines.extend(f"- {link}" for link in attachment_links)
        lines.append("")

    return "\n".join(lines)


def make_mail_stem(mail, received: datetime) -> str:
    date_part = received.strftime("%Y%m%d_%H%M%S")
    sender = safe_filename(get_com_text(mail, "SenderName") or "Unknown sender", max_length=40)
    subject = safe_filename(get_com_text(mail, "Subject") or "No subject", max_length=80)
    return f"{date_part}_{sender}_{subject}"


def export_mail(mail, output_dir: Path, manifest: dict) -> bool:
    received = get_mail_time(mail)
    if not received:
        raise RuntimeError("郵件沒有可用的 ReceivedTime。")

    message_key = get_message_key(mail)
    messages = manifest.setdefault("messages", {})
    existing_record = messages.get(message_key)
    if existing_record and existing_record.get("attachment_export_version") == ATTACHMENT_EXPORT_VERSION:
        return False

    mail_stem = make_mail_stem(mail, received)
    if existing_record and existing_record.get("file"):
        md_path = output_dir / existing_record["file"]
    else:
        md_path = unique_path(output_dir / f"{mail_stem}.md")
    attachment_dir = output_dir / ATTACHMENT_DIR_NAME

    attachment_links, cid_map, saved_files = save_attachments(mail, attachment_dir, mail_stem)
    folder_path = get_com_text(mail.Parent, "FolderPath")
    markdown = build_mail_markdown(mail, folder_path, attachment_links, cid_map)
    md_path.write_text(markdown, encoding="utf-8")

    messages[message_key] = {
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "received_time": received.isoformat(timespec="seconds"),
        "subject": get_com_text(mail, "Subject"),
        "sender": get_com_text(mail, "SenderName"),
        "file": md_path.name,
        "attachments": saved_files,
        "attachment_export_version": ATTACHMENT_EXPORT_VERSION,
    }
    return True


def parse_cli_date(value: str, *, end_of_day: bool = False) -> datetime:
    try:
        base = datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"日期格式需為 YYYY-MM-DD：{value}") from exc
    return datetime.combine(base.date(), time.max if end_of_day else time.min)


def export_by_date_range(
    output_dir: Path,
    start: datetime | None = None,
    end: datetime | None = None,
    scan_set_file: Path = SCAN_SET_FILE,
) -> None:
    if start is None:
        start = parse_date("請輸入開始日期")
    if end is None:
        end = parse_date("請輸入結束日期", default=start, end_of_day=True)
    if end < start:
        print("結束日期不能早於開始日期。")
        return

    scan_targets = load_scan_targets(scan_set_file)
    if not scan_targets:
        print(f"{scan_set_file} 沒有任何資料夾名稱，請先新增要掃描的資料夾。")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(output_dir)
    stats = ExportStats()

    print(f"\n正在連線 Outlook，並依 {scan_set_file} 掃描指定資料夾...")
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        for folder in iter_all_mail_folders(namespace):
            folder_path = get_com_text(folder, "FolderPath")
            if not folder_matches_scan_targets(folder, scan_targets):
                stats.skipped_unlisted_folder += 1
                continue

            try:
                items = folder.Items
                count = items.Count
            except Exception:
                continue

            if count == 0:
                continue

            print(f"掃描：{folder_path}")
            for index in range(1, count + 1):
                try:
                    item = items.Item(index)
                    if getattr(item, "Class", None) != MAIL_ITEM_CLASS:
                        continue

                    stats.scanned += 1
                    received = get_mail_time(item)
                    if not received or received < start or received > end:
                        stats.skipped_outside_date += 1
                        continue

                    if export_mail(item, output_dir, manifest):
                        stats.exported += 1
                        save_manifest(output_dir, manifest)
                    else:
                        stats.skipped_existing += 1

                    if stats.scanned % 50 == 0:
                        save_manifest(output_dir, manifest)
                except KeyboardInterrupt:
                    raise
                except Exception:
                    stats.errors += 1
                    print(f"  郵件處理失敗：{traceback.format_exc(limit=1).strip()}")

        save_manifest(output_dir, manifest)
    finally:
        pythoncom.CoUninitialize()

    print("\n匯出完成")
    print(f"掃描郵件：{stats.scanned}")
    print(f"新增匯出：{stats.exported}")
    print(f"已匯出略過：{stats.skipped_existing}")
    print(f"日期外略過：{stats.skipped_outside_date}")
    print(f"未列入掃描的資料夾略過：{stats.skipped_unlisted_folder}")
    print(f"錯誤：{stats.errors}")
    print(f"輸出資料夾：{output_dir.resolve()}")


def show_settings(output_dir: Path) -> None:
    manifest = load_manifest(output_dir) if output_dir.exists() else {"messages": {}}
    print("\n目前設定")
    print(f"輸出資料夾：{output_dir.resolve()}")
    print(f"附件資料夾：{(output_dir / ATTACHMENT_DIR_NAME).resolve()}")
    print(f"掃描清單：{SCAN_SET_FILE.resolve()}")
    print(f"已記錄匯出郵件數：{len(manifest.get('messages', {}))}")
    print("檔名格式：日期時間_寄信人_主旨.md")


def change_output_dir(current: Path) -> Path:
    raw = input(f"請輸入輸出資料夾 [{current}]: ").strip()
    if not raw:
        return current
    return Path(raw).expanduser()


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="匯出 Outlook 郵件為 Markdown；未提供參數時會啟動原本的互動選單。",
    )
    parser.add_argument("-s", "--start", help="開始日期，格式 YYYY-MM-DD")
    parser.add_argument("-e", "--end", help="結束日期，格式 YYYY-MM-DD")
    parser.add_argument("-f", "--file", help="要載入的掃描清單 ini 檔，例如 scan_set.ini")
    return parser


def run_cli(args: argparse.Namespace) -> int:
    missing = [
        option
        for option, value in (("-s/--start", args.start), ("-e/--end", args.end), ("-f/--file", args.file))
        if not value
    ]
    if missing:
        print(f"CLI 模式缺少必要參數：{', '.join(missing)}", file=sys.stderr)
        return 2

    scan_set_file = Path(args.file).expanduser()
    if not scan_set_file.exists():
        print(f"找不到掃描清單檔案：{scan_set_file}", file=sys.stderr)
        return 2

    try:
        start = parse_cli_date(args.start)
        end = parse_cli_date(args.end, end_of_day=True)
    except argparse.ArgumentTypeError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    export_by_date_range(DEFAULT_OUTPUT_DIR, start=start, end=end, scan_set_file=scan_set_file)
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv:
        parser = build_arg_parser()
        args = parser.parse_args(argv)
        return run_cli(args)

    output_dir = DEFAULT_OUTPUT_DIR

    while True:
        print("\nOutlook Markdown 匯出選單")
        print("1. 依日期範圍匯出 scan_set.ini 指定資料夾郵件")
        print("2. 查看目前設定與匯出紀錄")
        print("3. 修改輸出資料夾")
        print("4. 離開")

        choice = input("請選擇功能 [1-4]: ").strip()
        if choice == "1":
            try:
                export_by_date_range(output_dir)
            except KeyboardInterrupt:
                print("\n已中止。已完成的匯出紀錄會保留。")
            pause()
        elif choice == "2":
            show_settings(output_dir)
            pause()
        elif choice == "3":
            output_dir = change_output_dir(output_dir)
        elif choice == "4":
            print("再見。")
            return 0
        else:
            print("請輸入 1 到 4。")


if __name__ == "__main__":
    raise SystemExit(main())
