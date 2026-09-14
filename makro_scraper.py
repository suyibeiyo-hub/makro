#!/usr/bin/env python3
"""Makro.co.za 抓取核心与 PySide6 GUI。"""

from __future__ import annotations

import argparse
import os
import logging
import math
import random
import secrets
import sys
import time
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from typing import Any, Callable
from urllib.parse import parse_qs, parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
API_BASE = "https://www.makro.co.za"
PAGE_API = f"{API_BASE}/fccng/api/4/page/fetch"
SELLER_API = f"{API_BASE}/fccng/api/3/page/dynamic/product-sellers"
DEFAULT_PINCODE = "2157"
SORT_OPTIONS = ("price_asc", "price_desc")


def url_page_number(url: str) -> int:
    """读取 URL 中的 page 参数，没有或无效时返回第 1 页。"""
    raw_page = (parse_qs(urlsplit(url).query).get("page") or [""])[0]
    try:
        return max(1, int(raw_page))
    except (TypeError, ValueError):
        return 1


def set_page(url: str, page: int, sort_order: str | None = None) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    if page == 1:
        query.pop("page", None)
    else:
        query["page"] = str(page)
    if sort_order in SORT_OPTIONS:
        query["sort"] = sort_order
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


def product_url(product: dict[str, Any]) -> str:
    action = (product.get("productInfo") or {}).get("action") or {}
    url = action.get("url")
    if not url:
        value = (product.get("productInfo") or {}).get("value") or {}
        url = value.get("baseUrl") or ""
    return urljoin(API_BASE, url)


def post_json(session: requests.Session, endpoint: str, payload: Any, attempts: int = 3) -> dict[str, Any]:
    last = None
    for attempt in range(1, attempts + 1):
        try:
            time.sleep(random.uniform(3.0, 5.0))
            r = session.post(endpoint, json=payload, timeout=30)
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as exc:
            last = exc
            if attempt < attempts:
                time.sleep(1.5 * attempt)
    raise RuntimeError(f"请求失败 {endpoint}: {last}")


def fetch_page(session: requests.Session, page_url: str, page: int, first_context: dict[str, Any] | None,
               session_ids: tuple[str, str]) -> dict[str, Any]:
    context: dict[str, Any] = {"fetchSeoData": True}
    if page > 1:
        context.update({"paginatedFetch": True, "pageNumber": page})
        if first_context:
            context["paginationContextMap"] = deepcopy(first_context)
    payload = {
        "pageUri": urlsplit(page_url).path + ("?" + urlsplit(page_url).query if urlsplit(page_url).query else ""),
        "pageContext": context,
        "requestContext": {"type": "BROWSE_PAGE", "ssid": session_ids[0], "sqid": session_ids[1]},
    }
    return post_json(session, PAGE_API, payload)


def get_products(response: dict[str, Any]) -> list[dict[str, Any]]:
    products: list[dict[str, Any]] = []
    for slot in (response.get("RESPONSE") or {}).get("slots", []) or []:
        widget = slot.get("widget") or {}
        if widget.get("type") == "PRODUCT_SUMMARY":
            products.extend((widget.get("data") or {}).get("products", []) or [])
    return products


def fetch_sellers(session: requests.Session, product_id: str, pincode: str) -> list[dict[str, Any]]:
    response = post_json(session, SELLER_API, {
        "requestContext": {"productId": product_id},
        "locationContext": {"pincode": pincode},
    })
    root = response.get("RESPONSE") or {}
    fallback = ((root.get("pageContext") or {}).get("trackingDataV2") or {})
    rows = []
    data_root = root.get("data") or {}
    for item in (data_root.get("product_seller_detail_1") or {}).get("data", []) or []:
        value = item.get("value") or {}
        seller_value = ((value.get("sellerInfo") or {}).get("value") or {})
        pricing = ((value.get("pricing") or {}).get("value") or {})
        final_price = (pricing.get("finalPrice") or {}).get("value")
        mrp = (pricing.get("mrp") or {}).get("value")
        delivery = value.get("deliveryMessages") or []
        delivery_text = next(((x or {}).get("text") for x in delivery if (x or {}).get("text")), None)
        if not delivery_text:
            delivery_info = value.get("deliveryInfo") or {}
            delivery_text = (delivery_info.get("primaryOption") or {}).get("text")
        delivery_info = value.get("deliveryInfo") or {}
        primary_option = delivery_info.get("primaryOption") or {}
        rows.append({
            "seller": seller_value.get("name") or fallback.get("sellerName") or "",
            "seller_id": seller_value.get("id") or fallback.get("sellerId") or "",
            "price": final_price,
            "mrp": mrp,
            "delivery": ("FREE " if any((x or {}).get("freeDelivery") for x in delivery) else "") + (delivery_text or ""),
            "delivery_timestamp": primary_option.get("deliveryTimeStamp"),
            "listing_id": value.get("listingId") or fallback.get("listingId") or "",
        })
    if not rows and fallback:
        rows.append({
            "seller": fallback.get("sellerName", ""), "seller_id": fallback.get("sellerId", ""),
            "price": None, "mrp": None, "delivery": fallback.get("slaText", ""),
            "delivery_timestamp": fallback.get("slaTime"),
            "listing_id": fallback.get("listingId", ""),
        })
    return rows


def fetch_product_metadata(session: requests.Session, page_url: str) -> dict[str, Any]:
    """读取详情页标题、品牌和商品层面的 MRP/FSP。"""
    parts = urlsplit(page_url)
    page_uri = parts.path + ("?" + parts.query if parts.query else "")
    response = post_json(session, PAGE_API, {
        "pageUri": page_uri,
        "pageContext": {"fetchSeoData": True},
    })
    page_data = ((response.get("RESPONSE") or {}).get("pageData") or {})
    page_context = page_data.get("pageContext") or {}
    result: dict[str, Any] = {"brand": page_context.get("brand", ""), "title": "", "price": None, "mrp": None}
    for slot in (response.get("RESPONSE") or {}).get("slots", []) or []:
        widget = slot.get("widget") or {}
        data = widget.get("data") or {}
        if widget.get("type") == "PRODUCT_PAGE_SUMMARY":
            title_value = ((data.get("titleComponent") or {}).get("value") or {})
            result["title"] = title_value.get("title", "")
            pricing = (((data.get("pricing") or {}).get("value")) or {})
            result["price"] = (pricing.get("finalPrice") or {}).get("value")
            result["mrp"] = (pricing.get("mrp") or {}).get("value")
            if result["mrp"] is None:
                for price in pricing.get("prices", []) or []:
                    if price.get("priceType") == "MRP":
                        result["mrp"] = price.get("value")
                        break
            break
    return result


def price_in_rand(value: Any, tracking: dict[str, Any]) -> float | None:
    # Makro 页面 tracking 中的价格以 cents 表示：14900 -> R149.00。
    raw = tracking.get("price") or tracking.get("fsp")
    if raw is not None:
        try:
            return float(raw)
        except (TypeError, ValueError):
            pass
    try:
        return float(value) * 100 if value is not None else None
    except (TypeError, ValueError):
        return None


def seller_price_in_rand(value: Any) -> float | None:
    """商家接口的 value 是主货币单位，站点展示层再按 cents 展示。"""
    try:
        return float(value) * 100 if value is not None else None
    except (TypeError, ValueError):
        return None


def product_group_key(row: dict[str, Any]) -> tuple[str, str]:
    """返回稳定的商品键，确保同一详情页的多个卖家归并到一组。"""
    product_id = str(row.get("商品ID") or "").strip()
    if product_id:
        return ("id", product_id)

    detail_url = str(row.get("商品网址") or row.get("详情页地址") or "").strip()
    if detail_url:
        parts = urlsplit(detail_url)
        pid = (parse_qs(parts.query).get("pid") or [""])[0].strip()
        if pid:
            return ("pid", pid)
        # 去掉 lid、store 等跟踪参数，避免同一详情页因链接参数不同而重复。
        normalized = urlunsplit((parts.scheme.lower(), parts.netloc.lower(), parts.path.rstrip("/") or "/", "", ""))
        return ("url", normalized)

    # 没有商品 ID 和地址时，不把多个未知商品错误地归为同一条。
    return ("row", str(id(row)))


def delivery_days(timestamp: Any, delivery_text: str = "") -> float | None:
    try:
        return math.ceil(max(0.0, (float(timestamp) - time.time() * 1000) / 86_400_000))
    except (TypeError, ValueError):
        pass
    # 某些商品没有时间戳，只有类似“Delivery by 7 Oct, Wednesday”的文字。
    match = __import__("re").search(r"\b(\d{1,2})\s+([A-Za-z]{3,9})\b", delivery_text or "")
    if not match:
        return None
    day, month = match.groups()
    year = datetime.now().year
    parsed = None
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            parsed = datetime.strptime(f"{day} {month} {year}", fmt)
            break
        except ValueError:
            continue
    if parsed is None:
        return None
    if parsed.date() < datetime.now().date():
        parsed = parsed.replace(year=year + 1)
    return math.ceil(max(0.0, (parsed - datetime.now()).total_seconds() / 86_400))


def choose_lowest_offer(rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    priced: list[tuple[float, dict[str, Any]]] = []
    for row in rows:
        price = effective_price(row)
        if price is not None:
            priced.append((price, row))
    return min(priced, key=lambda item: item[0])[1] if priced else None


def effective_price(row: dict[str, Any]) -> float | None:
    """筛选/比价用价格：优惠价优先，否则原价，最后才回退到价格。"""
    for field in ("优惠价", "原价", "价格"):
        value = row.get(field)
        if value in (None, ""):
            continue
        try:
            number = float(value)
        except (TypeError, ValueError):
            continue
        if math.isfinite(number):
            return number
    return None


def deduplicate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """同一商品只保留有效价格最低的卖家记录，保留原始出现顺序。"""
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(product_group_key(row), []).append(row)

    result: list[dict[str, Any]] = []
    for offers in grouped.values():
        best = choose_lowest_offer(offers)
        result.append(best if best is not None else offers[0])
    return result


def scrape(url: str, pincode: str, max_pages: int | None, workers: int,
           progress_callback: Callable[[str], None] | None = None,
           page_start: int | None = None, page_end: int | None = None,
           sort_order: str | None = None) -> list[dict[str, Any]]:
    """抓取商品；page_start/page_end/sort_order 非空时覆盖 URL 参数。"""
    page_start = url_page_number(url) if page_start is None else max(1, int(page_start))
    if page_end is not None:
        page_end = max(1, int(page_end))
        if page_end < page_start:
            raise ValueError("结束页不能小于开始页")
    if sort_order not in (None, *SORT_OPTIONS):
        raise ValueError(f"排序参数必须是 {SORT_OPTIONS[0]} 或 {SORT_OPTIONS[1]}")
    # 兼容旧的 max_pages 参数：它表示从开始页起最多抓多少页。
    if page_end is None and max_pages is not None:
        page_end = page_start + max(1, int(max_pages)) - 1

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Origin": API_BASE,
        "Referer": url,
        "Content-Type": "application/json",
        "X-User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36 FKUA/website/42/website/Desktop",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "Sec-CH-UA": '"Chromium";v="152", "Not?A_Brand";v="24", "Google Chrome";v="152"',
        "Sec-CH-UA-Mobile": "?0",
        "Sec-CH-UA-Platform": '"macOS"',
    })
    time.sleep(random.uniform(3.0, 5.0))
    session.get(url, timeout=30)  # 建立站点 cookie；接口本身仍使用 JSON POST。
    session_ids = (secrets.token_hex(10) + str(int(time.time() * 1000)),
                   secrets.token_hex(9) + str(int(time.time() * 1000)))

    # 也支持直接输入商品详情页：?pid=...，无需经过分类页翻页。
    direct_pid = (parse_qs(urlsplit(url).query).get("pid") or [None])[0]
    if direct_pid:
        metadata = fetch_product_metadata(session, url)
        base_price = seller_price_in_rand(metadata.get("price"))
        base_mrp = seller_price_in_rand(metadata.get("mrp"))
        offers = fetch_sellers(session, direct_pid, pincode)
        rows = []
        for offer in offers:
            price = seller_price_in_rand(offer.get("price")) or base_price
            mrp = seller_price_in_rand(offer.get("mrp")) or base_mrp
            rows.append({
                "商品ID": direct_pid,
                "商品名称": metadata.get("title", ""),
                "品牌": metadata.get("brand", ""),
                "商品网址": url,
                "页码": 1,
                "商家": offer.get("seller", ""),
                "价格": price,
                "原价": mrp,
                "优惠价": price if price is not None and mrp is not None and price < mrp else None,
                "配送": offer.get("delivery", ""),
                "送达天数": delivery_days(offer.get("delivery_timestamp"), offer.get("delivery", "")),
                "Listing ID": offer.get("listing_id", ""),
                "Seller ID": offer.get("seller_id", ""),
            })
        if progress_callback:
            progress_callback("已处理第 1 条商品")
        return deduplicate_rows(rows)

    all_rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    first_context = None
    # 从第 N 页开始时，先请求第 1 页只获取分页上下文，确保接口能正确跳到 N 页。
    if page_start > 1:
        context_response = fetch_page(session, set_page(url, 1, sort_order), 1, None, session_ids)
        context_data = (context_response.get("RESPONSE") or {}).get("pageData") or {}
        first_context = context_data.get("paginationContextMap", {}).get("federator")
    previous_page_ids: tuple[str, ...] | None = None
    processed_products = 0
    page = page_start
    while page_end is None or page <= page_end:
        page_url = set_page(url, page, sort_order)
        response = fetch_page(session, page_url, page, first_context, session_ids)
        page_data = (response.get("RESPONSE") or {}).get("pageData") or {}
        if page == 1 and first_context is None:
            first_context = page_data.get("paginationContextMap", {}).get("federator")
        products = get_products(response)
        if not products:
            logging.info("第 %s 页没有商品，停止。", page)
            break
        current_page_ids = tuple(
            ((p.get("productInfo") or {}).get("value") or {}).get("id", "")
            for p in products
        )
        if current_page_ids == previous_page_ids:
            logging.info("第 %s 页与上一页重复，停止翻页。", page)
            if progress_callback:
                progress_callback(f"第 {page} 页内容重复，已安全停止")
            break
        previous_page_ids = current_page_ids
        logging.info("第 %s 页发现 %s 个商品。", page, len(products))

        def one(product: dict[str, Any]) -> list[dict[str, Any]]:
            product_info = product.get("productInfo") or {}
            info = product_info.get("value") or {}
            pid = info.get("id", "")
            tracking = product_info.get("tracking") or {}
            pricing_info = info.get("pricing") or {}
            base = {
                "商品ID": pid,
                "商品名称": (info.get("titles") or {}).get("title", ""),
                "品牌": info.get("productBrand", ""),
                "商品网址": product_url(product),
                "页码": page,
            }
            try:
                offers = fetch_sellers(session, pid, pincode)
            except Exception as exc:
                logging.warning("商品 %s 获取商家失败：%s", pid, exc)
                offers = [{}]
            result = []
            for offer in offers:
                row = dict(base)
                row.update({
                    "商家": offer.get("seller", ""),
                    "价格": seller_price_in_rand(offer.get("price")) if offer.get("price") is not None else price_in_rand((pricing_info.get("finalPrice") or {}).get("value"), tracking),
                    "原价": seller_price_in_rand(offer.get("mrp")) if offer.get("mrp") is not None else price_in_rand((pricing_info.get("mrp") or {}).get("value"), tracking),
                    "配送": offer.get("delivery", ""),
                    "送达天数": delivery_days(offer.get("delivery_timestamp"), offer.get("delivery", "")),
                    "Listing ID": offer.get("listing_id", info.get("listingId", "")),
                    "Seller ID": offer.get("seller_id", ""),
                })
                if row["价格"] is not None and row["原价"] is not None and row["价格"] < row["原价"]:
                    row["优惠价"] = row["价格"]
                else:
                    row["优惠价"] = None
                result.append(row)
            return result

        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [pool.submit(one, p) for p in products]
            for future in as_completed(futures):
                for row in future.result():
                    key = (row.get("商品ID", ""), row.get("Listing ID", ""))
                    if key not in seen:
                        seen.add(key)
                        all_rows.append(row)
                if progress_callback:
                    processed_products += 1
                    progress_callback(f"已处理第 {processed_products} 条商品")
        if not page_data.get("hasMorePages", False):
            break
        page += 1
    return deduplicate_rows(all_rows)


def filter_rows(rows: list[dict[str, Any]], min_days: float, min_price_rand: float,
                enabled: bool = True) -> list[dict[str, Any]]:
    """每个产品先取最低价卖家，再按送达天数和价格筛选。"""
    by_product: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        by_product.setdefault(product_group_key(row), []).append(row)
    result = []
    for offers in by_product.values():
        best = choose_lowest_offer(offers)
        if not best:
            continue
        selected_price = effective_price(best)
        if selected_price is None:
            continue
        price_rand = float(selected_price) / 100
        days = best.get("送达天数")
        if not enabled or (days is not None and days > min_days and price_rand > min_price_rand):
            result.append({
                "详情页地址": best.get("商品网址", ""),
                "价格(R)": round(price_rand, 2),
                "预计天数": int(days) if days is not None else "",
            })
    return sorted(result, key=lambda x: x["价格(R)"], reverse=True)


def export_xlsx(raw_rows: list[dict[str, Any]], filtered_rows: list[dict[str, Any]], output: str) -> None:
    raw_columns = ["商品名称", "品牌", "价格(R)", "原价(R)", "优惠价(R)", "配送", "预计天数", "详情页地址"]
    filtered_columns = ["详情页地址", "价格(R)", "预计天数"]
    wb = Workbook()
    ws_raw = wb.active
    ws_raw.title = "原始数据"
    ws_filtered = wb.create_sheet("筛选后的数据")

    for ws, columns, data in [(ws_raw, raw_columns, raw_rows), (ws_filtered, filtered_columns, filtered_rows)]:
        ws.append(columns)
        for row in data:
            values = []
            for column in columns:
                if column == "详情页地址":
                    value = row.get("详情页地址", row.get("商品网址", ""))
                elif column in ("价格(R)", "原价(R)", "优惠价(R)"):
                    if column == "价格(R)" and "价格(R)" in row:
                        value = row.get("价格(R)", "")
                    else:
                        source_column = {"价格(R)": "价格", "原价(R)": "原价", "优惠价(R)": "优惠价"}[column]
                        raw = row.get(source_column)
                        value = round(float(raw) / 100, 2) if raw is not None else ""
                elif column == "预计天数":
                    value = row.get("预计天数", row.get("送达天数", ""))
                else:
                    value = row.get(column, "")
                values.append(value)
            ws.append(values)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"
        ws.sheet_view.showGridLines = False
        for col in range(1, len(columns) + 1):
            ws.column_dimensions[get_column_letter(col)].width = min(max(14, max(len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 30) + 1)) + 2), 60)
        for col_index, column in enumerate(columns, 1):
            if column in ("价格(R)", "原价(R)", "优惠价(R)"):
                for cell in ws.iter_cols(min_col=col_index, max_col=col_index, min_row=2):
                    for item in cell:
                        item.number_format = '"R "#,##0.00'
    wb.save(output)


def cli_main() -> int:
    parser = argparse.ArgumentParser(description="抓取 Makro 分类页商品并筛选导出 Excel")
    parser.add_argument("url", help="Makro 分类页 URL，例如 https://www.makro.co.za/all/home-furnishing/pr?sid=all,jra")
    parser.add_argument("-o", "--output", default="makro_products.xlsx")
    parser.add_argument("--pincode", default=DEFAULT_PINCODE, help="配送邮编，默认 2157")
    parser.add_argument("--min-days", type=float, default=15)
    parser.add_argument("--min-price", type=float, default=300, help="实际兰特金额")
    parser.add_argument("--sort", choices=SORT_OPTIONS, default=None, help="覆盖 URL 的排序参数")
    parser.add_argument("--page-start", type=int, default=None, help="开始页；不设置则使用 URL 的 page")
    parser.add_argument("--page-end", type=int, default=None, help="结束页；不设置则一直翻页直到没有下一页")
    parser.add_argument("--max-pages", type=int, default=None, help="兼容旧参数：从开始页起最多抓多少页")
    parser.add_argument("--workers", type=int, default=5)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    rows = scrape(args.url, args.pincode, args.max_pages, max(1, min(args.workers, 10)),
                  page_start=args.page_start, page_end=args.page_end, sort_order=args.sort)
    filtered = filter_rows(rows, args.min_days, args.min_price)
    export_xlsx(rows, filtered, args.output)
    print(f"完成：抓取 {len(rows)} 条卖家记录，筛选 {len(filtered)} 个产品 -> {args.output}")
    return 0


def gui_main() -> int:
    from PySide6.QtCore import QObject, QThread, Signal, Slot, Qt
    from PySide6.QtWidgets import (QApplication, QCheckBox, QFormLayout, QHBoxLayout,
                                   QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton,
                                   QPlainTextEdit, QDoubleSpinBox, QSpinBox, QComboBox, QVBoxLayout,
                                   QWidget, QGroupBox)

    class Worker(QObject):
        progress = Signal(str)
        finished = Signal(str, int, int)
        failed = Signal(str)

        def __init__(self, url: str, pincode: str, page_start: int | None, page_end: int | None,
                     sort_order: str | None, min_days: float, min_price: float, output: str):
            super().__init__()
            self.url, self.pincode = url, pincode
            self.page_start, self.page_end, self.sort_order = page_start, page_end, sort_order
            self.min_days, self.min_price, self.output = min_days, min_price, output

        @Slot()
        def run(self):
            try:
                rows = scrape(self.url, self.pincode, None, 5,
                              progress_callback=self.progress.emit,
                              page_start=self.page_start, page_end=self.page_end,
                              sort_order=self.sort_order)
                self.progress.emit(f"已处理完成，共 {len(rows)} 条卖家记录")
                filtered = filter_rows(rows, self.min_days, self.min_price, True)
                export_xlsx(rows, filtered, self.output)
                self.finished.emit(self.output, len(rows), len(filtered))
            except Exception as exc:
                self.failed.emit(str(exc))

    class Window(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Makro 延迟送达商品筛选")
            self.resize(920, 620)
            self.thread = None
            self.worker = None
            root = QWidget(); self.setCentralWidget(root)
            root.setStyleSheet("""
                QWidget { font-size: 14px; }
                QGroupBox { font-weight: bold; border: 1px solid #D7DEE8; border-radius: 8px; margin-top: 10px; padding: 12px; }
                QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #1F4E78; }
                QLineEdit, QDoubleSpinBox, QSpinBox { min-height: 32px; border: 1px solid #B9C5D3; border-radius: 5px; padding: 3px 8px; }
                QPushButton { min-height: 38px; border-radius: 6px; padding: 0 18px; background: #1F4E78; color: white; font-weight: bold; }
                QPushButton:hover { background: #173A5B; }
                QPlainTextEdit { border: 1px solid #B9C5D3; border-radius: 5px; }
            """)
            layout = QVBoxLayout(root); layout.setContentsMargins(22, 18, 22, 18); layout.setSpacing(8)
            title = QLabel("Makro 商品筛选工具"); title.setStyleSheet("font-size: 22px; font-weight: bold; color: #1F4E78; margin-bottom: 4px;"); layout.addWidget(title)

            source_box = QGroupBox("数据来源"); source_form = QFormLayout(source_box); source_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
            self.url = QLineEdit(); self.url.setMinimumWidth(680); self.url.setPlaceholderText("输入 Makro 分类页或商品详情页 URL")
            self.pin = QLineEdit(DEFAULT_PINCODE)
            self.pin.setMaximumWidth(180)
            source_form.addRow("页面 URL", self.url); source_form.addRow("配送邮编", self.pin); layout.addWidget(source_box)

            condition_box = QGroupBox("筛选条件（固定启用）"); condition_form = QFormLayout(condition_box)
            self.days = QDoubleSpinBox(); self.days.setRange(0, 3650); self.days.setValue(15); self.days.setSuffix(" 天")
            self.price = QDoubleSpinBox(); self.price.setRange(0, 1_000_000); self.price.setValue(300); self.price.setPrefix("R ")
            condition_form.addRow("预计送达大于", self.days); condition_form.addRow("价格大于", self.price); layout.addWidget(condition_box)

            page_box = QGroupBox("分页与排序（可覆盖 URL 参数）"); page_form = QFormLayout(page_box)
            self.sort = QComboBox()
            self.sort.addItem("跟随 URL 参数", "")
            self.sort.addItem("价格从低到高（price_asc）", "price_asc")
            self.sort.addItem("价格从高到低（price_desc）", "price_desc")
            self.page_start = QLineEdit(); self.page_start.setMaximumWidth(180); self.page_start.setPlaceholderText("留空使用 URL 的 page")
            self.page_end = QLineEdit(); self.page_end.setMaximumWidth(180); self.page_end.setPlaceholderText("留空则抓到没有下一页")
            page_form.addRow("排序参数 sort", self.sort)
            page_form.addRow("开始页 page_start", self.page_start)
            page_form.addRow("结束页 page_end", self.page_end)
            page_hint = QLabel("例如：开始页 10、结束页 20，只抓第 10 到 20 页；留空的设置不会覆盖 URL。")
            page_hint.setStyleSheet("color: #667085; font-size: 12px;")
            page_form.addRow("", page_hint); layout.addWidget(page_box)

            self.start = QPushButton("保存 Excel"); self.start.clicked.connect(self.start_job); layout.addWidget(self.start)
            self.status = QLabel("准备就绪"); layout.addWidget(self.status)
            self.log = QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(100); layout.addWidget(self.log)

        def start_job(self):
            if not self.url.text().strip():
                QMessageBox.warning(self, "缺少网址", "请输入 Makro 分类页 URL")
                return
            url = self.url.text().strip()
            try:
                page_start = int(self.page_start.text().strip()) if self.page_start.text().strip() else None
                page_end = int(self.page_end.text().strip()) if self.page_end.text().strip() else None
            except ValueError:
                QMessageBox.warning(self, "页码无效", "开始页和结束页必须是整数。")
                return
            if page_start is not None and page_start < 1:
                QMessageBox.warning(self, "页码无效", "开始页必须大于或等于 1。")
                return
            if page_end is not None and page_end < 1:
                QMessageBox.warning(self, "页码无效", "结束页必须大于或等于 1。")
                return
            resolved_start = page_start if page_start is not None else url_page_number(url)
            if page_end is not None and page_end < resolved_start:
                QMessageBox.warning(self, "页码无效", "结束页不能小于开始页。")
                return
            sort_order = self.sort.currentData() or None
            self.start.setEnabled(False); self.status.setText("处理中，请等待..."); self.log.clear()
            run_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path.cwd()
            output = str(run_dir / f"makro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            self.thread = QThread(); self.worker = Worker(url, self.pin.text().strip(), page_start, page_end, sort_order, self.days.value(), self.price.value(), output)
            self.worker.moveToThread(self.thread); self.thread.started.connect(self.worker.run)
            self.worker.finished.connect(self.done); self.worker.failed.connect(self.error)
            self.worker.progress.connect(self.add_progress)
            self.worker.finished.connect(self.thread.quit); self.worker.failed.connect(self.thread.quit); self.thread.finished.connect(self.thread.deleteLater); self.thread.start()

        def add_progress(self, message):
            self.log.appendPlainText(message)

        def done(self, path, total, count):
            self.start.setEnabled(True); self.status.setText(f"完成：已保存 {count} 个产品"); self.log.appendPlainText(f"已保存 {count} 条，文件：{path}")
            QMessageBox.information(self, "完成", f"已保存 {count} 个产品到：\n{path}")

        def error(self, message):
            self.start.setEnabled(True); self.status.setText("失败"); self.log.appendPlainText(message); QMessageBox.critical(self, "请求失败", message)

    app = QApplication(sys.argv); win = Window(); win.show(); return app.exec()


if __name__ == "__main__":
    sys.exit(gui_main())
