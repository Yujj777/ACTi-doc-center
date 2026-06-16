import argparse
import re
import time
from pathlib import Path
from typing import Optional

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def pick_column(df: pd.DataFrame, candidates: list[str], fallback_index: int) -> str:
    norm = {clean_text(c).lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in norm:
            return norm[c.lower()]
    if len(df.columns) > fallback_index:
        return df.columns[fallback_index]
    raise ValueError(f"Missing required column candidates={candidates}, current={list(df.columns)}")


def normalize_warranty(raw: str) -> str:
    raw = clean_text(raw)
    if not raw:
        return ""
    m = re.search(r"\d+\s*年", raw)
    if m:
        return clean_text(m.group(0))
    m = re.search(r"\d+\s*year", raw, re.IGNORECASE)
    if m:
        return clean_text(m.group(0))
    return clean_text(re.split(r"[.。;；]", raw)[0])


def extract_by_label_from_soup(soup: BeautifulSoup, labels: list[str]) -> str:
    label_set = {x.lower() for x in labels}
    for tr in soup.select("tr"):
        tds = tr.select("td")
        if len(tds) < 2:
            continue
        left = clean_text(tds[0].get_text(" ", strip=True)).lower()
        if left in label_set:
            return clean_text(tds[1].get_text(" ", strip=True))
    return ""


def parse_html_source(html: str) -> tuple[str, str, str]:
    soup = BeautifulSoup(html, "html.parser")

    name = ""
    for selector in [
        "#popupRootDivId #popupHeaderSpec",
        "#popupHeaderSpec",
        "#popupRootDivId #selfModelName",
        "#selfModelName",
    ]:
        node = soup.select_one(selector)
        if node:
            name = clean_text(node.get_text(" ", strip=True))
            if name:
                break

    if not name:
        md = soup.select_one("meta[name='description']")
        if md and md.get("content"):
            name = clean_text(md["content"])
    if not name:
        t = soup.find("title")
        if t:
            name = clean_text(t.get_text(" ", strip=True).replace("| ACTi Corporation", ""))

    warranty_raw = ""
    w = soup.select_one("#Specifications_6621045F-F6EB-5C2D-2D44-24943C0C1F38")
    if w:
        warranty_raw = clean_text(w.get_text(" ", strip=True))
    if not warranty_raw:
        warranty_raw = extract_by_label_from_soup(soup, ["保固", "Warranty"])
    warranty = normalize_warranty(warranty_raw)

    ptype = ""
    t = soup.select_one("#Specifications_GROUPING-LEVE-L300-0000-000000000000")
    if t:
        ptype = clean_text(t.get_text(" ", strip=True))
    if not ptype:
        ptype = extract_by_label_from_soup(soup, ["類型", "Type"])

    return name, warranty, ptype


def build_driver(headless: bool, user_data_dir: Optional[Path] = None) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    if user_data_dir:
        opts.add_argument(f"--user-data-dir={str(user_data_dir)}")
    return webdriver.Chrome(options=opts)


def fetch_rendered_html(driver, url: str, timeout_s: int = 20) -> str:
    if "?tab=specifications" not in url:
        url = f"{url}?tab=specifications"
    driver.get(url)
    WebDriverWait(driver, timeout_s).until(lambda d: d.execute_script("return document.readyState") == "complete")
    # give dynamic popup script a moment
    time.sleep(1.2)
    return driver.page_source


def find_local_html(model: str, html_dir: Optional[Path]) -> Optional[Path]:
    if not html_dir or not html_dir.exists():
        return None
    direct = [
        html_dir / f"{model} _ ACTi Corporation.html",
        html_dir / f"{model}_ACTi Corporation.html",
        html_dir / f"{model}.html",
    ]
    for p in direct:
        if p.exists():
            return p
    model_lower = model.lower()
    for p in html_dir.glob("*.html"):
        n = p.name.lower()
        if model_lower in n and "acti corporation" in n:
            return p
    return None


def crawl_excel(
    input_path: Path,
    output_path: Path,
    html_dir: Optional[Path],
    headless: bool,
    retries: int,
    save_every: int,
    debug: bool,
    manual_login: bool = False,
    user_data_dir: Optional[Path] = None,
) -> None:
    df = pd.read_excel(input_path, engine="openpyxl")
    model_col = pick_column(df, ["model"], 0)
    website_col = pick_column(df, ["website", "link"], 1)

    rows: list[dict] = []
    failed: list[dict] = []
    total = len(df)

    driver = build_driver(headless=headless, user_data_dir=user_data_dir)
    try:
        if manual_login:
            login_url = "https://www.acti.com/zh-tw/member/login"
            driver.get(login_url)
            print("Manual login mode: please login in opened browser, then press Enter here to continue...")
            input()

        for i, row in df.iterrows():
            model = clean_text(row.get(model_col, ""))
            website = clean_text(row.get(website_col, ""))
            product_name, warranty, ptype = "", "", ""
            source = "none"
            err = ""

            if website.startswith("http"):
                for attempt in range(retries + 1):
                    try:
                        html = fetch_rendered_html(driver, website)
                        product_name, warranty, ptype = parse_html_source(html)
                        source = "website"
                        if product_name or warranty or ptype:
                            break
                        raise RuntimeError("Empty parsed result")
                    except Exception as e:
                        err = str(e)
                        if attempt < retries:
                            time.sleep(1.0)

            # local html fallback
            if (not product_name and not warranty and not ptype) or product_name.lower() == "model not found":
                local_file = find_local_html(model, html_dir)
                if local_file:
                    try:
                        html = local_file.read_text(encoding="utf-8", errors="ignore")
                        product_name, warranty, ptype = parse_html_source(html)
                        source = f"local:{local_file.name}"
                    except Exception as e:
                        err = f"local parse failed: {e}"

            if not (product_name or warranty or ptype):
                failed.append({"Model": model, "Website": website, "Error": err or "empty result"})

            rows.append(
                {
                    "Model": model,
                    "Website": website,
                    "產品名稱": product_name,
                    "保固": warranty,
                    "類型": ptype,
                }
            )

            if debug:
                print(f"[{i + 1}/{total}] {model} source={source} -> 名稱='{product_name}' 保固='{warranty}' 類型='{ptype}'")
            else:
                print(f"[{i + 1}/{total}] {model}: done")

            if save_every > 0 and ((i + 1) % save_every == 0 or (i + 1) == total):
                pd.DataFrame(rows).to_excel(output_path, index=False, engine="openpyxl")
                print(f"  [autosave] {i + 1}/{total} saved -> {output_path.name}")
    finally:
        driver.quit()

    result_df = pd.DataFrame(rows, columns=["Model", "Website", "產品名稱", "保固", "類型"])
    result_df.to_excel(output_path, index=False, engine="openpyxl")

    if failed:
        failed_path = output_path.with_name(f"{output_path.stem}_failed.xlsx")
        pd.DataFrame(failed).to_excel(failed_path, index=False, engine="openpyxl")
        print(f"Failed rows exported: {failed_path}")


def resolve_input_excel_path(requested: Path) -> Path:
    """If exact path missing, try Windows duplicate-extension names like Product.xlsx.xlsx.xlsx."""
    if requested.exists():
        return requested
    parent = requested.parent
    stem = requested.stem
    for alt in (
        parent / f"{stem}.xlsx.xlsx",
        parent / f"{stem}.xlsx.xlsx.xlsx",
        parent / f"{stem}.xlsx.xlsx.xlsx.xlsx",
    ):
        if alt.exists():
            print(f"Note: using input file (actual name on disk): {alt.name}")
            return alt
    matches = sorted(parent.glob(f"{stem}*.xlsx*"), key=lambda p: len(p.name))
    if len(matches) == 1:
        print(f"Note: using input file (matched): {matches[0].name}")
        return matches[0]
    return requested


def main() -> None:
    parser = argparse.ArgumentParser(description="ACTi product crawler (rewritten robust version)")
    parser.add_argument("--input", default="Product.xlsx")
    parser.add_argument("--output", default="Product_output.xlsx")
    parser.add_argument("--local-html-dir", default="")
    parser.add_argument("--no-headless", action="store_true")
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--save-every", type=int, default=10)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--manual-login", action="store_true")
    parser.add_argument(
        "--user-data-dir",
        default="",
        help="Optional Chrome profile folder for persisted login session.",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    input_path = resolve_input_excel_path(Path(args.input).resolve())
    output_path = Path(args.output).resolve()
    html_dir = Path(args.local_html_dir).resolve() if args.local_html_dir else None
    user_data_dir = Path(args.user_data_dir).resolve() if args.user_data_dir else None
    if not input_path.exists():
        hint = ""
        try:
            nearby = sorted(script_dir.glob("*.xlsx"))[:8]
            if nearby:
                hint = "\n  Excel files next to this script:\n    " + "\n    ".join(str(p.name) for p in nearby)
        except OSError:
            pass
        raise FileNotFoundError(
            f"Input not found: {input_path}\n"
            f"  Tip: use --input with full path, or place your .xlsx in:\n    {script_dir}"
            f"{hint}"
        )

    crawl_excel(
        input_path=input_path,
        output_path=output_path,
        html_dir=html_dir,
        headless=not args.no_headless,
        retries=max(0, args.retries),
        save_every=max(0, args.save_every),
        debug=args.debug,
        manual_login=args.manual_login,
        user_data_dir=user_data_dir,
    )
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
