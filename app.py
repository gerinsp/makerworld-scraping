import argparse, time, re, html, unicodedata, zipfile
from urllib.parse import quote_plus, urljoin
from pathlib import Path
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook, Workbook

# =========================
# Config / constants
# =========================

PROFILE_DIR = ".mw_profile"
STATE_FILE  = "makerworld_state.json"

STEALTH_JS = r"""() => {
  Object.defineProperty(navigator,'webdriver',{get:()=>undefined});
  const q=navigator.permissions&&navigator.permissions.query;
  if(q){navigator.permissions.query = p=>p.name==='notifications'
    ? Promise.resolve({state: Notification.permission}) : q(p);}
  Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']});
  Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4]});
  Object.defineProperty(navigator,'hardwareConcurrency',{get:()=>8});
  const gp=WebGLRenderingContext.prototype.getParameter;
  WebGLRenderingContext.prototype.getParameter=function(p){
    if(p===37445) return "Intel Inc."; if(p===37446) return "Intel Iris OpenGL Engine";
    return gp.call(this,p);
  };
  window.chrome={runtime:{}};
}"""

# =========================
# Helpers: text & SEO
# =========================

def _norm_text(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = re.sub(r"[^\w\s]", "", s)
    return s

def looks_like_cf(html_text:str)->bool:
    h=(html_text or "").lower()
    return any(k in h for k in ["just a moment","verifying you are human","cloudflare","cf-ray","cf-chl"])

def seo_title(t:str)->str:
    t=re.sub(r"\s+"," ", (t or "").strip())
    return (t + " – Cable Winder Organizer, Portable, 3D Print")[:255]

_REMOVE_LINE_PATTERNS = [
    re.compile(r"^\s*print\s*profile\s*:.*$", re.I),
    re.compile(r"^\s*sumber\s*desain\s*:.*$", re.I),
    re.compile(r"^\s*design\s*source\s*:.*$", re.I),
    re.compile(r"^\s*source\s*:.*$", re.I),
    re.compile(r"\bFAQ\b\s*$", re.I),
]

def _clean_desc_text(desc: str) -> str:
    if not desc:
        return ""
    text = html.unescape(desc).replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    cleaned = []
    for ln in lines:
        if any(p.search(ln) for p in _REMOVE_LINE_PATTERNS):
            continue
        cleaned.append(ln)
    return "\n".join(cleaned)

def seo_desc(desc:str)->str:
    parts=[]
    base = _clean_desc_text(desc)
    if base:
        parts.append(base)

    parts += [
        "Kegunaan: penggulung kabel USB/Type-C/Lightning.",
        "Material: PLA/ABS (sesuai profil cetak)."
    ]
    # no profiles, no 'Sumber desain'
    final = "\n\n".join([p for p in parts if p]).strip()
    # Safety cap
    return final[:3000]

def auto_scroll(page):
    page.evaluate("() => window.scrollTo(0,0)")
    for _ in range(14):
        page.evaluate("() => window.scrollBy(0,900)")
        time.sleep(0.12)

def get_gallery_urls(page)->list[str]:
    g=page.query_selector(".photo_show")
    if not g: return []
    for th in g.query_selector_all("img, .swiper-slide img, picture img"):
        try:
            th.scroll_into_view_if_needed()
            th.click(timeout=150)
            time.sleep(0.05)
        except: 
            pass
    urls=[]
    for im in g.query_selector_all("img"):
        try:
            u=im.get_attribute("src") or im.get_attribute("data-src")
            if not u: continue
            if u.startswith("//"): u="https:"+u
            if u.startswith("/"):  u=urljoin("https://makerworld.com", u)
            low=u.lower()
            if any(b in low for b in ["avatar","logo","icon","/comment","emote","placeholder",".svg",".ico"]): 
                continue
            if not any(ext in low for ext in [".jpg",".jpeg",".png",".webp",".avif"]): 
                continue
            if u not in urls: urls.append(u)
            if len(urls)==8: break
        except: 
            pass
    return urls

# =========================
# XLSX sanitizer & header mapping
# =========================

SHEETVIEWS_RE = re.compile(rb"<sheetViews[\s\S]*?</sheetViews>")
IMG_LABEL_RE  = re.compile(r"^foto\s+(sampul|produk\s+\d+)$") 

HEADER_ALIASES = {
    "Kategori": {"kategori", "category", "product category", "category id"},
    "Nama Produk": {"nama produk", "product name", "name", "nama"},
    "Deskripsi Produk": {"deskripsi produk", "product description", "description"},
    "Foto Produk": {"foto produk", "images", "image urls", "product images", "photo", "photos", "gambar", "url gambar"},
    "Harga": {"harga", "price"},
    "Stok": {"stok", "stock", "quantity"},
    "SKU Induk": {"sku induk", "sku", "parent sku", "model sku"},
    "Berat (gram)": {"berat gram", "berat", "berat produk", "weight", "weight g", "weight gram"},
    "Panjang Paket (cm)": {"panjang paket cm", "panjang cm", "panjang", "length", "length cm"},
    "Lebar Paket (cm)": {"lebar paket cm", "lebar cm", "lebar", "width", "width cm"},
    "Tinggi Paket (cm)": {"tinggi paket cm", "tinggi cm", "tinggi", "height", "height cm"},
}

def sanitize_xlsx(src_path:str)->str:
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {src_path}")
    dst = src.with_name(src.stem + "_sanitized.xlsx")
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                data = SHEETVIEWS_RE.sub(b"", data)
            zout.writestr(item, data)
    return str(dst)

def create_minimal_template(path:str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    ws.append(["Kategori","Nama Produk","Deskripsi Produk",
               "Harga","Stok","SKU Induk",
               "Berat (gram)","Panjang Paket (cm)","Lebar Paket (cm)","Tinggi Paket (cm)",
               "Foto Produk"])
    wb.save(path)
    return path

def _find_header_positions(ws, search_rows=50, search_cols=160):
    for start_r in range(1, search_rows+1):
        for h_rows in (1, 2, 3):
            merged = []
            for c in range(1, search_cols+1):
                parts=[]
                for rr in range(start_r, min(start_r+h_rows, search_rows+1)):
                    v = ws.cell(rr, c).value
                    if v is not None and str(v).strip():
                        parts.append(str(v))
                merged.append(" ".join(parts) if parts else "")
            normalized = [_norm_text(v) for v in merged]
            col_to_text = {i+1: t for i,t in enumerate(normalized) if t}

            found={}
            for target, aliases in HEADER_ALIASES.items():
                alias_norm = {_norm_text(a) for a in aliases}
                for c, txt in col_to_text.items():
                    if txt in alias_norm:
                        found[target]=c
                        break

            image_cols=[]
            for c, txt in col_to_text.items():
                if IMG_LABEL_RE.match(txt):
                    image_cols.append(c)

            if ("Kategori" in found and "Nama Produk" in found and
                ("Deskripsi Produk" in found or image_cols or "Foto Produk" in found)):
                return (start_r + h_rows - 1, found, sorted(image_cols))
    return (None, {}, [])

def write_rows_to_shopee_template(template_path:str, out_path:str, rows:list[dict], sheet_name: str | None = None):
    try:
        patched = sanitize_xlsx(template_path)
    except Exception as e:
        print("[Template] gagal sanitize:", e)
        patched = create_minimal_template("_fallback_template.xlsx")

    try:
        wb = load_workbook(patched)
    except Exception as e:
        print("[openpyxl] gagal load:", e)
        patched = create_minimal_template("_fallback_template.xlsx")
        wb = load_workbook(patched)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb["Template"] if "Template" in wb.sheetnames else wb[wb.sheetnames[0]]

    hdr_row, colmap_found, image_cols = _find_header_positions(ws)
    if not hdr_row:
        print("[Header] tidak ketemu header valid. Buat sheet 'Template' minimal.")
        if "Template" in wb.sheetnames:
            del wb["Template"]
        ws = wb.create_sheet("Template")
        ws.append(["Kategori","Nama Produk","Deskripsi Produk",
                   "Harga","Stok","SKU Induk",
                   "Berat (gram)","Panjang Paket (cm)","Lebar Paket (cm)","Tinggi Paket (cm)",
                   "Foto Produk"])
        hdr_row = 1
        colmap_found = {h: i+1 for i, h in enumerate(["Kategori","Nama Produk","Deskripsi Produk",
                                                      "Harga","Stok","SKU Induk",
                                                      "Berat (gram)","Panjang Paket (cm)","Lebar Paket (cm)","Tinggi Paket (cm)",
                                                      "Foto Produk"])}
        image_cols = []

    def ensure_col(label):
        if label in colmap_found and colmap_found[label]:
            return colmap_found[label]
        new_idx = ws.max_column + 1
        ws.cell(hdr_row, new_idx).value = label
        colmap_found[label] = new_idx
        return new_idx

    for req in ["Kategori","Nama Produk","Deskripsi Produk"]:
        ensure_col(req)
    for maybe in ["Harga","Stok","SKU Induk","Berat (gram)","Panjang Paket (cm)","Lebar Paket (cm)","Tinggi Paket (cm)","Foto Produk"]:
        ensure_col(maybe)

    img_cols=[]
    for c in range(1, ws.max_column+1):
        lab = _norm_text(ws.cell(hdr_row, c).value)
        if IMG_LABEL_RE.match(lab):
            img_cols.append(c)
    img_cols = sorted(img_cols)

    def row_has_data(rr):
        for c in range(1, ws.max_column+1):
            if str(ws.cell(rr, c).value or "").strip():
                return True
        return False
    r = hdr_row + 1
    while row_has_data(r): r += 1

    def C(label): return colmap_found[label]

    for row in rows:
        ws.cell(r, C("Kategori")).value          = str(row["category_id"])
        ws.cell(r, C("Nama Produk")).value       = row["name"]
        ws.cell(r, C("Deskripsi Produk")).value  = row["description"]

        if "price" in row and "Harga" in colmap_found:
            ws.cell(r, C("Harga")).value = int(row["price"])
        if "stock" in row and "Stok" in colmap_found:
            ws.cell(r, C("Stok")).value  = int(row["stock"])
        if "sku" in row and "SKU Induk" in colmap_found:
            ws.cell(r, C("SKU Induk")).value = row["sku"]

        if "weight_kg" in row:
            if "Berat (gram)" in colmap_found:
                ws.cell(r, C("Berat (gram)")).value = int(round(row["weight_kg"]*1000))
        if "dims_cm" in row:
            L,W,H = row["dims_cm"]
            if "Panjang Paket (cm)" in colmap_found: ws.cell(r, C("Panjang Paket (cm)")).value = int(L)
            if "Lebar Paket (cm)"   in colmap_found: ws.cell(r, C("Lebar Paket (cm)")).value   = int(W)
            if "Tinggi Paket (cm)"  in colmap_found: ws.cell(r, C("Tinggi Paket (cm)")).value  = int(H)

        urls = row.get("image_urls", [])[:8]
        if img_cols:
            for idx, c in enumerate(img_cols):
                if idx < len(urls):
                    ws.cell(r, c).value = urls[idx]
        else:
            if "Foto Produk" in colmap_found:
                ws.cell(r, C("Foto Produk")).value = ",".join(urls)

        r += 1

    wb.save(out_path)
    print("[Shopee] file jadi →", out_path)

# =========================
# Scraping MakerWorld
# =========================

def scrape(keyword:str, max_results:int, headless:bool, proxy:str|None):
    results=[]
    with sync_playwright() as p:
        for engine in ("chromium","webkit","firefox"):
            ctx=None
            try:
                bt=getattr(p, engine)
                kwargs={"headless": headless, "args":["--no-sandbox","--disable-blink-features=AutomationControlled"]}
                if proxy: kwargs["proxy"]={"server":proxy}
                ctx=bt.launch_persistent_context(
                    user_data_dir=PROFILE_DIR,
                    **kwargs, locale="en-US", timezone_id="Asia/Jakarta",
                    viewport={"width":1400,"height":900}, device_scale_factor=2
                )
                ctx.add_init_script(STEALTH_JS)
                page=ctx.new_page(); page.set_default_navigation_timeout(120000)

                url=f"https://makerworld.com/en/search/models?keyword={quote_plus(keyword)}"
                print(f"[{engine}] open", url)
                page.goto(url, wait_until="domcontentloaded")
                try: page.wait_for_load_state("networkidle", timeout=15000)
                except: pass

                if looks_like_cf(page.content()):
                    print(f"[{engine}] Cloudflare detected")
                    if headless and not proxy:
                        try: ctx.storage_state(path=STATE_FILE)
                        except: pass
                        ctx.close()
                        # retry sekali non-headless
                        return scrape(keyword, max_results, headless=False, proxy=proxy)
                    time.sleep(8)

                links=[]; seen=set()
                for a in page.query_selector_all("a[href*='/models/'], a[href*='/model/']"):
                    h=a.get_attribute("href")
                    if not h: continue
                    if h.startswith("/"): h=urljoin("https://makerworld.com", h)
                    if h not in seen: seen.add(h); links.append(h)
                    if len(links)>=max_results: break
                print(f"[{engine}] links:", len(links))

                for link in links[:max_results]:
                    page.goto(link, wait_until="domcontentloaded")
                    try: page.wait_for_load_state("networkidle", timeout=12000)
                    except: pass

                    title=""
                    try:
                        h1=page.query_selector("h1")
                        if h1: title=h1.inner_text().strip()
                    except: pass
                    if not title:
                        try:
                            og=page.query_selector("meta[property='og:title']")
                            if og: title=(og.get_attribute("content") or "").strip()
                        except: pass

                    desc=""
                    try:
                        md=page.query_selector("meta[name='description']")
                        if md: desc=(md.get_attribute("content") or "").strip()
                    except: pass

                    auto_scroll(page)
                    gal=get_gallery_urls(page)

                    results.append({
                        "title": title or link,
                        "description": desc,
                        "url": link,
                        "image_urls": gal
                    })
                try: ctx.storage_state(path=STATE_FILE)
                except: pass
                ctx.close()
                if results: break
            except Exception as e:
                print(f"[{engine}] error:", e)
                try:
                    if ctx: ctx.close()
                except: pass
    return results

# =========================
# Main
# =========================

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("-k","--keyword", required=True)
    ap.add_argument("-m","--max", type=int, default=1)
    ap.add_argument("--template", required=True)
    ap.add_argument("-o","--out", default="shopee_ready.xlsx")
    ap.add_argument("--category-id", required=True)
    ap.add_argument("--brand", default="No Brand")
    ap.add_argument("--price", type=float, default=45000)
    ap.add_argument("--stock", type=int, default=20)
    ap.add_argument("--weight-kg", type=float, default=0.15)
    ap.add_argument("--dims", default="10,10,3")  # L,W,H
    ap.add_argument("--sku-prefix", dest="sku_prefix", default="MW")
    ap.add_argument("--proxy", default=None)
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--sheet", default=None, help="Nama sheet bila template punya banyak tab")
    args=ap.parse_args()

    dims=tuple(int(x) for x in args.dims.split(",")[:3])

    print("[*] scraping …")
    recs=scrape(args.keyword, args.max, args.headless, args.proxy)
    if not recs:
        print("[!] Tidak ada hasil. Coba jalankan tanpa --headless atau pakai proxy residensial.")
        return

    rows=[]
    for i, r in enumerate(recs, start=1):
        rows.append({
            "category_id": args.category_id,
            "name": seo_title(r["title"]),
            "description": seo_desc(r["description"]),
            "image_urls": r.get("image_urls", [])[:8],
            "price": args.price,
            "stock": args.stock,
            "weight_kg": args.weight_kg,
            "dims_cm": dims,
            "sku": f"{args.sku_prefix}-{i:04d}",
        })

    write_rows_to_shopee_template(args.template, args.out, rows, sheet_name=args.sheet)
    print("✅ Selesai.")

if __name__ == "__main__":
    main()
