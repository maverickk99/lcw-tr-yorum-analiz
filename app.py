"""
Ürün Yorum Analiz Stüdyosu
Site ve mobil uygulama ürün yorumlarının dışa aktarımını (CSV / Excel) okuyup
puan, şikâyet teması ve ürün bazında alarm analizi üretir.
Veri sadece oturum belleğinde işlenir; hiçbir yere kaydedilmez.
"""
import io
import re
import random
from collections import Counter
from datetime import date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Ürün Yorum Analiz Stüdyosu", page_icon="🛍️", layout="wide")

# ---------------------------------------------------------------------------
# 1. ERİŞİM ŞİFRESİ
# Şifreyi Streamlit "Secrets" içine APP_PASSWORD = "..." olarak yazarsan oradan okunur.
# ---------------------------------------------------------------------------
def get_password():
    try:
        return st.secrets["APP_PASSWORD"]
    except Exception:
        return "LCW2026"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🛍️ Ürün Yorum Analiz Stüdyosu")
    st.caption("Site ve mobil uygulama ürün yorumları için tema, puan ve alarm analizi")
    left, _ = st.columns((1, 2))
    with left:
        pw = st.text_input("Erişim şifresi", type="password")
        if st.button("Giriş", type="primary"):
            if pw == get_password():
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Hatalı şifre.")
    st.stop()

# ---------------------------------------------------------------------------
# 2. SÖZLÜKLER
# ---------------------------------------------------------------------------
# Kelime kökleri: yorumda bu kökle BAŞLAYAN bir kelime varsa tema işaretlenir.
THEMES = {
    "Beden / Kalıp": ["beden", "kalıp", "kalib", "dar", "bol", "küçük", "küçü", "büyük", "büyü",
                      "uzun", "kısa", "oturm", "otur", "numara", "sıkı", "geniş"],
    "Kumaş / Kalite": ["kumaş", "kalite", "ince", "kalın", "terlet", "sentetik", "pamuk",
                       "tüylen", "tüy", "boncuk", "sert", "yumuşa"],
    "Yıkama / Dayanıklılık": ["yıka", "çekti", "çekme", "solma", "soldu", "sökül", "söküldü",
                              "dikiş", "yırtıl", "delik", "boya", "akıttı", "bozul"],
    "Renk / Görsel Farkı": ["renk", "rengi", "renkt", "fotoğraf", "görsel", "resim",
                            "farklı", "soluk", "ton"],
    "Kargo / Teslimat": ["kargo", "teslim", "geç", "gecik", "paket", "ulaş", "kurye"],
    "İade / Değişim": ["iade", "değişim", "değiştir", "geri"],
    "Fiyat / Değer": ["fiyat", "ucuz", "pahalı", "indirim", "para", "değer"],
}

POS_WORDS = ["güzel", "harika", "mükemmel", "beğendim", "beğen", "memnun", "tavsiye", "süper",
             "rahat", "kaliteli", "şık", "teşekkür", "bayıldım", "uygun", "tam oldu", "çok iyi"]
NEG_WORDS = ["kötü", "berbat", "rezalet", "beğenmedim", "memnun değil", "pişman", "iade",
             "sökül", "yırtıl", "solma", "soldu", "çekti", "olmadı", "değil", "maalesef",
             "hayal kırıklığı", "vasat", "tavsiye etmem", "almayın"]

STOP_WORDS = set("""ve ile bir bu ama için de da ne o çok daha gibi kadar olarak var yok mi mı mu mü
ben sen biz siz onlar şu her şey bi yani işte mesela aslında zaten fakat ancak veya ya hem ki en
sonra önce şimdi böyle şöyle nasıl neden kim hangi olan oldu olduğu ürün ürünü ürünün aldım
aldık geldi gerçekten biraz bence tam hiç sadece bile diye ise iyi güzel""".split())

# Dışa aktarımdaki sütun adlarını otomatik eşleştirmek için olası adlar
COLUMN_HINTS = {
    "yorum": ["yorum", "comment", "review", "metin", "text", "açıklama", "icerik", "içerik"],
    "puan": ["puan", "rating", "yıldız", "yildiz", "star", "score", "skor"],
    "urun_kodu": ["ürün kodu", "urun kodu", "ürünkodu", "sku", "product code", "productid",
                  "product_id", "urun_kodu", "model", "stok kodu"],
    "urun_adi": ["ürün adı", "urun adi", "ürün", "product name", "productname", "urun_adi", "ad"],
    "kategori": ["kategori", "category", "grup", "department", "reyon"],
    "tarih": ["tarih", "date", "created", "oluşturma", "yorum tarihi"],
    "kanal": ["kanal", "channel", "platform", "kaynak", "source", "cihaz"],
    "yazar": ["yazan", "müşteri", "musteri", "kullanıcı", "kullanici", "author", "customer",
              "isim", "ad soyad", "name", "email", "e-posta", "telefon"],
}

# ---------------------------------------------------------------------------
# 3. YARDIMCI FONKSİYONLAR
# ---------------------------------------------------------------------------
def tr_lower(text):
    """Türkçe doğru küçük harf: I -> ı, İ -> i."""
    return str(text).replace("I", "ı").replace("İ", "i").lower()

def guess_column(columns, key):
    # Hem Türkçe hem İngilizce küçük harf biçimini dene (PRODUCT_ID -> product_id / product_ıd)
    norm = {c: {tr_lower(c).strip(), str(c).lower().strip().replace("i̇", "i")} for c in columns}
    for hint in COLUMN_HINTS[key]:
        for col, forms in norm.items():
            if hint in forms:
                return col
    for hint in COLUMN_HINTS[key]:
        for col, forms in norm.items():
            if any(hint in f for f in forms):
                return col
    return None

def tag_themes(text):
    words = re.findall(r"[a-zçğıöşü]+", tr_lower(text))
    found = []
    for theme, roots in THEMES.items():
        if any(w.startswith(r) for w in words for r in roots):
            found.append(theme)
    return found

def text_sentiment(text):
    t = tr_lower(text)
    pos = sum(t.count(w) for w in POS_WORDS)
    neg = sum(t.count(w) for w in NEG_WORDS)
    if neg > pos:
        return "Olumsuz"
    if pos > neg:
        return "Olumlu"
    return "Nötr"

def sentiment(row, has_rating):
    if has_rating and pd.notna(row["puan"]):
        r = row["puan"]
        if r <= 2:
            return "Olumsuz"
        if r >= 4:
            return "Olumlu"
        return "Nötr"
    return text_sentiment(row["yorum"])

def top_words(texts, n=15):
    words = []
    for t in texts:
        words += [w for w in re.findall(r"[a-zçğıöşü]{3,}", tr_lower(t)) if w not in STOP_WORDS]
    return Counter(words).most_common(n)

def read_upload(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=None, engine="python", encoding=enc)
        except Exception:
            continue
    raise ValueError("Dosya okunamadı. CSV (UTF-8) veya Excel olarak kaydedip tekrar deneyin.")

def make_sample_data(n=600, seed=7):
    """Arayüzü denemek için SAHTE örnek veri. Gerçek müşteri yorumu değildir."""
    rnd = random.Random(seed)
    products = [
        ("P-1001", "Slim Fit Jean Pantolon", "Erkek Pantolon"),
        ("P-1002", "Basic Bisiklet Yaka Tişört", "Erkek Tişört"),
        ("P-2001", "Desenli Midi Elbise", "Kadın Elbise"),
        ("P-2002", "Örme Kazak", "Kadın Triko"),
        ("P-2003", "Yüksek Bel Tayt", "Kadın Spor"),
        ("P-3001", "Kapüşonlu Sweatshirt", "Çocuk Üst Giyim"),
        ("P-3002", "Pijama Takımı", "Çocuk İç Giyim"),
        ("P-4001", "Şişme Mont", "Erkek Dış Giyim"),
    ]
    good = ["Çok güzel, tam oldu, tavsiye ederim", "Kumaşı kaliteli ve rahat, memnunum",
            "Fiyatına göre harika, bayıldım", "Kargo hızlı geldi, ürün şık",
            "Rengi fotoğraftaki gibi, beğendim"]
    mid = ["Fena değil ama biraz ince", "Kalıbı biraz bol, bir beden küçük alın",
           "İdare eder, fiyatı uygun", "Rengi fotoğraftan biraz farklı ama olur"]
    bad = ["Kalıbı çok dar, iade ettim", "İlk yıkamada çekti, pişman oldum",
           "Dikişleri söküldü, kalitesi kötü", "Rengi görseldeki gibi değil, soluk geldi",
           "Kargo çok gecikti, paket hasarlı geldi", "Kumaş çok ince ve terletiyor, almayın",
           "Bir yıkamada tüylendi, boncuk boncuk oldu", "Beden tablosu yanlış, çok küçük geldi"]
    # Bazı ürünler kasıtlı olarak sorunlu, böylece alarm listesi dolu görünür
    weights = {"P-1001": (0.4, 0.2, 0.4), "P-2002": (0.35, 0.2, 0.45), "P-3002": (0.45, 0.25, 0.3)}
    rows = []
    start = date.today() - timedelta(days=180)
    for _ in range(n):
        code, pname, cat = rnd.choice(products)
        g, m, b = weights.get(code, (0.7, 0.2, 0.1))
        bucket = rnd.choices(["g", "m", "b"], weights=(g, m, b))[0]
        if bucket == "g":
            txt, r = rnd.choice(good), rnd.choice([4, 5, 5])
        elif bucket == "m":
            txt, r = rnd.choice(mid), 3
        else:
            txt, r = rnd.choice(bad), rnd.choice([1, 1, 2])
        rows.append({
            "Ürün Kodu": code, "Ürün Adı": pname, "Kategori": cat,
            "Puan": r, "Yorum": txt,
            "Tarih": start + timedelta(days=rnd.randint(0, 180)),
            "Kanal": rnd.choices(["Web", "iOS", "Android"], weights=(0.4, 0.25, 0.35))[0],
            "Müşteri Adı": f"Müşteri {rnd.randint(1000, 9999)}",
        })
    return pd.DataFrame(rows)

def to_excel(sheets):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=name[:31])
            ws = writer.sheets[name[:31]]
            for col_cells in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200])
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(10, width + 2), 60)
            ws.freeze_panes = "A2"
    return out.getvalue()

# ---------------------------------------------------------------------------
# 4. VERİ YÜKLEME
# ---------------------------------------------------------------------------
st.title("🛍️ Ürün Yorum Analiz Stüdyosu")
st.caption("Site ve mobil uygulama yorum dışa aktarımını yükleyin; ürün, tema ve kanal bazında içgörü alın.")

with st.expander("ℹ️ Hangi dosyayı yüklemeliyim?"):
    st.markdown("""
- E-ticaret / CRM ekibinden alınan yorum dışa aktarımı (**CSV veya Excel**).
- En az **yorum metni** sütunu gerekir. Puan, ürün kodu/adı, kategori, tarih ve kanal (Web / iOS / Android) varsa analiz zenginleşir.
- Sütun adları farklıysa aşağıda elle eşleştirebilirsiniz.
- **Gizlilik:** Müşteri adı, e-posta, telefon gibi sütunlar varsayılan olarak analizden ve çıktıdan **çıkarılır**. Veri sadece bu oturumda bellekte işlenir.
""")

c_up, c_demo = st.columns((3, 1))
with c_up:
    uploaded = st.file_uploader("Yorum dosyası", type=["csv", "xlsx", "xls"])
with c_demo:
    st.write("")
    st.write("")
    if st.button("🧪 Sahte örnek veriyle dene"):
        st.session_state.use_sample = True

if uploaded is not None:
    st.session_state.use_sample = False
    try:
        raw = read_upload(uploaded)
    except Exception as e:
        st.error(str(e))
        st.stop()
elif st.session_state.get("use_sample"):
    raw = make_sample_data()
    st.info("Şu an **sahte örnek veri** gösteriliyor (gerçek müşteri yorumu değildir).")
else:
    st.stop()

if raw.empty:
    st.error("Dosya boş görünüyor.")
    st.stop()

# ---------------------------------------------------------------------------
# 5. SÜTUN EŞLEŞTİRME
# ---------------------------------------------------------------------------
cols = list(raw.columns)
NONE = "— yok —"
labels = {
    "yorum": "Yorum metni *", "puan": "Puan (1-5)", "urun_kodu": "Ürün kodu",
    "urun_adi": "Ürün adı", "kategori": "Kategori", "tarih": "Tarih", "kanal": "Kanal",
}
with st.expander("🔧 Sütun eşleştirme", expanded=guess_column(cols, "yorum") is None):
    mapping = {}
    grid = st.columns(4)
    for i, (key, label) in enumerate(labels.items()):
        guess = guess_column(cols, key)
        options = [NONE] + cols
        idx = options.index(guess) if guess in options else 0
        mapping[key] = grid[i % 4].selectbox(label, options, index=idx, key=f"map_{key}")
    pii_guess = [c for c in cols if any(h in tr_lower(c) for h in COLUMN_HINTS["yazar"])
                 and c not in mapping.values()]
    pii_cols = st.multiselect("Kişisel veri sütunları (çıkarılacak)", cols, default=pii_guess)

if mapping["yorum"] == NONE:
    st.warning("Lütfen yorum metni sütununu seçin.")
    st.stop()

df = pd.DataFrame()
for key, col in mapping.items():
    df[key] = raw[col] if col != NONE else None

df["yorum"] = df["yorum"].fillna("").astype(str).str.strip()
df = df[df["yorum"].str.len() > 0].copy()
has_rating = mapping["puan"] != NONE
if has_rating:
    df["puan"] = pd.to_numeric(df["puan"], errors="coerce")
if mapping["tarih"] != NONE:
    df["tarih"] = pd.to_datetime(df["tarih"], errors="coerce", dayfirst=True)
if mapping["urun_kodu"] == NONE and mapping["urun_adi"] == NONE:
    df["urun"] = "Tüm ürünler"
else:
    code = df["urun_kodu"].astype(str) if mapping["urun_kodu"] != NONE else ""
    name = df["urun_adi"].astype(str) if mapping["urun_adi"] != NONE else ""
    df["urun"] = (code + " - " + name) if mapping["urun_kodu"] != NONE and mapping["urun_adi"] != NONE \
        else (code if mapping["urun_kodu"] != NONE else name)
for opt in ("kategori", "kanal"):
    df[opt] = df[opt].fillna("Belirtilmemiş").astype(str) if mapping[opt] != NONE else "Belirtilmemiş"

# ---------------------------------------------------------------------------
# 6. FİLTRELER
# ---------------------------------------------------------------------------
st.sidebar.header("Filtreler")
kanal_sec = st.sidebar.multiselect("Kanal", sorted(df["kanal"].unique()))
kat_sec = st.sidebar.multiselect("Kategori", sorted(df["kategori"].unique()))
if mapping["tarih"] != NONE and df["tarih"].notna().any():
    dmin, dmax = df["tarih"].min().date(), df["tarih"].max().date()
    tarih_sec = st.sidebar.date_input("Tarih aralığı", (dmin, dmax), min_value=dmin, max_value=dmax)
    if isinstance(tarih_sec, (list, tuple)) and len(tarih_sec) == 2:
        df = df[(df["tarih"].dt.date >= tarih_sec[0]) & (df["tarih"].dt.date <= tarih_sec[1])]
if kanal_sec:
    df = df[df["kanal"].isin(kanal_sec)]
if kat_sec:
    df = df[df["kategori"].isin(kat_sec)]
min_yorum = st.sidebar.number_input("Alarm için min. yorum sayısı", 1, 500, 10)
alarm_neg = st.sidebar.slider("Alarm eşiği: olumsuz oranı (%)", 10, 90, 35)

if df.empty:
    st.warning("Filtrelere uyan yorum yok.")
    st.stop()

# ---------------------------------------------------------------------------
# 7. ANALİZ
# ---------------------------------------------------------------------------
df["duygu"] = df.apply(lambda r: sentiment(r, has_rating), axis=1)
df["temalar"] = df["yorum"].apply(tag_themes)
df["temalar_metin"] = df["temalar"].apply(lambda t: ", ".join(t) if t else "Diğer")
neg = df[df["duygu"] == "Olumsuz"]

# Özet kartları
k1, k2, k3, k4 = st.columns(4)
k1.metric("Yorum sayısı", f"{len(df):,}".replace(",", "."))
k2.metric("Ürün sayısı", df["urun"].nunique())
k3.metric("Ortalama puan", f"{df['puan'].mean():.2f}" if has_rating and df["puan"].notna().any() else "-")
k4.metric("Olumsuz oran", f"%{len(neg) / len(df) * 100:.1f}")

# Tema analizi
theme_rows = []
for theme in list(THEMES) + ["Diğer"]:
    mask = df["temalar"].apply(lambda t: theme in t) if theme != "Diğer" else df["temalar"].apply(lambda t: not t)
    sub = df[mask]
    if sub.empty:
        continue
    theme_rows.append({
        "Tema": theme,
        "Yorum Sayısı": len(sub),
        "Olumsuz Yorum": int((sub["duygu"] == "Olumsuz").sum()),
        "Olumsuz Oran (%)": round((sub["duygu"] == "Olumsuz").mean() * 100, 1),
        "Ort. Puan": round(sub["puan"].mean(), 2) if has_rating else None,
    })
theme_df = pd.DataFrame(theme_rows).sort_values("Olumsuz Yorum", ascending=False)

# Ürün özeti
def main_complaint(sub):
    c = Counter(t for ts in sub.loc[sub["duygu"] == "Olumsuz", "temalar"] for t in ts)
    return ", ".join(f"{k} ({v})" for k, v in c.most_common(2)) or "-"

prod_rows = []
for urun, sub in df.groupby("urun"):
    prod_rows.append({
        "Ürün": urun,
        "Kategori": sub["kategori"].mode().iat[0],
        "Yorum Sayısı": len(sub),
        "Ort. Puan": round(sub["puan"].mean(), 2) if has_rating else None,
        "Olumsuz Oran (%)": round((sub["duygu"] == "Olumsuz").mean() * 100, 1),
        "Ana Şikâyet Temaları": main_complaint(sub),
        "Web": int((sub["kanal"] == "Web").sum()),
        "Mobil": int(sub["kanal"].isin(["iOS", "Android", "Mobil", "App"]).sum()),
    })
prod_df = pd.DataFrame(prod_rows).sort_values("Olumsuz Oran (%)", ascending=False)
alarm_df = prod_df[(prod_df["Yorum Sayısı"] >= min_yorum) & (prod_df["Olumsuz Oran (%)"] >= alarm_neg)]

# ---------------------------------------------------------------------------
# 8. GÖRÜNÜM
# ---------------------------------------------------------------------------
st.divider()
st.subheader(f"🚨 Alarm listesi ({len(alarm_df)} ürün)")
st.caption(f"En az {min_yorum} yorumu olan ve olumsuz oranı %{alarm_neg} üzerindeki ürünler")
if alarm_df.empty:
    st.success("Eşiği aşan ürün yok.")
else:
    st.dataframe(alarm_df, width="stretch", hide_index=True)

t1, t2, t3, t4, t5 = st.tabs(["🧩 Şikâyet temaları", "📦 Ürün özeti", "📱 Kanal karşılaştırma",
                              "📈 Zaman", "💬 Yorumlar"])
with t1:
    a, b = st.columns((1, 1))
    with a:
        st.markdown("**Temaya göre olumsuz yorum sayısı**")
        st.bar_chart(theme_df.set_index("Tema")["Olumsuz Yorum"], horizontal=True)
    with b:
        st.markdown("**Olumsuz yorumlarda en sık geçen kelimeler**")
        tw = top_words(neg["yorum"]) if not neg.empty else []
        if tw:
            st.bar_chart(pd.DataFrame(tw, columns=["Kelime", "Frekans"]).set_index("Kelime"), horizontal=True)
        else:
            st.info("Olumsuz yorum yok.")
    st.dataframe(theme_df, width="stretch", hide_index=True)
with t2:
    st.dataframe(prod_df, width="stretch", hide_index=True)
with t3:
    ch = df.groupby("kanal").agg(
        Yorum=("yorum", "size"),
        Olumsuz_Oran=("duygu", lambda s: round((s == "Olumsuz").mean() * 100, 1)),
    )
    if has_rating:
        ch["Ort_Puan"] = df.groupby("kanal")["puan"].mean().round(2)
    st.dataframe(ch.reset_index().rename(columns={"kanal": "Kanal"}), width="stretch", hide_index=True)
    if has_rating:
        dist = df.pivot_table(index="puan", columns="kanal", values="yorum", aggfunc="size", fill_value=0)
        st.markdown("**Puan dağılımı (kanal bazında)**")
        st.bar_chart(dist)
with t4:
    if mapping["tarih"] != NONE and df["tarih"].notna().any():
        w = df.dropna(subset=["tarih"]).set_index("tarih").groupby(pd.Grouper(freq="W"))
        trend = pd.DataFrame({
            "Yorum": w.size(),
            "Olumsuz Oran (%)": w["duygu"].apply(lambda s: (s == "Olumsuz").mean() * 100 if len(s) else None),
        })
        st.markdown("**Haftalık olumsuz yorum oranı (%)**")
        st.line_chart(trend["Olumsuz Oran (%)"])
        st.markdown("**Haftalık yorum sayısı**")
        st.bar_chart(trend["Yorum"])
    else:
        st.info("Tarih sütunu seçilmediği için zaman analizi yok.")
with t5:
    tema_filtre = st.selectbox("Temaya göre filtrele", ["Tümü"] + list(THEMES) + ["Diğer"])
    duygu_filtre = st.selectbox("Duygu", ["Tümü", "Olumsuz", "Nötr", "Olumlu"])
    view = df
    if tema_filtre != "Tümü":
        view = view[view["temalar_metin"].str.contains(tema_filtre, regex=False)]
    if duygu_filtre != "Tümü":
        view = view[view["duygu"] == duygu_filtre]
    show_cols = ["urun", "kategori", "kanal", "tarih", "puan", "duygu", "temalar_metin", "yorum"]
    st.dataframe(view[show_cols], width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# 9. EXCEL ÇIKTISI (kişisel veri sütunları hariç)
# ---------------------------------------------------------------------------
export = df[["urun", "kategori", "kanal", "tarih", "puan", "duygu", "temalar_metin", "yorum"]].rename(columns={
    "urun": "Ürün", "kategori": "Kategori", "kanal": "Kanal", "tarih": "Tarih", "puan": "Puan",
    "duygu": "Duygu", "temalar_metin": "Temalar", "yorum": "Yorum"})
if "Tarih" in export and export["Tarih"].notna().any():
    export["Tarih"] = pd.to_datetime(export["Tarih"], errors="coerce").dt.date
# Tema matrisi: her yorum için tema başına 0/1 (MAXQDA / SPSS için)
for theme in THEMES:
    export[theme] = df["temalar"].apply(lambda t: int(theme in t)).values

st.divider()
st.download_button(
    "📥 Analiz raporunu indir (.xlsx)",
    data=to_excel({
        "1_Alarm_Listesi": alarm_df,
        "2_Urun_Ozeti": prod_df,
        "3_Tema_Analizi": theme_df,
        "4_Tum_Yorumlar": export,
    }),
    file_name="urun_yorum_analizi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    width="stretch",
)
if pii_cols:
    st.caption(f"🔒 Çıktıya dahil edilmeyen kişisel veri sütunları: {', '.join(map(str, pii_cols))}")
