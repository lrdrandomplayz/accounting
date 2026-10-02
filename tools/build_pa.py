"""Bina buku kerja Excel untuk Prinsip Akaun (Bahasa Melayu).

Jalankan dari akar repositori:  python3 tools/build_pa.py
Menulis prinsip-akaun/excel/*.xlsx dan tools/manifest_pa.json (digunakan oleh tools/build_docs.py).
"""
import json
from pathlib import Path

from openpyxl import Workbook

from engine import D, F_ANS, F_DATA, F_INPUT, INT, NUM, T, V, Page

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "prinsip-akaun" / "excel"
TAB_README, TAB_WORKED, TAB_Q, TAB_A = "1F3864", "ED7D31", "2F5597", "70AD47"


def q(t):
    return f'"{t}"'


# =====================================================================================
# Persamaan Perakaunan
# =====================================================================================
PP_COLS = ["Urus niaga", "Tunai", "Bank", "Inventori", "Perabot", "Akaun Belum Terima",
           "Akaun Belum Bayar", "Pinjaman", "Ekuiti Pemilik"]
PP_KEYS = ["tu", "ba", "in", "pe", "abt", "abb", "pin", "ep"]
PP_CONTOH = [
    ("1. Memulakan perniagaan: RM20 000 ke bank dan RM2 000 tunai", {"ba": 20000, "tu": 2000, "ep": 22000}),
    ("2. Membeli perabot dengan cek RM3 000", {"pe": 3000, "ba": -3000}),
    ("3. Membeli barang niaga secara kredit daripada Kedai Siti RM5 000", {"in": 5000, "abb": 5000}),
    ("4. Menjual barang niaga berkos RM2 000 secara kredit kepada Encik Lim, RM3 200", {"in": -2000, "abt": 3200, "ep": 1200}),
    ("5. Menerima pinjaman bank RM10 000", {"ba": 10000, "pin": 10000}),
    ("6. Membayar sewa kedai secara tunai RM800", {"tu": -800, "ep": -800}),
    ("7. Menerima cek RM3 000 daripada Encik Lim", {"ba": 3000, "abt": -3000}),
    ("8. Membayar Kedai Siti dengan cek RM4 800 dan menerima diskaun RM200", {"ba": -4800, "abb": -5000, "ep": 200}),
    ("9. Pemilik mengambil tunai RM500 untuk kegunaan sendiri", {"tu": -500, "ep": -500}),
    ("10. Pemilik menambah modal: RM5 000 dimasukkan ke bank", {"ba": 5000, "ep": 5000}),
]


def pp_contoh(p):
    p.text("Contoh: Perniagaan Aiman, Januari 2025. Setiap urus niaga memberi kesan kepada sekurang-kurangnya dua "
           "item, tetapi Aset = Liabiliti + Ekuiti Pemilik sentiasa seimbang.")
    rows = []
    for i, (label, eff) in enumerate(PP_CONTOH):
        rows.append((label, [D(eff.get(k, 0), f"{k}{i}") for k in PP_KEYS]))
    rows.append(("Baki akhir", [V("+".join(f"{{{k}{i}}}" for i in range(len(PP_CONTOH))), key=f"{k}_t", line="total")
                                for k in PP_KEYS]))
    p.tag("pa.pp.contoh").table(PP_COLS, rows, 7, 3, checks=False, total_rows=("Baki akhir",))
    p.note("Aset: Tunai hingga Akaun Belum Terima. Liabiliti: Akaun Belum Bayar dan Pinjaman. Angka negatif bermaksud "
           "berkurang.")
    p.gap()
    p.tag("pa.pp.semak").calc([
        ("ta", "Jumlah Aset (RM)", "{tu_t}+{ba_t}+{in_t}+{pe_t}+{abt_t}"),
        ("tle", "Jumlah Liabiliti + Ekuiti Pemilik (RM)", "{abb_t}+{pin_t}+{ep_t}"),
        ("bz", "Perbezaan (mesti 0)", "{ta}-{tle}")])
    p.part("Penyata Kedudukan Kewangan Perniagaan Aiman pada 31 Januari 2025")
    p.tag("pa.pp.pkk").statement(["Perniagaan Aiman", "Penyata Kedudukan Kewangan pada 31 Januari 2025"], [
        ("", "Aset", {}, "h"),
        ("", "Perabot", {2: "{pe_t}"}),
        ("", "Inventori", {2: "{in_t}"}),
        ("", "Akaun Belum Terima", {2: "{abt_t}"}),
        ("", "Bank", {2: "{ba_t}"}),
        ("", "Tunai", {2: V("{tu_t}", line="under")}),
        ("", "Jumlah Aset", {3: V("{ta}", line="total")}, "b"),
        ("", "Liabiliti", {}, "h"),
        ("", "Pinjaman bank", {2: "{pin_t}"}),
        ("", "Akaun Belum Bayar", {2: V("{abb_t}", line="under")}),
        ("", "Jumlah Liabiliti", {3: "{pin_t}+{abb_t}"}),
        ("", "Ekuiti Pemilik", {3: V("{ep_t}", line="under")}),
        ("", "Jumlah Liabiliti dan Ekuiti Pemilik", {3: V("{tle}", line="total")}, "b")])


S1_TX = [
    ("1. Nurul memulakan perniagaan dengan RM15 000 di bank dan sebuah kenderaan bernilai RM25 000.", (40000, 0, 40000)),
    ("2. Membeli barang niaga secara kredit RM3 600 daripada Syarikat Bina.", (3600, 3600, 0)),
    ("3. Membeli mesin RM6 000. RM2 000 dibayar dengan cek dan bakinya secara kredit.", (4000, 4000, 0)),
    ("4. Menjual barang niaga berkos RM1 500 secara tunai dengan harga RM2 300.", (800, 0, 800)),
    ("5. Membayar gaji pekerja RM900 dengan cek.", (-900, 0, -900)),
    ("6. Membayar Syarikat Bina RM3 400 dengan cek dan menerima diskaun RM200.", (-3400, -3600, 200)),
    ("7. Pemilik mengambil barang niaga berkos RM300 untuk kegunaan sendiri.", (-300, 0, -300)),
    ("8. Menerima komisen RM450 secara tunai.", (450, 0, 450)),
]


def pp_s1(p):
    p.text("Perniagaan Nurul memulakan operasi pada 1 Mac 2025. Urus niaga berikut berlaku pada bulan Mac 2025.")
    p.part("(a) Lengkapkan jadual berikut untuk menunjukkan kesan bersih setiap urus niaga ke atas Aset, Liabiliti "
           "dan Ekuiti Pemilik. Gunakan + untuk bertambah, - untuk berkurang dan 0 jika tiada kesan. [8 markah]")
    rows = [(label, [V(str(a), key=f"a{i}", signed=True), V(str(l), key=f"l{i}", signed=True),
                     V(str(e), key=f"e{i}", signed=True)]) for i, (label, (a, l, e)) in enumerate(S1_TX)]
    p.table(["Urus niaga", "Aset (RM)", "Liabiliti (RM)", "Ekuiti Pemilik (RM)"], rows, 13, 5)
    n = range(len(S1_TX))
    p.part("(b) Hitung jumlah Aset, Liabiliti dan Ekuiti Pemilik pada 31 Mac 2025. [2 markah]")
    p.calc([("ta", "Jumlah Aset (RM)", "+".join(f"{{a{i}}}" for i in n)),
            ("tl", "Jumlah Liabiliti (RM)", "+".join(f"{{l{i}}}" for i in n)),
            ("te", "Ekuiti Pemilik (RM)", "+".join(f"{{e{i}}}" for i in n))])


def pp_s2(p):
    p.text("Jawab semua bahagian. Tunjukkan jalan kira.")
    p.part("(a) Aset sebuah perniagaan berjumlah RM96 500 dan liabilitinya RM31 200. Hitung ekuiti pemilik. [1 markah]")
    p.data([("a1", "Aset (RM)", 96500), ("l1", "Liabiliti (RM)", 31200)])
    p.calc([("ep1", "Ekuiti pemilik (RM)", "{a1}-{l1}")])
    p.part("(b) Maklumat berikut diperoleh daripada Perniagaan Hafiz bagi tahun berakhir 31 Disember 2025. "
           "Hitung untung bersih dan ekuiti pemilik pada 31 Disember 2025. [3 markah]")
    p.data([("m0", "Modal pada 1 Januari 2025 (RM)", 48000), ("mt", "Modal tambahan (RM)", 5000),
            ("h", "Hasil (RM)", 36500), ("b", "Belanja (RM)", 21800), ("am", "Ambilan (RM)", 6200)], md=True)
    p.calc([("ub", "Untung bersih (RM)", "{h}-{b}"),
            ("ep2", "Ekuiti pemilik pada 31 Disember 2025 (RM)", "{m0}+{mt}+{ub}-{am}")])
    p.part("(c) Liabiliti Perniagaan Hafiz pada 31 Disember 2025 ialah RM18 300. Hitung jumlah aset. [1 markah]")
    p.data([("l2", "Liabiliti (RM)", 18300)])
    p.calc([("a2", "Jumlah aset (RM)", "{ep2}+{l2}")])
    p.part("(d) Baki berikut diambil daripada buku Kedai Ros. Hitung baki tunai. [3 markah]")
    p.data([("bk", "Bank (RM)", 12000), ("inv", "Inventori (RM)", 9500), ("ken", "Kenderaan (RM)", 30000),
            ("abt", "Akaun Belum Terima (RM)", 4200), ("lia", "Liabiliti (RM)", 15000),
            ("epr", "Ekuiti Pemilik (RM)", 42000)], md=True)
    p.calc([("tla", "Jumlah Liabiliti + Ekuiti Pemilik (RM)", "{lia}+{epr}"),
            ("tun", "Baki tunai (RM)", "{tla}-({bk}+{inv}+{ken}+{abt})")])


# =====================================================================================
# Lejar dan Akaun Kawalan
# =====================================================================================
def ak_contoh(p):
    p.text("Contoh: Perniagaan Mesra, tahun berakhir 31 Disember 2025. Maklumat diambil daripada jurnal-jurnal dan "
           "Buku Tunai.")
    p.data([("o", "Akaun Belum Terima pada 1 Januari 2025 (RM)", 12400), ("js", "Jualan kredit (RM)", 48600),
            ("bk", "Tunai dan cek diterima daripada pelanggan (RM)", 41250), ("dd", "Diskaun diberi (RM)", 1150),
            ("pj", "Pulangan jualan (RM)", 860), ("hl", "Hutang lapuk (RM)", 540),
            ("ctl", "Cek pelanggan tak laku (RM)", 700), ("ko", "Kontra (RM)", 380),
            ("fa", "Faedah dikenakan atas akaun tertunggak (RM)", 60)], title="Maklumat: Akaun Belum Terima")
    close = "{o}+{js}+{ctl}+{fa}-({bk}+{dd}+{pj}+{hl}+{ko})"
    p.tag("pa.ak.akbt").ledger("Akaun Kawalan Belum Terima", [
        (("2025 Jan 1", "Baki b/b", "", "{o}"), ("2025 Dis 31", "Bank", "BT1", "{bk}")),
        (("Dis 31", "Jualan", "JJ1", "{js}"), ("Dis 31", "Diskaun diberi", "BT1", "{dd}")),
        (("Dis 31", "Bank (cek tak laku)", "BT1", "{ctl}"), ("Dis 31", "Pulangan jualan", "JPJ1", "{pj}")),
        (("Dis 31", "Faedah", "JA1", "{fa}"), ("Dis 31", "Hutang lapuk", "JA1", "{hl}")),
        (None, ("Dis 31", "Kontra", "JA1", "{ko}")),
        (None, ("Dis 31", "Baki h/b", "", close)),
        "TOTAL",
        (("2026 Jan 1", "Baki b/b", "", close), None)])
    p.gap()
    p.data([("o2", "Akaun Belum Bayar pada 1 Januari 2025 (RM)", 9300), ("bl", "Belian kredit (RM)", 36800),
            ("bb", "Cek dibayar kepada pembekal (RM)", 34100), ("dt", "Diskaun diterima (RM)", 920),
            ("pb", "Pulangan belian (RM)", 640), ("ko2", "Kontra (RM)", 380),
            ("am", "Angkutan masuk dikenakan oleh pembekal (RM)", 150)], title="Maklumat: Akaun Belum Bayar")
    close2 = "{o2}+{bl}+{am}-({bb}+{dt}+{pb}+{ko2})"
    p.tag("pa.ak.akbb").ledger("Akaun Kawalan Belum Bayar", [
        (("2025 Dis 31", "Bank", "BT1", "{bb}"), ("2025 Jan 1", "Baki b/b", "", "{o2}")),
        (("Dis 31", "Diskaun diterima", "BT1", "{dt}"), ("Dis 31", "Belian", "JB1", "{bl}")),
        (("Dis 31", "Pulangan belian", "JPB1", "{pb}"), ("Dis 31", "Angkutan masuk", "JA1", "{am}")),
        (("Dis 31", "Kontra", "JA1", "{ko2}"), None),
        (("Dis 31", "Baki h/b", "", close2), None),
        "TOTAL",
        (None, ("2026 Jan 1", "Baki b/b", "", close2))])


def ak_kontra(p):
    p.text("Contoh catatan kontra: Pembekal Yani ialah pembekal dan juga pelanggan Perniagaan Wawan. Pada 3 Januari "
           "2016, Perniagaan Wawan membeli barang niaga RM1 500 daripada Pembekal Yani. Pada 5 Januari 2016, "
           "Perniagaan Wawan menjual barang niaga RM450 kepada Pembekal Yani. Baki yang lebih kecil (RM450) "
           "dipindahkan ke baki yang lebih besar.")
    p.data([("b", "Belian daripada Pembekal Yani (RM)", 1500), ("j", "Jualan kepada Pembekal Yani (RM)", 450)])
    p.part("Lejar Belian")
    p.tag("pa.ak.kontra_belian").ledger("Akaun Belum Bayar: Pembekal Yani", [
        (("2016 Jan 31", "Kontra", "JA1", "MIN({b},{j})"), ("2016 Jan 3", "Belian", "JB1", "{b}")),
        (("Jan 31", "Baki h/b", "", "{b}-MIN({b},{j})"), None),
        "TOTAL",
        (None, ("2016 Feb 1", "Baki b/b", "", "{b}-MIN({b},{j})"))])
    p.part("Lejar Jualan")
    p.tag("pa.ak.kontra_jualan").ledger("Akaun Belum Terima: Pembekal Yani", [
        (("2016 Jan 5", "Jualan", "JJ1", "{j}"), ("2016 Jan 31", "Kontra", "JA1", "MIN({b},{j})")),
        "TOTAL"])


def ak_s3(p):
    p.text("Encik Nizam ialah pemilik Kedai Alatulis Nizam. Maklumat berikut diperoleh daripada buku perniagaannya "
           "bagi tahun berakhir 30 Jun 2025.")
    p.data([(None, "Baki"),
            ("o1", "Akaun Belum Terima pada 1 Julai 2024 (RM)", 15200), ("c1", "Akaun Belum Terima pada 30 Jun 2025 (RM)", 18450),
            ("o2", "Akaun Belum Bayar pada 1 Julai 2024 (RM)", 9800), ("c2", "Akaun Belum Bayar pada 30 Jun 2025 (RM)", 11250),
            (None, "Ringkasan Akaun Bank"),
            ("bt", "Diterima daripada Akaun Belum Terima (RM)", 72300), ("bb", "Dibayar kepada Akaun Belum Bayar (RM)", 54600),
            (None, "Maklumat tambahan"),
            ("dd", "Diskaun diberi (RM)", 2100), ("dt", "Diskaun diterima (RM)", 1350),
            ("pj", "Pulangan jualan (RM)", 950), ("pb", "Pulangan belian (RM)", 720),
            ("hl", "Hutang lapuk dihapuskan (RM)", 800), ("ctl", "Cek pelanggan tak laku (RM)", 1200),
            ("ko", "Kontra antara Lejar Jualan dan Lejar Belian (RM)", 600),
            ("am", "Angkutan masuk dikenakan oleh pembekal (RM)", 300)], md=True)
    js = "{c1}+{bt}+{dd}+{pj}+{hl}+{ko}-{o1}-{ctl}"
    bl = "{c2}+{bb}+{dt}+{pb}+{ko}-{o2}-{am}"
    p.part("(a) Hitung jualan kredit dan belian kredit bagi tahun berakhir 30 Jun 2025. [2 markah]")
    p.calc([("js", "Jualan kredit (RM)", js), ("bl", "Belian kredit (RM)", bl)])
    p.part("(b) Sediakan Akaun Kawalan Belum Terima bagi tahun berakhir 30 Jun 2025. [6 markah]")
    p.ledger("Akaun Kawalan Belum Terima", [
        (("2024 Jul 1", "Baki b/b", "", "{o1}"), ("2025 Jun 30", "Bank", "BT1", "{bt}")),
        (("2025 Jun 30", "Jualan", "JJ1", "{js}"), ("Jun 30", "Diskaun diberi", "BT1", "{dd}")),
        (("Jun 30", "Bank (cek tak laku)", "BT1", "{ctl}"), ("Jun 30", "Pulangan jualan", "JPJ1", "{pj}")),
        (None, ("Jun 30", "Hutang lapuk", "JA1", "{hl}")),
        (None, ("Jun 30", "Kontra", "JA1", "{ko}")),
        (None, ("Jun 30", "Baki h/b", "", "{c1}")),
        "TOTAL",
        (("2025 Jul 1", "Baki b/b", "", "{c1}"), None)])
    p.part("(c) Sediakan Akaun Kawalan Belum Bayar bagi tahun berakhir 30 Jun 2025. [5 markah]")
    p.ledger("Akaun Kawalan Belum Bayar", [
        (("2025 Jun 30", "Bank", "BT1", "{bb}"), ("2024 Jul 1", "Baki b/b", "", "{o2}")),
        (("Jun 30", "Diskaun diterima", "BT1", "{dt}"), ("2025 Jun 30", "Belian", "JB1", "{bl}")),
        (("Jun 30", "Pulangan belian", "JPB1", "{pb}"), ("Jun 30", "Angkutan masuk", "JA1", "{am}")),
        (("Jun 30", "Kontra", "JA1", "{ko}"), None),
        (("Jun 30", "Baki h/b", "", "{c2}"), None),
        "TOTAL",
        (None, ("2025 Jul 1", "Baki b/b", "", "{c2}"))])


def ak_s4(p):
    p.text("Kedai Lina ialah pelanggan dan juga pembekal kepada Perniagaan Danial. Urus niaga berikut berlaku pada "
           "bulan Januari 2025.")
    p.bullets(["1 Januari: Kedai Lina berhutang RM2 000 kepada Perniagaan Danial.",
               "6 Januari: Menjual barang niaga secara kredit kepada Kedai Lina RM1 800.",
               "12 Januari: Kedai Lina memulangkan barang niaga RM200.",
               "15 Januari: Membeli barang niaga secara kredit daripada Kedai Lina RM900.",
               "20 Januari: Menerima cek RM1 500 daripada Kedai Lina.",
               "31 Januari: Baki dalam Lejar Belian dipindahkan ke Lejar Jualan (kontra)."])
    p.data([("b0", "Baki pada 1 Januari (RM)", 2000), ("js", "Jualan kredit (RM)", 1800), ("pj", "Pulangan jualan (RM)", 200),
            ("bl", "Belian kredit (RM)", 900), ("ck", "Cek diterima (RM)", 1500)])
    ko = "MIN({bl},{b0}+{js}-{pj}-{ck})"
    p.part("(a) Sediakan Akaun Belum Terima Kedai Lina dalam Lejar Jualan dalam bentuk berbaki. [5 markah]")
    p.table(["Tarikh dan butir (2025)", "Debit (RM)", "Kredit (RM)", "Baki (RM)"], [
        ("Jan 1  Baki b/b", [None, None, V("{b0}", key="x0")]),
        ("Jan 6  Jualan", [V("{js}"), None, V("{x0}+{js}", key="x1")]),
        ("Jan 12  Pulangan jualan", [None, V("{pj}"), V("{x1}-{pj}", key="x2")]),
        ("Jan 20  Bank", [None, V("{ck}"), V("{x2}-{ck}", key="x3")]),
        ("Jan 31  Kontra", [None, V(ko), V(f"{{x3}}-{ko}")])], 10, 6)
    p.part("(b) Sediakan Akaun Belum Bayar Kedai Lina dalam Lejar Belian. [2 markah]")
    p.ledger("Akaun Belum Bayar: Kedai Lina", [
        (("2025 Jan 31", "Kontra", "JA1", ko), ("2025 Jan 15", "Belian", "JB1", "{bl}")),
        "TOTAL"])
    p.part("(c) Di sebelah manakah catatan kontra direkodkan dalam Akaun Kawalan Belum Terima? [1 markah]")
    p.choice([("s", "Sebelah", q("Kredit"), ["Debit", "Kredit"])])
    p.part("(d) Nyatakan dua implikasi Akaun Kawalan. [2 markah]")
    p.written("Mana-mana dua:\nSukar mengesan kesilapan ketinggalan catatan.\nSukar mengesan kesilapan menjumlahkan "
              "Akaun Belum Terima dan Akaun Belum Bayar.\nSukar mengesan kesilapan memindahkan catatan.", lines=4)


SIDES = ["Debit AKBT", "Kredit AKBT", "Debit AKBB", "Kredit AKBB"]
SUMBER = ["Jurnal Jualan", "Jurnal Belian", "Jurnal Pulangan Jualan", "Jurnal Pulangan Belian", "Buku Tunai", "Jurnal Am"]
S5 = [("Jualan kredit", "Debit AKBT", "Jurnal Jualan"),
      ("Belian kredit", "Kredit AKBB", "Jurnal Belian"),
      ("Pulangan jualan", "Kredit AKBT", "Jurnal Pulangan Jualan"),
      ("Pulangan belian", "Debit AKBB", "Jurnal Pulangan Belian"),
      ("Cek diterima daripada pelanggan", "Kredit AKBT", "Buku Tunai"),
      ("Diskaun diterima", "Debit AKBB", "Buku Tunai"),
      ("Hutang lapuk dihapuskan", "Kredit AKBT", "Jurnal Am"),
      ("Cek pelanggan tak laku", "Debit AKBT", "Buku Tunai"),
      ("Kontra (dalam Akaun Kawalan Belum Bayar)", "Debit AKBB", "Jurnal Am"),
      ("Faedah dikenakan oleh pembekal", "Kredit AKBB", "Jurnal Am")]


def ak_s5(p):
    p.text("Bagi setiap item, pilih sebelah yang betul dalam Akaun Kawalan Belum Terima (AKBT) atau Akaun Kawalan "
           "Belum Bayar (AKBB), dan sumber maklumatnya.")
    rows = []
    for i, (item, side, src) in enumerate(S5):
        rows.append((f"s{i}", f"{i + 1}. {item}: sebelah", q(side), SIDES))
        rows.append((f"r{i}", f"{i + 1}. {item}: sumber maklumat", q(src), SUMBER))
    p.choice(rows)


# =====================================================================================
# Kertas 1 (objektif)
# =====================================================================================
K1 = [
    ("Encik Amri memulakan perniagaan dengan memasukkan RM30 000 ke dalam akaun bank perniagaan. Apakah kesan urus "
     "niaga ini?",
     ["Aset bertambah RM30 000; ekuiti pemilik bertambah RM30 000", "Aset bertambah RM30 000; liabiliti bertambah RM30 000",
      "Aset berkurang RM30 000; ekuiti pemilik bertambah RM30 000", "Liabiliti bertambah RM30 000; ekuiti pemilik berkurang RM30 000"], "A"),
    ("Kedai Nora membayar hutang kepada pembekal dengan cek RM2 400 dan menerima diskaun tunai RM100. Apakah kesan "
     "urus niaga ini ke atas aset, liabiliti dan ekuiti pemilik?",
     ["Aset berkurang RM2 500; liabiliti berkurang RM2 400; ekuiti pemilik berkurang RM100",
      "Aset berkurang RM2 400; liabiliti berkurang RM2 500; ekuiti pemilik bertambah RM100",
      "Aset bertambah RM2 400; liabiliti berkurang RM2 500; ekuiti pemilik bertambah RM100",
      "Aset berkurang RM2 400; liabiliti berkurang RM2 400; tiada kesan ke atas ekuiti pemilik"], "B"),
    ("Jumlah aset sebuah perniagaan ialah RM85 000 dan jumlah liabilitinya RM23 000. Berapakah ekuiti pemilik?",
     ["RM23 000", "RM108 000", "RM62 000", "RM85 000"], "C"),
    ("Modal awal Perniagaan Seri ialah RM40 000. Untung bersih tahun itu RM12 000 dan ambilan RM5 000. Berapakah "
     "ekuiti pemilik pada akhir tahun?", ["RM57 000", "RM52 000", "RM35 000", "RM47 000"], "D"),
    ("Urus niaga manakah yang tidak mengubah jumlah aset?",
     ["Membeli perabot dengan cek", "Membayar sewa secara tunai", "Menerima pinjaman bank",
      "Pemilik mengambil tunai untuk kegunaan sendiri"], "A"),
    ("Barang niaga berkos RM1 500 dijual secara kredit dengan harga RM2 100. Apakah kesan urus niaga ini?",
     ["Aset bertambah RM2 100; ekuiti pemilik bertambah RM2 100", "Aset bertambah RM600; ekuiti pemilik bertambah RM600",
      "Aset berkurang RM1 500; liabiliti bertambah RM2 100", "Aset bertambah RM600; liabiliti bertambah RM600"], "B"),
    ("Manakah peraturan merekod urus niaga ke dalam lejar yang betul?",
     ["Aset didebitkan apabila bertambah", "Liabiliti didebitkan apabila bertambah", "Hasil didebitkan apabila bertambah",
      "Ekuiti pemilik didebitkan apabila bertambah"], "A"),
    ("Perniagaan Danish membeli sebuah van secara kredit daripada Syarikat Motor Jaya. Apakah akaun yang didebitkan "
     "dan dikreditkan?",
     ["Debit Belian; Kredit Syarikat Motor Jaya", "Debit Syarikat Motor Jaya; Kredit Kenderaan",
      "Debit Kenderaan; Kredit Syarikat Motor Jaya", "Debit Kenderaan; Kredit Bank"], "C"),
    ("Antara berikut, yang manakah akaun nominal?", ["Premis", "Sewa diterima", "Akaun Belum Bayar", "Bank"], "B"),
    ("Lejar Jualan digunakan untuk merekodkan",
     ["semua akaun pembekal kredit", "akaun hasil dan belanja", "semua akaun pelanggan kredit",
      "akaun aset bukan semasa"], "C"),
    ("Akaun Kawalan Belum Terima menunjukkan butiran berikut. Debit: Baki b/b RM2 150; P RM18 400. Kredit: Bank "
     "RM15 600; Pulangan jualan RM420; Q RM300; Baki h/b RM4 230. Apakah P dan Q?",
     ["P: Jualan; Q: Hutang lapuk", "P: Kontra; Q: Jualan", "P: Diskaun diterima; Q: Belian",
      "P: Pulangan belian; Q: Faedah"], "A"),
    ("Apakah sumber maklumat bagi jualan kredit dalam Akaun Kawalan Belum Terima?",
     ["Buku Tunai", "Jurnal Jualan", "Jurnal Belian", "Jurnal Pulangan Belian"], "B"),
    ("Akaun Kawalan Belum Bayar berbaki debit apabila",
     ["belian kredit bertambah", "pembekal mengenakan faedah", "pulangan belian berkurang",
      "perniagaan terlebih bayar kepada pembekal"], "D"),
    ("Bagaimanakah catatan kontra RM500 direkodkan dalam Akaun Kawalan Belum Bayar?",
     ["Dikreditkan sebagai Kontra RM500", "Didebitkan sebagai Bank RM500", "Didebitkan sebagai Kontra RM500",
      "Dikreditkan sebagai Belian RM500"], "C"),
    ("Yang manakah implikasi Akaun Kawalan?",
     ["Memudahkan penyediaan imbangan duga", "Mengesan kesilapan dalam lejar khas dengan cepat",
      "Mencegah penyelewengan", "Sukar mengesan kesilapan ketinggalan catatan"], "D"),
]


def k1(p):
    p.text("Kertas 1: soalan objektif. Setiap soalan diikuti oleh empat pilihan jawapan A, B, C dan D. Pilih satu "
           "jawapan sahaja. [15 markah]")
    p.mcq([(f"m{i}", stem, opts, ans) for i, (stem, opts, ans) in enumerate(K1, 1)])


# =====================================================================================
# Assembly
# =====================================================================================
BOOKS = [
    ("Persamaan_Perakaunan.xlsx", "pa_pp", "Persamaan Perakaunan",
     [("Contoh", pp_contoh)],
     [("S1", "Soalan 1 (Kertas 2): Kesan urus niaga [10 markah]", pp_s1, True),
      ("S2", "Soalan 2 (Kertas 2): Pengiraan persamaan perakaunan [8 markah]", pp_s2, False)]),
    ("Akaun_Kawalan.xlsx", "pa_ak", "Lejar dan Akaun Kawalan",
     [("Contoh AK", ak_contoh), ("Contoh Kontra", ak_kontra)],
     [("S3", "Soalan 3 (Kertas 2): Akaun Kawalan Belum Terima dan Belum Bayar [13 markah]", ak_s3, False),
      ("S4", "Soalan 4 (Kertas 2): Catatan kontra dan lejar bentuk berbaki [10 markah]", ak_s4, False),
      ("S5", "Latihan tambahan: sebelah dan sumber maklumat", ak_s5, False)]),
    ("Kertas_1_Objektif.xlsx", "pa_k1", "Kertas 1: Soalan Objektif",
     [],
     [("K1", "Kertas 1: 15 soalan objektif [15 markah]", k1, False)]),
]


def readme(wb, topic, worked, questions):
    p = Page(wb, "Baca Saya", "w", tab=TAB_README, lang="ms")
    p.title(topic, "Prinsip Akaun. Percuma untuk digunakan dan dikongsi. Bukan untuk dijual.")
    p.part("Cara menggunakan buku kerja ini")
    for line in ["1. Baca nota dahulu. Kemudian buka tab Contoh (oren): tukar mana-mana angka biru dan lihat jawapan berubah.",
                 "2. Cuba tab soalan (biru). Taip jawapan dalam sel kuning.",
                 "3. Lajur kecil di sebelah setiap sel kuning menunjukkan ✓ (betul) atau ✗ (cuba lagi). "
                 "Markah anda di bahagian atas.",
                 "4. Tarikh dan butir diberikan supaya helaian dapat menyemak angka anda. Cuba tulis akaun di atas "
                 "kertas dahulu.",
                 "5. Tab Jawapan (hijau) menunjukkan jawapan penuh dalam format peperiksaan."]:
        p.text(line, record=False)
    p.part("Kunci warna")
    r = p.r
    p.box(r, 0, 5, 1000, color="0000FF", fill_=F_DATA, fmt=NUM, align="right")
    p.box(r, 6, 25, "Angka biru: maklumat soalan. Guru boleh menukarnya untuk membuat versi baharu.")
    p.box(r + 1, 0, 5, 1000, fmt=NUM, align="right")
    p.box(r + 1, 6, 25, "Angka hitam: formula. Jangan taip di atasnya.")
    p.box(r + 2, 0, 5, None, fill_=F_INPUT)
    p.box(r + 2, 6, 25, "Sel kuning: taip jawapan anda di sini (nombor sahaja; 0 jika tiada).")
    p.box(r + 3, 0, 5, 1000, bold=True, fill_=F_ANS, fmt=NUM, align="right")
    p.box(r + 3, 6, 25, "Sel hijau: jawapan yang betul (tab Jawapan).")
    p.r += 4
    p.part("Helaian dalam buku kerja ini")
    rows = [(n, [T("Contoh (tab oren)")]) for n, _ in worked]
    rows += [(sheet, [T(f"{heading}. Jawapan: {sheet} Jawapan (tab hijau).")]) for sheet, heading, _f, _s in questions]
    p.table(["Helaian", "Kandungan"], rows, 7, 24, checks=False)
    p.finish()


def build():
    OUT.mkdir(parents=True, exist_ok=True)
    manifest = {}
    for fname, chap, topic, worked, questions in BOOKS:
        wb = Workbook()
        wb.remove(wb.active)
        readme(wb, topic, worked, questions)
        sheets = []
        for name, fn in worked:
            p = Page(wb, name, "w", tab=TAB_WORKED, lang="ms")
            p.title(name)
            p.status()
            fn(p)
            p.finish()
            sheets.append({"sheet": name, "mode": "w", "blocks": p.blocks})
        pages = []
        for sheet, heading, fn, signed in questions:
            ans = f"{sheet} Jawapan"
            pq = Page(wb, sheet, "q", twin=ans, tab=TAB_Q, lang="ms", signed=signed)
            pq.title(heading)
            pq.status()
            fn(pq)
            pages.append((pq, sheet, ans, heading, fn, signed))
        for pq, sheet, ans, heading, fn, signed in pages:
            pa = Page(wb, ans, "a", twin=sheet, tab=TAB_A, lang="ms", signed=signed)
            pa.title(heading + ": Jawapan")
            pa.status()
            fn(pa)
            assert pa.r == pq.r, f"layout mismatch in {sheet}"
            pq.finish()
            pa.finish()
            meta = {"qid": sheet, "level": 0, "section": "pa", "label": heading}
            sheets.append({"sheet": sheet, "mode": "q", **meta, "blocks": pq.blocks})
            sheets.append({"sheet": ans, "mode": "a", **meta, "blocks": pa.blocks})
        wb.active = 0
        wb.save(OUT / fname)
        manifest[fname] = {"chapter": chap, "path": "prinsip-akaun/excel", "lang": "ms", "sheets": sheets}
        print("wrote", fname)
    (ROOT / "tools" / "manifest_pa.json").write_text(json.dumps(manifest, indent=1))


if __name__ == "__main__":
    build()
