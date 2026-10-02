# Prinsip Akaun

Nota, fail Excel dan soalan gaya SPM bagi Prinsip Perakaunan dalam Bahasa Melayu. Nota mempunyai terjemahan bahasa Inggeris. *The notes include English translations.* Percuma untuk digunakan, dicetak dan dikongsi. Bukan untuk dijual. Semua amaun dalam Ringgit Malaysia (RM).

## Kandungan

| Topik | Nota | Excel | PDF |
|---|---|---|---|
| Persamaan Perakaunan | [nota](nota/Persamaan_Perakaunan.md) | [xlsx](excel/Persamaan_Perakaunan.xlsx) | [pdf](pdf/Persamaan_Perakaunan.pdf) |
| Lejar dan Akaun Kawalan | [nota](nota/Lejar_dan_Akaun_Kawalan.md) | [xlsx](excel/Akaun_Kawalan.xlsx) | [pdf](pdf/Lejar_dan_Akaun_Kawalan.pdf) |
| Glosari: Bahasa Melayu, English, 中文 | [glosari](nota/Glosari_PA.md) | | [pdf](pdf/Glosari_PA.pdf) |
| Kertas 1: 15 soalan objektif | [soalan](soalan/Soalan_SPM.md) | [xlsx](excel/Kertas_1_Objektif.xlsx) | [pdf](pdf/Soalan_SPM.pdf) |

* Soalan gaya SPM (Kertas 1 dan Kertas 2): [Soalan_SPM.md](soalan/Soalan_SPM.md) ([pdf](pdf/Soalan_SPM.pdf))
* Skema jawapan: [Skema_Jawapan.md](soalan/Skema_Jawapan.md) ([pdf](pdf/Skema_Jawapan.pdf))
* Pautan ke kertas percubaan SPM sebenar ada di bahagian Rujukan dalam `Soalan_SPM.md`.

## Cara belajar

1. Baca nota. Setiap topik ada objektif, istilah, contoh lengkap, kesilapan biasa dan semak kefahaman.
2. Buka fail Excel. Dalam tab oren **Contoh**, tukar mana-mana angka biru dan lihat jawapan berubah.
3. Jawab tab biru **S1** hingga **S5** dan **K1**. Taip jawapan dalam sel kuning. Lajur di sebelah menunjukkan ✓ (betul) atau ✗ (cuba lagi). Markah anda di bahagian atas.
4. Semak jawapan penuh dalam tab hijau **Jawapan**.
5. Jawab soalan yang sama di atas kertas, kemudian semak dengan skema jawapan.

## Format

| Item | Lajur |
|---|---|
| Lejar bentuk T | Tarikh, Butir, Folio, Amaun di setiap sebelah (Dt di kiri, Kt di kanan). Baris pertama menunjukkan tahun, dengan RM di bawah Amaun |
| Lejar bentuk berbaki | Tarikh, Butir, Debit, Kredit, Baki |
| Persamaan perakaunan | Urus niaga, Aset, Liabiliti, Ekuiti Pemilik. Tanda negatif (atau kurungan) bermaksud berkurang |

Baki b/b = baki bawa turun (awal tempoh). Baki h/b = baki hantar bawah (akhir tempoh).

Kekunci warna: angka biru ialah data yang boleh diubah, angka hitam ialah formula, sel kuning untuk jawapan anda, sel hijau menunjukkan jawapan betul.

Fail Excel dijana oleh `tools/build_pa.py`. Nota dan soalan dikemas kini oleh `tools/build_docs.py`.
