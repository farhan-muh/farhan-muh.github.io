---
title: "Menggunakan Referensi Absolut di Microsoft Excel"
description: "Referensi absolut di Excel dan bedanya dengan referensi relatif saat rumus ditarik dengan fill handle."
author: ["Muhamad Farhan"]
date: 2026-11-17
series: ["Rumus Excel"]
draft: false
series_order: 5
---

## Pembuka {#pembuka}

Pada artikel ini kita akan belajar cara menggunakan referensi absolut dan mencari tahu perbedaannya dengan referensi relatif.

Tapi referensi absolut itu apa sih? Misalnya saat kita menarik fill handle, ada sel referensi yang kita tidak mau alamatnya ikut bergeser. Nah di situlah kita pakai referensi absolut.

Terus kenapa kita perlu itu? Bukannya bagus kalau alamat sel itu mengikuti alur rumus yang kita buat? Kalau kita sedang menghitung satu baris data pada suatu tabel, lalu rumusnya ditarik ke bawah dengan fill handle, baris berikutnya ikut terisi. Alamat sel pada rumus yang dihitung juga ikut turun menyesuaikan. Ini yang disebut referensi relatif.

Namun ada kondisi seperti ini. Kadang agar mudah dibaca, beberapa nilai yang dihitung dipindahkan ke sel tersendiri. Contohnya suatu koperasi punya tabel perhitungan harga barang, semua barang saat ini sedang ada diskon 10%. Nilai diskon itu mereka simpan di luar tabel utama pada sel referensinya sendiri. Dengan begitu nilai diskon saat ini, lebih mudah diperiksa daripada diketik langsung pada rumus kan? Dan kalau mau diubah, kita tidak perlu mengubah langsung rumusnya.


## Cara Menulis Referensi Absolut {#cara-menulis-referensi-absolut}

Kita menggunakan tanda `$` (aku menyebutnya S dolar) untuk menandai sel yang menggunakan referensi absolut. Misalnya `$A$1`. Kalau sudah pakai `$` artinya sel itu sudah dikunci atau digembok, jadi tidak boleh bergeser kalau rumusnya ditarik.

Ada tiga variasi penulisan `$` untuk mengunci sel:

1.  **Kunci kolom dan baris**: tanda dolarnya ditulis di depan nama kolom dan baris, seperti ini `$A$1`. Jadi mau ditarik ke mana pun selnya akan tetap terkunci di sana.
2.  **Kunci barisnya saja**: tanda dolar ditulis di depan nama barisnya saja, seperti ini `A$1`.
3.  **Kunci kolomnya saja**: tanda dolar ditulis di depan nama kolomnya saja seperti ini `$A1`.

Hanya variasi pertama yaitu penguncian kolom dan baris yang disebut referensi absolut. Sementara kunci barisnya saja atau kolomnya saja disebut referensi campuran. Artikel ini akan mempersempit bahasan pada referensi absolut saja.


### Cara cepat dengan tombol **F4** {#cara-cepat-dengan-tombol-f4}

Daripada kita mengetik langsung `$` pada alamat sel, yang rawan salah ketik dan sedikit repot, kita pakai saja tombol **F4**.

Caranya seperti ini:

1.  Kita habis ketik rumus seperti ini `=C6*F2`.
2.  Terus `F2`-nya mau dikunci.
3.  Langsung saja tekan **F4** saat kursor ada di `F2`. Nanti jadinya seperti ini `$F$2`.

Pernah saya tekan **F4** tapi tidak berhasil menjadi referensi absolut, masih jadi sel biasa. Saya belum tahu penyebabnya apa, tapi saya sudah ketemu solusinya. Coba blok dulu `F2`-nya aja baru tekan **F4** lagi. Insya Allah berhasil.

> [!NOTE] Tombol **F4** tidak bisa digunakan untuk referensi absolut
>
> Pada sebagian laptop tombol **F1** sampai **F12** digabungkan dengan tombol fungsi tambahan.
>
> Jika tombol **F4** di laptop kamu ternyata berfungsi untuk menurunkan volume atau kecerahan layar, maka untuk menggunakan referensi absolut, kita perlu menekan tombol **Fn** bersamaan dengan tombol **F4**.
>
> Tombol **Fn** biasanya ada di pojok kiri bawah.

<!--quoteend-->

> [!IMPORTANT] Siklus Tombol **F4**
> Cukup tekan satu kali saja untuk referensi absolut seperti ini `$A$1`. Kalau ditekan lebih dari sekali akan jadi referensi campuran atau malah tidak terkunci.
>
> Begini urutannya: `$A$1` -&gt; `A$1` -&gt; `$A1` -&gt; `A1`.


## Menggunakan Referensi Absolut Pada Suatu Rumus {#menggunakan-referensi-absolut-pada-suatu-rumus}

Sebelumnya kita sempat bahas kalau koperasi sedang ada diskon 10% untuk setiap barang yang mereka jual. Mari kita hitung harga barang setelah diberikan diskon.

**Instruksi**: Buatlah tabel di bawah ini. Tabel utama dimulai dari sel `A1`, sementara tabel diskon dimulai dari sel `F1`.

<div class="table-caption">
  <span class="table-number">Table 1:</span>
  Tabel perhitungan harga diskon barang
</div>

| Nama Barang | Harga Awal | Diskon 10% | Harga Setelah Diskon |
|-------------|------------|------------|----------------------|
| Pensil      | 2000       |            |                      |
| Pulpen      | 2500       |            |                      |
| Buku Tulis  | 5000       |            |                      |
| Penggaris   | 2000       |            |                      |
| Penghapus   | 1000       |            |                      |

<div class="table-caption">
  <span class="table-number">Table 2:</span>
  Tabel besaran diskon
</div>

|        |     |
|--------|-----|
| Diskon | 10% |

Beginilah tampilannya di Excel:

{{< figure src="tabel-awal.webp" caption="<span class=\"figure-number\">Figure 1: </span>Tabel awal sebelum pakai rumus" >}}

Pertama-tama kita kerjakan kolom `Diskon 10%` dahulu:

1.  Pada kolom `Diskon 10%` yaitu di sel `C2`, ketiklah sama dengan `=`.
2.  Klik `Harga Awal` Pensil di sel `B2`, lalu lanjut ketik kali dari keyboard sehingga jadi `=B2*`.
3.  Klik `Diskon` 10% nya di sel `G1`, rumusnya sekarang jadi `=B2*G1`.
4.  Selagi kursornya sedang di `G1`, tekan **F4** sekali. Rumusnya jadi seperti ini sekarang `=B2*$G$1`.
5.  Tekan **Enter**, dan hasilnya adalah `200`.
6.  Tarik fill handle untuk mengisi baris berikutnya.

Berikut hasil setelah semua baris pada kolom `Diskon 10%` terisi:

{{< figure src="tabel-full-diskon.webp" caption="<span class=\"figure-number\">Figure 2: </span>Tabel setelah semua diskon terisi" >}}

Sekarang kita tinggal mengisi kolom `Harga Setelah Diskon`. Rumusnya itu didapat dari:

> Harga Awal - Diskon 10%

Ini hanyalah pengurangan biasa nilai dari dua sel, mirip seperti yang sudah pernah kita pelajari. Jadi rumus `Harga Setelah Diskon` untuk pensil seperti ini: `=B2-C2` dan hasilnya adalah `1800`. Tarik fill handle untuk mengisi rumus di baris berikutnya, sehingga hasil akhirnya akan jadi seperti ini:

{{< figure src="tabel-full.webp" caption="<span class=\"figure-number\">Figure 3: </span>Tabel yang sudah lengkap" >}}

Kalau kita lihat lagi rumus untuk `Harga Setelah Diskon`, kita cukup memakai referensi relatif ya, jadi tidak ada dolar di sana. Dibuat seperti itu karena kita memang ingin alamat sel dalam rumus itu mengikuti ke bawah ketika ditarik fill handle.


## Kesalahan Umum {#kesalahan-umum}


### Lupa mengunci alamat sel dengan referensi absolut {#lupa-mengunci-alamat-sel-dengan-referensi-absolut}

Kita baru saja belajar cara mengunci alamat sel di artikel ini, jadi tidak mungkin lupa dong? Tapi kalau misalnya kita lupa mengunci alamat sel pada latihan yang lain kira-kira apa yang terjadi ya?

Mari kita coba ubah rumus mencari besaran diskon jadi tidak dikunci seperti ini `=B2*G1`. Kemudian coba tarik ke bawah dan lihat hasilnya. Empat baris di bawahnya punya hasil `0`.

Kalau dicek rumus diskon pulpen sekarang jadi `=B3*G2`. Dan di sel `G2` ini tidak ada nilainya atau `0`. Maka dari itu hasil diskonnya adalah `0`.


### Salah mengunci alamat sel {#salah-mengunci-alamat-sel}

Kita perlu tahu sel mana yang harus dikunci dan mana yang tidak. Misalnya kita salah mengunci sel jadinya seperti ini `=$B$2*G1`. Di baris pertama hasilnya masih benar yaitu `200`, tapi kalau ditarik fill handle-nya ke bawah, maka hasil empat baris di bawahnya adalah `0`. Penyebab kita dapat hasil `0` karena di baris berikutnya rumusnya berubah menjadi `=$B$2*G2`, `G1` yang harusnya dikunci malah turun ke `G2` yang merupakan sel kosong.

Atau sekarang malah dikunci dua-duanya jadi seperti ini `=$B$2*$G$1`. Hasil baris pertama masih benar `200`, tapi kalau ditarik ke bawah empat baris berikutnya juga `200`.

Kedua hasil di atas bukan yang kita inginkan, jadi kita perlu pastikan sel yang benar-benar perlu dikunci.


## Penutup {#penutup}

Kita akan sering menemukan penggunaan referensi absolut dalam suatu rumus di Excel. Apalagi kalau nanti kita menggunakan rumus seperti `VLOOKUP`, `HLOOKUP`, atau `XLOOKUP`. Dan mungkin masih banyak lagi rumus lain yang perlu referensi absolut.
