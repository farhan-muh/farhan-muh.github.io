---
title: "Menggunakan Fungsi IF Tunggal untuk Uji Logika di Excel"
description: "Fungsi IF tunggal di Excel untuk uji logika, misalnya menentukan siswa lulus atau tidak lulus."
author: ["Muhamad Farhan"]
date: 2026-11-12
series: ["Rumus Excel"]
draft: false
series_order: 4
---

## Pembuka {#pembuka}

Pada artikel ini kita akan menggunakan fungsi `IF` tunggal. Dengan fungsi `IF` kita bisa membuat Excel mengambil keputusan dari data yang ada berdasarkan kondisi tertentu. Contoh penggunaannya adalah menentukan keterangan lulus atau tidak lulusnya siswa dari tabel nilai yang tersedia berdasarkan batas nilai yang sudah diputuskan.


## Konsep Dasar {#konsep-dasar}


### Melakukan Uji Logika dengan Operator Perbandingan {#melakukan-uji-logika-dengan-operator-perbandingan}

Uji logika adalah proses memeriksa suatu kondisi terpenuhi atau tidak. Hasil pemeriksaannya hanya boleh ada dua yaitu **TRUE** bila kondisinya terpenuhi, dan **FALSE** bila kondisinya tidak terpenuhi.

Dalam uji logika ada beberapa operator perbandingan yang umum digunakan:

<div class="table-caption">
  <span class="table-number">Table 1:</span>
  Tabel operator perbandingan
</div>

| Operator | Keterangan              |
|----------|-------------------------|
| =        | Sama dengan             |
| &gt;     | Lebih dari              |
| &lt;     | Kurang dari             |
| &gt;=    | Lebih dari sama dengan  |
| &lt;=    | Kurang dari sama dengan |
| &lt;&gt; | Tidak sama dengan       |

Untuk lebih jelasnya coba lihat tabel di bawah ini:

{{< figure src="tabel-kebenaran.webp" caption="<span class=\"figure-number\">Figure 1: </span>Tabel kebenaran" >}}

Pada tabel di atas kita ingin membandingkan dua angka menggunakan enam operator perbandingan. Yang bernilai `TRUE` adalah kondisi yang benar dan `FALSE` adalah yang salah.

Kita ambil contoh `70` dan `65`:

<div class="table-caption">
  <span class="table-number">Table 2:</span>
  Contoh proses uji logika
</div>

| Operator | Apakah ... ?   | Hasil        | TRUE/FALSE |
|----------|----------------|--------------|------------|
| =        | 70 = 65        | tidak, salah | FALSE      |
| &gt;     | 70 &gt; 65     | iya, benar   | TRUE       |
| &lt;     | 70 &lt; 65     | tidak, salah | FALSE      |
| &gt;=    | 70 &gt;= 65    | iya, benar   | TRUE       |
| &lt;=    | 70 &lt;= 65    | tidak, salah | FALSE      |
| &lt;&gt; | 70 &lt;&gt; 65 | iya, benar   | TRUE       |


### Notasi Fungsi IF {#notasi-fungsi-if}

Setelah kita memahami konsep uji logika di Excel, sekarang kita masuk pada notasi fungsi `IF`. Fungsi `IF` punya tiga argumen yang ditulis dalam notasi seperti berikut:

```text
=IF(kondisi_uji; nilai_jika_benar; nilai_jika_salah)
```

Keterangan:

<div class="table-caption">
  <span class="table-number">Table 3:</span>
  Penjelasan tiga argumen fungsi if
</div>

| Argumen          | Keterangan                                         |
|------------------|----------------------------------------------------|
| kondisi_uji      | Kondisi yang ingin diperiksa oleh Excel            |
| nilai_jika_benar | Hasil yang ditampilkan jika kondisi bernilai TRUE  |
| nilai_jika_salah | Hasil yang ditampilkan jika kondisi bernilai FALSE |

Agar lebih jelas mari kita terapkan pada contoh sederhana. Pada tabel sebelumnya ditampilkan hasil uji logika dari berbagai operator perbandingan. Sekarang kita coba ambil simbol operator `>` untuk membandingkan tiga angka yaitu `70`, `65`, dan `80`.

**Instruksi**: Cobalah ketik dua rumus ini pada sel yang berbeda kemudian lihat hasilnya apa.

> =IF(70&gt;65;"Lebih Besar";"Tidak Lebih Besar")
>
> =IF(70&gt;80;"Lebih Besar";"Tidak Lebih Besar")

Rumus di atas akan menghasilkan salah satu dari dua nilai yaitu `Lebih Besar` atau `Tidak Lebih Besar`. Bila hasilnya tidak keluar, coba periksa apakah susunan rumusnya sudah sesuai? Apakah tanda baca seperti titik koma `;` dan tanda kutipnya `"` sudah lengkap dan letaknya sudah sesuai? Bila semua itu terpenuhi _Insya Allah_ hasilnya akan keluar.

Pada rumus pertama kondisi yang ingin diuji adalah apakah 70 lebih besar dari 65? Jika ternyata 70 itu memang benar lebih besar dari 65 maka hasilnya adalah `Lebih Besar`, namun kalau salah maka hasilnya adalah `Tidak Lebih Besar`. Kita tahu bahwa memang benar 70 itu lebih besar dari 65, maka di Excel hasil yang seharusnya adalah `Lebih Besar`.

Rumus kedua kondisinya berubah sekarang kita ingin menguji apakah 70 lebih besar dari 80? Kita sudah tahu ya, kalau 70 tidak lebih besar dari 80, maka di Excel hasilnya adalah `Tidak Lebih Besar`.

Dari dua contoh di atas kita sudah tahu cara kerja fungsi `IF`, bila kondisi uji pada argumen pertama terpenuhi, maka nilai jika benar pada argumen kedua akan ditampilkan, namun bila kondisi uji tidak terpenuhi maka nilai jika salah pada argumen ketiga yang ditampilkan.

> [!IMPORTANT] Koma (,) atau Titik Koma (;)
>
> Rumus di atas menggunakan titik koma sebagai pemisah antara argumen. Namun ada juga yang menggunakan koma sebagai tanda pemisah.
>
> Cek sub-bab [Menggunakan Pemisah Argumen yang Tidak Sesuai](#menggunakan-pemisah-argumen-yang-tidak-sesuai)


## Menggunakan Fungsi IF {#menggunakan-fungsi-if}


### IF Dengan Syarat Angka {#if-dengan-syarat-angka}

Setelah kita paham konsep uji logika dan notasi fungsi `IF`, mari kita terapkan itu pada kasus sederhana.

**Kasus**: Seorang guru ingin memberikan keterangan lulus atau tidak lulus untuk siswa-siswa yang ada di tabel berikut. Guru menyatakan bahwa yang lulus adalah siswa yang memiliki nilai 80 ke atas sedangkan yang di bawahnya dinyatakan tidak lulus.

**Instruksi**: Buatlah tabel berikut di sheet baru pada Excel. Kamu boleh mulai dari sel **A1**.

<div class="table-caption">
  <span class="table-number">Table 4:</span>
  Tabel nilai siswa
</div>

| Nama Siswa | Nilai | Keterangan |
|------------|-------|------------|
| Aldi       | 84    |            |
| Bella      | 79    |            |
| Cintia     | 80    |            |
| Dion       | 56    |            |
| Eki        | 90    |            |

Berdasarkan kasus di atas kita bisa membuat notasi fungsi `IF`-nya seperti ini:

```text
=IF(nilai_siswa>=80;"Lulus";"Tidak Lulus")
```

Sekarang kita masukkan notasi rumusnya di Excel.

1.  Pada sel keterangan `Aldi` yaitu **C2**, ketiklah `=IF(`.
2.  Kemudian klik nilai `Aldi` yaitu `84`, rumusnya sekarang menjadi `=IF(B2`.
3.  Lanjutkan rumusnya menjadi `=IF(B2>=80;"Lulus";"Tidak Lulus")`. Kemudian **Enter**.
4.  Sekarang muncul hasilnya untuk `Aldi` adalah `Lulus`.
5.  Tarik Fill Handle untuk mengisi keterangan siswa yang lainnya.

Hasil akhir tabelnya sebagai berikut:

{{< figure src="tabel-nilai-siswa-full.webp" caption="<span class=\"figure-number\">Figure 2: </span>Tabel nilai siswa yang sudah diisi rumus" >}}


### IF Dengan Syarat Teks {#if-dengan-syarat-teks}

Selain membandingkan angka, fungsi `IF` juga bisa digunakan untuk mencocokkan teks.

**Kasus**: Guru tadi ingin menambahkan keterangan remedial pada siswa yang tidak lulus sementara yang lulus tidak remedial.

**Instruksi**: Tambahkan kolom `Remedial?` di samping kolom `Keterangan` pada tabel sebelumnya.

<div class="table-caption">
  <span class="table-number">Table 5:</span>
  Tabel awal keterangan remedial
</div>

| Nama Siswa | Nilai | Keterangan  | Remedial? |
|------------|-------|-------------|-----------|
| Aldi       | 84    | Lulus       |           |
| Bella      | 79    | Tidak Lulus |           |
| Cintia     | 80    | Lulus       |           |
| Dion       | 56    | Tidak Lulus |           |
| Eki        | 90    | Lulus       |           |

Berdasarkan kasus di atas kita bisa membuat notasi rumusnya jadi seperti ini:

```text
=IF(keterangan="Lulus";"Tidak";"Ya")
```

Sekarang kita masukkan notasi rumus tersebut ke Excel:

1.  Pada sel `Remedial?` `Aldi` yaitu **D2**, ketiklah `=IF(`.
2.  Kemudian klik `Keterangan` `Aldi` yaitu `Lulus`, sehingga rumusnya menjadi `=IF(C2`.
3.  Lanjutkan rumusnya menjadi `=IF(C2="Lulus";"Tidak";"Ya")`, lalu **Enter**.
4.  Sekarang muncul hasil untuk `Aldi` yaitu `Tidak`.
5.  Tarik Fill Handle untuk mengisi baris di bawahnya.

Hasil akhir tabelnya akan menjadi seperti ini:

{{< figure src="tabel-remedial-full.webp" caption="<span class=\"figure-number\">Figure 3: </span>Hasil keterangan remedial siswa" >}}


## Kesalahan Umum {#kesalahan-umum}


### Lupa Menggunakan Tanda Kutip Pada Teks {#lupa-menggunakan-tanda-kutip-pada-teks}

Setiap teks yang ditulis langsung sebagai argumen fungsi `IF` wajib diapit tanda kutip dua (`" "`). Baik yang dijadikan sebagai pembanding atau teks yang ingin ditampilkan sebagai hasil.

Contoh yang tepat pada rumus sebelumnya:

```text
=IF(B2>=80;"Lulus";"Tidak Lulus")
```

Kata `Lulus` dan `Tidak Lulus` wajib diapit tanda kutip dua, salah satu dari dua kata itu akan ditampilkan berdasarkan kondisinya `TRUE` atau `FALSE`. Sementara angka `80` tidak menggunakan kutip.

Contoh yang kurang tepat:

```text
=IF(B2>=80;Lulus;Tidak Lulus)
```

Hasil rumus di atas adalah `#NAME?`.


### Salah Menggunakan Operator Perbandingan {#salah-menggunakan-operator-perbandingan}

Pada contoh di atas guru menginginkan siswa yang nilainya dari 80 ke atas dinyatakan lulus. Makanya kita pakai operator pembanding `>=` bukan `>`.

Ada sedikit perbedaan dari kedua operator tersebut yaitu bila menggunakan `>=` maka hitungannya `80` akan disertakan, tapi kalau memakai `>` `80` tidak disertakan. Bila kita pakai `>` maka siswa yang nilainya `80` pas akan dinyatakan tidak lulus.


### Salah Menentukan Sel Referensi {#salah-menentukan-sel-referensi}

Pada latihan sebelumnya kita diminta untuk langsung klik sel referensinya agar tidak salah ketik. Dengan klik langsung kita jadi tahu sel mana yang saat ini sedang ditunjuk.

Bila mau diketik biasa boleh, tapi diperiksa juga apakah alamat sel yang diketik itu sudah menunjukkan ke sel yang tepat.


### Menggunakan Pemisah Argumen yang Tidak Sesuai {#menggunakan-pemisah-argumen-yang-tidak-sesuai}

Masalah pemisah argumen antara koma dan titik koma sebelumnya sudah dibahas. Kedua tanda itu bergantung pada pengaturan pemisah daftar di sistem operasi, yang kemudian diikuti oleh Excel.

Cara paling mudah untuk mengetahui pemisah argumen yang digunakan yaitu dengan melihat bantuan rumus saat kita mengetik rumus di Excel. Itu biasanya muncul di dekat sel yang sedang kita ketik rumusnya. Di sana tampil pemisah argumennya. Bila yang tampil adalah koma, ganti semua titik koma yang digunakan pada rumus di artikel ini dengan koma.


### Ada Spasi Yang Tidak Terlihat {#ada-spasi-yang-tidak-terlihat}

Spasi walau tidak kita lihat tapi itu juga dianggap sebagai karakter. Misalnya di sel yang sedang diperiksa tertulis `"Lulus "` tapi di rumusnya kita menulis `"Lulus"`. Excel menganggap mereka berbeda jadi hasil yang ditampilkan adalah nilai jika salahnya yang dalam rumus remedial tadi adalah `Ya`.
