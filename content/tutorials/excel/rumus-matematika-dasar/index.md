---
title: "Menggunakan Rumus Matematika Dasar di Microsoft Excel"
author: ["Muhamad Farhan"]
date: 2026-10-27
series: ["Rumus Excel"]
draft: false
series_order: 1
---

## Pembuka {#pembuka}

Pada artikel ini kita akan belajar cara menggunakan rumus di Excel untuk melakukan operasi perhitungan matematika dasar. Kita akan melakukan penjumlahan, pengurangan, perkalian dan pembagian.


## Konsep Dasar {#konsep-dasar}


### Sel di Excel {#sel-di-excel}

Dokumen Excel itu berbentuk tabel, dia terdiri dari **kolom** dan **baris**. Tiap **kolom** dinamai dengan huruf seperti A, B, C, dst. Sementara **baris** dinamai dengan angka seperti 1, 2, 3, dst. Tempat pertemuan kolom dan baris di Excel disebut dengan **sel**. Di dalam kotak **sel** ini kita memasukkan data.

Sel di Excel punya alamat. Alamatnya diambil dari gabungan nama kolom dan baris di mana sel itu berada. Misalnya sel **A1**, berarti dia terletak di kolom **A** baris **1**. Alamat sel ini yang nanti kita gunakan dalam rumus.

Kita bisa melihat alamat sel yang sedang aktif dari **Name Box** yang ada di sebelah kiri layar, di bawah tab menu, tepat di atas kolom **A**.

{{< figure src="gambar-sel-name-box.webp" caption="<span class=\"figure-number\">Figure 1: </span>Tampilan awal Excel menunjukkan sel dan **Name Box**" >}}


### Semua Rumus Diawali Sama Dengan (=) {#semua-rumus-diawali-sama-dengan}

Tanda sama dengan (=) menandakan bahwa kita sedang menggunakan rumus. Tanpa sama dengan, Excel akan menganggap itu sebagai teks biasa.

Pada sel **A1** cobalah ketik `7+5`, kemudian tekan **Tab** untuk pindah ke sel sebelah kanan. Pada sel **A1** akan ditampilkan apa adanya sebagai teks `7+5`.

Kemudian pada sel **B1**, cobalah ketik `=7+5`, lalu tekan **Tab** lagi. Sekarang pada sel **B1** akan tampil angka `12`. Angka `12` di sini adalah hasil penjumlahan `7` dengan `5`.

Excel menampilkan hasil dari rumusnya, bukan lagi notasi rumusnya. Jika ingin melihat kembali rumus yang dibuat, cobalah lihat pada **Formula Bar** yang ada di atas nama-nama kolom dan tepat di bawah tab menu.

{{< figure src="gambar-rumus-sederhana.webp" caption="<span class=\"figure-number\">Figure 2: </span>Tampilan penggunaan rumus excel" >}}


## Empat Operator Matematika Dasar {#empat-operator-matematika-dasar}

Setelah kita mengenali konsep dasar di Excel, mari kita menerapkan itu dalam operasi matematika dasar. Pada bab ini kita akan belajar cara menjumlahkan, mengurangi, mengalikan, dan membagi.


### Mengetik Tabel {#mengetik-tabel}

**Instruksi**: Ketiklah tabel berikut di Excel. Data pertama boleh dimulai dari sel **A1**.

| Angka 1 | Angka 2 | Penjumlahan | Pengurangan | Perkalian | Pembagian |
|---------|---------|-------------|-------------|-----------|-----------|
| 133     | 19      |             |             |           |           |
| 190     | 16      |             |             |           |           |
| 152     | 99      |             |             |           |           |

Setelah selesai diketik, maka tampilan tabelnya di Excel adalah sebagai berikut:

{{< figure src="tabel-tanpa-rumus.webp" caption="<span class=\"figure-number\">Figure 3: </span>Tampilan tabel awal" >}}


### Melakukan Penjumlahan {#melakukan-penjumlahan}

Sekarang kita akan melakukan penjumlahan pada sel **C2**. Kita akan menjumlahkan dua angka yaitu 133 dan 19.

**Instruksi**: Pada sel **C2** ketiklah sama dengan ( = ). Kemudian klik sel **A2** yang berisi angka `133`, lalu ketik `+`, selanjutnya klik sel **B2** yang berisi angka `19`.

Maka pada sel **C2** akan muncul rumus `=A2+B2`. Setelah itu tekan **Tab** atau **Enter** supaya Excel memproses rumusnya. Sekarang di sel **C2** akan menampilkan hasil penjumlahannya yaitu `152`.


### Melakukan Pengurangan {#melakukan-pengurangan}

Sekarang kita lanjut operasi pengurangan. Pengurangan akan kita lakukan pada sel **D2**. Kita akan mengurangi dua angka sebelumnya yaitu 133 dan 19.

**Instruksi**: Pada sel **D2** ketiklah sama dengan ( = ). Kemudian klik sel **A2** yang berisi angka `133`, lalu ketik `-`, selanjutnya klik sel **B2** yang berisi angka `19`.

Maka pada sel **D2** akan muncul rumus `=A2-B2`. Setelah itu tekanlah **Tab** atau **Enter**. Sekarang pada sel **D2** akan ditampilkan hasil pengurangannya yaitu `114`.


### Percobaan Mandiri untuk Perkalian dan Pembagian {#percobaan-mandiri-untuk-perkalian-dan-pembagian}

Sekarang cobalah mengisi kolom perkalian dan pembagian. Caranya mirip seperti tadi, cukup ganti simbol operator matematikanya. Untuk perkalian simbolnya `*`, sementara pembagian `/`.

Maka pada masing-masing sel akan punya rumus seperti ini:

-   `=A2*B2`, untuk perkalian.
-   `=A2/B2`, untuk pembagian.

Setelah semua kolom pada baris pertama terisi dengan rumus, maka tampilannya akan seperti berikut:

{{< figure src="tabel-dengan-rumus.webp" caption="<span class=\"figure-number\">Figure 4: </span>Tampilan tabel yang sudah diisi rumus" >}}


### Menyalin Rumus ke Baris Di Bawahnya {#menyalin-rumus-ke-baris-di-bawahnya}

Di Excel kita tidak perlu mengetik rumusnya satu per satu di setiap sel. Kita cukup kerjakan satu rumus saja dan pastikan itu sudah benar. Kemudian rumus itu bisa disalin ke baris di bawahnya.

**Instruksi**:

1.  Klik pada sel yang sudah terisi dengan rumus. Misalnya sel **C2** tadi.
2.  Pada sel yang terpilih itu, perhatikan kotak kecil yang ada di pojok kanan bawah sel. Itu adalah _Fill Handle_.
3.  Klik dan tahan pada Fill Handle. Kemudian tarik ke bawah sampai sel **C4**. Lalu lepaskan.

Sekarang semua baris pada kolom penjumlahan sudah terisi dengan hasil penjumlahan dari dua angka di sebelahnya.

Kalau kamu cek rumus di masing-masing sel, rumusnya sudah berubah menjadi `=A3+B3` dan `=A4+B4`. Ini adalah hal yang normal. Pada saat menyalin rumus, Excel menyesuaikan sel mana yang sedang dihitung.

Sekarang cobalah melakukan hal yang sama pada tiga kolom lainnya. Kalau sudah selesai, hasilnya akan jadi seperti ini:

{{< figure src="tabel-rumus-full.webp" caption="<span class=\"figure-number\">Figure 5: </span>Tampilan tabel yang sudah terisi semuanya dengan rumus" >}}

> [!IMPORTANT] Urutan Perhitungan
>
> Excel patuh pada urutan perhitungan matematika. Perkalian dan pembagian akan dikerjakan lebih dahulu daripada penjumlahan dan pengurangan. Jika ingin melakukan penjumlahan atau pengurangan lebih dahulu, gunakan tanda kurung `()`.
>
> Misalnya `2+3*4` hasilnya adalah `14`. Kalau `(2+3)*4` hasilnya adalah `20`.


## Kesalahan Umum {#kesalahan-umum}


### Lupa Mengetik Sama Dengan (=) Di Awal Rumus {#lupa-mengetik-sama-dengan-----di-awal-rumus}

Seperti yang sudah dijelaskan sebelumnya, tanda sama dengan ( = ) menyatakan bahwa kita sedang menggunakan rumus. Tanpa tanda sama dengan, Excel hanya akan menganggap itu sebagai teks biasa.


### Salah Mengetik Alamat Sel {#salah-mengetik-alamat-sel}

Semua rumus yang dikerjakan pada bab **Empat Operator Matematika Dasar** menganggap kalau kita mulai membuat tabel dari sel **A1**. Jika kamu membuat tabel dari sel lain, maka rumusnya akan berubah dan perlu disesuaikan.

Karena itu saya sarankan kamu tidak mengetik manual alamat sel dalam rumus karena bisa ada kesalahan ketik. Tapi kamu bisa klik langsung pada sel tempat nilai yang mau dihitung menggunakan rumus tersebut.

Sebenarnya ini tidak hanya berlaku untuk rumus matematika sederhana ini, tapi semua rumus yang lain. Penting kita mengetahui sel mana yang mau digunakan dalam rumus.


### Salah Simbol Operasi Perhitungan {#salah-simbol-operasi-perhitungan}

Kalau simbol penjumlahan dan pengurangan sudah jelas dengan `+` dan `-` ya. Tapi kadang kita lupa kalau simbol untuk perkalian dan pembagian adalah `*` dan `/`, tertukar dengan `x` dan `:`.
