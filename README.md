# ZAIN.NET — Skripsi Jadi Artikel • Scrib Universal AI Lokal V1.2

Versi GitHub Pages statis. Tidak memerlukan API token atau backend.

## Perbaikan V1.2
- Detektor struktur skripsi universal: tidak lagi bergantung pada satu nama BAB.
- Menghindari BAB/ABSTRAK palsu yang berasal dari Daftar Isi.
- Mengenali variasi BAB IV seperti Paparan Data, Temuan Penelitian, Hasil Penelitian, Hasil dan Pembahasan, dan Pembahasan.
- Mengenali Metode / Metodologi Penelitian dan Penutup / Kesimpulan / Simpulan.
- Menggunakan urutan BAB, style Word, bold, isi paragraf, dan kepadatan kata kunci sebagai sinyal.
- Pemetaan Struktur manual: pengguna bisa memilih sendiri bagian Pendahuluan, Metode, Hasil/Pembahasan, dan Kesimpulan bila format kampus berbeda.
- Nilai deteksi tidak lagi 100% bila bagian utama tidak ditemukan.
- Format artikel V1.1 tetap dipertahankan: nama penulis bold, ABSTRAK/ABSTRACT di tengah, abstrak single spacing, Kesimpulan/Saran bold, English Abstract AI lokal.

## Upload ke GitHub Pages
Upload semua file pada folder ini langsung ke root repository, lalu aktifkan GitHub Pages dari branch `main` dan `/ (root)`.

## Catatan
English Abstract menggunakan model terjemahan lokal di browser. Pada penggunaan pertama browser perlu mengunduh model, lalu dapat menggunakan cache.
