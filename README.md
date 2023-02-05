# Smart Absensi Pengenalan Wajah
Aplikasi ini dibangun dengan bantuan library open-cv dan metode haarcascade frontalface untuk melakukan pengenalan wajah. 
Tersedia fitur-fitur yang ditambahkan dalam aplikasi ini untuk memberikan sistem keamanan pada saat melakukan absensi. Fitur yang tersedia seperti:
1. Deteksi kedipan mata yang diidentifikasi menggunakan metode haarcascade guna memberikan keamanan pada deteksi gambar/foto wajah.
2. Hitungan mundur pada saat kedipan mata menggunakan library randint untuk memberikan jeda waktu kapan pengguna harus mengedipkan mata. Hal tersebut dapat mencegah terjadinya deteksi pada video kedip mata.
3. Screenshoot layar camera saat melakukan absensi dengan bantuan library open cv. Hal tersebut mencegah terjadinya video call jarak jauh. Sehingga dengan adanya screenshoot gambar dapat dilihat backround atau latar belakang yang dapat disesuaikan dengan lokasi camera dipasang.

Ketiga fitur keamanan tersebut dapat meminimalisir kecurangan yang dapat terjadi. Namun dalam proses training dan pengenalan wajah, sistem ini masih melakukan kesalahan identifikasi saat melakukan absensi. Namun hal tersebut dapat dihindari dengan melakukan pengambilan gambar dataset diruangan yang terang atau sama sebanyak lebih dari 100 per user.

Aplikasi saya bagikan secara free untuk bisa dikembangkan agar lebih baik lagi. Note: Tidak untuk diperjual belikan, tanpa seizin pemilik.
anggadanangpratama0408@gmail.com
