1. Instal Library yang Dibutuhkan
Buka terminal / Command Prompt / PowerShell, lalu jalankan:

pip install requests pandas numpy scikit-learn joblib streamlit


Jika terjadi masalah PATH pada Windows, gunakan:

python -m pip install requests pandas numpy scikit-learn joblib streamlit

2. Cara Menjalankan Aplikasi
 2.1 Training Model Machine Learning

Langkah pertama adalah melatih model SVM menggunakan data prakiraan cuaca BMKG.

Masuk ke folder proyek, lalu jalankan:

python bmkg_train_svm.py


Jika berhasil:

Data akan diambil dari API BMKG

Dataset akan disimpan sebagai bmkg_cuaca_dataset.csv

Model SVM dan scaler akan tersimpan sebagai:

bmkg_svm_model.pkl

bmkg_svm_scaler.pkl

Output evaluasi seperti confusion matrix dan accuracy akan tampil di terminal.

2.2 Menjalankan Aplikasi UI (Streamlit)

Setelah model berhasil dibuat, jalankan aplikasi antarmuka pengguna dengan perintah:

python -m streamlit run app_cuaca_svm.py


Browser akan otomatis terbuka dengan alamat seperti:

http://localhost:8501

3. Cara Menggunakan Aplikasi
Aplikasi menyediakan dua mode input:
3.1 Mode Input Manual

Pilih menu Input Manual

Masukkan parameter cuaca:

Suhu (°C)

Kelembapan (%)

Kecepatan angin (km/jam)

Tutupan awan (%)

Klik tombol Prediksi Cuaca

Sistem akan menampilkan:

Hasil prediksi: Hujan / Tidak Hujan

Probabilitas prediksi

3.2 Mode Ambil Data dari BMKG

Pilih menu Ambil dari BMKG (kode adm4)

Masukkan kode wilayah administrasi BMKG (adm4)

Contoh kode wilayah:

Bojong Rawalumbu → 32.75.05.1001

Pengasinan → 32.75.05.1002

Sepanjang Jaya → 32.75.05.1003

Bojong Menteng → 32.75.05.1004

Klik Ambil dari BMKG & Prediksi

Sistem akan:
Mengambil data cuaca terbaru dari BMKG
Menampilkan lokasi dan waktu prakiraan
Menampilkan hasil prediksi cuaca berdasarkan model SVM

6. Catatan Penting

Aplikasi ini tidak menggantikan prakiraan resmi BMKG, melainkan sebagai contoh implementasi machine learning berbasis data cuaca.

Model hanya dilatih menggunakan data prakiraan 3 harian sehingga hasil prediksi dapat berbeda dengan kondisi cuaca aktual.

Akurasi model dapat ditingkatkan dengan:

Menambah jumlah wilayah (adm4)
Menggabungkan data historis yang lebih panjang
Mengoptimalkan parameter model SVM

7. Penutup

Dengan mengikuti langkah-langkah di atas, aplikasi prediksi cuaca berbasis Support Vector Machine dapat dijalankan dengan baik. Aplikasi ini menunjukkan bagaimana data real-time dari BMKG dapat dimanfaatkan dalam sistem machine learning yang interaktif dan mudah digunakan.