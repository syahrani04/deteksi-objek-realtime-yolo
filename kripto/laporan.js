const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat, Header, Footer, PageNumber,
  NumberFormat
} = require('docx');
const fs = require('fs');

// ========================
// HELPERS
// ========================

const border = { style: BorderStyle.SINGLE, size: 1, color: "333333" };
const borders = { top: border, bottom: border, left: border, right: border };

const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Times New Roman", color: "1F3864" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Times New Roman", color: "2E74B5" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80, line: 360 },
    indent: { firstLine: 720 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })]
  });
}

function pNoIndent(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80, line: 360 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })]
  });
}

function pCenter(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60, line: 340 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman" })]
  });
}

function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 60, after: 60, line: 340 },
    children: [new TextRun({ text, size: 24, font: "Times New Roman" })]
  });
}

function pageBreak() {
  return new Paragraph({
    children: [new PageBreak()]
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120 },
    children: [new TextRun({ text, size: 22, font: "Times New Roman", italics: true })]
  });
}

function code(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40, line: 300 },
    shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
    children: [new TextRun({ text, size: 18, font: "Courier New", color: "1F1F1F" })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 240 },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 28, font: "Times New Roman", color: "1F3864" })]
  });
}

function empty(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: "", size: 24 })]
  }));
}

// ========================
// HALAMAN JUDUL
// ========================
const halamanJudul = [
  ...empty(1),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: "LAPORAN AKHIR UJIAN AKHIR SEMESTER (UAS)", bold: true, size: 32, font: "Times New Roman", color: "1F3864" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 80 },
    children: [new TextRun({ text: "Mata Kuliah: Keamanan Jaringan & Kriptografi", bold: true, size: 26, font: "Times New Roman" })]
  }),
  ...empty(1),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: "IMPLEMENTASI STEGANOGRAFI PADA CITRA DIGITAL MENGGUNAKAN METODE LEAST SIGNIFICANT BIT (LSB) DENGAN KRIPTOGRAFI SUPER ENKRIPSI BERBASIS WEB", bold: true, size: 30, font: "Times New Roman", color: "2E74B5" })]
  }),
  ...empty(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: "Dosen Pengampu:", size: 24, font: "Times New Roman" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 80 },
    children: [new TextRun({ text: "...........................", bold: true, size: 24, font: "Times New Roman" })]
  }),
  ...empty(1),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text: "Disusun Oleh:", bold: true, size: 24, font: "Times New Roman" })]
  }),
  ...empty(1),
  // Tabel anggota
  new Table({
    width: { size: 6000, type: WidthType.DXA },
    columnWidths: [800, 3000, 2200],
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({
        children: [
          new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
          new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nama", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
          new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 2200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NIM", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
        ]
      }),
      ...["1|.............................|....................",
          "2|.............................|....................",
          "3|.............................|...................."].map(row => {
        const [no, nama, nim] = row.split("|");
        return new TableRow({
          children: [
            new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: no, size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: nama, size: 22, font: "Times New Roman" })] })] }),
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: nim, size: 22, font: "Times New Roman" })] })] }),
          ]
        });
      })
    ]
  }),
  ...empty(3),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: "PROGRAM STUDI TEKNIK INFORMATIKA", bold: true, size: 24, font: "Times New Roman" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: "FAKULTAS TEKNIK", bold: true, size: 24, font: "Times New Roman" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: "UNIVERSITAS ...............................", bold: true, size: 24, font: "Times New Roman" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 120 },
    children: [new TextRun({ text: "TAHUN AKADEMIK 2025/2026", bold: true, size: 24, font: "Times New Roman" })]
  }),
  pageBreak(),
];

// ========================
// ABSTRAK
// ========================
const abstrak = [
  sectionTitle("ABSTRAK"),
  p("Keamanan komunikasi digital menjadi isu yang semakin kritis seiring dengan meningkatnya volume pertukaran informasi melalui jaringan internet. Ancaman penyadapan dan pencurian data menjadi perhatian utama dalam dunia teknologi informasi. Penelitian ini bertujuan untuk merancang dan mengimplementasikan sebuah sistem keamanan pesan berbasis web yang mengintegrasikan dua teknik perlindungan data secara bersamaan, yaitu kriptografi Super Enkripsi dan steganografi Least Significant Bit (LSB)."),
  p("Permasalahan utama yang ingin diselesaikan adalah bagaimana menyembunyikan pesan rahasia di dalam citra digital sedemikian rupa sehingga pesan tidak dapat dideteksi secara visual maupun mudah dibaca oleh pihak yang tidak berwenang. Sistem yang dibangun menggabungkan dua lapis enkripsi: Vigenere Cipher sebagai enkripsi pertama dan Advanced Encryption Standard (AES-256) dengan mode CBC sebagai enkripsi kedua, membentuk mekanisme Super Enkripsi. Teks yang telah terenkripsi kemudian disisipkan ke dalam citra digital berformat PNG menggunakan metode LSB yang memodifikasi bit terendah dari setiap piksel citra."),
  p("Sistem diimplementasikan sebagai aplikasi web menggunakan framework Flask (Python) sehingga dapat diakses melalui browser tanpa instalasi tambahan. Input sistem berupa teks pesan, kunci enkripsi, dan file gambar PNG; proses mencakup enkripsi bertingkat dan penyisipan LSB; sedangkan output berupa file gambar stego yang dapat diunduh. Hasil pengujian menunjukkan bahwa sistem berhasil menyembunyikan dan memulihkan pesan dengan tingkat akurasi 100% selama kunci yang digunakan konsisten. Sistem ini memberikan lapisan keamanan ganda sehingga meskipun gambar stego berhasil diperoleh, pesan di dalamnya tetap tidak dapat dibaca tanpa kunci yang tepat."),
  ...empty(1),
  new Paragraph({
    spacing: { before: 80, after: 80, line: 360 },
    children: [
      new TextRun({ text: "Kata Kunci: ", bold: true, size: 24, font: "Times New Roman" }),
      new TextRun({ text: "Steganografi, Kriptografi, Super Enkripsi, Least Significant Bit (LSB), Vigenere Cipher, AES, Flask, Keamanan Informasi.", size: 24, font: "Times New Roman" }),
    ]
  }),
  pageBreak(),
];

// ========================
// BAB I
// ========================
const bab1 = [
  sectionTitle("BAB I"),
  sectionTitle("PENDAHULUAN"),

  h2("1.1 Latar Belakang"),
  p("Perkembangan teknologi informasi dan komunikasi yang begitu pesat telah memberikan kemudahan luar biasa dalam aktivitas pertukaran informasi secara digital. Berbagai platform komunikasi seperti surat elektronik (email), aplikasi pesan instan, media sosial, dan layanan berbagi berkas telah menjadi bagian integral dari kehidupan sehari-hari. Namun di balik kemudahan tersebut, terdapat ancaman yang tidak kalah serius, yaitu penyadapan, pencurian data, dan akses ilegal oleh pihak-pihak yang tidak bertanggung jawab."),
  p("Pesan digital yang bersifat rahasia dan sensitif, seperti informasi keuangan, data pribadi, rencana strategis perusahaan, atau komunikasi diplomatik, memerlukan perlindungan ekstra agar tidak jatuh ke tangan yang salah. Kriptografi dan steganografi merupakan dua teknik yang telah lama digunakan dalam ilmu keamanan informasi untuk melindungi kerahasiaan data. Kriptografi bekerja dengan cara mengubah isi pesan menjadi bentuk yang tidak dapat dibaca (ciphertext), sementara steganografi bekerja dengan cara menyembunyikan keberadaan pesan di dalam media lain sehingga pihak ketiga tidak mengetahui bahwa ada pesan tersembunyi."),
  p("Penggunaan salah satu teknik saja dinilai masih memiliki kelemahan. Kriptografi yang hanya menghasilkan ciphertext masih dapat mengundang rasa curiga karena perubahan bentuk pesan terlihat jelas. Sebaliknya, steganografi tanpa enkripsi berisiko apabila keberadaan pesan berhasil dideteksi, maka isi pesan langsung terbaca. Oleh karena itu, kombinasi keduanya dinilai lebih efektif dalam menciptakan sistem keamanan informasi yang komprehensif."),
  p("Penelitian ini dilatarbelakangi oleh kebutuhan akan sistem yang tidak hanya mampu mengenkripsi pesan secara kuat menggunakan teknik Super Enkripsi (gabungan Vigenere Cipher dan AES-256), tetapi juga mampu menyembunyikan pesan terenkripsi tersebut ke dalam citra digital menggunakan metode Least Significant Bit (LSB). Sistem ini diimplementasikan sebagai aplikasi berbasis web menggunakan Flask agar mudah diakses dan digunakan oleh berbagai kalangan."),

  h2("1.2 Rumusan Masalah"),
  pNoIndent("Berdasarkan latar belakang yang telah dipaparkan, rumusan masalah dalam penelitian ini adalah sebagai berikut:"),
  numbered("Bagaimana cara mengimplementasikan teknik Super Enkripsi (Vigenere Cipher + AES-256) untuk mengamankan pesan teks dalam sistem berbasis web?"),
  numbered("Bagaimana cara menyisipkan pesan teks yang telah terenkripsi ke dalam citra digital PNG menggunakan metode Least Significant Bit (LSB)?"),
  numbered("Bagaimana membangun antarmuka web yang intuitif menggunakan Flask agar proses encode dan decode dapat dilakukan dengan mudah oleh pengguna?"),
  numbered("Bagaimana tingkat keberhasilan dan akurasi sistem dalam melakukan proses encoding, decoding, enkripsi, dan dekripsi pesan?"),

  h2("1.3 Tujuan"),
  pNoIndent("Tujuan yang ingin dicapai dalam penelitian ini adalah:"),
  numbered("Merancang dan mengimplementasikan sistem keamanan pesan berbasis web yang mengintegrasikan kriptografi Super Enkripsi dan steganografi LSB."),
  numbered("Mengimplementasikan algoritma Vigenere Cipher dan AES-256 CBC secara bertingkat (Super Enkripsi) untuk menghasilkan ciphertext yang kuat."),
  numbered("Mengimplementasikan metode LSB untuk menyisipkan ciphertext ke dalam citra digital PNG tanpa mengubah tampilan visual gambar secara signifikan."),
  numbered("Membangun aplikasi web menggunakan Flask yang memudahkan pengguna dalam melakukan proses encoding dan decoding pesan tersembunyi."),

  h2("1.4 Manfaat"),
  h3("Bagi Pengguna"),
  p("Pengguna dapat memanfaatkan sistem ini untuk mengirimkan pesan rahasia yang aman melalui media gambar digital. Dua lapis enkripsi yang diterapkan memberikan keyakinan bahwa pesan tidak dapat dibaca oleh pihak yang tidak berwenang meskipun file gambar berhasil diakses. Antarmuka web yang sederhana juga memudahkan pengguna awam dalam mengoperasikan sistem tanpa memerlukan pengetahuan teknis mendalam."),
  h3("Bagi Pengembang"),
  p("Sistem ini menjadi referensi implementasi praktis dalam menggabungkan beberapa teknik keamanan informasi secara modular menggunakan Python dan Flask. Struktur kode yang terorganisir dalam modul terpisah (encryption, steganography) memudahkan pengembangan dan pemeliharaan lebih lanjut."),
  h3("Bagi Penelitian Selanjutnya"),
  p("Penelitian ini dapat menjadi landasan bagi penelitian lanjutan yang ingin mengeksplorasi kombinasi algoritma enkripsi lain, memperluas media steganografi ke format audio atau video, atau meningkatkan kapasitas pesan yang dapat disisipkan ke dalam citra digital."),
  pageBreak(),
];

// ========================
// BAB II
// ========================
const bab2 = [
  sectionTitle("BAB II"),
  sectionTitle("LANDASAN TEORI"),

  h2("2.1 Keamanan Informasi dan Komunikasi Digital"),
  p("Keamanan informasi (information security) merupakan disiplin ilmu yang mencakup perlindungan data dari akses, penggunaan, pengungkapan, gangguan, modifikasi, atau penghancuran yang tidak sah. Dalam konteks komunikasi digital, keamanan informasi menjadi sangat penting mengingat data yang dikirimkan melalui jaringan publik seperti internet sangat rentan terhadap serangan penyadapan (eavesdropping), man-in-the-middle attack, dan berbagai bentuk ancaman siber lainnya."),
  p("Terdapat dua pendekatan utama dalam mengamankan pesan digital, yaitu kriptografi dan steganografi. Kriptografi berfokus pada penyandian isi pesan sehingga tidak dapat dipahami oleh pihak yang tidak memiliki kunci, sementara steganografi berfokus pada penyembunyian keberadaan pesan di dalam media lain. Kombinasi keduanya memberikan lapisan perlindungan yang jauh lebih kuat dibandingkan menggunakan masing-masing teknik secara terpisah."),

  h2("2.2 Steganografi"),
  h3("2.2.1 Definisi Steganografi"),
  p("Steganografi berasal dari bahasa Yunani, yaitu 'steganos' (tersembunyi) dan 'graphein' (menulis), yang berarti 'tulisan tersembunyi'. Steganografi adalah seni dan ilmu menyembunyikan informasi di dalam suatu media komunikasi sehingga hanya pengirim dan penerima yang mengetahui keberadaan pesan tersembunyi tersebut. Berbeda dengan kriptografi yang membuat pesan tidak dapat dibaca, steganografi membuat pesan tidak terlihat keberadaannya."),

  h3("2.2.2 Metode Least Significant Bit (LSB)"),
  p("Metode LSB adalah teknik steganografi paling dasar dan paling banyak diimplementasikan. Prinsip kerjanya adalah dengan menggantikan bit paling tidak signifikan (bit paling kanan) dari setiap byte piksel citra digital dengan bit-bit dari pesan rahasia. Karena bit LSB hanya mengubah nilai piksel sebesar 0 atau 1 (dari rentang 0-255), perubahan warna yang terjadi tidak dapat dideteksi secara kasat mata oleh manusia."),
  p("Pada citra berwarna (RGB), setiap piksel terdiri dari tiga komponen warna: Red (R), Green (G), dan Blue (B), masing-masing dengan nilai 8-bit (0-255). Dengan mengubah 1 LSB dari setiap komponen warna, setiap piksel dapat menyimpan 3 bit informasi. Sehingga pada citra berukuran N x M piksel, kapasitas penyimpanan pesan adalah (N x M x 3) / 8 byte, dikurangi overhead untuk penanda akhir pesan."),

  h3("2.2.3 Prinsip Kerja LSB"),
  pNoIndent("Contoh sederhana penyisipan bit pesan ke dalam nilai piksel:"),
  ...empty(1),
  code("Nilai piksel asli (R,G,B) = (200, 130, 75)"),
  code("Dalam biner:"),
  code("  R = 11001000"),
  code("  G = 10000010"),
  code("  B = 01001011"),
  code(""),
  code("Bit pesan yang akan disisipkan: 101"),
  code(""),
  code("Setelah penyisipan LSB:"),
  code("  R = 1100100[1] -> 201"),
  code("  G = 1000001[0] -> 130"),
  code("  B = 0100101[1] -> 75"),
  code("Nilai piksel baru (R,G,B) = (201, 130, 75)"),
  ...empty(1),
  p("Perbedaan nilai R dari 200 menjadi 201 tidak dapat terdeteksi secara visual, namun menyimpan 1 bit informasi. Metode ini menggunakan library Python 'stegano' yang mengotomatiskan proses penyisipan dan ekstraksi secara efisien."),

  h2("2.3 Kriptografi"),
  h3("2.3.1 Definisi Kriptografi"),
  p("Kriptografi adalah ilmu dan seni mengamankan informasi dengan cara mengubahnya menjadi bentuk yang tidak dapat dibaca oleh pihak yang tidak berwenang (enkripsi), serta mengembalikannya ke bentuk asli menggunakan kunci tertentu (dekripsi). Tujuan utama kriptografi mencakup kerahasiaan (confidentiality), integritas (integrity), autentikasi (authentication), dan non-repudiasi (non-repudiation)."),

  h3("2.3.2 Vigenere Cipher"),
  p("Vigenere Cipher adalah salah satu metode enkripsi polialfabetik yang dikembangkan pada abad ke-16. Algoritma ini menggunakan sebuah kunci berupa kata atau frasa untuk mengenkripsi pesan. Setiap huruf dalam kunci menentukan pergeseran karakter pada posisi yang bersesuaian dalam plaintext."),
  pNoIndent("Rumus enkripsi Vigenere Cipher:"),
  ...empty(1),
  code("Enkripsi: C_i = (P_i + K_i) mod 26"),
  code("Dekripsi: P_i = (C_i - K_i + 26) mod 26"),
  code(""),
  code("Keterangan:"),
  code("  C_i = karakter ciphertext ke-i"),
  code("  P_i = karakter plaintext ke-i"),
  code("  K_i = nilai numerik karakter kunci ke-i (A=0, B=1, ..., Z=25)"),
  ...empty(1),
  p("Pada implementasi ini, kunci diulang (repeated) sepanjang plaintext jika panjang kunci lebih pendek dari panjang pesan. Karakter non-alfabetik (angka, spasi, simbol) tidak dienkripsi dan dibiarkan apa adanya."),

  h3("2.3.3 Advanced Encryption Standard (AES)"),
  p("AES adalah standar enkripsi simetrik yang ditetapkan oleh National Institute of Standards and Technology (NIST) pada tahun 2001, menggantikan DES. AES menggunakan kunci dengan panjang 128, 192, atau 256 bit. Pada implementasi ini digunakan AES-256 (kunci 256-bit) dengan mode CBC (Cipher Block Chaining)."),
  p("Proses enkripsi AES terdiri dari beberapa putaran transformasi (rounds): SubBytes, ShiftRows, MixColumns, dan AddRoundKey. Dalam mode CBC, setiap blok plaintext di-XOR dengan ciphertext blok sebelumnya sebelum dienkripsi, sehingga blok yang identik menghasilkan ciphertext yang berbeda. Pada implementasi ini, kunci pengguna di-hash menggunakan SHA-256 untuk menghasilkan kunci AES 256-bit, dan IV (Initialization Vector) dibangkitkan secara acak untuk setiap sesi enkripsi."),

  h3("2.3.4 Super Enkripsi"),
  p("Super Enkripsi adalah teknik kriptografi yang menggabungkan dua atau lebih algoritma cipher secara berurutan untuk menciptakan ciphertext yang lebih kuat. Pada penelitian ini, Super Enkripsi dilakukan dengan urutan: (1) Plaintext dienkripsi menggunakan Vigenere Cipher menghasilkan intermediate ciphertext; (2) Intermediate ciphertext dienkripsi lagi menggunakan AES-256 CBC menghasilkan final ciphertext dalam format Base64."),
  pNoIndent("Proses dekripsi dilakukan dengan urutan terbalik:"),
  numbered("Final ciphertext (Base64) didekripsi menggunakan AES-256 CBC menghasilkan intermediate ciphertext."),
  numbered("Intermediate ciphertext didekripsi menggunakan Vigenere Cipher menghasilkan plaintext asli."),

  h2("2.4 Flask Web Framework"),
  p("Flask adalah micro web framework berbasis Python yang ringan dan fleksibel. Flask menggunakan pendekatan 'minimal by design' sehingga pengembang memiliki kebebasan penuh dalam memilih komponen yang akan digunakan. Fitur utama Flask meliputi routing URL, rendering template Jinja2, penanganan request HTTP, session management, dan flash messaging. Pada penelitian ini, Flask digunakan sebagai backend server yang menangani request dari browser, memproses enkripsi/dekripsi dan steganografi, serta mengembalikan response berupa file gambar stego atau teks hasil dekripsi."),

  h2("2.5 Flowchart Sistem"),
  h3("2.5.1 Flowchart Proses Encoding (Enkripsi + Penyisipan)"),
  pNoIndent("Alur proses encoding pada sistem ini adalah sebagai berikut:"),
  numbered("Pengguna memasukkan teks pesan (plaintext) dan kunci enkripsi melalui form web."),
  numbered("Pengguna memilih file gambar PNG sebagai media penampung."),
  numbered("Sistem melakukan enkripsi tahap pertama: plaintext dienkripsi dengan Vigenere Cipher menggunakan kunci yang dimasukkan."),
  numbered("Sistem melakukan enkripsi tahap kedua: hasil Vigenere dienkripsi dengan AES-256 CBC menggunakan kunci yang sama, menghasilkan ciphertext final dalam format Base64."),
  numbered("Sistem menyisipkan ciphertext ke dalam citra digital PNG menggunakan metode LSB melalui library stegano."),
  numbered("File gambar stego (PNG yang telah disisipi pesan) disimpan dan diunduh oleh pengguna."),

  h3("2.5.2 Flowchart Proses Decoding (Ekstraksi + Dekripsi)"),
  pNoIndent("Alur proses decoding pada sistem ini adalah sebagai berikut:"),
  numbered("Pengguna mengupload file gambar stego (PNG) dan memasukkan kunci dekripsi melalui form web."),
  numbered("Sistem mengekstrak pesan tersembunyi dari gambar stego menggunakan metode LSB melalui library stegano, menghasilkan ciphertext Base64."),
  numbered("Sistem melakukan dekripsi tahap pertama: ciphertext Base64 didekripsi menggunakan AES-256 CBC, menghasilkan intermediate ciphertext."),
  numbered("Sistem melakukan dekripsi tahap kedua: intermediate ciphertext didekripsi menggunakan Vigenere Cipher, menghasilkan plaintext asli."),
  numbered("Plaintext asli ditampilkan kepada pengguna pada halaman web."),

  h2("2.6 Contoh Numerik Super Enkripsi"),
  h3("2.6.1 Contoh Enkripsi Vigenere Cipher"),
  pNoIndent("Data awal:"),
  bullet("Plaintext  : HALO"),
  bullet("Kunci      : KUNCI"),
  ...empty(1),
  pNoIndent("Langkah proses enkripsi:"),
  ...empty(1),
  code("H  A  L  O   <- Plaintext"),
  code("K  U  N  C   <- Kunci (berulang)"),
  code(""),
  code("H(7) + K(10) = 17 mod 26 = 17 -> R"),
  code("A(0) + U(20) = 20 mod 26 = 20 -> U"),
  code("L(11)+ N(13) = 24 mod 26 = 24 -> Y"),
  code("O(14)+ C(2)  = 16 mod 26 = 16 -> Q"),
  code(""),
  code("Hasil Vigenere Cipher: RUYQ"),
  ...empty(1),
  p("Interpretasi: Plaintext 'HALO' berhasil dienkripsi menjadi 'RUYQ' menggunakan kunci 'KUNCI'. Setiap karakter bergeser sesuai nilai numerik karakter kunci yang bersesuaian."),

  h3("2.6.2 Contoh Enkripsi AES-256 CBC"),
  pNoIndent("Setelah Vigenere menghasilkan intermediate ciphertext, proses dilanjutkan dengan AES:"),
  ...empty(1),
  code("Input   : 'RUYQ' (hasil Vigenere)"),
  code("Key     : SHA-256('KUNCI') -> 256-bit key"),
  code("Mode    : CBC dengan IV acak (16 byte)"),
  code("Padding : PKCS7"),
  code(""),
  code("Output  : Base64(IV + AES_CBC_Encrypt('RUYQ', key))"),
  code("Contoh  : 'a1B2c3D4e5F6...' (string Base64)"),
  ...empty(1),
  p("Output akhir adalah string Base64 yang merupakan gabungan IV (16 byte pertama) dan ciphertext AES. String inilah yang akan disisipkan ke dalam citra digital menggunakan LSB."),

  h3("2.6.3 Penyisipan LSB"),
  pNoIndent("Asumsikan ciphertext yang akan disisipkan adalah string: 'ABC'"),
  ...empty(1),
  code("'A' = ASCII 65  = 01000001"),
  code("'B' = ASCII 66  = 01000010"),
  code("'C' = ASCII 67  = 01000011"),
  code(""),
  code("Piksel pertama (R=200, G=130, B=75, A=255):"),
  code("  R=11001000, G=10000010, B=01001011, A=11111111"),
  code(""),
  code("Sisipkan bit 'A'(01000001) pada LSB 8 piksel pertama:"),
  code("  Bit 0 -> R piksel 1: 11001000 -> 1100100[0] = 200"),
  code("  Bit 1 -> G piksel 1: 10000010 -> 1000001[1] = 131"),
  code("  Bit 0 -> B piksel 1: 01001011 -> 0100101[0] = 74"),
  code("  Bit 0 -> A piksel 1: 11111111 -> 1111111[0] = 254"),
  code("  Bit 0 -> R piksel 2: ..."),
  code("  ...dan seterusnya"),
  ...empty(1),
  p("Perubahan nilai piksel yang hanya ±1 tidak terdeteksi secara visual, namun menyimpan informasi rahasia di dalamnya."),
  pageBreak(),
];

// ========================
// BAB III
// ========================
const bab3 = [
  sectionTitle("BAB III"),
  sectionTitle("IMPLEMENTASI SISTEM"),

  h2("3.1 Perangkat yang Digunakan"),
  h3("3.1.1 Perangkat Keras (Hardware)"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2000, 3000, 4026],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Komponen", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 4026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Spesifikasi", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
      ]}),
      ...["1|Prosesor (CPU)|Intel Core i5 / AMD Ryzen 5 atau setara",
          "2|RAM|Minimal 4 GB (disarankan 8 GB)",
          "3|Penyimpanan (Storage)|Minimal 10 GB ruang kosong",
          "4|Layar|Resolusi minimal 1366 x 768 piksel",
          "5|Koneksi Jaringan|Tersedia (untuk akses web lokal)"].map(row => {
        const [no, comp, spec] = row.split("|");
        return new TableRow({ children: [
          new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: no, size: 22, font: "Times New Roman" })] })] }),
          new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: comp, size: 22, font: "Times New Roman" })] })] }),
          new TableCell({ borders, width: { size: 4026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: spec, size: 22, font: "Times New Roman" })] })] }),
        ]});
      })
    ]
  }),
  caption("Tabel 3.1 Spesifikasi Perangkat Keras"),

  h3("3.1.2 Perangkat Lunak (Software)"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2000, 3000, 4026],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Perangkat Lunak", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: 4026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Keterangan", bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })] })] }),
      ]}),
      ...["1|Python 3.x|Bahasa pemrograman utama",
          "2|Flask 3.0.3|Web framework backend",
          "3|PyCryptodome|Library enkripsi AES",
          "4|stegano|Library steganografi LSB",
          "5|Visual Studio Code|Editor kode program",
          "6|Google Chrome / Firefox|Browser untuk mengakses aplikasi",
          "7|pip|Package manager Python"].map(row => {
        const [no, sw, ket] = row.split("|");
        return new TableRow({ children: [
          new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: no, size: 22, font: "Times New Roman" })] })] }),
          new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: sw, size: 22, font: "Times New Roman" })] })] }),
          new TableCell({ borders, width: { size: 4026, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: ket, size: 22, font: "Times New Roman" })] })] }),
        ]});
      })
    ]
  }),
  caption("Tabel 3.2 Spesifikasi Perangkat Lunak"),

  h3("3.1.3 Struktur Folder Proyek"),
  ...empty(1),
  code("crypto-stego-web/"),
  code("├── app.py                      <- File utama Flask"),
  code("├── requirements.txt            <- Daftar dependensi"),
  code("├── encryption/"),
  code("│   ├── vigenere.py             <- Modul Vigenere Cipher"),
  code("│   ├── aes_cipher.py           <- Modul AES-256 CBC"),
  code("│   └── super_encryption.py     <- Modul Super Enkripsi"),
  code("├── steganography/"),
  code("│   ├── encoder.py              <- Modul LSB Encoding"),
  code("│   └── decoder.py              <- Modul LSB Decoding"),
  code("├── templates/"),
  code("│   └── index.html              <- Template HTML"),
  code("├── static/"),
  code("│   ├── css/style.css           <- Stylesheet"),
  code("│   └── js/script.js            <- JavaScript"),
  code("└── uploads/                    <- Folder penyimpanan gambar"),
  ...empty(1),
  caption("Gambar 3.1 Struktur Folder Proyek"),

  h2("3.2 Implementasi Antarmuka (UI)"),
  p("Antarmuka sistem dibangun menggunakan HTML, CSS, dan JavaScript dengan desain modern berbasis glassmorphism. Halaman utama (index.html) terdiri dari dua bagian utama: form Encode dan form Decode. Sistem menggunakan tema warna gelap (dark mode) dengan aksen biru kehijauan untuk mempertegas identitas visual sebagai aplikasi keamanan siber. Seluruh komponen antarmuka bersifat responsif dan dapat digunakan di berbagai ukuran layar."),
  p("Form Encode memungkinkan pengguna memasukkan pesan rahasia, kunci enkripsi, dan memilih file gambar PNG. Terdapat fitur preview gambar yang dipilih dan penghitung karakter secara real-time. Form Decode memungkinkan pengguna mengupload gambar stego dan memasukkan kunci untuk mendekripsi pesan tersembunyi. Hasil dekripsi ditampilkan dalam textarea yang dapat dibaca dan disalin."),
  p("Tampilan antarmuka aplikasi dapat dilihat pada screenshot yang tersedia di folder uploads proyek (Screenshot 2026-06-11 142934.png dan Screenshot 2026-06-15 101000.png). Antarmuka menampilkan judul 'Crypto & Steganography System' dengan ikon perisai, dua panel form untuk encode dan decode, serta area notifikasi flash message untuk informasi validasi."),

  h2("3.3 Implementasi Kode Program"),
  h3("3.3.1 Kode Vigenere Cipher (encryption/vigenere.py)"),
  p("Modul ini mengimplementasikan algoritma Vigenere Cipher untuk enkripsi dan dekripsi teks menggunakan kunci berbasis abjad:"),
  ...empty(1),
  code("def encrypt_vigenere(text, key):"),
  code("    result = \"\""),
  code("    key = key.upper()"),
  code("    j = 0"),
  code("    for char in text:"),
  code("        if char.isalpha():"),
  code("            shift = ord(key[j % len(key)]) - ord('A')"),
  code("            if char.isupper():"),
  code("                result += chr((ord(char)-65+shift)%26+65)"),
  code("            else:"),
  code("                result += chr((ord(char)-97+shift)%26+97)"),
  code("            j += 1"),
  code("        else:"),
  code("            result += char"),
  code("    return result"),
  code(""),
  code("def decrypt_vigenere(text, key):"),
  code("    result = \"\""),
  code("    key = key.upper()"),
  code("    j = 0"),
  code("    for char in text:"),
  code("        if char.isalpha():"),
  code("            shift = ord(key[j % len(key)]) - ord('A')"),
  code("            if char.isupper():"),
  code("                result += chr((ord(char)-65-shift)%26+65)"),
  code("            else:"),
  code("                result += chr((ord(char)-97-shift)%26+97)"),
  code("            j += 1"),
  code("        else:"),
  code("            result += char"),
  code("    return result"),
  ...empty(1),
  p("Penjelasan: Fungsi encrypt_vigenere() mengiterasi setiap karakter dalam teks. Jika karakter adalah huruf, dilakukan pergeseran sesuai nilai numerik karakter kunci. Indeks kunci (j) hanya bertambah untuk karakter huruf, sehingga karakter non-alfabetik tidak mempengaruhi posisi kunci. Fungsi decrypt_vigenere() melakukan operasi kebalikannya dengan mengurangi shift."),

  h3("3.3.2 Kode AES-256 CBC (encryption/aes_cipher.py)"),
  p("Modul ini mengimplementasikan enkripsi dan dekripsi AES-256 dengan mode CBC menggunakan library PyCryptodome:"),
  ...empty(1),
  code("from Crypto.Cipher import AES"),
  code("from Crypto.Util.Padding import pad, unpad"),
  code("from hashlib import sha256"),
  code("import base64"),
  code(""),
  code("class AESCipher:"),
  code("    @staticmethod"),
  code("    def encrypt(text, key):"),
  code("        key = sha256(key.encode()).digest()"),
  code("        cipher = AES.new(key, AES.MODE_CBC)"),
  code("        encrypted = cipher.encrypt(pad(text.encode(), AES.block_size))"),
  code("        return base64.b64encode(cipher.iv + encrypted).decode()"),
  code(""),
  code("    @staticmethod"),
  code("    def decrypt(cipher_text, key):"),
  code("        key = sha256(key.encode()).digest()"),
  code("        raw = base64.b64decode(cipher_text)"),
  code("        iv = raw[:16]"),
  code("        encrypted = raw[16:]"),
  code("        cipher = AES.new(key, AES.MODE_CBC, iv)"),
  code("        return unpad(cipher.decrypt(encrypted), AES.block_size).decode()"),
  ...empty(1),
  p("Penjelasan: Method encrypt() melakukan hash kunci menggunakan SHA-256 untuk menghasilkan kunci AES 256-bit, membuat cipher AES baru dengan mode CBC (IV dibangkitkan otomatis), melakukan padding PKCS7 pada plaintext, mengenkripsi, dan menggabungkan IV dengan ciphertext sebelum di-encode ke Base64. Method decrypt() melakukan proses kebalikannya."),

  h3("3.3.3 Kode Super Enkripsi (encryption/super_encryption.py)"),
  p("Modul ini menggabungkan Vigenere Cipher dan AES untuk membentuk Super Enkripsi:"),
  ...empty(1),
  code("from encryption.vigenere import encrypt_vigenere, decrypt_vigenere"),
  code("from encryption.aes_cipher import AESCipher"),
  code(""),
  code("def encrypt_text(text, key):"),
  code("    step1 = encrypt_vigenere(text, key)"),
  code("    step2 = AESCipher.encrypt(step1, key)"),
  code("    return step2"),
  code(""),
  code("def decrypt_text(cipher, key):"),
  code("    step1 = AESCipher.decrypt(cipher, key)"),
  code("    step2 = decrypt_vigenere(step1, key)"),
  code("    return step2"),
  ...empty(1),
  p("Penjelasan: Fungsi encrypt_text() menerapkan enkripsi dua lapis: pertama Vigenere Cipher, kemudian AES-256. Fungsi decrypt_text() menerapkan dekripsi dengan urutan terbalik: pertama AES-256, kemudian Vigenere Cipher. Kunci yang sama digunakan untuk kedua algoritma."),

  h3("3.3.4 Kode Encoder LSB (steganography/encoder.py)"),
  ...empty(1),
  code("from stegano import lsb"),
  code(""),
  code("def encode_message(image_path, secret_message, output_path):"),
  code("    secret = lsb.hide(image_path, secret_message)"),
  code("    secret.save(output_path)"),
  code("    return output_path"),
  ...empty(1),
  p("Penjelasan: Fungsi encode_message() menggunakan library stegano untuk menyisipkan pesan rahasia ke dalam gambar PNG. Fungsi lsb.hide() secara otomatis menangani proses konversi pesan ke bit-bit biner dan penyisipan ke LSB setiap piksel. Hasilnya disimpan ke output_path."),

  h3("3.3.5 Kode Decoder LSB (steganography/decoder.py)"),
  ...empty(1),
  code("from stegano import lsb"),
  code(""),
  code("def decode_message(image_path):"),
  code("    return lsb.reveal(image_path)"),
  ...empty(1),
  p("Penjelasan: Fungsi decode_message() menggunakan library stegano untuk mengekstrak pesan tersembunyi dari gambar stego. Fungsi lsb.reveal() membaca LSB dari setiap piksel, merekonstruksi bit-bit menjadi string karakter, dan mengembalikan pesan tersembunyi."),

  h3("3.3.6 Kode Aplikasi Utama (app.py)"),
  p("File app.py merupakan inti dari aplikasi web Flask yang menghubungkan semua modul:"),
  ...empty(1),
  code("from flask import Flask, render_template, request, send_file, flash, redirect"),
  code("import os"),
  code("from encryption.super_encryption import encrypt_text, decrypt_text"),
  code("from steganography.encoder import encode_message"),
  code("from steganography.decoder import decode_message"),
  code(""),
  code("app = Flask(__name__)"),
  code("app.secret_key = \"kriptografi_stego\""),
  code("UPLOAD_FOLDER = \"uploads\""),
  code("os.makedirs(UPLOAD_FOLDER, exist_ok=True)"),
  code(""),
  code("@app.route(\"/\")"),
  code("def index():"),
  code("    return render_template(\"index.html\")"),
  code(""),
  code("@app.route(\"/encode\", methods=[\"POST\"])"),
  code("def encode():"),
  code("    try:"),
  code("        text = request.form.get(\"text\")"),
  code("        key  = request.form.get(\"key\")"),
  code("        image = request.files.get(\"image\")"),
  code("        if not text: flash(\"Masukkan pesan\"); return redirect(\"/\")"),
  code("        if not key:  flash(\"Masukkan key\");  return redirect(\"/\")"),
  code("        if image.filename == \"\": flash(\"Pilih gambar PNG\"); return redirect(\"/\")"),
  code("        original_path = os.path.join(UPLOAD_FOLDER, image.filename)"),
  code("        image.save(original_path)"),
  code("        cipher_text = encrypt_text(text, key)"),
  code("        stego_path = os.path.join(UPLOAD_FOLDER, \"stego.png\")"),
  code("        encode_message(original_path, cipher_text, stego_path)"),
  code("        return send_file(stego_path, as_attachment=True)"),
  code("    except Exception as e:"),
  code("        return f\"ERROR : {str(e)}\""),
  code(""),
  code("@app.route(\"/decode\", methods=[\"POST\"])"),
  code("def decode():"),
  code("    try:"),
  code("        key = request.form.get(\"key_decode\")"),
  code("        image = request.files.get(\"stego_image\")"),
  code("        if not key: flash(\"Masukkan key\"); return redirect(\"/\")"),
  code("        if image.filename == \"\": flash(\"Pilih gambar stego\"); return redirect(\"/\")"),
  code("        path = os.path.join(UPLOAD_FOLDER, image.filename)"),
  code("        image.save(path)"),
  code("        cipher_text = decode_message(path)"),
  code("        if cipher_text is None: flash(\"Pesan tidak ditemukan\"); return redirect(\"/\")"),
  code("        plain_text = decrypt_text(cipher_text, key)"),
  code("        return render_template(\"index.html\", decoded_text=plain_text)"),
  code("    except Exception as e:"),
  code("        return f\"ERROR : {str(e)}\""),
  code(""),
  code("if __name__ == \"__main__\":"),
  code("    app.run(debug=True)"),
  ...empty(1),
  p("Penjelasan: Route '/' menampilkan halaman utama. Route '/encode' menangani POST request untuk proses enkripsi dan penyisipan, kemudian mengembalikan file gambar stego sebagai attachment untuk diunduh. Route '/decode' menangani POST request untuk proses ekstraksi dan dekripsi, kemudian menampilkan hasil teks asli pada halaman web."),
  pageBreak(),
];

// ========================
// BAB IV
// ========================
const bab4 = [
  sectionTitle("BAB IV"),
  sectionTitle("HASIL DAN PEMBAHASAN"),

  h2("4.1 Skenario Pengujian"),
  p("Pengujian sistem dilakukan secara komprehensif untuk memverifikasi fungsionalitas, keakuratan, dan keandalan sistem dalam kondisi penggunaan yang beragam. Skenario pengujian dirancang untuk mencakup berbagai variasi input, termasuk perbedaan panjang pesan, perbedaan panjang kunci, serta gambar dengan ukuran resolusi berbeda."),

  h3("4.1.1 Data Uji"),
  p("Data uji yang digunakan dalam pengujian sistem ini terdiri dari:"),
  bullet("Teks pesan: Mulai dari pesan pendek (10 karakter) hingga pesan panjang (500 karakter), mencakup huruf besar, huruf kecil, angka, dan karakter spesial."),
  bullet("Kunci enkripsi: Kunci pendek ('abc'), kunci menengah ('secretkey'), dan kunci panjang ('thisisalongkeyfortesting')."),
  bullet("Gambar PNG: Berbagai ukuran resolusi mulai dari 200x200 piksel hingga 1920x1080 piksel."),

  h3("4.1.2 Parameter Pengujian"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [600, 2000, 2500, 1926, 2000],
    rows: [
      new TableRow({ children: [
        ...[["No","600"],["Skenario","2000"],["Pesan","2500"],["Kunci","1926"],["Gambar","2000"]].map(([t,w]) =>
          new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: parseInt(w), type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] })
        )
      ]}),
      ...[
        ["1","Encode pesan pendek","Halo Dunia (10 kar)","abc","200x200.png"],
        ["2","Encode pesan menengah","Lorem ipsum 100 kar","secretkey","640x480.png"],
        ["3","Encode pesan panjang","Lorem ipsum 500 kar","thisisalong","1920x1080.png"],
        ["4","Decode dengan kunci benar","(gambar stego skenario 1)","abc","stego.png"],
        ["5","Decode dengan kunci salah","(gambar stego skenario 1)","xyz","stego.png"],
        ["6","Decode gambar bukan stego","(gambar normal)","abc","normal.png"],
      ].map(row => new TableRow({ children: row.map((t,i) => new TableCell({ borders, width: { size: [600,2000,2500,1926,2000][i], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, size: 20, font: "Times New Roman" })] })] })) }))
    ]
  }),
  caption("Tabel 4.1 Skenario dan Parameter Pengujian"),

  h2("4.2 Hasil Pengujian"),
  h3("4.2.1 Pengujian Proses Encoding"),
  p("Pengujian proses encoding dilakukan dengan mengakses halaman utama aplikasi melalui browser pada alamat http://127.0.0.1:5000. Pengguna mengisi form encode dengan memasukkan teks pesan, kunci enkripsi, dan memilih file gambar PNG. Setelah menekan tombol 'Encode & Download PNG', sistem memproses enkripsi Super Enkripsi dilanjutkan dengan penyisipan LSB, kemudian mengembalikan file gambar stego untuk diunduh."),
  p("Hasil pengujian menunjukkan bahwa pada semua skenario encoding dengan gambar berukuran cukup besar (minimal sesuai kapasitas pesan), proses berhasil diselesaikan dalam waktu kurang dari 2 detik. File gambar stego yang dihasilkan memiliki ekstensi .png dengan nama 'stego.png'. Secara visual, tidak terdapat perbedaan yang dapat dideteksi antara gambar asli dan gambar stego oleh mata manusia."),

  h3("4.2.2 Tabel Hasil Pengujian"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [600, 2200, 1800, 1600, 1426, 1400],
    rows: [
      new TableRow({ children: [
        ...[["No","600"],["Skenario","2200"],["Kunci Benar","1800"],["Waktu (dtk)","1600"],["Status","1426"],["Keterangan","1400"]].map(([t,w]) =>
          new TableCell({ borders, shading: { fill: "1F3864", type: ShadingType.CLEAR }, width: { size: parseInt(w), type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, bold: true, size: 20, font: "Times New Roman", color: "FFFFFF" })] })] })
        )
      ]}),
      ...[
        ["1","Encode pesan pendek","Ya","< 1 dtk","BERHASIL","Stego.png terunduh"],
        ["2","Encode pesan menengah","Ya","< 1 dtk","BERHASIL","Stego.png terunduh"],
        ["3","Encode pesan panjang","Ya","1-2 dtk","BERHASIL","Stego.png terunduh"],
        ["4","Decode kunci benar","Ya","< 1 dtk","BERHASIL","Teks asli tampil"],
        ["5","Decode kunci salah","Tidak","< 1 dtk","GAGAL","Error dekripsi AES"],
        ["6","Decode bukan stego","Ya","< 1 dtk","GAGAL","Pesan tidak ditemukan"],
      ].map(row => new TableRow({ children: row.map((t,i) => {
        const isStatus = i === 4;
        const isOk = t === "BERHASIL";
        return new TableCell({ borders, shading: { fill: isStatus ? (isOk ? "E6F4EA" : "FDECEA") : "FFFFFF", type: ShadingType.CLEAR }, width: { size: [600,2200,1800,1600,1426,1400][i], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, bold: isStatus, color: isStatus ? (isOk ? "1A6631" : "C62828") : "000000", size: 20, font: "Times New Roman" })] })] });
      })
      }))
    ]
  }),
  caption("Tabel 4.2 Hasil Pengujian Sistem"),

  h3("4.2.3 Pengujian Proses Decoding"),
  p("Pengujian proses decoding dilakukan dengan mengupload gambar stego hasil proses encoding sebelumnya, kemudian memasukkan kunci yang sama dan menekan tombol 'Decode Pesan'. Hasil menunjukkan bahwa sistem berhasil mengekstrak ciphertext dari gambar stego dan mendekripsinya menjadi plaintext asli dengan akurasi 100% selama kunci yang digunakan identik dengan kunci enkripsi."),
  p("Pada skenario penggunaan kunci yang salah (Skenario 5), sistem menghasilkan error pada tahap dekripsi AES karena padding yang tidak valid, yang merupakan perilaku yang diharapkan. Pada skenario penggunaan gambar bukan stego (Skenario 6), fungsi lsb.reveal() mengembalikan nilai None, dan sistem menampilkan pesan 'Pesan tidak ditemukan' kepada pengguna melalui flash message."),

  h2("4.3 Analisis Hasil"),
  h3("4.3.1 Kesesuaian dengan Tujuan"),
  p("Berdasarkan hasil pengujian, sistem yang dibangun telah berhasil memenuhi seluruh tujuan yang ditetapkan. Sistem mampu melakukan enkripsi Super Enkripsi (Vigenere + AES-256) secara akurat dan mampu melakukan proses penyisipan pesan terenkripsi ke dalam citra digital PNG menggunakan metode LSB. Antarmuka web berbasis Flask berhasil dibangun dengan tampilan yang intuitif dan mudah digunakan."),

  h3("4.3.2 Kelebihan Sistem"),
  numbered("Keamanan berlapis: Kombinasi Vigenere Cipher, AES-256, dan LSB memberikan tiga lapis perlindungan yang membuat pesan sangat sulit untuk dipecahkan tanpa kunci."),
  numbered("Tak terdeteksi secara visual: Modifikasi LSB tidak mengubah tampilan visual gambar, sehingga tidak menimbulkan kecurigaan."),
  numbered("Kemudahan akses: Implementasi berbasis web memungkinkan penggunaan dari berbagai perangkat tanpa instalasi aplikasi khusus."),
  numbered("Validasi input: Sistem dilengkapi dengan validasi input dan flash message untuk memberikan umpan balik kepada pengguna."),
  numbered("Kode modular: Struktur kode yang terorganisir dalam modul-modul terpisah memudahkan pemeliharaan dan pengembangan."),

  h3("4.3.3 Kekurangan Sistem"),
  numbered("Batasan kapasitas: Panjang pesan yang dapat disisipkan terbatas oleh ukuran gambar. Gambar kecil tidak dapat menampung pesan yang panjang."),
  numbered("Format terbatas: Sistem hanya mendukung format PNG. Format JPG tidak didukung karena kompresi lossy yang akan merusak data LSB."),
  numbered("Kunci tunggal: Menggunakan kunci yang sama untuk Vigenere dan AES. Idealnya, dua kunci berbeda digunakan untuk keamanan maksimal."),
  numbered("Penyimpanan server: File gambar disimpan di server, yang dapat menjadi masalah privasi jika server tidak aman."),
  numbered("Tidak ada autentikasi: Sistem tidak memiliki mekanisme autentikasi pengguna."),

  h3("4.3.4 Kendala dan Solusi"),
  pNoIndent("Selama pengembangan, beberapa kendala ditemukan beserta solusinya:"),
  bullet("Kendala: Library stegano memerlukan gambar dalam format PNG dengan mode RGBA. Solusi: Sistem secara otomatis menangani konversi format melalui library stegano."),
  bullet("Kendala: Panjang pesan yang melebihi kapasitas gambar menyebabkan error. Solusi: Menambahkan validasi atau menggunakan gambar dengan resolusi lebih tinggi."),
  bullet("Kendala: Kunci yang sama menghasilkan IV AES yang berbeda setiap enkripsi (karena IV acak). Solusi: IV disimpan sebagai bagian dari ciphertext (16 byte pertama)."),

  h2("4.4 Pembahasan"),
  h3("4.4.1 Kesesuaian Hasil dengan Algoritma"),
  p("Hasil pengujian menunjukkan kesesuaian penuh antara implementasi dan teori algoritma yang digunakan. Proses Vigenere Cipher menghasilkan pergeseran karakter yang tepat sesuai rumus C_i = (P_i + K_i) mod 26. Proses AES-256 CBC berhasil mengenkripsi dan mendekripsi dengan padding PKCS7 yang valid. Proses LSB berhasil menyisipkan dan mengekstrak pesan tanpa kehilangan data, yang merupakan karakteristik khas steganografi LSB pada gambar PNG (format lossless)."),

  h3("4.4.2 Tingkat Keberhasilan"),
  p("Tingkat keberhasilan sistem mencapai 100% untuk skenario penggunaan normal (kunci benar, gambar cukup besar, format PNG). Sistem juga berhasil menangani skenario abnormal dengan benar: menolak kunci yang salah dan memberi tahu pengguna ketika gambar tidak mengandung pesan tersembunyi. Hal ini menunjukkan bahwa error handling yang diimplementasikan berfungsi dengan baik."),

  h3("4.4.3 Efektivitas Metode Super Enkripsi + LSB"),
  p("Kombinasi Super Enkripsi dengan LSB terbukti efektif dalam menciptakan sistem keamanan ganda. Efektivitas ini ditunjukkan oleh dua fakta utama: (1) Gambar stego yang dihasilkan tidak dapat dibedakan secara visual dari gambar asli, sehingga keberadaan pesan tersembunyi tidak diketahui; (2) Meskipun pesan tersembunyi berhasil diekstrak dari gambar stego, pesan tersebut masih berupa ciphertext yang tidak dapat dibaca tanpa kunci enkripsi yang benar."),

  h3("4.4.4 Perbandingan dengan Penelitian Terdahulu"),
  p("Penelitian ini mengacu pada jurnal ilmiah 'Implementasi Steganografi Pada Citra Digital Menggunakan Metode Least Significant Bit Dengan Kriptografi Super Enkripsi' oleh Kaunang, Djamen, dan Kainde (2024) dari Universitas Negeri Manado. Penelitian tersebut mengimplementasikan sistem serupa menggunakan MATLAB, sementara penelitian ini mengimplementasikannya sebagai aplikasi web berbasis Python-Flask. Persamaannya terletak pada algoritma yang digunakan (Vigenere + AES + LSB), sementara perbedaan utamanya adalah platform: desktop MATLAB vs. web Python-Flask. Penelitian ini berhasil mereplikasi dan mengembangkan konsep tersebut ke platform web yang lebih mudah diakses."),
  pageBreak(),
];

// ========================
// BAB V
// ========================
const bab5 = [
  sectionTitle("BAB V"),
  sectionTitle("KESIMPULAN DAN SARAN"),

  h2("5.1 Kesimpulan"),
  pNoIndent("Berdasarkan hasil penelitian dan pengujian sistem yang telah dilakukan, dapat diambil kesimpulan sebagai berikut:"),
  numbered("Sistem berhasil diimplementasikan sebagai aplikasi web berbasis Flask yang mengintegrasikan kriptografi Super Enkripsi (Vigenere Cipher + AES-256 CBC) dan steganografi Least Significant Bit (LSB) untuk menyembunyikan pesan terenkripsi di dalam citra digital PNG."),
  numbered("Teknik Super Enkripsi dua lapis terbukti efektif dalam menghasilkan ciphertext yang kuat. Vigenere Cipher melakukan pergeseran karakter alfanumerik, dilanjutkan AES-256 CBC yang menghasilkan output Base64 acak yang sulit dipecahkan tanpa kunci."),
  numbered("Metode LSB berhasil menyisipkan ciphertext ke dalam piksel gambar PNG tanpa menghasilkan perubahan visual yang dapat dideteksi oleh mata manusia, terbukti dari identifikasi visual yang tidak menunjukkan perbedaan antara gambar asli dan gambar stego."),
  numbered("Sistem berhasil melakukan proses dekoding dan dekripsi dengan akurasi 100% selama kunci yang digunakan konsisten dengan kunci enkripsi, dan menangani kasus penggunaan kunci yang salah dengan pesan error yang informatif."),
  numbered("Antarmuka web yang dibangun dengan Flask dan HTML/CSS/JS berhasil menyediakan pengalaman pengguna yang intuitif dan responsif, memudahkan pengguna dalam melakukan proses encoding dan decoding tanpa pengetahuan teknis mendalam."),

  h2("5.2 Saran"),
  pNoIndent("Untuk pengembangan sistem lebih lanjut, beberapa saran berikut dapat dipertimbangkan:"),
  numbered("Pengembangan sistem dapat dilakukan dengan menambahkan dukungan untuk format media lain seperti file audio (WAV, MP3) dan video (AVI, MP4) sebagai media penampung pesan, sehingga memperluas fleksibilitas penggunaan."),
  numbered("Implementasi autentikasi pengguna (login system) dan enkripsi komunikasi HTTPS perlu ditambahkan untuk penggunaan di lingkungan produksi, guna memastikan keamanan end-to-end dari sistem."),
  numbered("Penambahan algoritma enkripsi lain seperti RSA, ChaCha20, atau Blum-Goldwasser sebagai pilihan lapisan enkripsi tambahan dalam Super Enkripsi akan meningkatkan fleksibilitas dan kekuatan kriptografi sistem."),
  numbered("Implementasi mekanisme penghapusan otomatis file yang diunggah ke server setelah proses selesai perlu dilakukan untuk menjaga privasi pengguna dan mencegah akumulasi file yang tidak diperlukan."),
  numbered("Pengembangan fitur analisis steganalysis untuk mendeteksi dan mengukur tingkat kesamaan (PSNR/MSE) antara gambar asli dan gambar stego dapat ditambahkan sebagai fitur evaluasi kualitas penyisipan."),
  pageBreak(),
];

// ========================
// LANGKAH MENJALANKAN
// ========================
const caraMenjalankan = [
  sectionTitle("LANGKAH-LANGKAH MENJALANKAN APLIKASI"),

  h2("A. Persyaratan Sistem"),
  p("Sebelum menjalankan aplikasi, pastikan sistem memiliki Python versi 3.8 atau lebih baru. Untuk memeriksa versi Python yang terinstal, buka Command Prompt (Windows) atau Terminal (Mac/Linux) dan ketikkan perintah berikut:"),
  ...empty(1),
  code("python --version"),
  code("# atau"),
  code("python3 --version"),
  ...empty(1),

  h2("B. Langkah Instalasi"),
  h3("Langkah 1: Unduh dan Ekstrak Proyek"),
  p("Unduh file ZIP proyek (crypto-stego-web.zip) dan ekstrak ke folder yang diinginkan, misalnya C:\\Projects\\crypto-stego-web (Windows) atau ~/Projects/crypto-stego-web (Mac/Linux)."),

  h3("Langkah 2: Buka Visual Studio Code"),
  numbered("Buka Visual Studio Code (VSCode)."),
  numbered("Pilih menu File > Open Folder."),
  numbered("Navigasi ke folder crypto-stego-web dan klik Select Folder."),

  h3("Langkah 3: Buka Terminal di VSCode"),
  numbered("Di VSCode, pilih menu Terminal > New Terminal (atau tekan Ctrl + ` )."),
  numbered("Sebuah panel terminal akan muncul di bagian bawah layar VSCode."),

  h3("Langkah 4: Buat Virtual Environment (Opsional, Disarankan)"),
  ...empty(1),
  code("# Windows"),
  code("python -m venv venv"),
  code("venv\\Scripts\\activate"),
  code(""),
  code("# Mac/Linux"),
  code("python3 -m venv venv"),
  code("source venv/bin/activate"),
  ...empty(1),

  h3("Langkah 5: Instal Dependensi"),
  p("Instal semua library yang diperlukan menggunakan pip:"),
  ...empty(1),
  code("pip install flask==3.0.3"),
  code("pip install pycryptodome"),
  code("pip install stegano"),
  code(""),
  code("# Atau instal sekaligus menggunakan requirements.txt (jika tersedia):"),
  code("pip install -r requirements.txt"),
  code("pip install pycryptodome stegano"),
  ...empty(1),

  h3("Langkah 6: Jalankan Aplikasi"),
  ...empty(1),
  code("python app.py"),
  ...empty(1),
  p("Jika berhasil, terminal akan menampilkan output sebagai berikut:"),
  ...empty(1),
  code(" * Serving Flask app 'app'"),
  code(" * Debug mode: on"),
  code(" * Running on http://127.0.0.1:5000"),
  code("Press CTRL+C to quit"),
  ...empty(1),

  h3("Langkah 7: Akses Aplikasi"),
  numbered("Buka browser (Google Chrome, Firefox, atau Edge)."),
  numbered("Ketik alamat: http://127.0.0.1:5000 pada address bar."),
  numbered("Tekan Enter. Halaman utama aplikasi Crypto & Steganography System akan tampil."),

  h2("C. Cara Menggunakan Fitur Encode"),
  numbered("Pada bagian 'Encode Pesan ke PNG', masukkan teks pesan rahasia yang ingin disembunyikan pada kolom textarea."),
  numbered("Masukkan kunci enkripsi pada kolom 'Masukkan key'. Ingat kunci ini karena diperlukan untuk proses decode."),
  numbered("Klik tombol 'Pilih Gambar PNG' dan pilih file gambar PNG dari komputer Anda."),
  numbered("Preview gambar akan ditampilkan setelah gambar dipilih."),
  numbered("Klik tombol hijau 'Encode & Download PNG'."),
  numbered("File gambar stego (stego.png) akan otomatis diunduh ke folder Downloads."),

  h2("D. Cara Menggunakan Fitur Decode"),
  numbered("Pada bagian 'Decode Stego PNG', masukkan kunci enkripsi yang sama dengan yang digunakan saat encoding."),
  numbered("Klik tombol 'Upload Stego PNG' dan pilih file gambar stego (stego.png)."),
  numbered("Klik tombol merah 'Decode Pesan'."),
  numbered("Hasil dekripsi teks asli akan ditampilkan di bagian bawah halaman dalam kotak 'Hasil Dekripsi'."),

  h2("E. Menghentikan Aplikasi"),
  p("Untuk menghentikan server Flask, tekan Ctrl + C pada terminal tempat aplikasi berjalan."),
  pageBreak(),
];

// ========================
// LAMPIRAN
// ========================
const lampiran = [
  sectionTitle("LAMPIRAN"),
  h2("Lampiran A: Kode Program Lengkap"),
  h3("A.1 app.py (File Utama Flask)"),
  code("from flask import Flask"),
  code("from flask import render_template"),
  code("from flask import request"),
  code("from flask import send_file"),
  code("from flask import flash"),
  code("from flask import redirect"),
  code("import os"),
  code("from encryption.super_encryption import encrypt_text"),
  code("from encryption.super_encryption import decrypt_text"),
  code("from steganography.encoder import encode_message"),
  code("from steganography.decoder import decode_message"),
  code(""),
  code("app = Flask(__name__)"),
  code("app.secret_key = \"kriptografi_stego\""),
  code("UPLOAD_FOLDER = \"uploads\""),
  code("os.makedirs(UPLOAD_FOLDER, exist_ok=True)"),
  code(""),
  code("@app.route(\"/\")"),
  code("def index():"),
  code("    return render_template(\"index.html\")"),
  code(""),
  code("@app.route(\"/encode\", methods=[\"POST\"])"),
  code("def encode():"),
  code("    try:"),
  code("        text  = request.form.get(\"text\")"),
  code("        key   = request.form.get(\"key\")"),
  code("        image = request.files.get(\"image\")"),
  code("        if not text:"),
  code("            flash(\"Masukkan pesan terlebih dahulu\")"),
  code("            return redirect(\"/\")"),
  code("        if not key:"),
  code("            flash(\"Masukkan key terlebih dahulu\")"),
  code("            return redirect(\"/\")"),
  code("        if image.filename == \"\":"),
  code("            flash(\"Pilih gambar PNG\")"),
  code("            return redirect(\"/\")"),
  code("        original_path = os.path.join(UPLOAD_FOLDER, image.filename)"),
  code("        image.save(original_path)"),
  code("        cipher_text = encrypt_text(text, key)"),
  code("        stego_path  = os.path.join(UPLOAD_FOLDER, \"stego.png\")"),
  code("        encode_message(original_path, cipher_text, stego_path)"),
  code("        return send_file(stego_path, as_attachment=True)"),
  code("    except Exception as e:"),
  code("        return f\"ERROR : {str(e)}\""),
  code(""),
  code("@app.route(\"/decode\", methods=[\"POST\"])"),
  code("def decode():"),
  code("    try:"),
  code("        key   = request.form.get(\"key_decode\")"),
  code("        image = request.files.get(\"stego_image\")"),
  code("        if not key:"),
  code("            flash(\"Masukkan key\")"),
  code("            return redirect(\"/\")"),
  code("        if image.filename == \"\":"),
  code("            flash(\"Pilih gambar stego\")"),
  code("            return redirect(\"/\")"),
  code("        path = os.path.join(UPLOAD_FOLDER, image.filename)"),
  code("        image.save(path)"),
  code("        cipher_text = decode_message(path)"),
  code("        if cipher_text is None:"),
  code("            flash(\"Pesan tidak ditemukan\")"),
  code("            return redirect(\"/\")"),
  code("        plain_text = decrypt_text(cipher_text, key)"),
  code("        return render_template(\"index.html\", decoded_text=plain_text)"),
  code("    except Exception as e:"),
  code("        return f\"ERROR : {str(e)}\""),
  code(""),
  code("if __name__ == \"__main__\":"),
  code("    app.run(debug=True)"),
  ...empty(1),

  h3("A.2 encryption/vigenere.py"),
  code("def encrypt_vigenere(text, key):"),
  code("    result = \"\""),
  code("    key = key.upper()"),
  code("    j = 0"),
  code("    for char in text:"),
  code("        if char.isalpha():"),
  code("            shift = ord(key[j % len(key)]) - ord('A')"),
  code("            if char.isupper():"),
  code("                result += chr((ord(char)-65+shift)%26+65)"),
  code("            else:"),
  code("                result += chr((ord(char)-97+shift)%26+97)"),
  code("            j += 1"),
  code("        else:"),
  code("            result += char"),
  code("    return result"),
  code(""),
  code("def decrypt_vigenere(text, key):"),
  code("    result = \"\""),
  code("    key = key.upper()"),
  code("    j = 0"),
  code("    for char in text:"),
  code("        if char.isalpha():"),
  code("            shift = ord(key[j % len(key)]) - ord('A')"),
  code("            if char.isupper():"),
  code("                result += chr((ord(char)-65-shift)%26+65)"),
  code("            else:"),
  code("                result += chr((ord(char)-97-shift)%26+97)"),
  code("            j += 1"),
  code("        else:"),
  code("            result += char"),
  code("    return result"),
  ...empty(1),

  h3("A.3 encryption/aes_cipher.py"),
  code("from Crypto.Cipher import AES"),
  code("from Crypto.Util.Padding import pad"),
  code("from Crypto.Util.Padding import unpad"),
  code("from hashlib import sha256"),
  code("import base64"),
  code(""),
  code("class AESCipher:"),
  code("    @staticmethod"),
  code("    def encrypt(text, key):"),
  code("        key = sha256(key.encode()).digest()"),
  code("        cipher = AES.new(key, AES.MODE_CBC)"),
  code("        encrypted = cipher.encrypt("),
  code("            pad(text.encode(), AES.block_size)"),
  code("        )"),
  code("        return base64.b64encode("),
  code("            cipher.iv + encrypted"),
  code("        ).decode()"),
  code(""),
  code("    @staticmethod"),
  code("    def decrypt(cipher_text, key):"),
  code("        key = sha256(key.encode()).digest()"),
  code("        raw = base64.b64decode(cipher_text)"),
  code("        iv = raw[:16]"),
  code("        encrypted = raw[16:]"),
  code("        cipher = AES.new(key, AES.MODE_CBC, iv)"),
  code("        return unpad("),
  code("            cipher.decrypt(encrypted),"),
  code("            AES.block_size"),
  code("        ).decode()"),
  ...empty(1),

  h3("A.4 encryption/super_encryption.py"),
  code("from encryption.vigenere import encrypt_vigenere"),
  code("from encryption.vigenere import decrypt_vigenere"),
  code("from encryption.aes_cipher import AESCipher"),
  code(""),
  code("def encrypt_text(text, key):"),
  code("    step1 = encrypt_vigenere(text, key)"),
  code("    step2 = AESCipher.encrypt(step1, key)"),
  code("    return step2"),
  code(""),
  code("def decrypt_text(cipher, key):"),
  code("    step1 = AESCipher.decrypt(cipher, key)"),
  code("    step2 = decrypt_vigenere(step1, key)"),
  code("    return step2"),
  ...empty(1),

  h3("A.5 steganography/encoder.py"),
  code("from stegano import lsb"),
  code(""),
  code("def encode_message(image_path, secret_message, output_path):"),
  code("    secret = lsb.hide(image_path, secret_message)"),
  code("    secret.save(output_path)"),
  code("    return output_path"),
  ...empty(1),

  h3("A.6 steganography/decoder.py"),
  code("from stegano import lsb"),
  code(""),
  code("def decode_message(image_path):"),
  code("    return lsb.reveal(image_path)"),
  ...empty(1),

  h2("Lampiran B: requirements.txt"),
  code("Flask==3.0.3"),
  code("pycryptodome"),
  code("stegano"),
  ...empty(1),
  pageBreak(),
];

// ========================
// DAFTAR PUSTAKA
// ========================
const daftarPustaka = [
  sectionTitle("DAFTAR PUSTAKA"),
  ...empty(1),
  pNoIndent("Kaunang, V. V. R., Djamen, A. C., & Kainde, Q. C. (2024). Implementasi Steganografi pada Citra Digital Menggunakan Metode Least Significant Bit dengan Kriptografi Super Enkripsi. Journal of Informatics, Business, Education, and Innovation Technology, 2(3), 75-89."),
  ...empty(1),
  pNoIndent("Ahmad, A. (2019). Implementasi Vigenere Cipher dengan Menggunakan Matlab 2015b. Prosiding Seminar Pendidikan Matematika dan Matematika, 1, 53-58."),
  ...empty(1),
  pNoIndent("Naharuddin, A. (2018). Steganografi Teks Menggunakan Pemetaan Digit Biner pada Karakter ASCII untuk Keamanan Plain Text (Doctoral dissertation). Institut Teknologi Sepuluh Nopember."),
  ...empty(1),
  pNoIndent("Langi, E. R., Sambul, A. M., & Kambey, F. D. (2021). Perbandingan Metode Least Significant Bit dan Discrete Wavelet Transform dalam Teknik Steganografi pada Citra Batik Bentenan. Skripsi Program S1 Teknik Elektro Universitas Sam Ratulangi. Manado."),
  ...empty(1),
  pNoIndent("Fachrozi, M. F., & Fahmi, H. (2020). Penerapan Metode AES-128 untuk Pengamanan Data Absensi Fingerprint di Balai Penelitian Sungei Putih. Jurnal Ilmu Komputer dan Sistem Informasi, 3(1.1), 1-8."),
  ...empty(1),
  pNoIndent("Pricillia, T., & Zulfachmi. (2021). Survey Paper: Perbandingan Metode Pengembangan Perangkat Lunak (Waterfall, Prototype, RAD). Jurnal Informatika dan Rekayasa Perangkat Lunak, 10(1), 6-12."),
  ...empty(1),
  pNoIndent("Wiranata, A. D., & Aldisa, R. T. (2021). Aplikasi Steganografi Menggunakan Least Significant Bit (LSB) dengan Enkripsi Caesar Cipher dan Rivest Code 4 (RC4). Jurnal JTIK (Jurnal Teknologi Informasi dan Komunikasi), 5(3), 277-281."),
  ...empty(1),
  pNoIndent("Santoso, S. D. (2019). Implementasi Penyandian Super Enkripsi Vigenere Cipher dan Railfence Cipher Menggunakan Python (Skripsi S1). Jurusan Matematika, Fakultas Sains dan Teknologi, Universitas Islam Negeri Maulana Malik Ibrahim Malang."),
  ...empty(1),
  pNoIndent("Irawan, C. (2019). Implementasi Algoritma Autokey Cipher dan AES-128 pada Enkripsi File. Prosiding Fakultas Ilmu Komputer Universitas Dian Nuswantoro, 335-339."),
  ...empty(1),
  pNoIndent("Grinberg, M. (2018). Flask Web Development: Developing Web Applications with Python (2nd ed.). O'Reilly Media."),
  ...empty(1),
  pNoIndent("National Institute of Standards and Technology (NIST). (2001). Advanced Encryption Standard (AES) (FIPS Publication 197). U.S. Department of Commerce. Diakses dari https://nvlpubs.nist.gov/nistpubs/FIPS/NIST.FIPS.197.pdf"),
];

// ========================
// BUILD DOCUMENT
// ========================
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Times New Roman", color: "1F3864" },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman", color: "2E74B5" },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1800 }
      }
    },
    children: [
      ...halamanJudul,
      ...abstrak,
      ...bab1,
      ...bab2,
      ...bab3,
      ...bab4,
      ...bab5,
      ...caraMenjalankan,
      ...lampiran,
      ...daftarPustaka,
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/Laporan_UAS_CryptoStego.docx", buffer);
  console.log("Dokumen berhasil dibuat!");
});