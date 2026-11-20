# ujiancbt-bersama2027
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CBT-PRO | Sistem Ujian Berbasis Komputer</title>
    
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- Lucide Icons CDN -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- SheetJS (XLSX) CDN untuk Import/Export Excel -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <!-- Mammoth.js CDN untuk membaca file Word (.docx) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <!-- PDF.js CDN untuk membaca file PDF (.pdf) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        };
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    </script>

    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            overflow-x: hidden;
        }

        /* Mencegah seleksi teks saat siswa mengerjakan ujian */
        .no-select {
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
        }

        /* Animasi latar belakang layar terkunci anti-curang */
        @keyframes lockPulse {
            0%, 100% { background-color: rgba(127, 29, 29, 0.98); }
            50% { background-color: rgba(185, 28, 28, 0.98); }
        }
        .lock-overlay-bg {
            animation: lockPulse 2s infinite;
        }

        /* Aturan Cetak (Print A4 2 Kolom) Khusus Kartu Login Peserta */
        @media print {
            body * {
                visibility: hidden;
            }
            #printCardsContainer, #printCardsContainer * {
                visibility: visible;
            }
            #printCardsContainer {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                display: grid !important;
                grid-template-columns: repeat(2, 1fr) !important;
                gap: 12px !important;
                padding: 8px !important;
                background: #ffffff !important;
            }
            .no-print {
                display: none !important;
            }
            .login-card-item {
                break-inside: avoid;
                page-break-inside: avoid;
                border: 1.5px solid #0f172a !important;
                box-shadow: none !important;
            }
        }

        /* Scrollbar halus */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col">

    <!-- ===================================================================== -->
    <!-- 1. SISTEM TOAST & DIALOG KONFIRMASI (PENGGANTI ALERT & CONFIRM)       -->
    <!-- ===================================================================== -->
    <div id="toastContainer" class="fixed top-4 right-4 z-[9999] flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

    <div id="confirmModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9990] hidden items-center justify-center p-4">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center gap-3 mb-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                    <i data-lucide="alert-triangle" class="w-5 h-5"></i>
                </div>
                <h3 id="confirmTitle" class="text-lg font-bold text-slate-900">Konfirmasi</h3>
            </div>
            <p id="confirmMessage" class="text-sm text-slate-600 mb-6 leading-relaxed">Apakah Anda yakin?</p>
            <div class="flex justify-end gap-2.5">
                <button id="confirmCancelBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 transition">Batal</button>
                <button id="confirmOkBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-white bg-blue-600 hover:bg-blue-700 transition shadow-sm">Ya, Lanjutkan</button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 2. OVERLAY LAYAR TERKUNCI ANTI-CURANG (LOCK SCREEN PIN PENGUJI)       -->
    <!-- ===================================================================== -->
    <div id="antiCheatLockOverlay" class="fixed inset-0 lock-overlay-bg z-[99999] hidden flex-col items-center justify-center p-4 sm:p-6 text-white text-center select-none">
        <div class="max-w-lg w-full bg-red-950/95 border-2 border-red-400 rounded-3xl p-6 sm:p-8 shadow-2xl">
            <div class="w-20 h-20 rounded-full bg-red-600/30 border-2 border-red-400 flex items-center justify-center mx-auto mb-4 animate-bounce">
                <i data-lucide="lock" class="w-10 h-10 text-red-200"></i>
            </div>
            <span class="inline-block px-3 py-1 rounded-full bg-red-600 text-white text-xs font-extrabold uppercase tracking-wider mb-3">
                Sistem Proteksi Ujian • Pelanggaran Terdeteksi
            </span>
            <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight mb-2">LAYAR UJIAN TERKUNCI!</h1>
            <p id="lockReasonText" class="text-red-200 text-xs sm:text-sm mb-5 leading-relaxed">
                Sistem mendeteksi aktivitas keluar dari halaman ujian, berpindah tab/aplikasi lain, atau menekan tombol terlarang.
            </p>

            <div class="bg-red-900/70 border border-red-700 rounded-2xl p-4 mb-5 text-left text-xs space-y-2">
                <div class="flex justify-between">
                    <span class="text-red-300">Nama Peserta:</span>
                    <span id="lockStudentName" class="font-bold text-white">-</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">NISN / Kelas:</span>
                    <span id="lockStudentMeta" class="font-semibold text-white">-</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">Total Pelanggaran Tercatat:</span>
                    <span id="lockViolationBadge" class="font-extrabold text-amber-300">1 Kali</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">Laporan Real-Time:</span>
                    <span class="font-semibold text-emerald-300">Terkirim ke Pengawas & Admin</span>
                </div>
            </div>

            <div class="space-y-3">
                <label class="block text-xs font-bold uppercase tracking-wider text-red-200">
                    Masukkan Kode PIN Pengawas untuk Membuka Kunci
                </label>
                <input type="password" id="unlockPinField" placeholder="Masukkan Kode PIN Pengawas"
                    class="w-full px-4 py-3.5 rounded-xl bg-white text-slate-900 font-extrabold text-center text-lg tracking-widest focus:outline-none focus:ring-4 focus:ring-amber-400"
                    onkeypress="if(event.key === 'Enter') verifyPinUnlock()">
                <button onclick="verifyPinUnlock()" class="w-full py-3.5 px-6 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-sm transition shadow-lg flex items-center justify-center gap-2">
                    <i data-lucide="unlock" class="w-4 h-4"></i> BUKA KUNCI LAYAR UJIAN
                </button>
            </div>
            <p class="mt-4 text-[11px] text-red-300/90">
                Silakan lapor kepada Pengawas Ruangan untuk verifikasi dan pembukaan kunci ujian.
            </p>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 3. HEADER / NAVBAR UTAMA                                              -->
    <!-- ===================================================================== -->
    <header id="mainHeader" class="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16 sm:h-18 gap-3">
                <!-- Identitas Sekolah -->
                <div class="flex items-center gap-3 cursor-pointer min-w-0" onclick="navigateTo('portal')">
                    <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex items-center justify-center shadow-md shrink-0">
                        <i data-lucide="shield-check" class="w-6 h-6"></i>
                    </div>
                    <div class="min-w-0">
                        <div class="flex items-center gap-2">
                            <h1 id="headerSchoolTitle" class="font-extrabold text-sm sm:text-base text-slate-900 truncate">
                                CBT-PRO | Sistem Ujian Berbasis Komputer
                            </h1>
                            <span class="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0">
                                <span class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping"></span> LIVE
                            </span>
                        </div>
                        <p id="headerSchoolSubtitle" class="text-[11px] text-slate-500 truncate">
                            Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027
                        </p>
                    </div>
                </div>

                <!-- Tombol Aksi Header -->
                <div class="flex items-center gap-2 shrink-0">
                    <div id="headerGuestButtons" class="flex items-center gap-1.5 sm:gap-2">
                        <button onclick="openLoginModal('siswa')" class="px-3 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                            <i data-lucide="graduation-cap" class="w-4 h-4"></i>
                            <span>Login Siswa</span>
                        </button>
                        <button onclick="openLoginModal('pengawas')" class="px-3 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="user-check" class="w-4 h-4"></i>
                            <span>Pengawas</span>
                        </button>
                        <button onclick="openLoginModal('admin')" class="px-3 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="settings" class="w-4 h-4"></i>
                            <span>Admin</span>
                        </button>
                    </div>

                    <div id="headerAuthControls" class="hidden items-center gap-2">
                        <span id="headerUserBadge" class="hidden sm:inline-block px-3 py-1.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">
                            -
                        </span>
                        <button onclick="navigateTo('portal')" class="px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition">
                            Portal
                        </button>
                        <button onclick="logoutSession()" class="px-3 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1">
                            <i data-lucide="log-out" class="w-3.5 h-3.5"></i> Keluar
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <!-- ===================================================================== -->
    <!-- 4. KONTEN UTAMA (PORTAL, DASHBOARD ADMIN, PENGAWAS, UJIAN SISWA)      -->
    <!-- ===================================================================== -->
    <main class="flex-grow">

        <!-- ================================================================= -->
        <!-- VIEW A: HALAMAN PORTAL UTAMA (LANDING PORTAL)                     -->
        <!-- ================================================================= -->
        <section id="view-portal" class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-10">

            <!-- Hero Section + Portal Login Resmi + Salin Link Akses -->
            <div class="bg-gradient-to-br from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-6 sm:p-10 text-white shadow-xl relative overflow-hidden">
                <div class="absolute -top-24 -right-24 w-80 h-80 bg-blue-500/15 rounded-full blur-3xl pointer-events-none"></div>
                <div class="relative z-10 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
                    <div class="lg:col-span-7 space-y-5">
                        <div class="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-blue-500/20 border border-blue-400/30 text-blue-200 text-xs font-semibold">
                            <i data-lucide="shield-alert" class="w-4 h-4 text-amber-300"></i>
                            <span>Sistem Ujian Resmi • Token Wajib • Anti-Curang Otomatis</span>
                        </div>
                        <h2 class="text-2xl sm:text-4xl font-extrabold tracking-tight leading-tight">
                            Selamat Datang di Portal <span class="text-blue-400" id="heroExamTitle">Ujian Akhir Semester</span>
                        </h2>
                        <p class="text-slate-300 text-sm sm:text-base leading-relaxed">
                            Silakan masuk sesuai peran Anda. Peserta ujian wajib menggunakan <strong>NISN / Username</strong> dan <strong>Password</strong> resmi yang tertera pada <strong>Kartu Login Peserta</strong> serta <strong>Token Ujian Aktif</strong> dari Pengawas Ruangan.
                        </p>

                        <!-- Tombol Login 3 Peran -->
                        <div class="flex flex-wrap gap-3 pt-1">
                            <button onclick="openLoginModal('siswa')" class="px-5 py-3 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs sm:text-sm transition flex items-center gap-2 shadow-lg shadow-blue-950/60">
                                <i data-lucide="pen-tool" class="w-4 h-4"></i> Masuk Ujian Siswa
                            </button>
                            <button onclick="openLoginModal('pengawas')" class="px-4 py-3 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="eye" class="w-4 h-4"></i> Portal Pengawas
                            </button>
                            <button onclick="openLoginModal('admin')" class="px-4 py-3 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/20 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="sliders" class="w-4 h-4"></i> Administrator
                            </button>
                        </div>

                        <!-- Tombol Salin Link Akses -->
                        <div class="pt-3 border-t border-white/10 flex flex-wrap items-center gap-2 text-xs">
                            <span class="text-slate-400 font-medium mr-1">Bagikan Link Akses:</span>
                            <button onclick="copyPortalLink('siswa')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-blue-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Portal Siswa
                            </button>
                            <button onclick="copyPortalLink('pengawas')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-amber-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Pengawas
                            </button>
                            <button onclick="copyPortalLink('admin')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-indigo-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Administrator
                            </button>
                        </div>
                    </div>

                    <!-- Kartu Kata-Kata Motivasi Belajar Interaktif -->
                    <div class="lg:col-span-5">
                        <div class="bg-white/10 backdrop-blur-md border border-white/20 rounded-3xl p-6 space-y-4">
                            <div class="flex items-center justify-between">
                                <span class="inline-flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-amber-300">
                                    <i data-lucide="sparkles" class="w-4 h-4"></i> Motivasi Belajar & Kejujuran
                                </span>
                                <button onclick="nextMotivationQuote()" class="px-2.5 py-1 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white transition flex items-center gap-1" title="Ganti Kutipan Motivasi">
                                    <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i> Acak
                                </button>
                            </div>
                            <blockquote class="space-y-3">
                                <p id="motivationQuoteText" class="text-sm sm:text-base font-medium italic text-white leading-relaxed">
                                    "Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."
                                </p>
                                <footer id="motivationQuoteAuthor" class="text-xs font-bold text-blue-300">
                                    — Pesan Akademik Sekolah
                                </footer>
                            </blockquote>
                            <div class="pt-3 border-t border-white/10 grid grid-cols-3 gap-2 text-center">
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSoal" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Bank Soal</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSiswa" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Peserta Aktif</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatDurasi" class="text-lg font-extrabold text-amber-300">45m</div>
                                    <div class="text-[10px] text-slate-400">Durasi Ujian</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- BAGIAN PANDUAN SISTEM UJIAN CBT (SISWA, PENGAWAS, ADMIN) -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Tata Tertib & Alur Pelaksanaan</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Panduan Sistem Ujian CBT</h2>
                    </div>
                    <p class="text-xs text-slate-500">Pelajari prosedur resmi pelaksanaan ujian berbasis komputer</p>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <!-- Panduan Siswa -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                                <i data-lucide="user-check" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">Panduan Peserta / Siswa</h3>
                                <span class="text-xs text-blue-600 font-semibold">4 Langkah Pelaksanaan</span>
                            </div>
                        </div>
                        <ol class="space-y-2.5 text-xs text-slate-600 list-decimal pl-4 leading-relaxed">
                            <li><strong>Kartu Login Resmi:</strong> Siapkan Kartu Login Peserta yang berisi <em>NISN / Username</em> dan <em>Password</em> resmi Anda.</li>
                            <li><strong>Input Token Aktif:</strong> Masukkan <strong>Token Ujian Aktif</strong> dari pengawas setiap kali masuk sesi ujian.</li>
                            <li><strong>Proteksi Layar Penuh:</strong> Dilarang membuka tab/aplikasi lain atau mengecilkan layar. Sistem akan <strong>mengunci layar otomatis</strong> jika terjadi pelanggaran.</li>
                            <li><strong>Kirim Jawaban:</strong> Periksa navigasi nomor soal, lalu tekan <em>Selesai Ujian</em> untuk mengirim jawaban secara real-time.</li>
                        </ol>
                        <div class="p-3 rounded-xl bg-blue-50 border border-blue-100 text-[11px] text-blue-900 leading-relaxed">
                            <strong>Catatan Peserta:</strong><br>
                            Gunakan NISN/Username dan Password yang tercantum pada Kartu Login Peserta resmi serta Token Ujian Aktif dari Pengawas Ruang.
                        </div>
                    </div>

                    <!-- Panduan Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                                <i data-lucide="shield" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">Panduan Pengawas Ruang</h3>
                                <span class="text-xs text-amber-600 font-semibold">Tugas & Wewenang Pengawas</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Distribusi Token:</strong> Membagikan Token Ujian Aktif kepada peserta di ruangan saat sesi ujian dimulai.</li>
                            <li><strong>Monitoring Real-Time:</strong> Memantau status kehadiran, progres pengerjaan soal, serta catatan pelanggaran siswa secara langsung.</li>
                            <li><strong>Verifikasi Buka Kunci:</strong> Memberikan kode PIN pembuka kunci atau melakukan <em>Buka Kunci Layar</em> jarak jauh setelah memeriksa kondisi perangkat peserta.</li>
                            <li><strong>Unggah Soal Tambahan:</strong> Dapat membantu mengunggah naskah soal tambahan jika ditugaskan.</li>
                        </ul>
                        <div class="p-3 rounded-xl bg-amber-50 border border-amber-100 text-[11px] text-amber-900 leading-relaxed">
                            <strong>Akses Pengawas:</strong><br>
                            Masuk menggunakan Username dan Password Pengawas Ruang yang telah didaftarkan oleh Administrator.
                        </div>
                    </div>

                    <!-- Panduan Administrator -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                                <i data-lucide="cpu" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">Panduan Administrator</h3>
                                <span class="text-xs text-indigo-600 font-semibold">Manajemen Penuh Sistem</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Bank & Upload Soal:</strong> Import soal otomatis dari file <strong>Word (.docx)</strong>, <strong>PDF (.pdf)</strong>, maupun input manual dan edit soal.</li>
                            <li><strong>Akun Siswa & Kartu Login:</strong> Kelola akun siswa berbasis NISN, import template Excel, edit/reset password siswa, serta cetak kartu login siap bagi.</li>
                            <li><strong>Assign & Token:</strong> Atur durasi, acak soal/opsi, serta perbarui Token Ujian dan PIN keamanan.</li>
                            <li><strong>Download Nilai Excel:</strong> Unduh rekap hasil ujian real-time langsung dalam format <strong>Excel (.xlsx)</strong>.</li>
                        </ul>
                        <div class="p-3 rounded-xl bg-indigo-50 border border-indigo-100 text-[11px] text-indigo-900 leading-relaxed">
                            <strong>Panel Administrator:</strong><br>
                            Akses penuh pengelolaan bank soal, token ujian, akun peserta, kartu login, dan rekap nilai sekolah.
                        </div>
                    </div>
                </div>
            </div>

            <!-- BAGIAN GALERI FOTO KEGIATAN UJIAN -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-600">Dokumentasi & Suasana Ujian</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Galeri Kegiatan Ujian Berbasis Komputer</h2>
                    </div>
                    <p class="text-xs text-slate-500">Dokumentasi kegiatan dapat dikelola melalui menu Setelan Administrator</p>
                </div>

                <div id="portalGalleryGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
                    <!-- Diisi dinamis oleh JS -->
                </div>
            </div>

            <!-- BAGIAN DOWNLOAD & INSTALASI APLIKASI (ANDROID, IPHONE/IOS, KOMPUTER) -->
            <div class="bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm space-y-6">
                <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-5 border-b border-slate-100">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Aplikasi Multi-Platform</span>
                        <h2 class="text-xl font-extrabold text-slate-900 mt-1">Download & Pasang Aplikasi di HP Android, iPhone, maupun Komputer</h2>
                        <p class="text-xs sm:text-sm text-slate-500 mt-1">Pasang sebagai aplikasi mandiri layar penuh (Safe Exam WebApp) untuk kenyamanan dan keamanan ujian maksimal.</p>
                    </div>
                    <button onclick="downloadStandaloneApp()" class="px-4 py-3 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-2 self-start md:self-center shrink-0 shadow-sm">
                        <i data-lucide="download" class="w-4 h-4"></i> Unduh File Aplikasi (.HTML)
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <!-- Android -->
                    <div class="rounded-2xl border border-emerald-200 bg-emerald-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">HP / Tablet Android</h3>
                                    <span class="text-[11px] font-semibold text-emerald-700">Android 8.0+ (Chrome / WebApp)</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang aplikasi CBT langsung ke layar utama Android tanpa bilah alamat browser sehingga berjalan seperti aplikasi mandiri.
                            </p>
                        </div>
                        <button onclick="openInstallModal('android')" class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="download-cloud" class="w-4 h-4"></i> Install / Panduan Android
                        </button>
                    </div>

                    <!-- iPhone / iPad (iOS) -->
                    <div class="rounded-2xl border border-slate-200 bg-slate-50/70 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-slate-900 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="tablet-smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">iPhone & iPad (iOS)</h3>
                                    <span class="text-[11px] font-semibold text-slate-600">iOS Safari WebClip / PWA</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Gunakan fitur <em>"Tambahkan ke Layar Utama"</em> di Safari iPhone/iPad untuk menjalankan CBT dalam mode layar penuh yang stabil.
                            </p>
                        </div>
                        <button onclick="openInstallModal('ios')" class="w-full py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="smartphone" class="w-4 h-4"></i> Install / Panduan iPhone (iOS)
                        </button>
                    </div>

                    <!-- Komputer / Laptop -->
                    <div class="rounded-2xl border border-blue-200 bg-blue-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="monitor" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">Komputer / Laptop (PC & Mac)</h3>
                                    <span class="text-[11px] font-semibold text-blue-700">Windows / macOS / Chromebook</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang sebagai aplikasi desktop mandiri lewat Chrome/Edge atau unduh file paket aplikasi offline untuk laboratorium komputer.
                            </p>
                        </div>
                        <button onclick="openInstallModal('desktop')" class="w-full py-2.5 px-4 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="laptop" class="w-4 h-4"></i> Install / Download Komputer
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW B: DASHBOARD ADMINISTRATOR (6 MENU SIDEBAR TERPISAH)         -->
        <!-- ================================================================= -->
        <section id="view-admin" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                
                <!-- SIDEBAR NAVIGASI 6 MENU TERPISAH -->
                <aside class="lg:col-span-3 bg-white rounded-2xl border border-slate-200 p-4 shadow-sm space-y-2 lg:sticky lg:top-20 no-print">
                    <div class="px-3 py-2 border-b border-slate-100 mb-2">
                        <span class="text-[10px] font-extrabold uppercase tracking-wider text-indigo-600">Panel Kendali Utama</span>
                        <h2 class="font-extrabold text-slate-900 text-base">Menu Administrator</h2>
                    </div>

                    <nav class="flex lg:flex-col overflow-x-auto lg:overflow-visible gap-1.5 pb-1 lg:pb-0">
                        <button onclick="switchAdminMenu('monitoring')" id="adminNav-monitoring" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap">
                            <i data-lucide="activity" class="w-4 h-4 shrink-0"></i>
                            <span>1. Monitoring Siswa</span>
                        </button>
                        <button onclick="switchAdminMenu('soal')" id="adminNav-soal" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-text" class="w-4 h-4 shrink-0"></i>
                            <span>2. Bank & Input Soal</span>
                        </button>
                        <button onclick="switchAdminMenu('assign')" id="adminNav-assign" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="key-round" class="w-4 h-4 shrink-0"></i>
                            <span>3. Assign Ujian & Token</span>
                        </button>
                        <button onclick="switchAdminMenu('nilai')" id="adminNav-nilai" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4 shrink-0"></i>
                            <span>4. Data & Download Nilai</span>
                        </button>
                        <button onclick="switchAdminMenu('kartu')" id="adminNav-kartu" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="id-card" class="w-4 h-4 shrink-0"></i>
                            <span>5. Akun Siswa & Kartu</span>
                        </button>
                        <button onclick="switchAdminMenu('setelan')" id="adminNav-setelan" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="settings-2" class="w-4 h-4 shrink-0"></i>
                            <span>6. Setelan Keseluruhan</span>
                        </button>
                    </nav>

                    <!-- Info Token Aktif di Sidebar -->
                    <div class="hidden lg:block mt-4 pt-4 border-t border-slate-100">
                        <div class="bg-slate-900 text-white rounded-xl p-3.5">
                            <div class="flex items-center justify-between text-[10px] text-slate-400 uppercase font-bold">
                                <span>Token Ujian Aktif</span>
                                <span class="text-emerald-400">LIVE</span>
                            </div>
                            <div class="flex items-center justify-between mt-1">
                                <span id="sidebarTokenDisplay" class="font-mono font-extrabold text-xl text-amber-400 tracking-widest">------</span>
                                <button onclick="copyTextToClipboard(document.getElementById('sidebarTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white" title="Salin Token">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </aside>

                <!-- KONTEN KANAN ADMINISTRATOR -->
                <div class="lg:col-span-9 space-y-6">

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 1: MONITORING SISWA (REAL-TIME)                  -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-monitoring" class="space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Monitoring Peserta Ujian Real-Time</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Pantau aktivitas siswa secara langsung, buka kunci layar pelanggaran, atau reset sesi peserta.</p>
                            </div>
                            <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                                <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Sinkronisasi Otomatis
                            </span>
                        </div>

                        <!-- 4 Kartu Statistik Monitoring -->
                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-4">
                            <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm">
                                <span class="text-xs text-slate-500 font-medium">Total Peserta</span>
                                <div id="monStatTotal" class="text-2xl font-extrabold text-slate-900 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-blue-200 p-4 shadow-sm">
                                <span class="text-xs text-blue-600 font-semibold">Sedang Ujian</span>
                                <div id="monStatActive" class="text-2xl font-extrabold text-blue-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-emerald-200 p-4 shadow-sm">
                                <span class="text-xs text-emerald-600 font-semibold">Selesai Ujian</span>
                                <div id="monStatDone" class="text-2xl font-extrabold text-emerald-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-red-200 p-4 shadow-sm">
                                <span class="text-xs text-red-600 font-semibold">Terkunci / Curang</span>
                                <div id="monStatLocked" class="text-2xl font-extrabold text-red-600 mt-1">0</div>
                            </div>
                        </div>

                        <!-- Filter & Tabel Live Monitoring -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="monSearchInput" oninput="renderAdminMonitoring()" placeholder="Cari nama atau NISN siswa..." 
                                    class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-64 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <select id="monRoomFilter" onchange="renderAdminMonitoring()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas / Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Peserta & NISN</th>
                                            <th class="py-3 px-4 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-4 font-bold">Status Live</th>
                                            <th class="py-3 px-4 font-bold">Progres</th>
                                            <th class="py-3 px-4 font-bold">Skor</th>
                                            <th class="py-3 px-4 font-bold">Pelanggaran</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Cepat</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminMonitoringTbody" class="divide-y divide-slate-100">
                                        <!-- Diisi via JS -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 2: BANK & INPUT SOAL (MANUAL, WORD/PDF, EDIT)    -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-soal" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Bank Soal & Input Otomatis (Word / PDF / Manual)</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Unggah naskah soal dari Microsoft Word (.docx) atau PDF (.pdf), tambahkan manual, serta edit soal langsung di sistem.</p>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6">
                            <!-- Kolom Input Soal -->
                            <div class="xl:col-span-5 space-y-6">
                                <!-- Upload Word (.docx) & PDF (.pdf) -->
                                <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-3">
                                    <div class="flex items-center justify-between">
                                        <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                            <i data-lucide="file-up" class="w-4 h-4 text-indigo-600"></i> Upload Soal Word (.docx) / PDF (.pdf)
                                        </h3>
                                        <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-indigo-600 hover:underline">
                                            Unduh Contoh Format
                                        </button>
                                    </div>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Format standar di dalam file Word/PDF:<br>
                                        <code class="block bg-slate-50 p-2 rounded-lg border border-slate-200 text-slate-700 mt-1">
                                            1. Pertanyaan soal di sini?<br>
                                            A. Pilihan A &nbsp; B. Pilihan B<br>
                                            C. Pilihan C &nbsp; D. Pilihan D &nbsp; E. Pilihan E<br>
                                            Kunci: B
                                        </code>
                                    </p>
                                    <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-indigo-300 rounded-2xl bg-indigo-50/40 hover:bg-indigo-50 cursor-pointer transition p-4 text-center">
                                        <i data-lucide="upload-cloud" class="w-7 h-7 text-indigo-600 mb-1"></i>
                                        <span class="text-xs font-bold text-indigo-900">Klik untuk Pilih File .DOCX, .PDF, atau .TXT</span>
                                        <span class="text-[10px] text-slate-500 mt-0.5">Otomatis mengekstrak soal, opsi A-E, dan kunci jawaban</span>
                                        <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event)" class="hidden">
                                    </label>
                                    <div id="uploadParseStatus" class="hidden text-xs p-2.5 rounded-xl bg-blue-50 text-blue-700 font-medium"></div>
                                </div>

                                <!-- Form Input Soal Manual -->
                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Soal Manual</h3>
                                    <form onsubmit="handleAddManualQuestion(event)" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                            <textarea id="mQText" required rows="2" placeholder="Tuliskan pertanyaan..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"></textarea>
                                        </div>
                                        <div class="space-y-2">
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">A.</span>
                                                <input type="text" id="mQOptA" required placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">B.</span>
                                                <input type="text" id="mQOptB" required placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">C.</span>
                                                <input type="text" id="mQOptC" required placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">D.</span>
                                                <input type="text" id="mQOptD" required placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">E.</span>
                                                <input type="text" id="mQOptE" placeholder="Opsi E (Opsional - kosongkan jika 4 opsi)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-3 pt-1">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                                <select id="mQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                                                    <option value="A">Opsi A</option>
                                                    <option value="B">Opsi B</option>
                                                    <option value="C">Opsi C</option>
                                                    <option value="D">Opsi D</option>
                                                    <option value="E">Opsi E</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin</label>
                                                <input type="number" id="mQPoints" value="10" min="1" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Soal ke Bank Soal
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <!-- Daftar & Edit Soal -->
                            <div class="xl:col-span-7 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col">
                                <div class="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Bank Soal (<span id="adminBankSoalCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Klik tombol Edit untuk mengubah pertanyaan, opsi, maupun kunci jawaban</p>
                                    </div>
                                    <button onclick="confirmClearAllQuestions()" class="px-3 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                        Hapus Semua
                                    </button>
                                </div>
                                <div id="adminSoalListContainer" class="space-y-3 overflow-y-auto max-h-[680px] pr-1">
                                    <!-- Render daftar soal + tombol Edit & Hapus -->
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 3: ASSIGN UJIAN & TOKEN                          -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-assign" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Assign Paket Ujian & Manajemen Token</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Tentukan paket ujian, target peserta, durasi waktu, pengacakan soal, serta atur Token Wajib Masuk Ujian.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                            <!-- Manajemen Token Ujian -->
                            <div class="lg:col-span-5 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5 self-start">
                                <div>
                                    <span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">Wajib Dimasukkan Siswa</span>
                                    <h3 class="font-bold text-slate-900 text-base mt-2">Token Akses Ujian Aktif</h3>
                                    <p class="text-xs text-slate-500 mt-0.5">Setiap kali siswa masuk sesi ujian, siswa wajib mengisi token ini.</p>
                                </div>

                                <div class="bg-slate-900 text-white p-5 rounded-2xl flex items-center justify-between">
                                    <div>
                                        <span class="text-[10px] uppercase text-slate-400 font-bold block">Token Saat Ini</span>
                                        <span id="assignTokenDisplay" class="text-3xl font-mono font-extrabold text-amber-400 tracking-widest">CBT26X</span>
                                    </div>
                                    <button onclick="copyTextToClipboard(document.getElementById('assignTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-3 rounded-xl bg-white/10 hover:bg-white/20 text-white transition" title="Salin Token">
                                        <i data-lucide="copy" class="w-5 h-5"></i>
                                    </button>
                                </div>

                                <!-- Ubah Token Manual -->
                                <div class="space-y-2 pt-2 border-t border-slate-100">
                                    <label class="block text-xs font-bold text-slate-700">Ubah Token Secara Manual</label>
                                    <div class="flex gap-2">
                                        <input type="text" id="customTokenInput" maxlength="10" placeholder="Misal: UJIAN01" 
                                            class="flex-1 px-3.5 py-2.5 rounded-xl border border-slate-300 font-mono font-bold uppercase text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                        <button onclick="saveManualToken()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan Token
                                        </button>
                                    </div>
                                </div>

                                <button onclick="generateRandomToken()" class="w-full py-3 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition flex items-center justify-center gap-2 shadow-sm">
                                    <i data-lucide="refresh-cw" class="w-4 h-4"></i> Acak / Generate Token Baru Otomatis
                                </button>
                            </div>

                            <!-- Form Penugasan (Assign) Ujian -->
                            <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base">Konfigurasi Penugasan (Assign) ke Siswa</h3>
                                <form onsubmit="saveAssignSettings(event)" class="space-y-4">
                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Judul Paket Ujian</label>
                                            <input type="text" id="assignExamTitle" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Mata Pelajaran</label>
                                            <input type="text" id="assignSubject" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Assign ke Kelas / Ruangan</label>
                                            <select id="assignTargetGroup" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-semibold">
                                                <option value="ALL">Semua Kelas & Ruangan</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Durasi Pengerjaan (Menit)</label>
                                            <input type="number" id="assignDuration" min="1" max="360" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm font-bold">
                                        </div>
                                    </div>

                                    <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-3">
                                        <label class="flex items-center justify-between cursor-pointer">
                                            <div>
                                                <span class="text-xs sm:text-sm font-bold text-slate-800 block">Acak Urutan Soal (Randomize Questions)</span>
                                                <span class="text-[11px] text-slate-500">Setiap siswa mendapatkan urutan nomor soal yang berbeda</span>
                                            </div>
                                            <input type="checkbox" id="assignShuffleQuestions" class="w-5 h-5 rounded text-indigo-600 focus:ring-indigo-500">
                                        </label>
                                        <div class="border-t border-slate-200 pt-3">
                                            <label class="flex items-center justify-between cursor-pointer">
                                                <div>
                                                    <span class="text-xs sm:text-sm font-bold text-slate-800 block">Acak Pilihan Jawaban (Randomize Options)</span>
                                                    <span class="text-[11px] text-slate-500">Mengacak posisi opsi A, B, C, D, E pada setiap soal</span>
                                                </div>
                                                <input type="checkbox" id="assignShuffleOptions" class="w-5 h-5 rounded text-indigo-600 focus:ring-indigo-500">
                                            </label>
                                        </div>
                                    </div>

                                    <div class="flex justify-end pt-2">
                                        <button type="submit" class="px-6 py-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                            <i data-lucide="check-circle" class="w-4 h-4"></i> Terapkan Penugasan Ujian
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 4: DATA SISWA & DOWNLOAD NILAI (.XLSX)           -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-nilai" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Data Peserta & Rekap Nilai Ujian</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Nilai masuk secara real-time dan dapat diunduh langsung dalam format Microsoft Excel (.xlsx).</p>
                            </div>
                            <button onclick="downloadExcelResults()" class="px-5 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2 shadow-sm self-start">
                                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Download Nilai Excel (.XLSX)
                            </button>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <!-- Filter Kelas / Ruangan & Pencarian -->
                            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                <input type="text" id="nilaiSearchInput" oninput="renderAdminNilaiTable()" placeholder="Cari Nama atau NISN..." 
                                    class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                <select id="nilaiClassFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas</option>
                                </select>
                                <select id="nilaiRoomFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-3 font-bold">No</th>
                                            <th class="py-3 px-3 font-bold">Nama Peserta</th>
                                            <th class="py-3 px-3 font-bold">NISN</th>
                                            <th class="py-3 px-3 font-bold">Kelas</th>
                                            <th class="py-3 px-3 font-bold">Ruangan</th>
                                            <th class="py-3 px-3 font-bold">Status</th>
                                            <th class="py-3 px-3 font-bold">Benar</th>
                                            <th class="py-3 px-3 font-bold">Salah</th>
                                            <th class="py-3 px-3 font-bold">Nilai Akhir</th>
                                            <th class="py-3 px-3 font-bold">Pelanggaran</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminNilaiTbody" class="divide-y divide-slate-100">
                                        <!-- Diisi via JS -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 5: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN        -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-kartu" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Manajemen Akun Siswa & Kartu Login Resmi</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola akun siswa siap digunakan berbasis NISN, edit/reset password, import Excel, dan cetak Kartu Login A4.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-2">
                                <button onclick="exportStudentAccountsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Export Akun & Password (.XLSX)
                                </button>
                                <button onclick="window.print()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="printer" class="w-4 h-4"></i> Cetak Kartu Login (A4)
                                </button>
                            </div>
                        </div>

                        <!-- Baris 1: Form Input Manual / Import Excel + Tabel Manajemen Akun Siswa -->
                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6 no-print">
                            <!-- Kolom Kiri: Tambah Manual & Import Excel -->
                            <div class="xl:col-span-4 space-y-5">
                                <!-- Import Excel Template -->
                                <div class="bg-white rounded-2xl border border-emerald-200 p-5 shadow-sm space-y-3">
                                    <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                        <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-600"></i> Import Data Siswa via Excel
                                    </h3>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Gunakan template Excel dengan kolom: <strong>Nama, NISN, Kelas, Ruangan</strong>. Username otomatis menggunakan NISN dan Password dibuat otomatis.
                                    </p>
                                    <button onclick="downloadStudentExcelTemplate()" class="w-full py-2.5 px-4 rounded-xl bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-bold transition flex items-center justify-center gap-2">
                                        <i data-lucide="download" class="w-4 h-4"></i> Download Template Excel (.XLSX)
                                    </button>
                                    <label class="flex flex-col items-center justify-center w-full py-3.5 px-3 border-2 border-dashed border-emerald-300 rounded-xl bg-emerald-50/30 hover:bg-emerald-50 cursor-pointer transition text-center">
                                        <span class="text-xs font-bold text-emerald-900"><i data-lucide="upload" class="w-3.5 h-3.5 inline mr-1"></i> Upload File Excel (.XLSX / .CSV)</span>
                                        <input type="file" accept=".xlsx,.xls,.csv" onchange="handleStudentExcelUpload(event)" class="hidden">
                                    </label>
                                </div>

                                <!-- Tambah Akun Siswa Manual -->
                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Akun Siswa Manual</h3>
                                    <form onsubmit="handleAddManualStudent(event)" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Nama Lengkap *</label>
                                            <input type="text" id="kStuName" required placeholder="Contoh: Nadia Putri" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">NISN *</label>
                                            <input type="text" id="kStuNisn" required placeholder="Contoh: 0081234504" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-mono">
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kelas *</label>
                                                <input type="text" id="kStuClass" required placeholder="XII IPA 1" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Ruangan *</label>
                                                <input type="text" id="kStuRoom" required placeholder="Ruang 01" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Username (Opsional)</label>
                                                <input type="text" id="kStuUser" placeholder="Otomatis NISN" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Password (Opsional)</label>
                                                <input type="text" id="kStuPass" placeholder="Otomatis Acak" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Akun & Buat Kartu
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <!-- Kolom Kanan: Tabel Daftar Akun Siswa Siap Digunakan (Edit & Reset Password) -->
                            <div class="xl:col-span-8 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col space-y-4">
                                <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Akun Siswa Aktif (<span id="studentAccountCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Akun berikut langsung aktif untuk login ujian menggunakan Token Aktif</p>
                                    </div>
                                    <div class="flex items-center gap-2">
                                        <button onclick="toggleShowStudentPasswords()" id="btnTogglePassVis" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eye" class="w-3.5 h-3.5"></i> <span>Tampilkan Password</span>
                                        </button>
                                        <button onclick="confirmClearAllStudents()" class="px-3 py-1.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                            Kosongkan Daftar Siswa
                                        </button>
                                    </div>
                                </div>

                                <input type="text" id="accountSearchInput" oninput="renderAdminStudentAccountTable()" placeholder="Cari berdasarkan Nama, NISN, Kelas, atau Ruangan..." 
                                    class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full focus:ring-2 focus:ring-indigo-500 focus:outline-none">

                                <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[380px] overflow-y-auto">
                                    <table class="w-full text-left border-collapse text-xs">
                                        <thead>
                                            <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                                <th class="py-2.5 px-3 font-bold">Nama & NISN</th>
                                                <th class="py-2.5 px-3 font-bold">Kelas / Ruang</th>
                                                <th class="py-2.5 px-3 font-bold">Username</th>
                                                <th class="py-2.5 px-3 font-bold">Password</th>
                                                <th class="py-2.5 px-3 font-bold text-right">Kelola Akun</th>
                                            </tr>
                                        </thead>
                                        <tbody id="adminStudentAccountTbody" class="divide-y divide-slate-100">
                                            <!-- Diisi via JS -->
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <!-- Baris 2: Pratinjau Kartu Login Peserta Siap Cetak (Print A4 2 Kolom) -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex items-center justify-between no-print">
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm sm:text-base">Pratinjau Kartu Login Peserta Siap Cetak (<span id="cardTotalCount">0</span> Kartu)</h3>
                                    <p class="text-[11px] text-slate-500">Format cetak otomatis 2 kolom per halaman A4 lengkap dengan QR Code identitas peserta</p>
                                </div>
                                <button onclick="window.print()" class="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="printer" class="w-3.5 h-3.5"></i> Cetak Sekarang
                                </button>
                            </div>
                            <div id="printCardsContainer" class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <!-- Render Kartu Login Peserta -->
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 6: SETELAN KESELURUHAN (PENGAWAS, SISWA, SISTEM) -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-setelan" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Setelan Keseluruhan Sistem, Pengawas & Siswa</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Kelola identitas sekolah, kredensial administrator, akun pengawas ruangan, keamanan anti-curang siswa, galeri portal, serta backup database.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <!-- 1. Setelan Identitas Sekolah & Kredensial Admin -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="school" class="w-4 h-4 text-indigo-600"></i> Identitas Sekolah & Akun Administrator
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah / Lembaga</label>
                                        <input type="text" id="setSchoolName" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Subjudul / Alamat & Tahun Pelajaran</label>
                                        <input type="text" id="setSchoolSubtitle" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div class="pt-2 border-t border-slate-100 grid grid-cols-2 gap-2.5">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Username Admin</label>
                                            <input type="text" id="setAdminUsername" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Password Admin</label>
                                            <input type="text" id="setAdminPassword" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                    </div>
                                    <button onclick="saveSchoolIdentity()" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition">
                                        Simpan Identitas & Akun Admin
                                    </button>
                                </div>
                            </div>

                            <!-- 2. Setelan Keamanan & Perilaku Siswa -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="shield-alert" class="w-4 h-4 text-red-600"></i> Setelan Keamanan & Perilaku Siswa
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-red-700 mb-1">Kode PIN Pembuka Kunci Layar (Default: 123)</label>
                                        <input type="text" id="setUnlockPin" class="w-full px-3.5 py-2 rounded-xl border border-red-300 bg-red-50/40 font-mono font-extrabold text-red-700 text-sm">
                                    </div>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700 pt-1">
                                        <span>Wajib Mode Layar Penuh (Fullscreen)</span>
                                        <input type="checkbox" id="setRequireFullscreen" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Kunci Otomatis Jika Pindah Tab / Aplikasi</span>
                                        <input type="checkbox" id="setDetectTabSwitch" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Tampilkan Skor Akhir ke Siswa Setelah Selesai</span>
                                        <input type="checkbox" id="setShowScoreToStudent" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <button onclick="saveSecuritySettings()" class="px-5 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                        Simpan Setelan Keamanan
                                    </button>
                                </div>
                            </div>

                            <!-- 3. Manajemen Akun Pengawas -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="user-cog" class="w-4 h-4 text-amber-600"></i> Manajemen Akun Pengawas Ruangan
                                </h3>
                                <form onsubmit="handleAddProctor(event)" class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="procName" required placeholder="Nama Pengawas" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procRoom" required placeholder="Ruang Tugas (Misal: Ruang 01 / Semua)" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procUser" required placeholder="Username Login" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                    <div class="flex gap-2">
                                        <input type="text" id="procPass" required placeholder="Password" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                        <button type="submit" class="px-3 py-2 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold shrink-0">
                                            + Tambah
                                        </button>
                                    </div>
                                </form>
                                <div id="proctorListContainer" class="space-y-2 max-h-52 overflow-y-auto">
                                    <!-- Daftar pengawas -->
                                </div>
                            </div>

                            <!-- 4. Kelola Foto Kegiatan Portal & Backup Database -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="image" class="w-4 h-4 text-emerald-600"></i> Galeri Foto Portal & Manajemen Database
                                </h3>
                                <form onsubmit="handleAddGalleryItem(event)" class="space-y-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="galTitle" required placeholder="Judul Foto Kegiatan (Misal: Pelaksanaan Ujian Sesi 1)" class="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <div class="flex gap-2">
                                        <input type="url" id="galUrl" placeholder="URL Gambar Eksternal (Kosongkan untuk Ilustrasi SVG)" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                        <button type="submit" class="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shrink-0">
                                            + Tambah Foto
                                        </button>
                                    </div>
                                </form>

                                <div class="pt-3 border-t border-slate-100 space-y-2">
                                    <span class="text-xs font-bold text-slate-700 block">Cadangan, Pemulihan & Bersihkan Data Awal</span>
                                    <div class="flex flex-wrap gap-2">
                                        <button onclick="exportDatabaseJSON()" class="px-3.5 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="download" class="w-3.5 h-3.5"></i> Backup JSON
                                        </button>
                                        <label class="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer">
                                            <i data-lucide="upload" class="w-3.5 h-3.5"></i> Restore JSON
                                            <input type="file" accept=".json" onchange="importDatabaseJSON(event)" class="hidden">
                                        </label>
                                        <button onclick="confirmCleanStartEmpty()" class="px-3.5 py-2 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eraser" class="w-3.5 h-3.5"></i> Bersihkan Data Awal (Mulai Kosong)
                                        </button>
                                        <button onclick="confirmFactoryReset()" class="px-3.5 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> Reset Bawaan
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW C: DASHBOARD PENGAWAS RUANG                                  -->
        <!-- ================================================================= -->
        <section id="view-pengawas" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
            <!-- Header Pengawas -->
            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <div class="flex items-center gap-2">
                        <span class="px-2.5 py-0.5 rounded-md bg-amber-100 text-amber-800 text-xs font-bold">PORTAL PENGAWAS</span>
                        <span id="proctorAssignedRoomBadge" class="px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-700 text-xs font-semibold">Semua Ruangan</span>
                    </div>
                    <h2 id="proctorDashboardName" class="text-2xl font-extrabold text-slate-900 mt-1">Pengawas Ujian</h2>
                    <p class="text-xs text-slate-500 mt-0.5">Bagikan token ruang ke siswa, pantau pengerjaan secara langsung, dan buka kunci jika ada pelanggaran.</p>
                </div>

                <!-- Kartu Token Ruang Aktif -->
                <div class="bg-slate-900 text-white px-5 py-3.5 rounded-2xl flex items-center gap-4 self-start md:self-center">
                    <div>
                        <span class="text-[10px] uppercase text-slate-400 font-bold block">Token Masuk Siswa</span>
                        <span id="proctorTokenDisplay" class="text-2xl font-mono font-extrabold text-amber-400 tracking-widest">------</span>
                    </div>
                    <button onclick="copyTextToClipboard(document.getElementById('proctorTokenDisplay').innerText, 'Token berhasil disalin!')" class="px-3 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-bold transition flex items-center gap-1.5">
                        <i data-lucide="copy" class="w-4 h-4"></i> Salin
                    </button>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                <!-- Monitoring Siswa Real-Time -->
                <div class="lg:col-span-8 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-bold text-slate-900 text-base">Monitoring Peserta Ujian Real-Time</h3>
                            <p class="text-xs text-slate-500">Kode PIN Buka Kunci Layar: <strong id="proctorPinDisplay" class="font-mono text-red-600">123</strong></p>
                        </div>
                        <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                            <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Live Monitoring
                        </span>
                    </div>

                    <div id="proctorStudentCards" class="space-y-3">
                        <!-- Diisi via JS -->
                    </div>
                </div>

                <!-- Input Soal Tambahan oleh Pengawas -->
                <div class="lg:col-span-4 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4 self-start">
                    <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                        <i data-lucide="file-plus" class="w-4 h-4 text-amber-600"></i> Input Soal Tambahan
                    </h3>
                    <p class="text-xs text-slate-500 leading-relaxed">
                        Pengawas dapat menambahkan soal dari file Word (.docx) / PDF (.pdf) atau file teks secara cepat.
                    </p>
                    <label class="flex flex-col items-center justify-center w-full p-4 border-2 border-dashed border-amber-300 rounded-2xl bg-amber-50/40 hover:bg-amber-50 cursor-pointer transition text-center">
                        <i data-lucide="upload-cloud" class="w-6 h-6 text-amber-600 mb-1"></i>
                        <span class="text-xs font-bold text-amber-900">Upload Word (.DOCX) / PDF (.PDF)</span>
                        <input type="file" accept=".docx,.pdf,.txt" onchange="handleDocumentQuestionUpload(event)" class="hidden">
                    </label>
                    <div class="pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                        <span class="text-slate-600">Total Bank Soal Sistem:</span>
                        <span id="proctorTotalSoalBadge" class="font-extrabold text-slate-900">0 Soal</span>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW D: HALAMAN UJIAN SISWA (ANTI-CHEAT LOCKDOWN MODE)            -->
        <!-- ================================================================= -->
        <section id="view-siswa" class="hidden min-h-screen bg-slate-100 flex flex-col no-select">
            <!-- Header Ujian Siswa -->
            <div class="bg-slate-900 text-white px-4 sm:px-6 py-3.5 shadow-md sticky top-0 z-30">
                <div class="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-extrabold shrink-0">
                            <i data-lucide="shield-check" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h2 id="studentExamTitle" class="font-bold text-sm sm:text-base leading-tight">Paket Ujian</h2>
                            <p class="text-xs text-slate-300">
                                <strong id="studentExamName" class="text-white">-</strong> • NISN: <span id="studentExamNisn">-</span> • <span id="studentExamRoom">-</span>
                            </p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2 sm:gap-3">
                        <!-- Kontrol Ukuran Font (A-, A, A+) -->
                        <div class="hidden sm:flex items-center bg-slate-800 border border-slate-700 rounded-xl p-1">
                            <button onclick="changeExamFontSize('sm')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perkecil Teks">A-</button>
                            <button onclick="changeExamFontSize('base')" class="px-2 py-1 rounded-lg text-xs font-bold bg-slate-700" title="Ukuran Normal">A</button>
                            <button onclick="changeExamFontSize('lg')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perbesar Teks">A+</button>
                        </div>

                        <!-- Timer Hitung Mundur -->
                        <div class="bg-slate-800 border border-slate-700 px-3.5 py-1.5 rounded-xl flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-amber-400"></i>
                            <span id="studentTimerDisplay" class="font-mono font-extrabold text-base sm:text-lg text-amber-400">00:00</span>
                        </div>

                        <button onclick="confirmSubmitExam()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-bold transition shadow">
                            Selesai Ujian
                        </button>
                    </div>
                </div>
            </div>

            <!-- Area Soal & Navigasi -->
            <div class="max-w-7xl w-full mx-auto px-4 sm:px-6 py-6 grid grid-cols-1 lg:grid-cols-12 gap-6 flex-grow">
                <!-- Kartu Pertanyaan -->
                <div class="lg:col-span-8 bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm flex flex-col justify-between">
                    <div>
                        <div class="flex flex-wrap items-center justify-between gap-2 pb-4 mb-6 border-b border-slate-100">
                            <span class="px-3.5 py-1.5 rounded-xl bg-blue-50 text-blue-700 font-extrabold text-xs sm:text-sm">
                                SOAL NO. <span id="examCurrentNum">1</span> DARI <span id="examTotalNum">0</span>
                            </span>
                            <button id="doubtToggleBtn" onclick="toggleDoubtQuestion()" class="px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5">
                                <i data-lucide="flag" class="w-3.5 h-3.5"></i> <span>Tandai Ragu-Ragu</span>
                            </button>
                        </div>

                        <div id="examQuestionBody" class="text-base font-semibold text-slate-900 leading-relaxed mb-6">
                            Memuat soal ujian...
                        </div>

                        <div id="examOptionsList" class="space-y-3">
                            <!-- Opsi A-E dirender via JS -->
                        </div>
                    </div>

                    <!-- Navigasi Sebelumnya / Selanjutnya -->
                    <div class="flex items-center justify-between pt-8 mt-8 border-t border-slate-100">
                        <button id="btnPrevQuestion" onclick="goPrevQuestion()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs sm:text-sm font-bold transition flex items-center gap-2 disabled:opacity-40">
                            <i data-lucide="chevron-left" class="w-4 h-4"></i> Sebelumnya
                        </button>
                        <button id="btnNextQuestion" onclick="goNextQuestion()" class="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2">
                            Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>
                        </button>
                    </div>
                </div>

                <!-- Sidebar Navigasi Nomor -->
                <div class="lg:col-span-4 space-y-4">
                    <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Navigasi Nomor Soal</h3>
                        <div id="examNumberGrid" class="grid grid-cols-5 gap-2">
                            <!-- Tombol nomor soal -->
                        </div>
                        <div class="mt-5 pt-4 border-t border-slate-100 grid grid-cols-3 gap-2 text-[11px] text-slate-600">
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-blue-600 inline-block"></span> Dijawab</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-amber-400 inline-block"></span> Ragu</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-slate-100 border border-slate-300 inline-block"></span> Kosong</span>
                        </div>
                    </div>

                    <div class="bg-red-50 border border-red-200 rounded-2xl p-4 text-xs text-red-900 space-y-1">
                        <div class="font-bold flex items-center gap-1.5 text-red-700">
                            <i data-lucide="shield-alert" class="w-4 h-4"></i> Sistem Proteksi Aktif
                        </div>
                        <p class="leading-relaxed">
                            Dilarang berpindah tab, membuka aplikasi lain, mengecilkan layar, atau menekan tombol pengembang. Pelanggaran akan <strong>mengunci layar secara otomatis</strong>.
                        </p>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- ===================================================================== -->
    <!-- 5. MODAL LOGIN MULTI-PERAN (SISWA + TOKEN, PENGAWAS, ADMIN)           -->
    <!-- ===================================================================== -->
    <div id="loginModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-slate-100 relative">
            <button onclick="closeLoginModal()" class="absolute top-5 right-5 w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>

            <div class="mb-6">
                <span id="loginModalBadge" class="px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase">Portal Login</span>
                <h2 id="loginModalTitle" class="text-2xl font-extrabold text-slate-900 mt-2">Masuk Sistem</h2>
                <p id="loginModalDesc" class="text-xs text-slate-500 mt-1">Masukkan kredensial akun resmi Anda.</p>
            </div>

            <form onsubmit="processLogin(event)" class="space-y-4">
                <input type="hidden" id="loginRoleField" value="siswa">

                <div>
                    <label id="loginUsernameLabel" class="block text-xs font-bold text-slate-700 mb-1">NISN / Username</label>
                    <input type="text" id="loginUsernameField" required placeholder="Masukkan NISN atau Username..." 
                        class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Password</label>
                    <input type="password" id="loginPasswordField" required placeholder="Masukkan password..." 
                        class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <!-- Input Token Wajib Khusus Siswa -->
                <div id="loginTokenWrapper" class="space-y-1 pt-1">
                    <label class="block text-xs font-bold text-blue-700">
                        Token Ujian Aktif <span class="text-red-500">*</span>
                    </label>
                    <input type="text" id="loginTokenField" maxlength="10" placeholder="Masukkan Token dari Pengawas" 
                        class="w-full px-4 py-3 rounded-xl border-2 border-blue-300 bg-blue-50/40 font-mono font-extrabold uppercase tracking-widest text-sm text-blue-950 focus:ring-2 focus:ring-blue-600 focus:outline-none">
                    <p class="text-[11px] text-slate-500">
                        * Wajib diisi setiap kali masuk sesi ujian. Minta token kepada pengawas ruang.
                    </p>
                </div>

                <button type="submit" id="loginSubmitButton" class="w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2">
                    Masuk Sekarang
                </button>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 6. MODAL EDIT AKUN & RESET PASSWORD SISWA (MENU 5 ADMINISTRATOR)      -->
    <!-- ===================================================================== -->
    <div id="editStudentModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Akun & Password Siswa</h3>
                <button onclick="closeEditStudentModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedStudent(event)" class="space-y-3">
                <input type="hidden" id="editStuId">
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap *</label>
                    <input type="text" id="editStuName" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">NISN *</label>
                    <input type="text" id="editStuNisn" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm font-mono">
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kelas *</label>
                        <input type="text" id="editStuClass" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Ruangan *</label>
                        <input type="text" id="editStuRoom" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3 pt-1">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Username Login *</label>
                        <input type="text" id="editStuUser" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Password *</label>
                        <div class="flex gap-1.5">
                            <input type="text" id="editStuPass" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold text-blue-700">
                            <button type="button" onclick="generateRandomStudentPassInModal()" class="px-2.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold" title="Acak Password Baru">
                                <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
                            </button>
                        </div>
                    </div>
                </div>
                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditStudentModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 7. MODAL EDIT SOAL (MENU 2 ADMINISTRATOR)                             -->
    <!-- ===================================================================== -->
    <div id="editQuestionModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Soal Ujian</h3>
                <button onclick="closeEditQuestionModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedQuestion(event)" class="space-y-3">
                <input type="hidden" id="editQId">
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Pertanyaan</label>
                    <textarea id="editQText" required rows="3" class="w-full p-2.5 rounded-xl border border-slate-300 text-sm"></textarea>
                </div>
                <div class="space-y-2">
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">A.</span>
                        <input type="text" id="editQOptA" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">B.</span>
                        <input type="text" id="editQOptB" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">C.</span>
                        <input type="text" id="editQOptC" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">D.</span>
                        <input type="text" id="editQOptD" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">E.</span>
                        <input type="text" id="editQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3 pt-1">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kunci Jawaban</label>
                        <select id="editQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                            <option value="A">Opsi A</option>
                            <option value="B">Opsi B</option>
                            <option value="C">Opsi C</option>
                            <option value="D">Opsi D</option>
                            <option value="E">Opsi E</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Bobot Poin</label>
                        <input type="number" id="editQPoints" min="1" max="100" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                    </div>
                </div>
                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditQuestionModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 8. MODAL PANDUAN INSTALASI PERANGKAT (ANDROID / IOS / DESKTOP)        -->
    <!-- ===================================================================== -->
    <div id="installPlatformModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-100">
            <div class="flex items-center justify-between mb-3">
                <h3 id="installPlatformTitle" class="text-lg font-extrabold text-slate-900">Panduan Instalasi</h3>
                <button onclick="closeInstallModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <div id="installPlatformContent" class="text-xs sm:text-sm text-slate-600 space-y-3 leading-relaxed">
                <!-- Diisi via JS -->
            </div>
            <div class="mt-6 flex justify-between items-center gap-2">
                <button onclick="downloadStandaloneApp()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold">
                    Unduh File .HTML
                </button>
                <button onclick="closeInstallModal()" class="px-5 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-bold">
                    Tutup
                </button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 9. MODAL HASIL SELESAI UJIAN SISWA                                    -->
    <!-- ===================================================================== -->
    <div id="examFinishedModal" class="fixed inset-0 bg-slate-900/75 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl text-center space-y-4">
            <div class="w-16 h-16 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center mx-auto">
                <i data-lucide="check-circle-2" class="w-9 h-9"></i>
            </div>
            <h2 class="text-2xl font-extrabold text-slate-900">Ujian Telah Selesai!</h2>
            <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
                Terima kasih telah mengerjakan ujian dengan jujur. Seluruh jawaban Anda telah dikirim secara real-time ke server Administrator.
            </p>
            <div id="studentScoreSummaryBox" class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-1">
                <span class="text-xs text-slate-500 block">Nilai Akhir Anda</span>
                <span id="studentFinalScoreText" class="text-3xl font-extrabold text-blue-600">100</span>
                <span id="studentFinalDetailText" class="text-xs text-slate-600 block">Benar 5 dari 5 soal</span>
            </div>
            <button onclick="closeExamFinishedModal()" class="w-full py-3 px-6 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm transition">
                Kembali ke Halaman Portal
            </button>
        </div>
    </div>

    <!-- Footer -->
    <footer id="mainFooter" class="bg-white border-t border-slate-200 py-6 mt-auto no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
            <div>
                <strong id="footerSchoolName" class="text-slate-800">CBT-PRO | Sistem Ujian Berbasis Komputer</strong> &copy; 2026 • Siap Hosting GitHub Pages.
            </div>
            <div class="flex items-center gap-4">
                <span>Anti-Cheat Lockdown PIN</span>
                <span>•</span>
                <span>Real-Time Sync Database</span>
            </div>
        </div>
    </footer>

    <!-- ===================================================================== -->
    <!-- 10. JAVASCRIPT ENGINE LENGKAP                                         -->
    <!-- ===================================================================== -->
    <script>
        // =================================================================
        // A. KONSTANTA DATABASE LOKAL (LOCALSTORAGE) & DATA AWAL PRODUKSI
        // =================================================================
        const STORAGE_KEYS = {
            SETTINGS: 'cbt_prod_settings',
            QUESTIONS: 'cbt_prod_questions',
            STUDENTS: 'cbt_prod_students',
            PROCTORS: 'cbt_prod_proctors',
            LIVE_STATUS: 'cbt_prod_live_status',
            GALLERY: 'cbt_prod_gallery'
        };

        const DEFAULT_SETTINGS = {
            schoolName: 'CBT-PRO | Sistem Ujian Berbasis Komputer',
            schoolSubtitle: 'Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027',
            adminUsername: 'admin',
            adminPassword: 'admin123',
            examTitle: 'Ujian Akhir Semester (UAS) Berbasis Komputer',
            subjectName: 'Kompetensi Akademik Umum',
            targetGroup: 'ALL',
            durationMinutes: 45,
            activeToken: 'CBT26X',
            shuffleQuestions: true,
            shuffleOptions: true,
            unlockPin: '123',
            requireFullscreen: true,
            detectTabSwitch: true,
            showScoreToStudent: true
        };

        const DEFAULT_QUESTIONS = [
            {
                id: 'q_init_1',
                question: 'Komponen perangkat keras komputer yang berfungsi sebagai unit pemroses utama seluruh instruksi logika dan aritmatika adalah...',
                options: { A: 'RAM (Random Access Memory)', B: 'CPU (Central Processing Unit)', C: 'Solid State Drive (SSD)', D: 'Power Supply Unit', E: 'Network Interface Card' },
                answer: 'B',
                points: 20
            },
            {
                id: 'q_init_2',
                question: 'Protokol keamanan standar pada web yang mengenkripsi pertukaran data antara browser dan server ditandai dengan awalan...',
                options: { A: 'HTTP://', B: 'FTP://', C: 'HTTPS://', D: 'SMTP://', E: 'TELNET://' },
                answer: 'C',
                points: 20
            },
            {
                id: 'q_init_3',
                question: 'Jika sebuah segitiga siku-siku memiliki panjang alas 6 cm dan tinggi 8 cm, maka panjang sisi miringnya adalah...',
                options: { A: '10 cm', B: '12 cm', C: '14 cm', D: '15 cm', E: '16 cm' },
                answer: 'A',
                points: 20
            },
            {
                id: 'q_init_4',
                question: 'Teks Proklamasi Kemerdekaan Republik Indonesia dibacakan oleh Ir. Soekarno pada tanggal...',
                options: { A: '17 Agustus 1944', B: '17 Agustus 1945', C: '18 Agustus 1945', D: '28 Oktober 1928', E: '1 Juni 1945' },
                answer: 'B',
                points: 20
            },
            {
                id: 'q_init_5',
                question: 'Sikap utama yang harus dijunjung tinggi oleh setiap peserta didik dalam mengikuti ujian berbasis komputer adalah...',
                options: { A: 'Mencari celah membuka tab lain', B: 'Kejujuran dan kemandirian', C: 'Menyalin jawaban teman', D: 'Mengabaikan tata tertib ruang', E: 'Membagikan token ke luar ruang' },
                answer: 'B',
                points: 20
            }
        ];

        // Akun Siswa Siap Digunakan berbasis NISN resmi (tanpa label demo)
        const DEFAULT_STUDENTS = [
            { id: 'stu_001', name: 'Ahmad Fauzi', nisn: '0081234501', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234501', password: '84921' },
            { id: 'stu_002', name: 'Siti Nurhaliza', nisn: '0081234502', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234502', password: '37194' },
            { id: 'stu_003', name: 'Budi Santoso', nisn: '0081234503', className: 'XII IPS 2', room: 'Ruang 02', username: '0081234503', password: '62508' }
        ];

        const DEFAULT_PROCTORS = [
            { id: 'proc_1', name: 'Drs. Hendra Wijaya', room: 'Semua Ruangan', username: 'pengawas', password: 'pengawas123' },
            { id: 'proc_2', name: 'Rina Marlina, S.Pd', room: 'Ruang 01', username: 'pengawas_r1', password: '74125' }
        ];

        const DEFAULT_GALLERY = [
            { id: 'gal_1', title: 'Pelaksanaan Ujian Berbasis Komputer di Laboratorium 1', svgType: 'lab', url: '' },
            { id: 'gal_2', title: 'Pengawasan Ruang Ujian & Verifikasi Token Peserta', svgType: 'proctor', url: '' },
            { id: 'gal_3', title: 'Fokus & Kejujuran Akademik Siswa Saat Mengerjakan Soal', svgType: 'student', url: '' }
        ];

        const MOTIVATION_QUOTES = [
            {
                text: '"Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."',
                author: '— Pesan Akademik Sekolah'
            },
            {
                text: '"Jangan takut pada soal yang sulit; percayalah pada proses belajar yang sudah kamu lalui setiap hari. Kerjakan dengan tenang dan teliti!"',
                author: '— Motivasi Belajar CBT-PRO'
            },
            {
                text: '"Pendidikan adalah paspor menuju masa depan, karena hari esok dimiliki oleh mereka yang mempersiapkannya hari ini."',
                author: '— Malcolm X'
            },
            {
                text: '"Ujian bukan sekadar mengukur seberapa banyak yang kamu hafal, tetapi melatih integritas, fokus, dan ketangguhan dirimu."',
                author: '— Dewan Guru & Pengawas'
            }
        ];

        let currentQuoteIdx = 0;
        let showStudentPasswords = false;

        // Getter & Setter LocalStorage
        const getDB = (key, fallback) => {
            try {
                const raw = localStorage.getItem(key);
                return raw ? JSON.parse(raw) : fallback;
            } catch (e) {
                return fallback;
            }
        };
        const setDB = (key, value) => {
            localStorage.setItem(key, JSON.stringify(value));
        };

        function initializeAppDatabase() {
            if (!localStorage.getItem(STORAGE_KEYS.SETTINGS)) setDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (!localStorage.getItem(STORAGE_KEYS.QUESTIONS)) setDB(STORAGE_KEYS.QUESTIONS, DEFAULT_QUESTIONS);
            if (!localStorage.getItem(STORAGE_KEYS.STUDENTS)) setDB(STORAGE_KEYS.STUDENTS, DEFAULT_STUDENTS);
            if (!localStorage.getItem(STORAGE_KEYS.PROCTORS)) setDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
            if (!localStorage.getItem(STORAGE_KEYS.LIVE_STATUS)) setDB(STORAGE_KEYS.LIVE_STATUS, {});
            if (!localStorage.getItem(STORAGE_KEYS.GALLERY)) setDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
        }

        // =================================================================
        // B. STATE SESI PENGGUNA & MESIN UJIAN
        // =================================================================
        let activeSession = null; // { role: 'admin'|'pengawas'|'siswa', user: {...} }
        let currentAdminMenu = 'monitoring';

        let examSession = {
            isRunning: false,
            isLocked: false,
            questions: [],
            answers: {},
            doubts: {},
            currentIdx: 0,
            remainingSeconds: 0,
            timerId: null,
            violations: 0
        };

        // =================================================================
        // C. NAVIGASI SPA & RENDER UI GLOBAL
        // =================================================================
        function navigateTo(view) {
            ['portal', 'admin', 'pengawas', 'siswa'].forEach(v => {
                document.getElementById(`view-${v}`).classList.add('hidden');
            });
            document.getElementById(`view-${view}`).classList.remove('hidden');

            const header = document.getElementById('mainHeader');
            const footer = document.getElementById('mainFooter');
            const guestBtns = document.getElementById('headerGuestButtons');
            const authControls = document.getElementById('headerAuthControls');

            if (view === 'siswa') {
                header.classList.add('hidden');
                footer.classList.add('hidden');
            } else {
                header.classList.remove('hidden');
                footer.classList.remove('hidden');
                if (activeSession && (view === 'admin' || view === 'pengawas')) {
                    guestBtns.classList.add('hidden');
                    authControls.classList.remove('hidden');
                    authControls.classList.add('flex');
                    document.getElementById('headerUserBadge').innerText =
                        activeSession.role === 'admin' ? 'Administrator' : `Pengawas: ${activeSession.user.name}`;
                } else {
                    guestBtns.classList.remove('hidden');
                    authControls.classList.add('hidden');
                    authControls.classList.remove('flex');
                }
            }

            syncAllViews();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function syncAllViews() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
            const students = getDB(STORAGE_KEYS.STUDENTS, []);

            // Header & Footer Identitas Sekolah
            document.getElementById('headerSchoolTitle').innerText = settings.schoolName;
            document.getElementById('headerSchoolSubtitle').innerText = settings.schoolSubtitle;
            document.getElementById('footerSchoolName').innerText = settings.schoolName;

            // Portal Stats
            document.getElementById('heroExamTitle').innerText = settings.examTitle;
            document.getElementById('portalStatSoal').innerText = questions.length;
            document.getElementById('portalStatSiswa').innerText = students.length;
            document.getElementById('portalStatDurasi').innerText = `${settings.durationMinutes}m`;
            renderPortalGallery();

            // Admin Views
            if (activeSession?.role === 'admin') {
                document.getElementById('sidebarTokenDisplay').innerText = settings.activeToken;
                document.getElementById('assignTokenDisplay').innerText = settings.activeToken;
                populateFilterDropdowns();
                renderAdminMonitoring();
                renderAdminQuestionList();
                renderAdminNilaiTable();
                renderAdminStudentAccountTable();
                renderAdminLoginCards();
                renderAdminSettingsForms();
            }

            // Pengawas Views
            if (activeSession?.role === 'pengawas') {
                document.getElementById('proctorDashboardName').innerText = activeSession.user.name;
                document.getElementById('proctorAssignedRoomBadge').innerText = `Tugas: ${activeSession.user.room}`;
                document.getElementById('proctorTokenDisplay').innerText = settings.activeToken;
                document.getElementById('proctorPinDisplay').innerText = settings.unlockPin;
                document.getElementById('proctorTotalSoalBadge').innerText = `${questions.length} Soal`;
                renderProctorMonitoring();
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        // =================================================================
        // D. TOAST, KONFIRMASI & CLIPBOARD (TANPA ALERT/CONFIRM)
        // =================================================================
        function showToast(message, type = 'info') {
            const container = document.getElementById('toastContainer');
            const el = document.createElement('div');
            const palette = {
                success: 'bg-emerald-600 text-white border-emerald-500',
                error: 'bg-red-600 text-white border-red-500',
                warning: 'bg-amber-400 text-slate-950 border-amber-300',
                info: 'bg-slate-900 text-white border-slate-700'
            };
            el.className = `pointer-events-auto flex items-center gap-2.5 px-4 py-3 rounded-2xl shadow-xl border text-xs sm:text-sm font-semibold transition-all duration-300 ${palette[type] || palette.info}`;
            el.innerHTML = `<span class="flex-1">${message}</span>`;
            container.appendChild(el);
            setTimeout(() => {
                el.classList.add('opacity-0', '-translate-y-2');
                setTimeout(() => el.remove(), 300);
            }, 3500);
        }

        function showConfirmDialog(title, message, onYes) {
            const modal = document.getElementById('confirmModal');
            document.getElementById('confirmTitle').innerText = title;
            document.getElementById('confirmMessage').innerText = message;
            modal.classList.remove('hidden');
            modal.classList.add('flex');

            const okBtn = document.getElementById('confirmOkBtn');
            const cancelBtn = document.getElementById('confirmCancelBtn');

            const close = () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                okBtn.onclick = null;
                cancelBtn.onclick = null;
            };
            cancelBtn.onclick = close;
            okBtn.onclick = () => {
                close();
                onYes();
            };
        }

        function copyTextToClipboard(text, successMessage = 'Berhasil disalin!') {
            const ta = document.createElement('textarea');
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            try {
                document.execCommand('copy');
                showToast(successMessage, 'success');
            } catch (e) {
                showToast('Gagal menyalin teks.', 'error');
            }
            document.body.removeChild(ta);
        }

        function copyPortalLink(role) {
            const base = window.location.href.split('?')[0].split('#')[0];
            const fullLink = `${base}?role=${role}`;
            const label = role === 'admin' ? 'Link Administrator' : role === 'pengawas' ? 'Link Portal Pengawas' : 'Link Portal Siswa';
            copyTextToClipboard(fullLink, `${label} disalin! Siap dibagikan.`);
        }

        // =================================================================
        // E. PORTAL UTAMA: MOTIVASI, GALERI SVG & PANDUAN INSTALASI
        // =================================================================
        function nextMotivationQuote() {
            currentQuoteIdx = (currentQuoteIdx + 1) % MOTIVATION_QUOTES.length;
            const q = MOTIVATION_QUOTES[currentQuoteIdx];
            document.getElementById('motivationQuoteText').innerText = q.text;
            document.getElementById('motivationQuoteAuthor').innerText = q.author;
        }

        function getGalleryIllustrationSVG(type) {
            if (type === 'proctor') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-amber-50 to-orange-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#fde68a" stroke-width="2"/>
                    <circle cx="110" cy="95" r="28" fill="#f59e0b"/>
                    <path d="M65 165 C65 135, 155 135, 155 165" fill="#d97706"/>
                    <rect x="180" y="65" width="160" height="90" rx="10" fill="#1e293b"/>
                    <rect x="192" y="78" width="80" height="10" rx="4" fill="#fbbf24"/>
                    <rect x="192" y="96" width="130" height="8" rx="4" fill="#64748b"/>
                    <rect x="192" y="112" width="110" height="8" rx="4" fill="#38bdf8"/>
                    <circle cx="315" cy="132" r="12" fill="#10b981"/>
                </svg>`;
            }
            if (type === 'student') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-emerald-50 to-teal-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#a7f3d0" stroke-width="2"/>
                    <rect x="85" y="65" width="230" height="105" rx="12" fill="#0f172a"/>
                    <rect x="98" y="78" width="204" height="75" rx="6" fill="#eff6ff"/>
                    <rect x="112" y="90" width="140" height="10" rx="4" fill="#2563eb"/>
                    <rect x="112" y="108" width="170" height="8" rx="4" fill="#94a3b8"/>
                    <rect x="112" y="124" width="120" height="8" rx="4" fill="#10b981"/>
                    <rect x="65" y="170" width="270" height="10" rx="5" fill="#64748b"/>
                </svg>`;
            }
            return `
            <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-blue-50 to-indigo-100">
                <rect x="25" y="25" width="350" height="170" rx="16" fill="#ffffff" stroke="#bfdbfe" stroke-width="2"/>
                <rect x="55" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="63" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="95" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="75" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <rect x="225" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="233" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="265" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="245" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <circle cx="200" cy="45" r="10" fill="#10b981"/>
            </svg>`;
        }

        function renderPortalGallery() {
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            const grid = document.getElementById('portalGalleryGrid');
            if (!grid) return;

            grid.innerHTML = gallery.map(item => {
                const mediaHtml = item.url
                    ? `<img src="${item.url}" alt="${item.title}" onerror="this.outerHTML=\`${getGalleryIllustrationSVG(item.svgType || 'lab').replace(/"/g, '&quot;')}\`" class="w-full h-44 object-cover">`
                    : getGalleryIllustrationSVG(item.svgType || 'lab');

                return `
                    <div class="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition flex flex-col">
                        ${mediaHtml}
                        <div class="p-4 flex-1 flex flex-col justify-between">
                            <p class="text-xs sm:text-sm font-bold text-slate-800 leading-snug">${item.title}</p>
                            <span class="text-[11px] text-slate-400 mt-2 inline-flex items-center gap-1">
                                <i data-lucide="camera" class="w-3 h-3"></i> Dokumentasi CBT Sekolah
                            </span>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function openInstallModal(platform) {
            const modal = document.getElementById('installPlatformModal');
            const title = document.getElementById('installPlatformTitle');
            const content = document.getElementById('installPlatformContent');

            if (platform === 'android') {
                title.innerText = 'Instalasi Aplikasi di HP / Tablet Android';
                content.innerHTML = `
                    <p>Pasang <strong>CBT-PRO</strong> di perangkat Android agar berjalan dalam mode layar penuh tanpa bilah navigasi browser:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka link portal ujian ini melalui <strong>Google Chrome</strong> di HP Android Anda.</li>
                        <li>Ketuk ikon menu <strong>Titik Tiga (⋮)</strong> di pojok kanan atas layar.</li>
                        <li>Pilih <strong>"Tambahkan ke Layar Utama"</strong> atau <strong>"Install Aplikasi"</strong>.</li>
                        <li>Konfirmasi dengan menekan <strong>Install / Tambah</strong>. Ikon aplikasi CBT akan muncul di layar utama HP Anda.</li>
                    </ol>
                `;
            } else if (platform === 'ios') {
                title.innerText = 'Instalasi Aplikasi di iPhone / iPad (iOS)';
                content.innerHTML = `
                    <p>Untuk memasang aplikasi ujian pada perangkat iPhone atau iPad:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka halaman ujian ini menggunakan browser bawaan <strong>Safari</strong>.</li>
                        <li>Ketuk tombol <strong>Bagikan (Share)</strong> (ikon kotak dengan panah ke atas) di bagian bawah Safari.</li>
                        <li>Gulir menu ke bawah, lalu ketuk <strong>"Tambah ke Layar Utama" (Add to Home Screen)</strong>.</li>
                        <li>Ketuk <strong>Tambah</strong> di pojok kanan atas. Aplikasi siap dibuka langsung dari Home Screen iPhone/iPad.</li>
                    </ol>
                `;
            } else {
                title.innerText = 'Instalasi Aplikasi di Komputer (Windows / Mac)';
                content.innerHTML = `
                    <p>Untuk menjalankan aplikasi ujian secara mandiri di komputer laboratorium atau laptop:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka link aplikasi ini menggunakan <strong>Google Chrome</strong> atau <strong>Microsoft Edge</strong>.</li>
                        <li>Klik ikon <strong>Install App</strong> di ujung kanan bilah alamat URL (atau klik menu Titik Tiga &rarr; <em>Simpan dan Bagikan</em> &rarr; <em>Install halaman sebagai aplikasi</em>).</li>
                        <li>Atau klik tombol <strong>Unduh File .HTML</strong> di bawah untuk menjalankan aplikasi secara lokal tanpa internet.</li>
                    </ol>
                `;
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeInstallModal() {
            const modal = document.getElementById('installPlatformModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function downloadStandaloneApp() {
            const htmlDoc = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
            const blob = new Blob([htmlDoc], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'CBT-PRO-Aplikasi-Ujian.html';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('File aplikasi mandiri (.HTML) berhasil diunduh!', 'success');
        }

        // =================================================================
        // F. AUTENTIKASI 3 PERAN (SISWA + TOKEN, PENGAWAS, ADMIN)
        // =================================================================
        function openLoginModal(role) {
            const modal = document.getElementById('loginModal');
            const roleField = document.getElementById('loginRoleField');
            const badge = document.getElementById('loginModalBadge');
            const title = document.getElementById('loginModalTitle');
            const desc = document.getElementById('loginModalDesc');
            const userLabel = document.getElementById('loginUsernameLabel');
            const tokenWrap = document.getElementById('loginTokenWrapper');
            const submitBtn = document.getElementById('loginSubmitButton');

            document.getElementById('loginUsernameField').value = '';
            document.getElementById('loginPasswordField').value = '';
            document.getElementById('loginTokenField').value = '';
            roleField.value = role;

            if (role === 'admin') {
                badge.className = 'px-3 py-1 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold uppercase';
                badge.innerText = 'Administrator';
                title.innerText = 'Login Administrator';
                desc.innerText = 'Masukkan Username dan Password Administrator sekolah.';
                userLabel.innerText = 'Username Administrator';
                tokenWrap.classList.add('hidden');
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm transition shadow-lg mt-2';
            } else if (role === 'pengawas') {
                badge.className = 'px-3 py-1 rounded-full bg-amber-100 text-amber-800 text-xs font-bold uppercase';
                badge.innerText = 'Pengawas Ruang';
                title.innerText = 'Login Pengawas';
                desc.innerText = 'Masukkan Username dan Password Pengawas Ruang yang terdaftar.';
                userLabel.innerText = 'Username Pengawas';
                tokenWrap.classList.add('hidden');
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-bold text-sm transition shadow-lg mt-2';
            } else {
                badge.className = 'px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase';
                badge.innerText = 'Peserta Ujian';
                title.innerText = 'Login Ujian Siswa';
                desc.innerText = 'Masukkan NISN/Username, Password Kartu Login, dan Token Aktif.';
                userLabel.innerText = 'NISN / Username Peserta';
                tokenWrap.classList.remove('hidden');
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2';
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeLoginModal() {
            const modal = document.getElementById('loginModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function processLogin(e) {
            e.preventDefault();
            const role = document.getElementById('loginRoleField').value;
            const userInput = document.getElementById('loginUsernameField').value.trim();
            const passInput = document.getElementById('loginPasswordField').value.trim();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);

            if (role === 'admin') {
                const validUser = settings.adminUsername || 'admin';
                const validPass = settings.adminPassword || 'admin123';
                if (userInput === validUser && passInput === validPass) {
                    activeSession = { role: 'admin', user: { name: 'Administrator' } };
                    closeLoginModal();
                    showToast('Berhasil masuk sebagai Administrator.', 'success');
                    navigateTo('admin');
                } else {
                    showToast('Username atau Password Administrator tidak valid!', 'error');
                }
                return;
            }

            if (role === 'pengawas') {
                const proctors = getDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
                const foundProctor = proctors.find(p => p.username === userInput && p.password === passInput);
                if (foundProctor) {
                    activeSession = { role: 'pengawas', user: foundProctor };
                    closeLoginModal();
                    showToast(`Selamat bertugas, ${foundProctor.name}!`, 'success');
                    navigateTo('pengawas');
                } else {
                    showToast('Username atau Password Pengawas tidak terdaftar!', 'error');
                }
                return;
            }

            if (role === 'siswa') {
                const tokenInput = document.getElementById('loginTokenField').value.trim().toUpperCase();
                if (!tokenInput || tokenInput !== settings.activeToken.toUpperCase()) {
                    showToast('Token Ujian salah atau kosong! Minta token aktif ke pengawas ruang.', 'error');
                    return;
                }

                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                const student = students.find(s =>
                    (s.username.toLowerCase() === userInput.toLowerCase() || s.nisn === userInput) &&
                    s.password === passInput
                );

                if (!student) {
                    showToast('NISN/Username atau Password tidak cocok dengan Kartu Login!', 'error');
                    return;
                }

                // Cek target Assign Kelas/Ruangan
                if (settings.targetGroup && settings.targetGroup !== 'ALL') {
                    const matchTarget = student.className === settings.targetGroup || student.room === settings.targetGroup;
                    if (!matchTarget) {
                        showToast(`Ujian saat ini hanya di-assign untuk ${settings.targetGroup}.`, 'warning');
                        return;
                    }
                }

                // Cek apakah siswa sudah menyelesaikan ujian
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                if (liveMap[student.id]?.status === 'Selesai') {
                    showToast('Anda sudah menyelesaikan ujian ini. Hubungi pengawas/admin jika perlu reset sesi.', 'warning');
                    return;
                }

                const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
                if (questions.length === 0) {
                    showToast('Bank soal masih kosong. Hubungi Administrator.', 'warning');
                    return;
                }

                activeSession = { role: 'siswa', user: student };
                closeLoginModal();
                beginStudentExam(student);
            }
        }

        function logoutSession() {
            activeSession = null;
            showToast('Anda telah keluar dari dashboard.', 'info');
            navigateTo('portal');
        }

        // =================================================================
        // G. DASHBOARD ADMINISTRATOR: 6 MENU TERPISAH
        // =================================================================
        function switchAdminMenu(menuKey) {
            currentAdminMenu = menuKey;
            const menus = ['monitoring', 'soal', 'assign', 'nilai', 'kartu', 'setelan'];

            menus.forEach(m => {
                document.getElementById(`adminSection-${m}`).classList.add('hidden');
                const btn = document.getElementById(`adminNav-${m}`);
                btn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap';
            });

            document.getElementById(`adminSection-${menuKey}`).classList.remove('hidden');
            const activeBtn = document.getElementById(`adminNav-${menuKey}`);
            activeBtn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap';

            syncAllViews();
        }

        function populateFilterDropdowns() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const classes = [...new Set(students.map(s => s.className).filter(Boolean))];
            const rooms = [...new Set(students.map(s => s.room).filter(Boolean))];

            // Monitoring room/class filter
            const monFilter = document.getElementById('monRoomFilter');
            const currentMonVal = monFilter.value;
            monFilter.innerHTML = `<option value="">Semua Kelas & Ruangan</option>` +
                classes.map(c => `<option value="class:${c}">Kelas: ${c}</option>`).join('') +
                rooms.map(r => `<option value="room:${r}">Ruangan: ${r}</option>`).join('');
            monFilter.value = currentMonVal;

            // Nilai filters
            const nClass = document.getElementById('nilaiClassFilter');
            const curNClass = nClass.value;
            nClass.innerHTML = `<option value="">Semua Kelas</option>` + classes.map(c => `<option value="${c}">${c}</option>`).join('');
            nClass.value = curNClass;

            const nRoom = document.getElementById('nilaiRoomFilter');
            const curNRoom = nRoom.value;
            nRoom.innerHTML = `<option value="">Semua Ruangan</option>` + rooms.map(r => `<option value="${r}">${r}</option>`).join('');
            nRoom.value = curNRoom;

            // Assign target dropdown
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const assignTarget = document.getElementById('assignTargetGroup');
            assignTarget.innerHTML = `<option value="ALL">Semua Kelas & Ruangan</option>` +
                classes.map(c => `<option value="${c}">Khusus Kelas: ${c}</option>`).join('') +
                rooms.map(r => `<option value="${r}">Khusus Ruangan: ${r}</option>`).join('');
            assignTarget.value = settings.targetGroup || 'ALL';
        }

        // --- MENU 1: MONITORING SISWA ---
        function renderAdminMonitoring() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('monSearchInput')?.value || '').toLowerCase();
            const filterVal = document.getElementById('monRoomFilter')?.value || '';

            let activeCount = 0, doneCount = 0, lockedCount = 0;
            students.forEach(s => {
                const st = liveMap[s.id]?.status;
                if (st === 'Mengerjakan') activeCount++;
                else if (st === 'Selesai') doneCount++;
                else if (st === 'TERKUNCI') lockedCount++;
            });

            document.getElementById('monStatTotal').innerText = students.length;
            document.getElementById('monStatActive').innerText = activeCount;
            document.getElementById('monStatDone').innerText = doneCount;
            document.getElementById('monStatLocked').innerText = lockedCount;

            const filtered = students.filter(s => {
                const matchText = s.name.toLowerCase().includes(search) || s.nisn.includes(search);
                let matchGroup = true;
                if (filterVal.startsWith('class:')) matchGroup = s.className === filterVal.replace('class:', '');
                if (filterVal.startsWith('room:')) matchGroup = s.room === filterVal.replace('room:', '');
                return matchText && matchGroup;
            });

            const tbody = document.getElementById('adminMonitoringTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="py-8 text-center text-slate-400 text-xs">Tidak ada data peserta yang sesuai.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => {
                const st = liveMap[s.id] || { status: 'Belum Masuk', progress: '-', score: '-', violations: 0 };
                let badge = `<span class="px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-[11px] font-bold">Belum Masuk</span>`;
                if (st.status === 'Mengerjakan') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-blue-100 text-blue-700 text-[11px] font-bold animate-pulse">Mengerjakan</span>`;
                } else if (st.status === 'TERKUNCI') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-red-600 text-white text-[11px] font-bold">TERKUNCI</span>`;
                } else if (st.status === 'Selesai') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">Selesai</span>`;
                }

                return `
                    <tr class="hover:bg-slate-50 transition ${st.status === 'TERKUNCI' ? 'bg-red-50/60' : ''}">
                        <td class="py-3 px-4">
                            <div class="font-bold text-slate-900">${s.name}</div>
                            <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                        </td>
                        <td class="py-3 px-4 text-xs text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                        <td class="py-3 px-4">${badge}</td>
                        <td class="py-3 px-4 font-mono text-xs font-semibold">${st.progress}</td>
                        <td class="py-3 px-4 font-extrabold text-indigo-700">${st.score !== undefined ? st.score : '-'}</td>
                        <td class="py-3 px-4">
                            <span class="px-2 py-0.5 rounded-full text-xs font-bold ${st.violations > 0 ? 'bg-red-100 text-red-700' : 'bg-slate-100 text-slate-600'}">
                                ${st.violations || 0}x
                            </span>
                        </td>
                        <td class="py-3 px-4 text-right space-x-1">
                            ${st.status === 'TERKUNCI' ? `
                                <button onclick="remoteUnlockStudentById('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold transition">
                                    Buka Kunci
                                </button>
                            ` : ''}
                            <button onclick="resetStudentExamSession('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-bold transition" title="Reset Sesi Ujian Siswa">
                                Reset Sesi
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function remoteUnlockStudentById(studentId) {
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            if (liveMap[studentId]) {
                liveMap[studentId].status = 'Mengerjakan';
                liveMap[studentId].remoteUnlocked = true;
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                syncAllViews();
                showToast('Kunci layar peserta telah dibuka dari jarak jauh!', 'success');
            }
        }

        function resetStudentExamSession(studentId) {
            showConfirmDialog('Reset Sesi Ujian Peserta?', 'Status pengerjaan dan nilai peserta ini akan direset sehingga peserta dapat login & ujian ulang menggunakan token.', () => {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                delete liveMap[studentId];
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                syncAllViews();
                showToast('Sesi peserta berhasil direset.', 'info');
            });
        }

        // --- MENU 2: BANK & INPUT SOAL (MANUAL, WORD/PDF, EDIT) ---
        function handleAddManualQuestion(e) {
            e.preventDefault();
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
            const optE = document.getElementById('mQOptE').value.trim();

            const options = {
                A: document.getElementById('mQOptA').value.trim(),
                B: document.getElementById('mQOptB').value.trim(),
                C: document.getElementById('mQOptC').value.trim(),
                D: document.getElementById('mQOptD').value.trim()
            };
            if (optE) options.E = optE;

            const ans = document.getElementById('mQAnswer').value;
            if (ans === 'E' && !optE) {
                showToast('Opsi E kosong, silakan isi Opsi E atau pilih kunci A-D.', 'warning');
                return;
            }

            questions.push({
                id: 'q_' + Date.now(),
                question: document.getElementById('mQText').value.trim(),
                options,
                answer: ans,
                points: parseInt(document.getElementById('mQPoints').value, 10) || 10
            });

            setDB(STORAGE_KEYS.QUESTIONS, questions);
            e.target.reset();
            document.getElementById('mQPoints').value = 10;
            syncAllViews();
            showToast('Soal baru berhasil ditambahkan!', 'success');
        }

        function parseExtractedTextToQuestions(rawText) {
            const normalized = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            const parsedList = [];

            const lines = normalized.split('\n').map(l => l.trim()).filter(Boolean);
            lines.forEach((line, idx) => {
                if (line.includes('|')) {
                    const parts = line.split('|').map(p => p.trim());
                    if (parts.length >= 6) {
                        const hasE = parts.length >= 7 && ['A', 'B', 'C', 'D', 'E'].includes(parts[6].toUpperCase());
                        const ans = hasE ? parts[6].toUpperCase() : parts[5].toUpperCase();
                        if (['A', 'B', 'C', 'D', 'E'].includes(ans)) {
                            const opts = { A: parts[1], B: parts[2], C: parts[3], D: parts[4] };
                            if (hasE && parts[5]) opts.E = parts[5];
                            parsedList.push({
                                id: `q_imp_${Date.now()}_${idx}`,
                                question: parts[0],
                                options: opts,
                                answer: ans,
                                points: 10
                            });
                        }
                    }
                }
            });

            if (parsedList.length > 0) return parsedList;

            const blocks = normalized.split(/(?=\n?\s*\d+[\.\)]\s+)/g).map(b => b.trim()).filter(Boolean);
            blocks.forEach((block, idx) => {
                const keyMatch = block.match(/(?:Kunci|Jawaban)\s*[:=\-]\s*([A-Ea-e])/i);
                const optAMatch = block.match(/(?:^|\n|\s)A[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)B[\.\)]\s*)/i);
                const optBMatch = block.match(/(?:^|\n|\s)B[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)C[\.\)]\s*)/i);
                const optCMatch = block.match(/(?:^|\n|\s)C[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)D[\.\)]\s*)/i);
                const optDMatch = block.match(/(?:^|\n|\s)D[\.\)]\s*([\s\S]*?)(?=(?:(?:\n|\s)E[\.\)]\s*|(?:Kunci|Jawaban)\s*[:=\-]|$))/i);
                const optEMatch = block.match(/(?:^|\n|\s)E[\.\)]\s*([\s\S]*?)(?=(?:(?:Kunci|Jawaban)\s*[:=\-]|$))/i);

                if (optAMatch && optBMatch && optCMatch && optDMatch) {
                    const firstOptIdx = block.search(/(?:^|\n|\s)A[\.\)]\s*/i);
                    let qText = block.substring(0, firstOptIdx).replace(/^\d+[\.\)]\s*/, '').trim();
                    const opts = {
                        A: optAMatch[1].trim(),
                        B: optBMatch[1].trim(),
                        C: optCMatch[1].trim(),
                        D: optDMatch[1].trim()
                    };
                    if (optEMatch && optEMatch[1].trim()) {
                        opts.E = optEMatch[1].trim();
                    }
                    const ans = keyMatch ? keyMatch[1].toUpperCase() : 'A';
                    if (qText) {
                        parsedList.push({
                            id: `q_doc_${Date.now()}_${idx}`,
                            question: qText,
                            options: opts,
                            answer: ans,
                            points: 10
                        });
                    }
                }
            });

            return parsedList;
        }

        async function handleDocumentQuestionUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const statusBox = document.getElementById('uploadParseStatus');
            if (statusBox) {
                statusBox.classList.remove('hidden');
                statusBox.innerText = `Sedang membaca file "${file.name}"...`;
            }

            const ext = file.name.split('.').pop().toLowerCase();
            let extractedText = '';

            try {
                if (ext === 'docx') {
                    if (typeof mammoth === 'undefined') {
                        throw new Error('Library Mammoth.js tidak tersedia (mode offline). Gunakan file .TXT atau input manual.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    const result = await mammoth.extractRawText({ arrayBuffer });
                    extractedText = result.value;
                } else if (ext === 'pdf') {
                    if (typeof pdfjsLib === 'undefined') {
                        throw new Error('Library PDF.js tidak tersedia (mode offline). Gunakan file .TXT atau input manual.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                    for (let p = 1; p <= pdf.numPages; p++) {
                        const page = await pdf.getPage(p);
                        const content = await page.getTextContent();
                        const pageText = content.items.map(item => item.str).join(' ');
                        extractedText += '\n' + pageText;
                    }
                } else {
                    extractedText = await file.text();
                }

                const newQuestions = parseExtractedTextToQuestions(extractedText);
                if (newQuestions.length > 0) {
                    const current = getDB(STORAGE_KEYS.QUESTIONS, []);
                    setDB(STORAGE_KEYS.QUESTIONS, [...current, ...newQuestions]);
                    syncAllViews();
                    showToast(`Berhasil mengimpor ${newQuestions.length} soal dari ${file.name}!`, 'success');
                    if (statusBox) statusBox.innerText = `Sukses: ${newQuestions.length} soal ditambahkan dari ${file.name}.`;
                } else {
                    showToast('Tidak ditemukan pola soal yang valid dalam file. Unduh contoh format untuk panduan.', 'warning');
                    if (statusBox) statusBox.innerText = 'Format soal tidak dikenali. Pastikan menggunakan penomoran 1., opsi A-D, dan Kunci: A.';
                }
            } catch (err) {
                showToast(err.message || 'Gagal memproses file dokumen.', 'error');
                if (statusBox) statusBox.innerText = `Error: ${err.message}`;
            } finally {
                event.target.value = '';
            }
        }

        function downloadSampleQuestionFormat() {
            const sample = [
                "1. Ibu kota negara Indonesia adalah...",
                "A. Bandung",
                "B. Jakarta",
                "C. Surabaya",
                "D. Medan",
                "E. Makassar",
                "Kunci: B",
                "",
                "2. Hasil dari 25 x 4 adalah...",
                "A. 80",
                "B. 90",
                "C. 100",
                "D. 110",
                "Kunci: C"
            ].join('\n');

            const blob = new Blob([sample], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Contoh_Format_Soal_CBT.txt';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('Contoh format soal berhasil diunduh!', 'success');
        }

        function renderAdminQuestionList() {
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
            const container = document.getElementById('adminSoalListContainer');
            document.getElementById('adminBankSoalCount').innerText = questions.length;

            if (questions.length === 0) {
                container.innerHTML = `<div class="text-center py-12 text-xs text-slate-400">Belum ada soal tersimpan di Bank Soal.</div>`;
                return;
            }

            container.innerHTML = questions.map((q, idx) => {
                const optionKeys = Object.keys(q.options);
                return `
                    <div class="p-4 rounded-2xl border border-slate-200 bg-slate-50/60 hover:bg-white transition space-y-2.5">
                        <div class="flex items-start justify-between gap-3">
                            <div class="text-xs sm:text-sm font-bold text-slate-900">
                                <span class="text-indigo-600 mr-1">${idx + 1}.</span> ${q.question}
                                <span class="ml-2 px-2 py-0.5 rounded bg-slate-200 text-slate-700 text-[10px] font-semibold">${q.points || 10} Poin</span>
                            </div>
                            <div class="flex items-center gap-1.5 shrink-0">
                                <button onclick="openEditQuestionModal('${q.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-800 text-xs font-bold transition flex items-center gap-1">
                                    <i data-lucide="edit-3" class="w-3.5 h-3.5"></i> Edit
                                </button>
                                <button onclick="deleteSingleQuestion('${q.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Soal">
                                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-xs">
                            ${optionKeys.map(k => `
                                <div class="px-2.5 py-1.5 rounded-lg border ${q.answer === k ? 'bg-emerald-100 border-emerald-300 text-emerald-900 font-bold' : 'bg-white border-slate-200 text-slate-600'}">
                                    <strong>${k}.</strong> ${q.options[k]}
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `;
            }).join('');
        }

        function openEditQuestionModal(qId) {
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
            const q = questions.find(item => item.id === qId);
            if (!q) return;

            document.getElementById('editQId').value = q.id;
            document.getElementById('editQText').value = q.question;
            document.getElementById('editQOptA').value = q.options.A || '';
            document.getElementById('editQOptB').value = q.options.B || '';
            document.getElementById('editQOptC').value = q.options.C || '';
            document.getElementById('editQOptD').value = q.options.D || '';
            document.getElementById('editQOptE').value = q.options.E || '';
            document.getElementById('editQAnswer').value = q.answer;
            document.getElementById('editQPoints').value = q.points || 10;

            const modal = document.getElementById('editQuestionModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeEditQuestionModal() {
            const modal = document.getElementById('editQuestionModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function saveEditedQuestion(e) {
            e.preventDefault();
            const qId = document.getElementById('editQId').value;
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []);
            const idx = questions.findIndex(item => item.id === qId);
            if (idx === -1) return;

            const optE = document.getElementById('editQOptE').value.trim();
            const options = {
                A: document.getElementById('editQOptA').value.trim(),
                B: document.getElementById('editQOptB').value.trim(),
                C: document.getElementById('editQOptC').value.trim(),
                D: document.getElementById('editQOptD').value.trim()
            };
            if (optE) options.E = optE;

            const ans = document.getElementById('editQAnswer').value;
            if (ans === 'E' && !optE) {
                showToast('Opsi E kosong, pilih kunci A-D atau isi teks Opsi E.', 'warning');
                return;
            }

            questions[idx] = {
                ...questions[idx],
                question: document.getElementById('editQText').value.trim(),
                options,
                answer: ans,
                points: parseInt(document.getElementById('editQPoints').value, 10) || 10
            };

            setDB(STORAGE_KEYS.QUESTIONS, questions);
            closeEditQuestionModal();
            syncAllViews();
            showToast('Perubahan soal berhasil disimpan!', 'success');
        }

        function deleteSingleQuestion(qId) {
            const questions = getDB(STORAGE_KEYS.QUESTIONS, []).filter(q => q.id !== qId);
            setDB(STORAGE_KEYS.QUESTIONS, questions);
            syncAllViews();
            showToast('Soal dihapus dari bank soal.', 'info');
        }

        function confirmClearAllQuestions() {
            showConfirmDialog('Kosongkan Bank Soal?', 'Seluruh soal yang tersimpan akan dihapus permanen.', () => {
                setDB(STORAGE_KEYS.QUESTIONS, []);
                syncAllViews();
                showToast('Bank soal telah dikosongkan.', 'info');
            });
        }

        // --- MENU 3: ASSIGN UJIAN & TOKEN ---
        function generateRandomToken() {
            const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
            let token = '';
            for (let i = 0; i < 6; i++) token += chars.charAt(Math.floor(Math.random() * chars.length));
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = token;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast(`Token baru berhasil dibuat: ${token}`, 'success');
        }

        function saveManualToken() {
            const val = document.getElementById('customTokenInput').value.trim().toUpperCase();
            if (val.length < 4) {
                showToast('Token minimal terdiri dari 4 karakter!', 'warning');
                return;
            }
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = val;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            document.getElementById('customTokenInput').value = '';
            syncAllViews();
            showToast(`Token berhasil diubah menjadi: ${val}`, 'success');
        }

        function saveAssignSettings(e) {
            e.preventDefault();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.examTitle = document.getElementById('assignExamTitle').value.trim();
            settings.subjectName = document.getElementById('assignSubject').value.trim();
            settings.targetGroup = document.getElementById('assignTargetGroup').value;
            settings.durationMinutes = parseInt(document.getElementById('assignDuration').value, 10) || 45;
            settings.shuffleQuestions = document.getElementById('assignShuffleQuestions').checked;
            settings.shuffleOptions = document.getElementById('assignShuffleOptions').checked;

            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast('Pengaturan Assign Ujian berhasil disimpan & diterapkan!', 'success');
        }

        // --- MENU 4: DATA SISWA & DOWNLOAD NILAI EXCEL (.XLSX) ---
        function renderAdminNilaiTable() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('nilaiSearchInput')?.value || '').toLowerCase();
            const classFilter = document.getElementById('nilaiClassFilter')?.value || '';
            const roomFilter = document.getElementById('nilaiRoomFilter')?.value || '';

            const filtered = students.filter(s => {
                const matchText = s.name.toLowerCase().includes(search) || s.nisn.includes(search);
                const matchClass = !classFilter || s.className === classFilter;
                const matchRoom = !roomFilter || s.room === roomFilter;
                return matchText && matchClass && matchRoom;
            });

            const tbody = document.getElementById('adminNilaiTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" class="py-8 text-center text-slate-400 text-xs">Tidak ada data peserta yang ditemukan.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map((s, idx) => {
                const st = liveMap[s.id] || { status: 'Belum Masuk', correct: '-', wrong: '-', score: '-', violations: 0 };
                return `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-3 text-slate-500 font-semibold">${idx + 1}</td>
                        <td class="py-3 px-3 font-bold text-slate-900">${s.name}</td>
                        <td class="py-3 px-3 font-mono text-xs text-slate-600">${s.nisn}</td>
                        <td class="py-3 px-3 text-xs">${s.className}</td>
                        <td class="py-3 px-3 text-xs font-semibold">${s.room}</td>
                        <td class="py-3 px-3">
                            <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${st.status === 'Selesai' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}">
                                ${st.status}
                            </span>
                        </td>
                        <td class="py-3 px-3 font-bold text-emerald-700">${st.correct !== undefined ? st.correct : '-'}</td>
                        <td class="py-3 px-3 font-bold text-red-600">${st.wrong !== undefined ? st.wrong : '-'}</td>
                        <td class="py-3 px-3">
                            <span class="px-2.5 py-1 rounded-lg font-extrabold text-xs ${typeof st.score === 'number' ? 'bg-blue-100 text-blue-800' : 'bg-slate-100 text-slate-400'}">
                                ${st.score !== undefined ? st.score : '-'}
                            </span>
                        </td>
                        <td class="py-3 px-3 text-xs font-bold ${st.violations > 0 ? 'text-red-600' : 'text-slate-500'}">${st.violations || 0}x</td>
                    </tr>
                `;
            }).join('');
        }

        function downloadExcelResults() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);

            const rows = students.map((s, idx) => {
                const st = liveMap[s.id] || {};
                return {
                    'No': idx + 1,
                    'Nama Peserta': s.name,
                    'NISN': s.nisn,
                    'Kelas': s.className,
                    'Ruangan': s.room,
                    'Status Ujian': st.status || 'Belum Masuk',
                    'Jumlah Benar': st.correct !== undefined ? st.correct : 0,
                    'Jumlah Salah': st.wrong !== undefined ? st.wrong : 0,
                    'Nilai Akhir': st.score !== undefined ? st.score : 0,
                    'Pelanggaran': st.violations || 0,
                    'Waktu Selesai': st.finishedAt || '-'
                };
            });

            const fileName = `Rekap_Nilai_${settings.subjectName.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}`;

            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Hasil Nilai CBT');
                XLSX.writeFile(wb, `${fileName}.xlsx`);
                showToast('File rekap nilai Excel (.XLSX) berhasil diunduh!', 'success');
            } else {
                const headers = Object.keys(rows[0] || {});
                const csv = [headers.join(','), ...rows.map(r => headers.map(h => `"${r[h]}"`).join(','))].join('\n');
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = `${fileName}.csv`;
                a.click();
                showToast('File rekap nilai (.CSV) berhasil diunduh!', 'success');
            }
        }

        // --- MENU 5: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN ---
        function toggleShowStudentPasswords() {
            showStudentPasswords = !showStudentPasswords;
            const btn = document.getElementById('btnTogglePassVis');
            btn.innerHTML = showStudentPasswords
                ? `<i data-lucide="eye-off" class="w-3.5 h-3.5"></i> <span>Sembunyikan Password</span>`
                : `<i data-lucide="eye" class="w-3.5 h-3.5"></i> <span>Tampilkan Password</span>`;
            renderAdminStudentAccountTable();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function renderAdminStudentAccountTable() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const search = (document.getElementById('accountSearchInput')?.value || '').toLowerCase();
            document.getElementById('studentAccountCount').innerText = students.length;

            const filtered = students.filter(s =>
                s.name.toLowerCase().includes(search) ||
                s.nisn.toLowerCase().includes(search) ||
                s.className.toLowerCase().includes(search) ||
                s.room.toLowerCase().includes(search)
            );

            const tbody = document.getElementById('adminStudentAccountTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400 text-xs">Tidak ada akun siswa yang ditemukan.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => `
                <tr class="hover:bg-slate-50 transition">
                    <td class="py-2.5 px-3">
                        <div class="font-bold text-slate-900">${s.name}</div>
                        <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                    </td>
                    <td class="py-2.5 px-3 text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                    <td class="py-2.5 px-3 font-mono font-bold text-slate-900">${s.username}</td>
                    <td class="py-2.5 px-3 font-mono font-extrabold text-blue-700">
                        ${showStudentPasswords ? s.password : '•••••'}
                    </td>
                    <td class="py-2.5 px-3 text-right space-x-1">
                        <button onclick="openEditStudentModal('${s.id}')" class="px-2.5 py-1 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-800 text-[11px] font-bold transition">
                            Edit / Reset Pass
                        </button>
                        <button onclick="deleteSingleStudent('${s.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Akun Siswa">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5 inline"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function openEditStudentModal(stuId) {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const s = students.find(item => item.id === stuId);
            if (!s) return;

            document.getElementById('editStuId').value = s.id;
            document.getElementById('editStuName').value = s.name;
            document.getElementById('editStuNisn').value = s.nisn;
            document.getElementById('editStuClass').value = s.className;
            document.getElementById('editStuRoom').value = s.room;
            document.getElementById('editStuUser').value = s.username;
            document.getElementById('editStuPass').value = s.password;

            const modal = document.getElementById('editStudentModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeEditStudentModal() {
            const modal = document.getElementById('editStudentModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function generateRandomStudentPassInModal() {
            document.getElementById('editStuPass').value = String(Math.floor(10000 + Math.random() * 90000));
        }

        function saveEditedStudent(e) {
            e.preventDefault();
            const stuId = document.getElementById('editStuId').value;
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const idx = students.findIndex(item => item.id === stuId);
            if (idx === -1) return;

            students[idx] = {
                ...students[idx],
                name: document.getElementById('editStuName').value.trim(),
                nisn: document.getElementById('editStuNisn').value.trim(),
                className: document.getElementById('editStuClass').value.trim(),
                room: document.getElementById('editStuRoom').value.trim(),
                username: document.getElementById('editStuUser').value.trim(),
                password: document.getElementById('editStuPass').value.trim()
            };

            setDB(STORAGE_KEYS.STUDENTS, students);
            closeEditStudentModal();
            syncAllViews();
            showToast('Data akun & password siswa berhasil diperbarui!', 'success');
        }

        function exportStudentAccountsExcel() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            if (students.length === 0) {
                showToast('Daftar akun siswa masih kosong.', 'warning');
                return;
            }
            const rows = students.map((s, idx) => ({
                'No': idx + 1,
                'Nama Lengkap': s.name,
                'NISN': s.nisn,
                'Kelas': s.className,
                'Ruangan': s.room,
                'Username Login': s.username,
                'Password': s.password
            }));

            const fileName = `Daftar_Akun_Peserta_CBT_${new Date().toISOString().slice(0, 10)}`;
            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Akun Peserta');
                XLSX.writeFile(wb, `${fileName}.xlsx`);
                showToast('Daftar akun & password siswa (.XLSX) berhasil diunduh!', 'success');
            } else {
                const headers = Object.keys(rows[0]);
                const csv = [headers.join(','), ...rows.map(r => headers.map(h => `"${r[h]}"`).join(','))].join('\n');
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = `${fileName}.csv`;
                a.click();
                showToast('Daftar akun siswa (.CSV) berhasil diunduh!', 'success');
            }
        }

        function downloadStudentExcelTemplate() {
            const templateData = [
                { Nama: 'Dewi Lestari', NISN: '0081234504', Kelas: 'XII IPA 1', Ruangan: 'Ruang 01' },
                { Nama: 'Fajar Ramadhan', NISN: '0081234505', Kelas: 'XII IPA 2', Ruangan: 'Ruang 01' },
                { Nama: 'Gita Savitri', NISN: '0081234506', Kelas: 'XII IPS 1', Ruangan: 'Ruang 02' }
            ];

            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(templateData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Template Peserta');
                XLSX.writeFile(wb, 'Template_Import_Peserta_CBT.xlsx');
                showToast('Template Excel (Nama, NISN, Kelas, Ruangan) berhasil diunduh!', 'success');
            } else {
                const csv = 'Nama,NISN,Kelas,Ruangan\nDewi Lestari,0081234504,XII IPA 1,Ruang 01\nFajar Ramadhan,0081234505,XII IPA 2,Ruang 01';
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = 'Template_Import_Peserta_CBT.csv';
                a.click();
                showToast('Template CSV peserta berhasil diunduh!', 'success');
            }
        }

        async function handleStudentExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            try {
                let importedRows = [];
                if (typeof XLSX !== 'undefined') {
                    const data = await file.arrayBuffer();
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    importedRows = XLSX.utils.sheet_to_json(sheet);
                } else {
                    const text = await file.text();
                    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
                    for (let i = 1; i < lines.length; i++) {
                        const cols = lines[i].split(',').map(c => c.replace(/"/g, '').trim());
                        if (cols.length >= 4) {
                            importedRows.push({ Nama: cols[0], NISN: cols[1], Kelas: cols[2], Ruangan: cols[3] });
                        }
                    }
                }

                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                let added = 0;

                importedRows.forEach((row, idx) => {
                    const nama = String(row.Nama || row.nama || row.NAMA || '').trim();
                    const nisn = String(row.NISN || row.nisn || '').trim();
                    const kelas = String(row.Kelas || row.kelas || row.KELAS || '').trim();
                    const ruangan = String(row.Ruangan || row.ruangan || row.RUANGAN || '').trim();

                    if (nama && nisn) {
                        const exists = students.some(s => s.nisn === nisn);
                        if (!exists) {
                            students.push({
                                id: `stu_xl_${Date.now()}_${idx}`,
                                name: nama,
                                nisn: nisn,
                                className: kelas || 'Umum',
                                room: ruangan || 'Ruang 01',
                                username: nisn,
                                password: String(Math.floor(10000 + Math.random() * 90000))
                            });
                            added++;
                        }
                    }
                });

                setDB(STORAGE_KEYS.STUDENTS, students);
                syncAllViews();
                showToast(`Berhasil mengimpor ${added} akun siswa & kartu login otomatis!`, 'success');
            } catch (err) {
                showToast('Gagal membaca file Excel peserta. Pastikan format sesuai template.', 'error');
            } finally {
                event.target.value = '';
            }
        }

        function handleAddManualStudent(e) {
            e.preventDefault();
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const name = document.getElementById('kStuName').value.trim();
            const nisn = document.getElementById('kStuNisn').value.trim();
            const className = document.getElementById('kStuClass').value.trim();
            const room = document.getElementById('kStuRoom').value.trim();
            const customUser = document.getElementById('kStuUser').value.trim();
            const customPass = document.getElementById('kStuPass').value.trim();

            if (students.some(s => s.nisn === nisn)) {
                showToast('NISN sudah terdaftar di sistem!', 'error');
                return;
            }

            students.push({
                id: 'stu_' + Date.now(),
                name,
                nisn,
                className,
                room,
                username: customUser || nisn,
                password: customPass || String(Math.floor(10000 + Math.random() * 90000))
            });

            setDB(STORAGE_KEYS.STUDENTS, students);
            e.target.reset();
            syncAllViews();
            showToast(`Akun & Kartu Login untuk ${name} berhasil dibuat!`, 'success');
        }

        function deleteSingleStudent(studentId) {
            const students = getDB(STORAGE_KEYS.STUDENTS, []).filter(s => s.id !== studentId);
            setDB(STORAGE_KEYS.STUDENTS, students);
            syncAllViews();
            showToast('Akun siswa dihapus.', 'info');
        }

        function confirmClearAllStudents() {
            showConfirmDialog('Kosongkan Daftar Siswa?', 'Seluruh akun siswa dan kartu login akan dihapus agar Anda bisa mengimpor data peserta yang baru.', () => {
                setDB(STORAGE_KEYS.STUDENTS, []);
                syncAllViews();
                showToast('Daftar siswa berhasil dikosongkan.', 'info');
            });
        }

        function renderAdminLoginCards() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const container = document.getElementById('printCardsContainer');
            document.getElementById('cardTotalCount').innerText = students.length;

            if (students.length === 0) {
                container.innerHTML = `<div class="col-span-2 text-center py-12 text-xs text-slate-400">Belum ada data peserta. Silakan tambah manual atau import Excel.</div>`;
                return;
            }

            container.innerHTML = students.map((s, idx) => `
                <div class="login-card-item border-2 border-slate-300 rounded-2xl p-4 bg-white relative flex flex-col justify-between">
                    <button onclick="deleteSingleStudent('${s.id}')" class="no-print absolute top-3 right-3 w-7 h-7 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 flex items-center justify-center" title="Hapus Peserta">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                    </button>

                    <!-- Kop Kartu -->
                    <div class="flex items-center gap-2.5 pb-2.5 mb-2.5 border-b border-slate-200">
                        <div class="w-9 h-9 rounded-xl bg-slate-900 text-white flex items-center justify-center font-extrabold text-xs shrink-0">
                            CBT
                        </div>
                        <div class="min-w-0 pr-6">
                            <div class="text-[11px] font-extrabold uppercase tracking-wider text-slate-900 truncate">KARTU PESERTA UJIAN</div>
                            <div class="text-[10px] text-slate-500 truncate">${settings.schoolName}</div>
                        </div>
                    </div>

                    <!-- Identitas Peserta -->
                    <div class="space-y-1 text-xs mb-3">
                        <div class="flex justify-between">
                            <span class="text-slate-500">No. Urut:</span>
                            <span class="font-bold text-slate-800">${String(idx + 1).padStart(3, '0')}</span>
                        </div>
                        <div class="flex justify-between">
                            <span class="text-slate-500">Nama Peserta:</span>
                            <span class="font-extrabold text-slate-900">${s.name}</span>
                        </div>
                        <div class="flex justify-between">
                            <span class="text-slate-500">NISN:</span>
                            <span class="font-mono font-bold text-slate-700">${s.nisn}</span>
                        </div>
                        <div class="flex justify-between">
                            <span class="text-slate-500">Kelas / Ruangan:</span>
                            <span class="font-bold text-indigo-700">${s.className} • ${s.room}</span>
                        </div>
                    </div>

                    <!-- Kotak Kredensial & Visual QR SVG -->
                    <div class="bg-slate-50 border border-dashed border-slate-300 rounded-xl p-2.5 flex items-center justify-between gap-2">
                        <div class="space-y-1 text-xs">
                            <div>
                                <span class="text-[9px] font-bold uppercase text-slate-400 block">Username Login</span>
                                <span class="font-mono font-extrabold text-slate-900">${s.username}</span>
                            </div>
                            <div>
                                <span class="text-[9px] font-bold uppercase text-slate-400 block">Password</span>
                                <span class="font-mono font-extrabold text-blue-700">${s.password}</span>
                            </div>
                        </div>
                        <!-- Visual QR Code SVG Tajam -->
                        <svg viewBox="0 0 40 40" class="w-12 h-12 shrink-0 border border-slate-200 p-1 bg-white rounded-lg">
                            <rect x="2" y="2" width="12" height="12" fill="#0f172a"/>
                            <rect x="4" y="4" width="8" height="8" fill="#ffffff"/>
                            <rect x="6" y="6" width="4" height="4" fill="#0f172a"/>
                            <rect x="26" y="2" width="12" height="12" fill="#0f172a"/>
                            <rect x="28" y="4" width="8" height="8" fill="#ffffff"/>
                            <rect x="30" y="6" width="4" height="4" fill="#0f172a"/>
                            <rect x="2" y="26" width="12" height="12" fill="#0f172a"/>
                            <rect x="4" y="28" width="8" height="8" fill="#ffffff"/>
                            <rect x="6" y="30" width="4" height="4" fill="#0f172a"/>
                            <rect x="18" y="6" width="4" height="4" fill="#0f172a"/>
                            <rect x="18" y="18" width="6" height="6" fill="#2563eb"/>
                            <rect x="28" y="20" width="4" height="8" fill="#0f172a"/>
                            <rect x="18" y="30" width="10" height="4" fill="#0f172a"/>
                        </svg>
                    </div>
                </div>
            `).join('');
        }

        // --- MENU 6: SETELAN KESELURUHAN ---
        function renderAdminSettingsForms() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const proctors = getDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);

            // Isi form Assign juga agar sinkron
            document.getElementById('assignExamTitle').value = settings.examTitle;
            document.getElementById('assignSubject').value = settings.subjectName;
            document.getElementById('assignDuration').value = settings.durationMinutes;
            document.getElementById('assignShuffleQuestions').checked = settings.shuffleQuestions;
            document.getElementById('assignShuffleOptions').checked = settings.shuffleOptions;

            // Isi form Setelan
            document.getElementById('setSchoolName').value = settings.schoolName;
            document.getElementById('setSchoolSubtitle').value = settings.schoolSubtitle;
            document.getElementById('setAdminUsername').value = settings.adminUsername || 'admin';
            document.getElementById('setAdminPassword').value = settings.adminPassword || 'admin123';
            document.getElementById('setUnlockPin').value = settings.unlockPin;
            document.getElementById('setRequireFullscreen').checked = settings.requireFullscreen;
            document.getElementById('setDetectTabSwitch').checked = settings.detectTabSwitch;
            document.getElementById('setShowScoreToStudent').checked = settings.showScoreToStudent;

            const procContainer = document.getElementById('proctorListContainer');
            procContainer.innerHTML = proctors.map(p => `
                <div class="flex items-center justify-between p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                    <div>
                        <div class="font-bold text-slate-900">${p.name} <span class="text-amber-700 font-semibold">(${p.room})</span></div>
                        <div class="text-[11px] font-mono text-slate-500">User: ${p.username} • Pass: ${p.password}</div>
                    </div>
                    <button onclick="deleteProctor('${p.id}')" class="p-1.5 rounded-lg text-red-600 hover:bg-red-100 transition" title="Hapus Pengawas">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                    </button>
                </div>
            `).join('');
        }

        function saveSchoolIdentity() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const newAdminUser = document.getElementById('setAdminUsername').value.trim();
            const newAdminPass = document.getElementById('setAdminPassword').value.trim();
            if (!newAdminUser || !newAdminPass) {
                showToast('Username dan Password Administrator tidak boleh kosong!', 'error');
                return;
            }
            settings.schoolName = document.getElementById('setSchoolName').value.trim() || DEFAULT_SETTINGS.schoolName;
            settings.schoolSubtitle = document.getElementById('setSchoolSubtitle').value.trim() || DEFAULT_SETTINGS.schoolSubtitle;
            settings.adminUsername = newAdminUser;
            settings.adminPassword = newAdminPass;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast('Identitas sekolah & akun Administrator berhasil diperbarui!', 'success');
        }

        function saveSecuritySettings() {
            const pin = document.getElementById('setUnlockPin').value.trim();
            if (!pin) {
                showToast('Kode PIN pembuka kunci tidak boleh kosong!', 'error');
                return;
            }
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.unlockPin = pin;
            settings.requireFullscreen = document.getElementById('setRequireFullscreen').checked;
            settings.detectTabSwitch = document.getElementById('setDetectTabSwitch').checked;
            settings.showScoreToStudent = document.getElementById('setShowScoreToStudent').checked;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast('Setelan keamanan & perilaku siswa disimpan!', 'success');
        }

        function handleAddProctor(e) {
            e.preventDefault();
            const proctors = getDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
            const username = document.getElementById('procUser').value.trim();
            if (proctors.some(p => p.username === username)) {
                showToast('Username pengawas sudah terdaftar!', 'error');
                return;
            }
            proctors.push({
                id: 'proc_' + Date.now(),
                name: document.getElementById('procName').value.trim(),
                room: document.getElementById('procRoom').value.trim(),
                username,
                password: document.getElementById('procPass').value.trim()
            });
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            e.target.reset();
            syncAllViews();
            showToast('Akun pengawas baru berhasil ditambahkan!', 'success');
        }

        function deleteProctor(procId) {
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []).filter(p => p.id !== procId);
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            syncAllViews();
            showToast('Akun pengawas dihapus.', 'info');
        }

        function handleAddGalleryItem(e) {
            e.preventDefault();
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            gallery.unshift({
                id: 'gal_' + Date.now(),
                title: document.getElementById('galTitle').value.trim(),
                url: document.getElementById('galUrl').value.trim(),
                svgType: 'lab'
            });
            setDB(STORAGE_KEYS.GALLERY, gallery);
            e.target.reset();
            syncAllViews();
            showToast('Foto kegiatan berhasil ditambahkan ke Portal!', 'success');
        }

        function exportDatabaseJSON() {
            const dump = {
                settings: getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS),
                questions: getDB(STORAGE_KEYS.QUESTIONS, []),
                students: getDB(STORAGE_KEYS.STUDENTS, []),
                proctors: getDB(STORAGE_KEYS.PROCTORS, []),
                liveStatus: getDB(STORAGE_KEYS.LIVE_STATUS, {}),
                gallery: getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY)
            };
            const blob = new Blob([JSON.stringify(dump, null, 2)], { type: 'application/json' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `Backup_CBT_PRO_${new Date().toISOString().slice(0, 10)}.json`;
            a.click();
            showToast('Backup database JSON berhasil diunduh!', 'success');
        }

        async function importDatabaseJSON(event) {
            const file = event.target.files[0];
            if (!file) return;
            try {
                const text = await file.text();
                const data = JSON.parse(text);
                if (data.settings) setDB(STORAGE_KEYS.SETTINGS, data.settings);
                if (data.questions) setDB(STORAGE_KEYS.QUESTIONS, data.questions);
                if (data.students) setDB(STORAGE_KEYS.STUDENTS, data.students);
                if (data.proctors) setDB(STORAGE_KEYS.PROCTORS, data.proctors);
                if (data.liveStatus) setDB(STORAGE_KEYS.LIVE_STATUS, data.liveStatus);
                if (data.gallery) setDB(STORAGE_KEYS.GALLERY, data.gallery);
                syncAllViews();
                showToast('Database berhasil dipulihkan dari file backup!', 'success');
            } catch (err) {
                showToast('File JSON backup tidak valid.', 'error');
            } finally {
                event.target.value = '';
            }
        }

        function confirmCleanStartEmpty() {
            showConfirmDialog('Bersihkan Data Awal (Mulai Kosong)?', 'Seluruh data awal siswa, bank soal, dan riwayat nilai akan dikosongkan agar Anda dapat mengisi data peserta dan soal sekolah yang sebenarnya.', () => {
                setDB(STORAGE_KEYS.QUESTIONS, []);
                setDB(STORAGE_KEYS.STUDENTS, []);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                syncAllViews();
                showToast('Data awal dibersihkan! Sistem siap diisi data sekolah Anda.', 'success');
            });
        }

        function confirmFactoryReset() {
            showConfirmDialog('Reset Seluruh Data ke Bawaan?', 'Semua soal, peserta, pengawas, dan nilai saat ini akan dikembalikan ke data awal sistem.', () => {
                setDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
                setDB(STORAGE_KEYS.QUESTIONS, DEFAULT_QUESTIONS);
                setDB(STORAGE_KEYS.STUDENTS, DEFAULT_STUDENTS);
                setDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                setDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
                syncAllViews();
                showToast('Seluruh data berhasil direset ke kondisi bawaan.', 'success');
            });
        }

        // =================================================================
        // H. PORTAL PENGAWAS RUANG
        // =================================================================
        function renderProctorMonitoring() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const assignedRoom = activeSession?.user?.room || 'Semua Ruangan';

            const filtered = students.filter(s =>
                assignedRoom.toLowerCase().includes('semua') || s.room.toLowerCase() === assignedRoom.toLowerCase()
            );

            const container = document.getElementById('proctorStudentCards');
            if (filtered.length === 0) {
                container.innerHTML = `<div class="text-xs text-slate-400 py-6 text-center">Belum ada peserta pada ${assignedRoom}.</div>`;
                return;
            }

            container.innerHTML = filtered.map(s => {
                const st = liveMap[s.id] || { status: 'Belum Masuk', progress: '0/0', violations: 0 };
                let badge = `<span class="px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-xs font-bold">Belum Masuk</span>`;
                if (st.status === 'Mengerjakan') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold animate-pulse">Mengerjakan (${st.progress})</span>`;
                } else if (st.status === 'TERKUNCI') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-red-600 text-white text-xs font-bold animate-bounce">TERKUNCI (${st.violations}x)</span>`;
                } else if (st.status === 'Selesai') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold">Selesai</span>`;
                }

                return `
                    <div class="p-4 rounded-2xl border ${st.status === 'TERKUNCI' ? 'border-red-400 bg-red-50/70' : 'border-slate-200 bg-slate-50/50'} flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div>
                            <div class="font-bold text-slate-900 text-sm">${s.name} <span class="font-mono text-xs font-normal text-slate-500">(${s.nisn})</span></div>
                            <div class="text-xs text-slate-500">${s.className} • <strong>${s.room}</strong></div>
                        </div>
                        <div class="flex items-center gap-2 flex-wrap">
                            ${badge}
                            ${st.status === 'TERKUNCI' ? `
                                <button onclick="remoteUnlockStudentById('${s.id}')" class="px-3 py-1.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition shadow-sm">
                                    Buka Kunci Layar
                                </button>
                            ` : ''}
                            <button onclick="resetStudentExamSession('${s.id}')" class="px-2.5 py-1.5 rounded-xl bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 text-xs font-bold transition">
                                Reset Sesi
                            </button>
                        </div>
                    </div>
                `;
            }).join('');
        }

        // =================================================================
        // I. MESIN UJIAN SISWA (ACAK SOAL & OPSI) + ANTI-CHEAT LOCKDOWN
        // =================================================================
        function shuffleArrayItems(list) {
            const arr = [...list];
            for (let i = arr.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [arr[i], arr[j]] = [arr[j], arr[i]];
            }
            return arr;
        }

        function pushStudentLiveState(status, extra = {}) {
            if (!activeSession || activeSession.role !== 'siswa') return;
            const s = activeSession.user;
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            liveMap[s.id] = {
                ...(liveMap[s.id] || {}),
                status,
                progress: `${Object.keys(examSession.answers).length}/${examSession.questions.length}`,
                violations: examSession.violations,
                updatedAt: new Date().toLocaleTimeString('id-ID'),
                ...extra
            };
            setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
        }

        function beginStudentExam(student) {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            let rawQuestions = getDB(STORAGE_KEYS.QUESTIONS, []);

            if (settings.shuffleQuestions) {
                rawQuestions = shuffleArrayItems(rawQuestions);
            }

            const preparedQuestions = rawQuestions.map(q => {
                const origKeys = Object.keys(q.options);
                const displayOrder = settings.shuffleOptions ? shuffleArrayItems(origKeys) : origKeys;
                return {
                    ...q,
                    displayOrder
                };
            });

            examSession = {
                isRunning: true,
                isLocked: false,
                questions: preparedQuestions,
                answers: {},
                doubts: {},
                currentIdx: 0,
                remainingSeconds: settings.durationMinutes * 60,
                timerId: null,
                violations: 0
            };

            document.getElementById('studentExamTitle').innerText = `${settings.examTitle} (${settings.subjectName})`;
            document.getElementById('studentExamName').innerText = student.name;
            document.getElementById('studentExamNisn').innerText = student.nisn;
            document.getElementById('studentExamRoom').innerText = `${student.className} • ${student.room}`;
            document.getElementById('examTotalNum').innerText = preparedQuestions.length;

            navigateTo('siswa');
            renderExamQuestion();
            startExamCountdown();
            pushStudentLiveState('Mengerjakan', { remoteUnlocked: false });

            if (settings.requireFullscreen && document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }

            showToast('Ujian dimulai! Sistem proteksi anti-curang aktif.', 'warning');
        }

        function startExamCountdown() {
            if (examSession.timerId) clearInterval(examSession.timerId);
            const renderClock = () => {
                const m = Math.floor(examSession.remainingSeconds / 60);
                const s = examSession.remainingSeconds % 60;
                document.getElementById('studentTimerDisplay').innerText = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            };
            renderClock();
            examSession.timerId = setInterval(() => {
                if (examSession.isLocked) return;
                examSession.remainingSeconds--;
                renderClock();
                if (examSession.remainingSeconds <= 0) {
                    clearInterval(examSession.timerId);
                    showToast('Waktu ujian telah habis! Jawaban dikirim otomatis.', 'warning');
                    finalizeStudentExam();
                }
            }, 1000);
        }

        function changeExamFontSize(size) {
            const body = document.getElementById('examQuestionBody');
            body.classList.remove('text-sm', 'text-base', 'text-lg', 'text-xl');
            if (size === 'sm') body.classList.add('text-sm');
            else if (size === 'lg') body.classList.add('text-lg', 'sm:text-xl');
            else body.classList.add('text-base');
        }

        function renderExamQuestion() {
            const q = examSession.questions[examSession.currentIdx];
            if (!q) return;

            document.getElementById('examCurrentNum').innerText = examSession.currentIdx + 1;
            document.getElementById('examQuestionBody').innerText = q.question;

            const chosenOrigKey = examSession.answers[q.id];
            const labels = ['A', 'B', 'C', 'D', 'E'];
            const optionsContainer = document.getElementById('examOptionsList');

            optionsContainer.innerHTML = q.displayOrder.map((origKey, displayIdx) => {
                const displayLetter = labels[displayIdx] || origKey;
                const isSelected = chosenOrigKey === origKey;
                return `
                    <button type="button" onclick="pickExamOption('${q.id}', '${origKey}')"
                        class="w-full text-left p-4 rounded-2xl border-2 transition flex items-center gap-3.5 ${isSelected ? 'border-blue-600 bg-blue-50 shadow-sm' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}">
                        <span class="w-8 h-8 rounded-xl flex items-center justify-center font-extrabold text-sm shrink-0 ${isSelected ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-700'}">
                            ${displayLetter}
                        </span>
                        <span class="text-sm sm:text-base ${isSelected ? 'font-bold text-blue-950' : 'font-medium text-slate-700'}">
                            ${q.options[origKey]}
                        </span>
                    </button>
                `;
            }).join('');

            const doubtBtn = document.getElementById('doubtToggleBtn');
            const isDoubt = !!examSession.doubts[q.id];
            doubtBtn.className = isDoubt
                ? 'px-3.5 py-1.5 rounded-xl border-2 border-amber-500 bg-amber-400 text-slate-950 text-xs font-extrabold transition flex items-center gap-1.5'
                : 'px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5';

            document.getElementById('btnPrevQuestion').disabled = examSession.currentIdx === 0;
            const nextBtn = document.getElementById('btnNextQuestion');
            if (examSession.currentIdx === examSession.questions.length - 1) {
                nextBtn.innerHTML = `Selesai Ujian <i data-lucide="check-circle" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = confirmSubmitExam;
            } else {
                nextBtn.innerHTML = `Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = goNextQuestion;
            }

            renderExamNumberGrid();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function renderExamNumberGrid() {
            const grid = document.getElementById('examNumberGrid');
            grid.innerHTML = examSession.questions.map((q, idx) => {
                const answered = !!examSession.answers[q.id];
                const doubt = !!examSession.doubts[q.id];
                const current = idx === examSession.currentIdx;

                let cls = 'bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200';
                if (doubt) cls = 'bg-amber-400 text-slate-950 border-amber-500';
                else if (answered) cls = 'bg-blue-600 text-white border-blue-600';
                if (current) cls += ' ring-2 ring-offset-2 ring-slate-900';

                return `<button onclick="jumpExamQuestion(${idx})" class="h-10 rounded-xl border font-bold text-xs transition ${cls}">${idx + 1}</button>`;
            }).join('');
        }

        function pickExamOption(qId, origKey) {
            examSession.answers[qId] = origKey;
            pushStudentLiveState('Mengerjakan');
            renderExamQuestion();
        }

        function toggleDoubtQuestion() {
            const q = examSession.questions[examSession.currentIdx];
            if (!q) return;
            examSession.doubts[q.id] = !examSession.doubts[q.id];
            renderExamQuestion();
        }

        function goPrevQuestion() {
            if (examSession.currentIdx > 0) {
                examSession.currentIdx--;
                renderExamQuestion();
            }
        }

        function goNextQuestion() {
            if (examSession.currentIdx < examSession.questions.length - 1) {
                examSession.currentIdx++;
                renderExamQuestion();
            }
        }

        function jumpExamQuestion(idx) {
            examSession.currentIdx = idx;
            renderExamQuestion();
        }

        function confirmSubmitExam() {
            const answered = Object.keys(examSession.answers).length;
            const total = examSession.questions.length;
            const msg = answered < total
                ? `Anda baru menjawab ${answered} dari ${total} soal. Yakin ingin mengakhiri ujian dan mengirim jawaban?`
                : `Semua soal (${total}) sudah terjawab. Yakin ingin mengumpulkan jawaban sekarang?`;

            showConfirmDialog('Kumpulkan Jawaban Ujian?', msg, () => {
                finalizeStudentExam();
            });
        }

        function finalizeStudentExam() {
            if (!examSession.isRunning) return;
            examSession.isRunning = false;
            examSession.isLocked = false;
            if (examSession.timerId) clearInterval(examSession.timerId);

            let correct = 0;
            let earnedPoints = 0;
            let maxPoints = 0;

            examSession.questions.forEach(q => {
                const pts = q.points || 10;
                maxPoints += pts;
                if (examSession.answers[q.id] === q.answer) {
                    correct++;
                    earnedPoints += pts;
                }
            });

            const total = examSession.questions.length;
            const wrong = total - correct;
            const finalScore = maxPoints > 0 ? Math.round((earnedPoints / maxPoints) * 100) : 0;

            pushStudentLiveState('Selesai', {
                correct,
                wrong,
                score: finalScore,
                finishedAt: new Date().toLocaleString('id-ID')
            });

            if (document.fullscreenElement && document.exitFullscreen) {
                document.exitFullscreen().catch(() => {});
            }

            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const scoreBox = document.getElementById('studentScoreSummaryBox');
            if (settings.showScoreToStudent) {
                scoreBox.classList.remove('hidden');
                document.getElementById('studentFinalScoreText').innerText = finalScore;
                document.getElementById('studentFinalDetailText').innerText = `Benar ${correct} • Salah ${wrong} dari ${total} Soal`;
            } else {
                scoreBox.classList.add('hidden');
            }

            activeSession = null;
            navigateTo('portal');

            const finModal = document.getElementById('examFinishedModal');
            finModal.classList.remove('hidden');
            finModal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeExamFinishedModal() {
            const finModal = document.getElementById('examFinishedModal');
            finModal.classList.add('hidden');
            finModal.classList.remove('flex');
        }

        // --- SISTEM ANTI-CHEAT / LOCKDOWN OTOMATIS ---
        function activateExamLockdown(reason) {
            if (!examSession.isRunning || examSession.isLocked) return;
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (!settings.detectTabSwitch) return;

            examSession.isLocked = true;
            examSession.violations++;

            document.getElementById('lockReasonText').innerText = reason;
            document.getElementById('lockStudentName').innerText = activeSession?.user?.name || '-';
            document.getElementById('lockStudentMeta').innerText = `${activeSession?.user?.nisn || '-'} • ${activeSession?.user?.className || '-'}`;
            document.getElementById('lockViolationBadge').innerText = `${examSession.violations} Kali`;
            document.getElementById('unlockPinField').value = '';

            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.remove('hidden');
            overlay.classList.add('flex');

            pushStudentLiveState('TERKUNCI');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function verifyPinUnlock() {
            const pin = document.getElementById('unlockPinField').value.trim();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);

            if (pin === settings.unlockPin) {
                releaseExamLockdown();
                showToast('PIN Pengawas valid! Layar ujian dibuka kembali.', 'success');
            } else {
                showToast('Kode PIN salah! Hubungi pengawas ruang.', 'error');
                document.getElementById('unlockPinField').value = '';
            }
        }

        function releaseExamLockdown() {
            examSession.isLocked = false;
            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.add('hidden');
            overlay.classList.remove('flex');
            pushStudentLiveState('Mengerjakan', { remoteUnlocked: false });

            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (settings.requireFullscreen && document.documentElement.requestFullscreen && !document.fullscreenElement) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
        }

        document.addEventListener('visibilitychange', () => {
            if (document.hidden && examSession.isRunning) {
                activateExamLockdown('Terdeteksi berpindah tab browser atau membuka aplikasi lain saat ujian berlangsung!');
            }
        });

        window.addEventListener('blur', () => {
            if (examSession.isRunning) {
                activateExamLockdown('Layar ujian kehilangan fokus! Dilarang membuka jendela lain, split-screen, atau berbagi layar.');
            }
        });

        document.addEventListener('fullscreenchange', () => {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (!document.fullscreenElement && examSession.isRunning && !examSession.isLocked && settings.requireFullscreen) {
                activateExamLockdown('Anda keluar dari mode Layar Penuh (Fullscreen) tanpa izin!');
            }
        });

        document.addEventListener('contextmenu', (e) => {
            if (examSession.isRunning) {
                e.preventDefault();
                showToast('Klik kanan tidak diizinkan selama ujian!', 'warning');
            }
        });

        document.addEventListener('keydown', (e) => {
            if (!examSession.isRunning) return;
            if (
                e.key === 'F12' ||
                e.key === 'PrintScreen' ||
                (e.ctrlKey && e.shiftKey && ['I', 'i', 'J', 'j', 'C', 'c'].includes(e.key)) ||
                (e.ctrlKey && ['U', 'u', 'S', 's', 'P', 'p'].includes(e.key))
            ) {
                e.preventDefault();
                activateExamLockdown('Terdeteksi menekan tombol shortcut terlarang / Developer Tools!');
            }
        });

        // =================================================================
        // J. SINKRONISASI REAL-TIME ANTAR TAB & BOOTSTRAP APLIKASI
        // =================================================================
        window.addEventListener('storage', (e) => {
            if (Object.values(STORAGE_KEYS).includes(e.key)) {
                syncAllViews();

                if (e.key === STORAGE_KEYS.LIVE_STATUS && examSession.isRunning && examSession.isLocked && activeSession?.role === 'siswa') {
                    const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                    const myState = liveMap[activeSession.user.id];
                    if (myState && myState.remoteUnlocked) {
                        releaseExamLockdown();
                        showToast('Pengawas telah membuka kunci layar Anda dari jarak jauh!', 'success');
                    }
                }
            }
        });

        setInterval(() => {
            if (!examSession.isRunning) {
                syncAllViews();
            }
        }, 3000);

        window.addEventListener('DOMContentLoaded', () => {
            initializeAppDatabase();
            syncAllViews();
            if (typeof lucide !== 'undefined') lucide.createIcons();

            const params = new URLSearchParams(window.location.search);
            const roleParam = params.get('role') || params.get('portal');
            if (['admin', 'pengawas', 'siswa'].includes(roleParam)) {
                openLoginModal(roleParam);
            }
        });
    </script>
</body>
</html>
