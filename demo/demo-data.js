// Data contoh AgendaGuru (dibuat dari DataContoh.gs). Mode demo: tanpa server.
window.AGENDA_DEMO = true;
function buatDataContoh_() {
  // Acak yang selalu sama hasilnya (agar data contoh konsisten)
  let seed = 20261028;
  const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  const pick = arr => arr[Math.floor(rnd() * arr.length)];
  const pad = n => String(n).padStart(2, '0');
  const fmt = dt => dt.getFullYear() + '-' + pad(dt.getMonth() + 1) + '-' + pad(dt.getDate());

  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth() + 1;
  const tapel = m >= 7 ? y + '/' + (y + 1) : (y - 1) + '/' + y;
  const tahunAwal = Number(tapel.slice(0, 4));
  // Bulan semester ganjil yang sudah berjalan (Juli s.d. bulan ini; jika sedang genap, Juli–Desember)
  const bulanGanjil = [7, 8, 9, 10, 11, 12].filter(b => (m >= 7 ? b <= m : true));

  const subjects = ['Matematika', 'Bahasa Indonesia', 'IPAS', 'Pendidikan Pancasila', 'Bahasa Inggris', 'PJOK'];
  const classes = [{ ClassID: 'C1', ClassName: 'Kelas 5A' }, { ClassID: 'C2', ClassName: 'Kelas 5B' }];
  const nama = {
    C1: ['Ahmad Fauzan', 'Aisyah Putri Ramadhani', 'Bagas Pratama', 'Citra Lestari', 'Dimas Saputra', 'Dinda Ayu Kirana',
      'Fajar Nugroho', 'Gita Permata Sari', 'Hafiz Maulana', 'Intan Nuraini', 'Joko Susilo', 'Kayla Azzahra'],
    C2: ['Laras Wulandari', 'Muhammad Rizky', 'Nabila Khairunnisa', 'Oki Setiawan', 'Putri Melati', 'Raka Aditya',
      'Salsabila Zahra', 'Taufik Hidayat', 'Ulfa Rahmawati', 'Vino Ardiansyah', 'Wulan Sari', 'Yusuf Habibi']
  };
  const kota = ['Bandung', 'Cimahi', 'Sumedang', 'Garut', 'Bekasi'];
  const agama = ['Islam', 'Islam', 'Islam', 'Islam', 'Kristen', 'Katolik', 'Hindu'];
  const ayah = ['Budi', 'Agus', 'Hendra', 'Joko', 'Rudi', 'Eko', 'Dedi', 'Yanto', 'Sigit', 'Wahyu'];
  const ibu = ['Siti', 'Sri', 'Dewi', 'Rina', 'Lilis', 'Nur', 'Ani', 'Yuli', 'Endang', 'Ratna'];

  const students = [], kemampuan = {};
  let no = 1;
  classes.forEach(c => {
    nama[c.ClassID].forEach((n, i) => {
      const sid = 'S' + c.ClassID + pad(i + 1);
      const perempuan = /Aisyah|Citra|Dinda|Gita|Intan|Kayla|Laras|Nabila|Putri|Salsabila|Ulfa|Wulan/.test(n);
      students.push({
        StudentID: sid, ClassID: c.ClassID, StudentName: n,
        NIS: String(2400 + no), NISN: '01' + String(5000000 + no * 137).padStart(8, '0'),
        TempatLahir: pick(kota), TanggalLahir: (2015 + (i % 2)) + '-' + pad(1 + (i * 5) % 12) + '-' + pad(3 + (i * 7) % 25),
        JenisKelamin: perempuan ? 'P' : 'L', Agama: pick(agama),
        Alamat: 'Jl. ' + pick(['Melati', 'Kenanga', 'Mawar', 'Anggrek', 'Flamboyan']) + ' No. ' + (3 + i * 2) + ', ' + pick(kota),
        NamaAyah: pick(ayah) + ' ' + n.split(' ').slice(-1)[0], NamaIbu: pick(ibu) + ' ' + pick(['Rahayu', 'Wati', 'Lestari', 'Handayani']),
        NamaWali: '', ParentPhone: '0812' + String(30000000 + no * 4271).slice(0, 8)
      });
      // kemampuan dasar 55–95 agar ada siswa tuntas & belum tuntas
      kemampuan[sid] = 55 + Math.round(rnd() * 40);
      no++;
    });
  });

  const nilai = (sid, rentang) => String(Math.max(40, Math.min(100, Math.round(kemampuan[sid] + (rnd() - 0.5) * (rentang || 16)))));
  const hariSekolah = (bulan, jumlah) => {
    const tahun = bulan >= 7 ? tahunAwal : tahunAwal + 1;
    const out = []; let dt = new Date(tahun, bulan - 1, 1);
    while (out.length < jumlah && dt.getMonth() === bulan - 1) {
      const dow = dt.getDay();
      const lewat = dt > now;
      if (dow >= 1 && dow <= 5 && !lewat) out.push(fmt(dt));
      dt = new Date(dt.getTime() + 864e5 * (out.length % 2 ? 2 : 3));
    }
    return out;
  };

  // --- Absensi: 8 pertemuan per bulan ---
  const attendance = [];
  classes.forEach(c => bulanGanjil.forEach(b => {
    const dates = hariSekolah(b, 8);
    if (!dates.length) return;
    const data = {};
    students.filter(s => s.ClassID === c.ClassID).forEach(s => {
      data[s.StudentID] = dates.map(() => { const r = rnd(); return r < 0.88 ? 'H' : r < 0.94 ? 'S' : r < 0.97 ? 'I' : 'A'; });
    });
    attendance.push({ RecordID: 'A' + c.ClassID + pad(b), ClassID: c.ClassID, MonthYear: pad(b), DatesData: JSON.stringify(dates),
      AttendanceData: JSON.stringify(data), TahunPelajaran: tapel });
  }));

  // --- Nilai: formatif 2x per bulan, PTS (SumatifLM) & PAS (SumatifAS) semester ganjil ---
  const materi = {
    'Matematika': ['Bilangan cacah', 'Pecahan senilai', 'Operasi pecahan', 'Pengukuran sudut', 'Bangun datar', 'Data dan diagram'],
    'Bahasa Indonesia': ['Teks narasi', 'Ide pokok paragraf', 'Puisi anak', 'Teks prosedur', 'Wawancara', 'Surat pribadi'],
    'IPAS': ['Organ tubuh', 'Ekosistem', 'Cahaya dan bunyi', 'Energi', 'Keragaman budaya', 'Peta dan wilayah'],
    'Pendidikan Pancasila': ['Makna sila Pancasila', 'Hak dan kewajiban', 'Norma di masyarakat', 'Gotong royong', 'Keberagaman', 'Musyawarah'],
    'Bahasa Inggris': ['Greetings', 'My family', 'Daily activities', 'Animals', 'Shopping', 'Telling time'],
    'PJOK': ['Gerak dasar lari', 'Lompat jauh', 'Senam lantai', 'Bola besar', 'Kebugaran', 'Hidup sehat']
  };
  const grades = [];
  classes.forEach(c => {
    const stds = students.filter(s => s.ClassID === c.ClassID);
    subjects.forEach((subj, si) => {
      bulanGanjil.forEach((b, bi) => {
        const dates = hariSekolah(b, 6).filter((_, k) => k === 1 || k === 4);
        if (!dates.length) return;
        const gd = {}; stds.forEach(s => { gd[s.StudentID] = dates.map(() => nilai(s.StudentID)); });
        const cak = dates.map((_, k) => materi[subj][(bi * 2 + k) % materi[subj].length]);
        grades.push({ RecordID: 'G' + c.ClassID + si + pad(b), ClassID: c.ClassID, Subject: subj, Type: 'Formatif', MonthYear: pad(b),
          DatesData: JSON.stringify(dates), Cakupan: JSON.stringify(cak), GradesData: JSON.stringify(gd), Semester: '', TahunPelajaran: tapel });
      });
      ['SumatifLM', 'SumatifAS'].forEach(type => {
        const gd = {}; stds.forEach(s => { gd[s.StudentID] = [nilai(s.StudentID, 20)]; });
        grades.push({ RecordID: 'G' + c.ClassID + si + type, ClassID: c.ClassID, Subject: subj, Type: type, MonthYear: '12',
          DatesData: '[]', Cakupan: '[]', GradesData: JSON.stringify(gd), Semester: 'ganjil', TahunPelajaran: tapel });
      });
    });
  });

  // --- Kokurikuler: 1 projek per kelas ---
  const dims = [['Kewargaan', 'Kolaborasi'], ['Kreativitas', 'Kemandirian']];
  const proj = [
    { p: 'Gaya Hidup Berkelanjutan', desc: 'Siswa memilah sampah organik dan anorganik serta membuat kompos sederhana di sekolah.' },
    { p: 'Kearifan Lokal', desc: 'Siswa mengenal permainan tradisional daerah dan mempresentasikannya kepada teman.' }
  ];
  const saran = { 'Sangat Baik': 'Terus pertahankan inisiatif dan kepemimpinan ini.', 'Baik': 'Kembangkan terus potensi dan keaktifan ini.',
    'Cukup': 'Perlu ditingkatkan fokus dan konsistensi belajarnya.', 'Kurang': 'Sangat disarankan lebih aktif bertanya dan berkolaborasi.' };
  const koku = classes.map((c, ci) => {
    const data = {};
    students.filter(s => s.ClassID === c.ClassID).forEach(s => {
      const k = kemampuan[s.StudentID];
      const cat = k >= 85 ? 'Sangat Baik' : k >= 72 ? 'Baik' : k >= 62 ? 'Cukup' : 'Kurang';
      data[s.StudentID] = { Kategori: cat, Narasi: 'Ananda <strong>' + s.StudentName + '</strong> menunjukkan profil <strong>' + dims[ci].join(' dan ') +
        '</strong> yang <strong>' + cat.toLowerCase() + '</strong> dalam kegiatan ' + proj[ci].p + '. ' + proj[ci].desc + ' ' + saran[cat] };
    });
    return { RecordID: 'K' + c.ClassID, ClassID: c.ClassID, Tanggal: tahunAwal + '-09-1' + (5 + ci), Jenis: 'LintasMapel',
      Projek: proj[ci].p, Deskripsi: proj[ci].desc, Dimensi: JSON.stringify(dims[ci]), KokuData: JSON.stringify(data) };
  });

  // --- Ekstrakurikuler ---
  const ekskul = [];
  students.forEach((s, i) => {
    ekskul.push({ EkskulID: 'E' + s.StudentID + 'P', ClassID: s.ClassID, StudentID: s.StudentID, StudentName: s.StudentName,
      Semester: 'ganjil', TahunPelajaran: tapel, NamaEkskul: 'Pramuka', Predikat: kemampuan[s.StudentID] >= 75 ? 'Baik' : 'Cukup',
      Keterangan: 'Aktif mengikuti latihan rutin Pramuka Siaga.' });
    if (i % 3 === 0) ekskul.push({ EkskulID: 'E' + s.StudentID + 'X', ClassID: s.ClassID, StudentID: s.StudentID, StudentName: s.StudentName,
      Semester: 'ganjil', TahunPelajaran: tapel, NamaEkskul: i % 2 ? 'Seni Tari' : 'Futsal', Predikat: 'Sangat Baik',
      Keterangan: 'Menunjukkan bakat dan semangat yang tinggi.' });
  });

  // --- Jadwal pelajaran (Senin–Jumat) ---
  const hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'];
  const slot = [['07:30', '08:40'], ['08:40', '09:50'], ['10:10', '11:20']];
  const schedule = [];
  classes.forEach((c, ci) => hari.forEach((h, hi) => slot.forEach((sl, k) => {
    schedule.push({ RecordID: 'SCH' + c.ClassID + hi + k, Day: h, TimeStart: sl[0], TimeEnd: sl[1],
      Subject: subjects[(hi * 3 + k + ci) % subjects.length], ClassID: c.ClassID });
  })));

  // --- Jurnal harian (10 hari sekolah terakhir) ---
  const journal = [];
  let dt = new Date(now.getTime()); let j = 0;
  while (journal.length < 10 && j < 30) {
    dt = new Date(dt.getTime() - 864e5); j++;
    if (dt.getDay() === 0 || dt.getDay() === 6) continue;
    const c = classes[journal.length % 2], subj = subjects[journal.length % subjects.length];
    journal.push({ RecordID: 'J' + fmt(dt).replace(/-/g, '') + c.ClassID, Date: fmt(dt), ClassID: c.ClassID, Subject: subj,
      Content: 'Pembelajaran ' + materi[subj][journal.length % 6].toLowerCase() + ': diskusi kelompok, latihan soal, dan refleksi. Siswa antusias mengikuti kegiatan.',
      AttachmentUrl: '' });
  }

  // --- Poin siswa ---
  const poinKet = [['Juara 1 lomba cerdas cermat', 15, 0], ['Membantu teman yang kesulitan', 5, 0], ['Terlambat masuk kelas', 0, 5],
    ['Tidak mengerjakan PR', 0, 3], ['Menjadi petugas upacara', 10, 0], ['Menjaga kebersihan kelas', 5, 0]];
  const poin = [];
  students.filter((_, i) => i % 2 === 0).forEach((s, i) => {
    const pk = poinKet[i % poinKet.length], tgl = new Date(now.getTime() - (i + 2) * 3 * 864e5);
    poin.push({ RecordID: 'P' + s.StudentID, StudentID: s.StudentID, ClassID: s.ClassID, Tanggal: fmt(tgl),
      PoinReward: String(pk[1]), PoinPelanggaran: String(pk[2]), Keterangan: pk[0] });
  });

  // --- Catatan wali kelas (rapor semester ganjil) ---
  const raporNotes = students.map(s => ({
    RecordID: 'RN' + s.StudentID, StudentID: s.StudentID, ClassID: s.ClassID, Semester: 'ganjil', TahunPelajaran: tapel,
    CatatanWali: kemampuan[s.StudentID] >= 80 ? 'Ananda sangat aktif dan bertanggung jawab. Pertahankan prestasimu!'
      : kemampuan[s.StudentID] >= 68 ? 'Ananda cukup baik dalam belajar. Tingkatkan lagi ketelitian dan keaktifan di kelas.'
        : 'Ananda perlu pendampingan belajar di rumah, terutama membaca dan berhitung. Semangat terus!',
    Keputusan: ''
  }));

  // --- Contoh hasil Pengolah Nilai: Matematika Kelas 5A (bobot 40/30/30, nilai < KKM disesuaikan ke 75–80) ---
  const nilaiAkhir = [];
  {
    const cid = 'C1', subj = 'Matematika', kkm = 75;
    const rows = students.filter(s => s.ClassID === cid).map(s => {
      const f = grades.filter(g => g.ClassID === cid && g.Subject === subj && g.Type === 'Formatif')
        .reduce((a, g) => a.concat(JSON.parse(g.GradesData)[s.StudentID].map(Number)), []);
      const harian = f.length ? f.reduce((a, b) => a + b, 0) / f.length : null;
      const one = t => Number(JSON.parse(grades.find(g => g.ClassID === cid && g.Subject === subj && g.Type === t).GradesData)[s.StudentID][0]);
      const pts = one('SumatifLM'), pas = one('SumatifAS');
      // Sama persis dengan perhitungan menu Pengolah Nilai (tanpa pembulatan sebelum nilai akhir)
      const mentah = harian * 0.4 + pts * 0.3 + pas * 0.3;
      return { sid: s.StudentID, harian, pts, pas, mentah };
    });
    const below = rows.filter(r => r.mentah < kkm).map(r => r.mentah);
    const lo = below.length ? Math.min.apply(null, below) : 0, hi = below.length ? Math.max.apply(null, below) : 0;
    const data = {};
    rows.forEach(r => {
      let akhir = r.mentah >= kkm ? r.mentah : (hi > lo ? 75 + (r.mentah - lo) / (hi - lo) * 5 : 75);
      akhir = Math.round(Math.min(100, Math.max(0, akhir)));
      data[r.sid] = { harian: Math.round(r.harian * 10) / 10, pts: r.pts, pas: r.pas, hadir: '', mentah: Math.round(r.mentah * 10) / 10, akhir: akhir, manual: false };
    });
    nilaiAkhir.push({ RecordID: 'NA_CONTOH', ClassID: cid, Subject: subj, Semester: 'ganjil', TahunPelajaran: tapel,
      Pengaturan: JSON.stringify({ bobot: { harian: 40, pts: 30, pas: 30, hadir: 0 }, kkm: kkm, bulat: true, metode: 'kkm', rentangMin: 75, rentangMax: 80 }),
      NilaiData: JSON.stringify(data), UpdatedAt: new Date().toISOString() });
  }

  const config = {
    userSchool: 'SD Negeri Contoh 1', schoolAddress: 'Jl. Pendidikan No. 10, Bandung', schoolCity: 'Bandung',
    headmasterName: 'Dra. Siti Rahmawati, M.Pd.', headmasterNIP: '197003121995032001', raporFase: 'C'
  };

  return { tapel, classes, students, attendance, grades, koku, ekskul, schedule, journal, poin, raporNotes, nilaiAkhir, subjects, config };
}
