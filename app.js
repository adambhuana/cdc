/* ===== CDC Dashboard - Main Application ===== */
(function () {
    'use strict';

    // ===== State =====
    let mahasiswaData = [];
    let magangData = [];
    let belumMagangData = [];
    let reportData = [];
    let mergedData = [];
    let filteredData = [];
    let filteredMagangData = [];
    let currentPage = 1;
    let magangPage = 1;
    let selectedAngkatan = 'all';
    let selectedSemester = 'all';
    let selectedKelas = 'all';
    let currentSection = 'dashboard';
    const PAGE_SIZE = 15;
    const MAGANG_PAGE_SIZE = 20;

    // Kelas mapping
    const KELAS_MAP = {
        'Reg': 'Reguler',
        'Pro': 'Profesional',
        'Aksel': 'Akselerasi'
    };

    // Chart instances for destroy/recreate
    let barChartInstance = null;
    let doughnutChartInstance = null;
    let companyChartInstance = null;
    let positionChartInstance = null;

    // ===== DOM Refs =====
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    // ===== Init =====
    document.addEventListener('DOMContentLoaded', init);

    async function init() {
        try {
            await loadData();
            processData();
            buildAngkatanChips();
            buildAngkatanSelect();
            buildKelasChips();
            buildKelasSelect();
            buildMagangSelects();
            renderDashboard();
            renderTable();
            renderMagangTable();
            renderPositionSection();
            setupEvents();
            showApp();
        } catch (err) {
            console.error('Init error:', err);
            $('.loader-text').textContent = 'Gagal memuat data. Pastikan file Excel tersedia.';
            $('.loader-ring').style.borderTopColor = '#ef4444';
        }
    }

    // ===== Helpers =====
    function getAngkatan(nim) {
        const prefix = nim.substring(0, 2);
        return '20' + prefix;
    }

    function parseKelas(kelasPerkuliahan) {
        if (!kelasPerkuliahan) return '';
        const str = String(kelasPerkuliahan).trim();
        if (str.includes('Aksel')) return 'Aksel';
        if (str.includes('Pro')) return 'Pro';
        if (str.includes('Reg')) return 'Reg';
        return '';
    }

    function getKelasLabel(kode) {
        return KELAS_MAP[kode] || kode || '-';
    }

    function getFilteredData(data) {
        let result = data;
        // Filter by angkatan
        if (selectedAngkatan !== 'all') {
            result = result.filter(m => m.angkatan === selectedAngkatan);
        }
        // Filter by kelas
        if (selectedKelas !== 'all') {
            result = result.filter(m => m.kelas === selectedKelas);
        }
        // Filter by semester: only keep students who have at least one internship in that semester
        if (selectedSemester !== 'all') {
            result = result.map(m => {
                const filteredInternships = m.internships.filter(i => String(i.semester).trim() === selectedSemester);
                return {
                    ...m,
                    internships: filteredInternships,
                    jumlahMagang: filteredInternships.length
                };
            });
        }
        return result;
    }

    // ===== Data Loading =====
    async function loadData() {
        const [mhsRes, mgRes, blmRes, rptRes] = await Promise.all([
            fetch('data_mahasiswa.xlsx').then(r => r.arrayBuffer()),
            fetch('data_magang.xlsx').then(r => r.arrayBuffer()),
            fetch('Mahasiswa Belum Magang.xlsx').then(r => r.arrayBuffer()).catch(() => null),
            fetch('report_magang_terkini.xlsx').then(r => r.arrayBuffer()).catch(() => null)
        ]);

        const mhsWb = XLSX.read(mhsRes, { type: 'array' });
        const mgWb = XLSX.read(mgRes, { type: 'array' });

        const mhsSheet = mhsWb.Sheets[mhsWb.SheetNames[0]];
        const mgSheet = mgWb.Sheets[mgWb.SheetNames[0]];

        mahasiswaData = XLSX.utils.sheet_to_json(mhsSheet).map(row => ({
            nim: String(row['NIM'] || '').replace(/^'/, '').trim(),
            nama: String(row['Nama'] || '').trim(),
            kelasPerkuliahan: String(row['Kelas Perkuliahan'] || '').trim()
        })).filter(r => r.nim && r.nama);

        const rawMagang = XLSX.utils.sheet_to_json(mgSheet);
        magangData = rawMagang.map(row => {
            const name = String(row['Student Name'] || row['Student Name '] || '').trim();
            const major = String(row['Major'] || row['Major '] || '').trim();
            const company = String(row['Company Name'] || row['Company Name '] || '').trim();
            const position = String(row['Position'] || row['Position '] || '').trim();
            const semester = String(row['Internship Semester'] || row['Internship Semester '] || '').trim();
            const compensation = String(row['Compensation'] || row['Compensation '] || '').trim();
            const type = String(row['Type'] || row['Type '] || '').trim();
            const feedback = String(row["Employer's Feedback"] || row["Employer's Feedback "] || '').trim();
            const intake = String(row['Intake Year'] || row['Intake Year '] || '').trim();
            const evidence = String(row['Evidence'] || row['Evidence '] || '').trim();
            return { name, major, company, position, semester, compensation, type, feedback, intake, evidence };
        }).filter(r => r.name && r.company);

        // Load Mahasiswa Belum Magang data for Jumlah Apply (Data Science sheet only)
        if (blmRes) {
            const blmWb = XLSX.read(blmRes, { type: 'array' });
            // Use 'Data Science' sheet specifically
            const dsSheetName = blmWb.SheetNames.find(s => s.toLowerCase().includes('data science'));
            if (dsSheetName) {
                const blmSheet = blmWb.Sheets[dsSheetName];
                belumMagangData = XLSX.utils.sheet_to_json(blmSheet).map(row => ({
                    nim: String(row['NIM'] || '').replace(/^'/, '').trim(),
                    nama: String(row['Nama'] || '').trim(),
                    jumlahApply: parseInt(row['Jumlah Apply'] || '0', 10) || 0
                })).filter(r => r.nim);
            }
        }

        // Load report_magang_terkini.xlsx - Detail Mahasiswa sheet
        if (rptRes) {
            const rptWb = XLSX.read(rptRes, { type: 'array' });
            const detailSheetName = rptWb.SheetNames.find(s => s.toLowerCase().includes('detail mahasiswa'));
            if (detailSheetName) {
                const rptSheet = rptWb.Sheets[detailSheetName];
                const rawReport = XLSX.utils.sheet_to_json(rptSheet, { header: 1, defval: '' });
                // First row is header: No, NIM, Nama, Program Studi, Semester, Jumlah Magang, Detail Magang
                reportData = rawReport.slice(1).filter(r => r[1] && String(r[1]).trim()).map(r => ({
                    nim: String(r[1] || '').replace(/^'/, '').trim(),
                    nama: String(r[2] || '').trim(),
                    prodi: String(r[3] || '').trim(),
                    semester: String(r[4] || '').trim(),
                    jumlahMagang: parseInt(r[5] || '0', 10) || 0,
                    detailMagang: String(r[6] || '').trim()
                }));
            }
        }
    }

    // ===== Data Processing =====
    function normalizeName(n) {
        return n.toUpperCase()
            .replace(/[^A-Z\s]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    // Position normalization: case-insensitive + merge similar names
    const POSITION_ALIASES = {
        'frontend': 'Frontend Developer',
        'frontend developer': 'Frontend Developer',
        'front end': 'Frontend Developer',
        'front end developer': 'Frontend Developer',
        'frontend dev': 'Frontend Developer',
        'data analyst': 'Data Analyst',
        'data analist': 'Data Analyst',
        'data analysis': 'Data Analyst',
        'software developer': 'Software Developer',
        'software engineer': 'Software Engineer',
        'fullstack developer': 'Fullstack Developer',
        'full stack developer': 'Fullstack Developer',
        'social media specialist': 'Social Media Specialist',
        'socia media specialist': 'Social Media Specialist',
        'social media': 'Social Media Specialist',
        'graphic design': 'Graphic Designer',
        'graphic designer': 'Graphic Designer',
        'digital marketing': 'Digital Marketing',
        'python developer': 'Python Developer',
        'programmer': 'Programmer',
        'project manager': 'Project Manager',
        'technical writer': 'Technical Writer',
        'data engineer': 'Data Engineer',
    };

    function normalizePosition(pos) {
        const trimmed = pos.trim();
        if (!trimmed) return '';
        const lower = trimmed.toLowerCase();
        // Check exact alias match first
        if (POSITION_ALIASES[lower]) return POSITION_ALIASES[lower];
        // Check if any alias key is contained in the position or vice versa
        for (const [key, canonical] of Object.entries(POSITION_ALIASES)) {
            if (lower.includes(key) || key.includes(lower)) {
                return canonical;
            }
        }
        // No alias found: title-case the original
        return trimmed.replace(/\b\w/g, c => c.toUpperCase());
    }

    function processData() {
        // Filter magang to Data Science only
        const dsMagang = magangData.filter(m =>
            m.major.toLowerCase().includes('data science')
        );

        // Build lookup: normalized name -> magang records
        const magangByName = {};
        dsMagang.forEach(m => {
            const key = normalizeName(m.name);
            if (!magangByName[key]) magangByName[key] = [];
            magangByName[key].push(m);
        });

        // Merge
        mergedData = mahasiswaData.map(mhs => {
            const key = normalizeName(mhs.nama);
            let internships = magangByName[key] || [];

            // If no exact match, try partial matching
            if (internships.length === 0) {
                for (const [mk, records] of Object.entries(magangByName)) {
                    if (key.includes(mk) || mk.includes(key)) {
                        internships = records;
                        break;
                    }
                    const keyWords = key.split(' ');
                    const mkWords = mk.split(' ');
                    const matchCount = keyWords.filter(w => mkWords.includes(w)).length;
                    if (matchCount >= 2 && matchCount >= Math.min(keyWords.length, mkWords.length) * 0.6) {
                        internships = records;
                        break;
                    }
                }
            }

            // Find jumlah apply from belumMagangData
            const applyRecord = belumMagangData.find(b => b.nim === mhs.nim);
            const jumlahApply = applyRecord ? applyRecord.jumlahApply : 0;

            return {
                nim: mhs.nim,
                nama: mhs.nama,
                angkatan: getAngkatan(mhs.nim),
                kelas: parseKelas(mhs.kelasPerkuliahan),
                kelasPerkuliahan: mhs.kelasPerkuliahan,
                jumlahMagang: internships.length,
                jumlahApply: jumlahApply,
                internships: internships
            };
        });

        // Sync with report_magang_terkini data
        if (reportData.length > 0) {
            const mergedNIMs = new Set(mergedData.map(m => m.nim));

            // Update existing students with report data
            mergedData.forEach(m => {
                const rpt = reportData.find(r => r.nim === m.nim);
                if (rpt) {
                    // Update jumlahMagang from report if report has more up-to-date data
                    if (rpt.jumlahMagang > m.jumlahMagang && m.internships.length === 0) {
                        m.jumlahMagang = rpt.jumlahMagang;
                    }
                    // Add semester info from report
                    m.semesterMhs = rpt.semester;
                    // Parse detail magang from report if student has magang but no internship records
                    if (rpt.jumlahMagang > 0 && m.internships.length === 0 && rpt.detailMagang && rpt.detailMagang !== '-') {
                        const details = rpt.detailMagang.split('\n').filter(d => d.trim());
                        m.internships = details.map(d => {
                            const parts = d.split(' - ');
                            return {
                                name: m.nama,
                                major: 'Data Science',
                                company: parts[0] ? parts[0].trim() : d.trim(),
                                position: parts[1] ? parts[1].trim() : '',
                                semester: rpt.semester || '',
                                compensation: '',
                                type: '',
                                feedback: '',
                                intake: '',
                                evidence: ''
                            };
                        });
                        m.jumlahMagang = m.internships.length;
                    }
                }
            });

            // Add students from report that are not in data_mahasiswa
            reportData.forEach(rpt => {
                if (!mergedNIMs.has(rpt.nim)) {
                    let internships = [];
                    if (rpt.jumlahMagang > 0 && rpt.detailMagang && rpt.detailMagang !== '-') {
                        const details = rpt.detailMagang.split('\n').filter(d => d.trim());
                        internships = details.map(d => {
                            const parts = d.split(' - ');
                            return {
                                name: rpt.nama,
                                major: 'Data Science',
                                company: parts[0] ? parts[0].trim() : d.trim(),
                                position: parts[1] ? parts[1].trim() : '',
                                semester: rpt.semester || '',
                                compensation: '',
                                type: '',
                                feedback: '',
                                intake: '',
                                evidence: ''
                            };
                        });
                    }
                    const applyRecord = belumMagangData.find(b => b.nim === rpt.nim);
                    mergedData.push({
                        nim: rpt.nim,
                        nama: rpt.nama,
                        angkatan: getAngkatan(rpt.nim),
                        kelas: '',
                        kelasPerkuliahan: '',
                        jumlahMagang: rpt.jumlahMagang,
                        jumlahApply: applyRecord ? applyRecord.jumlahApply : 0,
                        semesterMhs: rpt.semester,
                        internships: internships
                    });
                }
            });
        }

        mergedData.sort((a, b) => b.jumlahMagang - a.jumlahMagang);
        filteredData = [...mergedData];
    }

    // ===== Angkatan Chips & Select =====
    function getAngkatanList() {
        const set = new Set(mergedData.map(m => m.angkatan));
        return Array.from(set).sort();
    }

    function buildAngkatanChips() {
        const container = $('#angkatanChips');
        const angkatanList = getAngkatanList();
        // Keep the "Semua" chip already in HTML, append year chips
        angkatanList.forEach(year => {
            const btn = document.createElement('button');
            btn.className = 'chip';
            btn.dataset.angkatan = year;
            btn.textContent = year;
            container.appendChild(btn);
        });
    }

    function buildAngkatanSelect() {
        const select = $('#filterAngkatan');
        const angkatanList = getAngkatanList();
        angkatanList.forEach(year => {
            const opt = document.createElement('option');
            opt.value = year;
            opt.textContent = year;
            select.appendChild(opt);
        });
    }

    // ===== Kelas Chips & Select =====
    function getKelasList() {
        const set = new Set(mergedData.map(m => m.kelas).filter(k => k));
        return Array.from(set).sort();
    }

    function buildKelasChips() {
        const container = $('#kelasChips');
        if (!container) return;
        const kelasList = getKelasList();
        kelasList.forEach(kode => {
            const btn = document.createElement('button');
            btn.className = 'chip';
            btn.dataset.kelas = kode;
            btn.textContent = getKelasLabel(kode);
            container.appendChild(btn);
        });
    }

    function buildKelasSelect() {
        const select = $('#filterKelas');
        if (!select) return;
        const kelasList = getKelasList();
        kelasList.forEach(kode => {
            const opt = document.createElement('option');
            opt.value = kode;
            opt.textContent = getKelasLabel(kode);
            select.appendChild(opt);
        });
    }

    // ===== Dashboard Rendering =====
    function renderDashboard() {
        const data = getFilteredData(mergedData);
        const total = data.length;
        const totalMagang = data.reduce((s, m) => s + m.jumlahMagang, 0);
        const sudah = data.filter(m => m.jumlahMagang > 0).length;
        const belum = total - sudah;
        const persen = total > 0 ? ((sudah / total) * 100).toFixed(1) : 0;

        // Set values directly (re-animate on filter change)
        animateCounter('totalMahasiswa', total);
        animateCounter('totalMagang', totalMagang);
        animateCounter('sudahMagang', sudah);
        animateCounter('belumMagang', belum);
        animatePercentage('persentaseMagang', parseFloat(persen));

        renderBarChart(data);
        renderDoughnutChart(sudah, belum);
        renderCompanyChart(data);
        renderPositionChart(data);
    }

    function animateCounter(id, target) {
        const el = document.getElementById(id);
        let current = 0;
        const step = Math.max(1, Math.ceil(target / 30));
        const timer = setInterval(() => {
            current += step;
            if (current >= target) { current = target; clearInterval(timer); }
            el.textContent = current;
        }, 25);
    }

    function animatePercentage(id, target) {
        const el = document.getElementById(id);
        let current = 0;
        const step = Math.max(0.5, target / 30);
        const timer = setInterval(() => {
            current += step;
            if (current >= target) { current = target; clearInterval(timer); }
            el.textContent = current.toFixed(1) + '%';
        }, 25);
    }

    function renderBarChart(data) {
        const groups = { '0': 0, '1': 0, '2': 0, '3+': 0 };
        data.forEach(m => {
            if (m.jumlahMagang === 0) groups['0']++;
            else if (m.jumlahMagang === 1) groups['1']++;
            else if (m.jumlahMagang === 2) groups['2']++;
            else groups['3+']++;
        });

        if (barChartInstance) barChartInstance.destroy();
        barChartInstance = new Chart($('#barChart'), {
            type: 'bar',
            data: {
                labels: ['0 Magang', '1 Magang', '2 Magang', '3+ Magang'],
                datasets: [{
                    label: 'Jumlah Mahasiswa',
                    data: Object.values(groups),
                    backgroundColor: [
                        'rgba(239,68,68,0.7)',
                        'rgba(245,158,11,0.7)',
                        'rgba(59,130,246,0.7)',
                        'rgba(16,185,129,0.7)'
                    ],
                    borderColor: ['#ef4444', '#f59e0b', '#3b82f6', '#10b981'],
                    borderWidth: 2,
                    borderRadius: 8,
                    barPercentage: 0.6,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#1a1f35', titleColor: '#f1f5f9',
                        bodyColor: '#94a3b8', borderColor: '#334155',
                        borderWidth: 1, cornerRadius: 8, padding: 12,
                    }
                },
                scales: {
                    x: { ticks: { color: '#64748b', font: { family: 'Inter' } }, grid: { display: false }, border: { display: false } },
                    y: { beginAtZero: true, ticks: { color: '#64748b', stepSize: 5, font: { family: 'Inter' } }, grid: { color: 'rgba(30,41,59,0.5)' }, border: { display: false } }
                }
            }
        });
    }

    function renderDoughnutChart(sudah, belum) {
        if (doughnutChartInstance) doughnutChartInstance.destroy();
        doughnutChartInstance = new Chart($('#doughnutChart'), {
            type: 'doughnut',
            data: {
                labels: ['Sudah Magang', 'Belum Magang'],
                datasets: [{
                    data: [sudah, belum],
                    backgroundColor: ['rgba(16,185,129,0.8)', 'rgba(239,68,68,0.6)'],
                    borderColor: ['#10b981', '#ef4444'],
                    borderWidth: 2,
                    hoverOffset: 8,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '65%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#94a3b8', padding: 20, font: { family: 'Inter', size: 13 }, usePointStyle: true, pointStyleWidth: 10 }
                    },
                    tooltip: {
                        backgroundColor: '#1a1f35', titleColor: '#f1f5f9',
                        bodyColor: '#94a3b8', borderColor: '#334155',
                        borderWidth: 1, cornerRadius: 8, padding: 12,
                        callbacks: {
                            label: function(ctx) {
                                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                                const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
                                return `${ctx.label}: ${ctx.parsed} (${pct}%)`;
                            }
                        }
                    }
                }
            }
        });
    }

    function renderCompanyChart(data) {
        const companyCount = {};
        data.forEach(m => {
            m.internships.forEach(i => {
                const c = i.company.trim();
                if (c) companyCount[c] = (companyCount[c] || 0) + 1;
            });
        });

        const sorted = Object.entries(companyCount).sort((a, b) => b[1] - a[1]).slice(0, 8);
        const colors = ['#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#06b6d4', '#ef4444', '#84cc16'];

        if (companyChartInstance) companyChartInstance.destroy();
        companyChartInstance = new Chart($('#companyChart'), {
            type: 'bar',
            data: {
                labels: sorted.map(s => s[0].length > 30 ? s[0].substring(0, 28) + '...' : s[0]),
                datasets: [{
                    label: 'Jumlah Mahasiswa',
                    data: sorted.map(s => s[1]),
                    backgroundColor: colors.map(c => c + 'bb'),
                    borderColor: colors,
                    borderWidth: 2,
                    borderRadius: 8,
                    barPercentage: 0.6,
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { backgroundColor: '#1a1f35', titleColor: '#f1f5f9', bodyColor: '#94a3b8', borderColor: '#334155', borderWidth: 1, cornerRadius: 8, padding: 12 }
                },
                scales: {
                    x: { beginAtZero: true, ticks: { color: '#64748b', stepSize: 1, font: { family: 'Inter' } }, grid: { color: 'rgba(30,41,59,0.5)' }, border: { display: false } },
                    y: { ticks: { color: '#94a3b8', font: { family: 'Inter', size: 12 } }, grid: { display: false }, border: { display: false } }
                }
            }
        });
        $('#companyChart').parentElement.style.height = '320px';
    }

    function renderPositionChart(data) {
        const posCount = {};
        data.forEach(m => {
            m.internships.forEach(i => {
                const p = normalizePosition(i.position);
                if (p) posCount[p] = (posCount[p] || 0) + 1;
            });
        });

        const sorted = Object.entries(posCount).sort((a, b) => b[1] - a[1]).slice(0, 8);
        const colors = ['#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#3b82f6', '#84cc16', '#ef4444'];

        if (positionChartInstance) positionChartInstance.destroy();
        positionChartInstance = new Chart($('#positionChart'), {
            type: 'bar',
            data: {
                labels: sorted.map(s => s[0].length > 25 ? s[0].substring(0, 23) + '...' : s[0]),
                datasets: [{
                    label: 'Jumlah Mahasiswa',
                    data: sorted.map(s => s[1]),
                    backgroundColor: colors.map(c => c + 'bb'),
                    borderColor: colors,
                    borderWidth: 2,
                    borderRadius: 8,
                    barPercentage: 0.6,
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { backgroundColor: '#1a1f35', titleColor: '#f1f5f9', bodyColor: '#94a3b8', borderColor: '#334155', borderWidth: 1, cornerRadius: 8, padding: 12 }
                },
                scales: {
                    x: { beginAtZero: true, ticks: { color: '#64748b', stepSize: 1, font: { family: 'Inter' } }, grid: { color: 'rgba(30,41,59,0.5)' }, border: { display: false } },
                    y: { ticks: { color: '#94a3b8', font: { family: 'Inter', size: 11 } }, grid: { display: false }, border: { display: false } }
                }
            }
        });
        $('#positionChart').parentElement.style.height = '320px';
    }

    // ===== Table Rendering =====
    function renderTable() {
        const tbody = $('#mahasiswaBody');
        const start = (currentPage - 1) * PAGE_SIZE;
        const pageData = filteredData.slice(start, start + PAGE_SIZE);

        tbody.innerHTML = pageData.map((m, i) => {
            const badgeClass = m.jumlahMagang === 0 ? 'badge-red' :
                               m.jumlahMagang === 1 ? 'badge-orange' :
                               m.jumlahMagang === 2 ? 'badge-blue' : 'badge-green';
            const kelasBadge = m.kelas === 'Pro' ? 'badge-kelas-pro' :
                               m.kelas === 'Aksel' ? 'badge-kelas-aksel' : 'badge-kelas-reg';
            const applyBadge = m.jumlahApply === 0 ? 'badge-red' :
                               m.jumlahApply <= 5 ? 'badge-orange' :
                               m.jumlahApply <= 10 ? 'badge-blue' :
                               m.jumlahApply <= 20 ? 'badge-purple' : 'badge-green';
            return `<tr>
                <td>${start + i + 1}</td>
                <td style="font-family:monospace;color:var(--accent-cyan)">${m.nim}</td>
                <td style="font-weight:600;color:var(--text-primary)">${m.nama}</td>
                <td><span class="badge badge-purple">${m.angkatan}</span></td>
                <td><span class="badge ${kelasBadge}">${getKelasLabel(m.kelas)}</span></td>
                <td><span class="badge ${badgeClass}">${m.jumlahMagang} magang</span></td>
                <td><span class="badge ${applyBadge}">${m.jumlahApply} apply</span></td>
                <td><button class="btn-detail" onclick="showDetail('${m.nim}')">Detail</button></td>
            </tr>`;
        }).join('');

        renderPagination();
    }

    function renderPagination() {
        const totalPages = Math.ceil(filteredData.length / PAGE_SIZE);
        const container = $('#pagination');
        if (totalPages <= 1) { container.innerHTML = ''; return; }

        let html = '';
        if (currentPage > 1) html += `<button class="page-btn" onclick="goToPage(${currentPage - 1})">&laquo;</button>`;
        for (let i = 1; i <= totalPages; i++) {
            if (i === 1 || i === totalPages || Math.abs(i - currentPage) <= 2) {
                html += `<button class="page-btn ${i === currentPage ? 'active' : ''}" onclick="goToPage(${i})">${i}</button>`;
            } else if (Math.abs(i - currentPage) === 3) {
                html += `<span style="color:var(--text-muted);padding:0 4px">...</span>`;
            }
        }
        if (currentPage < totalPages) html += `<button class="page-btn" onclick="goToPage(${currentPage + 1})">&raquo;</button>`;
        container.innerHTML = html;
    }

    // ===== Magang Selects =====
    function buildMagangSelects() {
        const dsMagang = magangData.filter(m => m.major.toLowerCase().includes('data science'));

        // Company select
        const companies = [...new Set(dsMagang.map(m => m.company).filter(c => c))].sort();
        const companySelect = $('#filterMagangCompany');
        if (companySelect) {
            companies.forEach(c => {
                const opt = document.createElement('option');
                opt.value = c;
                opt.textContent = c.length > 35 ? c.substring(0, 33) + '...' : c;
                companySelect.appendChild(opt);
            });
        }

        // Position select
        const positions = [...new Set(dsMagang.map(m => normalizePosition(m.position)).filter(p => p))].sort();
        const posSelect = $('#filterMagangPosition');
        if (posSelect) {
            positions.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = p;
                posSelect.appendChild(opt);
            });
        }

        // Type select
        const types = [...new Set(dsMagang.map(m => m.type).filter(t => t))].sort();
        const typeSelect = $('#filterMagangType');
        if (typeSelect) {
            types.forEach(t => {
                const opt = document.createElement('option');
                opt.value = t;
                opt.textContent = t;
                typeSelect.appendChild(opt);
            });
        }
    }

    function renderMagangTable() {
        const dsMagang = magangData.filter(m => m.major.toLowerCase().includes('data science'));

        // Apply filters
        const search = ($('#filterMagangSearch')?.value || '').toLowerCase().trim();
        const semFilter = $('#filterMagangSemester')?.value || 'all';
        const compFilter = $('#filterMagangCompany')?.value || 'all';
        const posFilter = $('#filterMagangPosition')?.value || 'all';
        const typeFilter = $('#filterMagangType')?.value || 'all';
        const sortVal = $('#sortMagang')?.value || 'name-asc';

        filteredMagangData = dsMagang.filter(m => {
            if (search && !m.name.toLowerCase().includes(search) && !m.company.toLowerCase().includes(search)) return false;
            if (semFilter !== 'all' && String(m.semester).trim() !== semFilter) return false;
            if (compFilter !== 'all' && m.company !== compFilter) return false;
            if (posFilter !== 'all' && normalizePosition(m.position) !== posFilter) return false;
            if (typeFilter !== 'all' && m.type !== typeFilter) return false;
            return true;
        });

        // Sort
        const [sortField, sortDir] = sortVal.split('-');
        filteredMagangData.sort((a, b) => {
            if (sortField === 'name') return sortDir === 'asc' ? a.name.localeCompare(b.name) : b.name.localeCompare(a.name);
            if (sortField === 'company') return a.company.localeCompare(b.company);
            if (sortField === 'semester') return Number(a.semester) - Number(b.semester);
            return 0;
        });

        // Paginate
        const start = (magangPage - 1) * MAGANG_PAGE_SIZE;
        const pageData = filteredMagangData.slice(start, start + MAGANG_PAGE_SIZE);

        const tbody = $('#magangBody');
        if (pageData.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="no-data">Tidak ada data magang ditemukan</td></tr>';
        } else {
            tbody.innerHTML = pageData.map((m, i) => {
                let evCell = '-';
                if (m.evidence && m.evidence !== 'undefined') {
                    if (m.evidence.startsWith('http')) {
                        evCell = `<a href="${m.evidence}" target="_blank" class="btn-evidence" title="Lihat Evidence">📎 Lihat</a>`;
                    } else {
                        evCell = `<span class="badge badge-green">✓ Ada</span>`;
                    }
                }
                return `<tr>
                <td>${start + i + 1}</td>
                <td style="font-weight:600;color:var(--text-primary)">${m.name}</td>
                <td>${m.company}</td>
                <td>${normalizePosition(m.position)}</td>
                <td><span class="badge badge-blue">Semester ${m.semester}</span></td>
                <td><span class="badge badge-purple">${m.type || '-'}</span></td>
                <td>${m.compensation || '-'}</td>
                <td>${evCell}</td>
            </tr>`;
            }).join('');
        }

        renderMagangPagination();
    }

    function renderMagangPagination() {
        const totalPages = Math.ceil(filteredMagangData.length / MAGANG_PAGE_SIZE);
        const container = $('#magangPagination');
        if (!container) return;
        if (totalPages <= 1) { container.innerHTML = ''; return; }

        let html = '';
        if (magangPage > 1) html += `<button class="page-btn" onclick="goToMagangPage(${magangPage - 1})">&laquo;</button>`;
        for (let i = 1; i <= totalPages; i++) {
            if (i === 1 || i === totalPages || Math.abs(i - magangPage) <= 2) {
                html += `<button class="page-btn ${i === magangPage ? 'active' : ''}" onclick="goToMagangPage(${i})">${i}</button>`;
            } else if (Math.abs(i - magangPage) === 3) {
                html += `<span style="color:var(--text-muted);padding:0 4px">...</span>`;
            }
        }
        if (magangPage < totalPages) html += `<button class="page-btn" onclick="goToMagangPage(${magangPage + 1})">&raquo;</button>`;
        container.innerHTML = html;
    }

    window.goToMagangPage = function(page) {
        magangPage = page;
        renderMagangTable();
        $('#sectionMagang .table-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
    };

    function applyMagangFilters() {
        magangPage = 1;
        renderMagangTable();
    }

    // ===== Position Section =====
    function getPositionData() {
        const posMap = {};
        mergedData.forEach(m => {
            m.internships.forEach(int => {
                const pos = normalizePosition(int.position);
                if (!pos) return;
                if (!posMap[pos]) posMap[pos] = [];
                posMap[pos].push({
                    nim: m.nim,
                    nama: m.nama,
                    angkatan: m.angkatan,
                    company: int.company,
                    semester: int.semester,
                    compensation: int.compensation,
                    feedback: int.feedback
                });
            });
        });
        return posMap;
    }

    function renderPositionSection() {
        const posMap = getPositionData();
        const search = ($('#filterPosisiSearch')?.value || '').toLowerCase().trim();
        const sortVal = $('#sortPosisi')?.value || 'count-desc';

        let entries = Object.entries(posMap);

        // Filter by search
        if (search) {
            entries = entries.filter(([pos]) => pos.toLowerCase().includes(search));
        }

        // Sort
        const [field, dir] = sortVal.split('-');
        entries.sort((a, b) => {
            if (field === 'name') return dir === 'asc' ? a[0].localeCompare(b[0]) : b[0].localeCompare(a[0]);
            return dir === 'asc' ? a[1].length - b[1].length : b[1].length - a[1].length;
        });

        const iconColors = [
            'rgba(6,182,212,0.15)', 'rgba(16,185,129,0.15)', 'rgba(245,158,11,0.15)',
            'rgba(236,72,153,0.15)', 'rgba(139,92,246,0.15)', 'rgba(59,130,246,0.15)',
            'rgba(132,204,22,0.15)', 'rgba(239,68,68,0.15)'
        ];
        const posIcons = ['💼', '📊', '💻', '🎨', '📝', '🔧', '📈', '🔬', '🎯', '📱'];

        const container = $('#positionCardsGrid');
        container.innerHTML = entries.map(([pos, students], idx) => {
            const icon = posIcons[idx % posIcons.length];
            const bgColor = iconColors[idx % iconColors.length];
            const studentListHtml = students.map(s => `
                <div class="position-student-item" onclick="showDetail('${s.nim}')" style="cursor:pointer">
                    <div class="position-student-info">
                        <span class="position-student-name">${s.nama}</span>
                        <div class="position-student-meta">
                            <span>🏢 ${s.company}</span>
                            <span>📅 Sem. ${s.semester}</span>
                            <span>🎓 ${s.angkatan}</span>
                        </div>
                    </div>
                </div>
            `).join('');

            return `
                <div class="position-card" id="posCard${idx}">
                    <div class="position-card-header" onclick="togglePositionCard(${idx})">
                        <div class="position-card-title">
                            <div class="position-card-icon" style="background:${bgColor}">${icon}</div>
                            <div>
                                <div class="position-card-name" title="${pos}">${pos}</div>
                                <div class="position-card-count">${students.length} mahasiswa</div>
                            </div>
                        </div>
                        <span class="position-card-badge">${students.length}</span>
                        <div class="position-card-toggle">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><polyline points="6 9 12 15 18 9"/></svg>
                        </div>
                    </div>
                    <div class="position-card-body">
                        <div class="position-card-content">
                            <ul class="position-student-list">
                                ${studentListHtml}
                            </ul>
                        </div>
                    </div>
                </div>
            `;
        }).join('');

        if (entries.length === 0) {
            container.innerHTML = '<div class="no-data">Tidak ada posisi ditemukan</div>';
        }
    }

    window.togglePositionCard = function(idx) {
        const card = document.getElementById('posCard' + idx);
        if (card) card.classList.toggle('expanded');
    };

    // ===== Filtering & Sorting =====
    function applyFilters() {
        const filterVal = $('#filterMagang').value;
        const sortVal = $('#sortBy').value;
        const search = ($('#searchInput').value || '').toLowerCase().trim();
        const angkatanFilter = $('#filterAngkatan').value;
        const semesterFilter = $('#filterSemester').value;
        const kelasFilter = $('#filterKelas').value;
        const applyFilter = $('#filterJumlahApply').value;

        filteredData = mergedData.map(m => {
            // If semester filter active, narrow internships
            if (semesterFilter !== 'all') {
                const filtIntern = m.internships.filter(i => String(i.semester).trim() === semesterFilter);
                return { ...m, internships: filtIntern, jumlahMagang: filtIntern.length };
            }
            return m;
        }).filter(m => {
            // Angkatan filter (table-level)
            if (angkatanFilter !== 'all' && m.angkatan !== angkatanFilter) return false;
            // Kelas filter (table-level)
            if (kelasFilter !== 'all' && m.kelas !== kelasFilter) return false;
            // Filter by magang count
            if (filterVal === '0' && m.jumlahMagang !== 0) return false;
            if (filterVal === '1' && m.jumlahMagang !== 1) return false;
            if (filterVal === '2+' && m.jumlahMagang < 2) return false;
            // Filter by jumlah apply
            if (applyFilter !== 'all') {
                if (applyFilter === '1-5' && (m.jumlahApply < 1 || m.jumlahApply > 5)) return false;
                if (applyFilter === '6-10' && (m.jumlahApply < 6 || m.jumlahApply > 10)) return false;
                if (applyFilter === '11-20' && (m.jumlahApply < 11 || m.jumlahApply > 20)) return false;
                if (applyFilter === '20+' && m.jumlahApply <= 20) return false;
            }
            // Search
            if (search) {
                return m.nama.toLowerCase().includes(search) || m.nim.includes(search);
            }
            return true;
        });

        // Sort
        const [field, dir] = sortVal.split('-');
        filteredData.sort((a, b) => {
            if (field === 'nama') return dir === 'asc' ? a.nama.localeCompare(b.nama) : b.nama.localeCompare(a.nama);
            return dir === 'asc' ? a.jumlahMagang - b.jumlahMagang : b.jumlahMagang - a.jumlahMagang;
        });

        currentPage = 1;
        renderTable();
    }

    // ===== Modal =====
    window.showDetail = function (nim) {
        const mhs = mergedData.find(m => m.nim === nim);
        if (!mhs) return;

        $('#modalTitle').textContent = mhs.nama;
        const body = $('#modalBody');

        let internshipHtml = '';
        if (mhs.internships.length === 0) {
            internshipHtml = '<div class="no-data">Belum ada pengalaman magang tercatat</div>';
        } else {
            internshipHtml = mhs.internships.map((int, i) => `
                <div class="internship-card">
                    <div class="company">${i + 1}. ${int.company}</div>
                    <div class="meta">
                        <span>📋 ${int.position || '-'}</span>
                        <span>📅 Semester ${int.semester || '-'}</span>
                        <span>💰 ${int.compensation || '-'}</span>
                        <span>📌 ${int.type || '-'}</span>
                    </div>
                    ${int.feedback && int.feedback !== 'undefined' ? `
                    <div class="feedback-box">
                        <div class="fb-label">💬 Feedback Perusahaan</div>
                        <div class="fb-text">"${int.feedback.replace(/^"|"$/g, '')}"</div>
                    </div>` : ''}
                </div>
            `).join('');
        }

        body.innerHTML = `
            <div class="detail-section">
                <h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg> Informasi Mahasiswa</h4>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="label">NIM</div>
                        <div class="value" style="font-family:monospace;color:var(--accent-cyan)">${mhs.nim}</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Nama Lengkap</div>
                        <div class="value">${mhs.nama}</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Angkatan</div>
                        <div class="value">${mhs.angkatan}</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Kelas</div>
                        <div class="value">${getKelasLabel(mhs.kelas)}</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Total Magang</div>
                        <div class="value">${mhs.jumlahMagang} pengalaman</div>
                    </div>
                    <div class="info-item">
                        <div class="label">Status</div>
                        <div class="value">${mhs.jumlahMagang > 0 ? '✅ Sudah Magang' : '⚠️ Belum Magang'}</div>
                    </div>
                </div>
            </div>
            <div class="detail-section">
                <h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg> Riwayat Magang (${mhs.jumlahMagang})</h4>
                ${internshipHtml}
            </div>
        `;

        $('#modalOverlay').classList.add('show');
        document.body.style.overflow = 'hidden';
    };

    function closeModal() {
        $('#modalOverlay').classList.remove('show');
        document.body.style.overflow = '';
    }

    // ===== Navigation =====
    function switchSection(section) {
        $$('.content-section').forEach(s => s.classList.remove('active'));
        $$('.nav-item').forEach(n => n.classList.remove('active'));

        $(`#section${section.charAt(0).toUpperCase() + section.slice(1)}`).classList.add('active');
        $(`[data-section="${section}"]`).classList.add('active');

        currentSection = section;

        const titles = {
            dashboard: ['Dashboard Statistik', 'Ringkasan data magang mahasiswa Sains Data'],
            mahasiswa: ['Data Mahasiswa', 'Daftar lengkap mahasiswa dan pengalaman magang'],
            magang: ['Data Magang', 'Seluruh data magang mahasiswa Sains Data'],
            posisi: ['Posisi Magang', 'Detail mahasiswa berdasarkan posisi/profesi magang']
        };
        $('#pageTitle').textContent = titles[section][0];
        $('#pageSubtitle').textContent = titles[section][1];
    }

    // ===== Pagination Global =====
    window.goToPage = function (page) {
        currentPage = page;
        renderTable();
        $('#sectionMahasiswa .table-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
    };

    // ===== PDF Export =====
    function createPDF(orientation) {
        if (!window.jspdf || !window.jspdf.jsPDF) {
            alert('Library jsPDF belum dimuat. Silakan refresh halaman dan coba lagi.');
            return null;
        }
        return new window.jspdf.jsPDF(orientation, 'mm', 'a4');
    }

    function savePDF(doc, filename) {
        doc.save(filename);
    }

    function addPDFHeader(doc, title, subtitle) {
        doc.setFontSize(18);
        doc.setTextColor(59, 130, 246);
        doc.text(title, 14, 20);
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        const dateStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
        doc.text('Diekspor pada: ' + dateStr + (subtitle ? '  |  ' + subtitle : ''), 14, 28);
    }

    function addAutoTable(doc, startY, head, body, headColor) {
        try {
            doc.autoTable({
                startY: startY,
                head: [head],
                body: body,
                styles: { fontSize: 8, cellPadding: 3 },
                headStyles: { fillColor: headColor, textColor: [255, 255, 255] },
                alternateRowStyles: { fillColor: [245, 247, 250] },
                margin: { left: 14, right: 14 },
            });
        } catch (e) {
            console.error('autoTable error:', e);
            // Fallback: simple text table
            doc.setFontSize(9);
            doc.setTextColor(40, 40, 40);
            let y = startY + 5;
            body.forEach(function(row, idx) {
                if (y > 270) { doc.addPage(); y = 20; }
                doc.text(row.map(function(c) { return String(c); }).join('  |  '), 14, y);
                y += 6;
            });
        }
    }

    function exportDashboardPDF() {
        try {
            var doc = createPDF('l');
            if (!doc) return;

            var data = getFilteredData(mergedData);
            var total = data.length;
            var totalMag = data.reduce(function(s, m) { return s + m.jumlahMagang; }, 0);
            var sudah = data.filter(function(m) { return m.jumlahMagang > 0; }).length;
            var belum = total - sudah;
            var persen = total > 0 ? ((sudah / total) * 100).toFixed(1) : '0';

            addPDFHeader(doc, 'Dashboard Statistik Magang - Sains Data', '');

            // Stats summary
            doc.setFontSize(12);
            doc.setTextColor(40, 40, 40);
            doc.text('Total Mahasiswa: ' + total, 14, 40);
            doc.text('Total Pengalaman Magang: ' + totalMag, 14, 48);
            doc.text('Sudah Magang: ' + sudah, 14, 56);
            doc.text('Belum Magang: ' + belum, 14, 64);
            doc.text('Persentase Magang: ' + persen + '%', 14, 72);

            // Table
            var tableData = data.map(function(m, i) {
                return [
                    i + 1, m.nim, m.nama, m.angkatan, getKelasLabel(m.kelas),
                    m.jumlahMagang > 0 ? 'Sudah' : 'Belum', m.jumlahMagang
                ];
            });

            addAutoTable(doc, 82,
                ['No', 'NIM', 'Nama', 'Angkatan', 'Kelas', 'Status', 'Jml Magang'],
                tableData, [59, 130, 246]
            );

            savePDF(doc, 'Dashboard_Statistik_Magang.pdf');
        } catch (err) {
            console.error('Export Dashboard PDF error:', err);
            alert('Gagal mengekspor PDF: ' + err.message);
        }
    }

    function exportMahasiswaPDF() {
        try {
            var doc = createPDF('l');
            if (!doc) return;

            addPDFHeader(doc, 'Data Mahasiswa - Sains Data', 'Total: ' + filteredData.length + ' mahasiswa');

            var tableData = filteredData.map(function(m, i) {
                return [i + 1, m.nim, m.nama, m.angkatan, getKelasLabel(m.kelas), m.jumlahMagang, m.jumlahApply];
            });

            addAutoTable(doc, 36,
                ['No', 'NIM', 'Nama Mahasiswa', 'Angkatan', 'Kelas', 'Jml Magang', 'Jml Apply'],
                tableData, [59, 130, 246]
            );

            savePDF(doc, 'Data_Mahasiswa.pdf');
        } catch (err) {
            console.error('Export Mahasiswa PDF error:', err);
            alert('Gagal mengekspor PDF: ' + err.message);
        }
    }

    function exportMagangPDF() {
        try {
            var doc = createPDF('l');
            if (!doc) return;

            addPDFHeader(doc, 'Data Magang - Sains Data', 'Total: ' + filteredMagangData.length + ' data magang');

            var tableData = filteredMagangData.map(function(m, i) {
                return [
                    i + 1, m.name, m.company, normalizePosition(m.position),
                    'Sem. ' + m.semester, m.type || '-', m.compensation || '-'
                ];
            });

            addAutoTable(doc, 36,
                ['No', 'Nama Mahasiswa', 'Perusahaan', 'Posisi', 'Semester', 'Tipe', 'Kompensasi'],
                tableData, [16, 185, 129]
            );

            savePDF(doc, 'Data_Magang.pdf');
        } catch (err) {
            console.error('Export Magang PDF error:', err);
            alert('Gagal mengekspor PDF: ' + err.message);
        }
    }

    function exportPosisiPDF() {
        try {
            var doc = createPDF('p');
            if (!doc) return;

            addPDFHeader(doc, 'Posisi Magang - Sains Data', '');

            var posMap = getPositionData();
            var entries = Object.entries(posMap).sort(function(a, b) { return b[1].length - a[1].length; });
            var tableData = [];
            entries.forEach(function(entry) {
                var pos = entry[0];
                var students = entry[1];
                students.forEach(function(s, i) {
                    tableData.push([
                        i === 0 ? pos : '',
                        i === 0 ? students.length : '',
                        s.nama, s.company, 'Sem. ' + s.semester, s.angkatan
                    ]);
                });
            });

            addAutoTable(doc, 36,
                ['Posisi', 'Jumlah', 'Nama Mahasiswa', 'Perusahaan', 'Semester', 'Angkatan'],
                tableData, [139, 92, 246]
            );

            savePDF(doc, 'Posisi_Magang.pdf');
        } catch (err) {
            console.error('Export Posisi PDF error:', err);
            alert('Gagal mengekspor PDF: ' + err.message);
        }
    }

    // ===== Events =====
    function setupEvents() {
        // Navigation
        $$('.nav-item').forEach(item => {
            item.addEventListener('click', e => {
                e.preventDefault();
                switchSection(item.dataset.section);
                $('#sidebar').classList.remove('open');
            });
        });

        // Mobile menu
        $('#menuToggle').addEventListener('click', () => {
            $('#sidebar').classList.toggle('open');
        });

        // Modal close
        $('#modalClose').addEventListener('click', closeModal);
        $('#modalOverlay').addEventListener('click', e => {
            if (e.target === $('#modalOverlay')) closeModal();
        });
        document.addEventListener('keydown', e => {
            if (e.key === 'Escape') closeModal();
        });

        // Angkatan chip buttons (dashboard level)
        $('#angkatanChips').addEventListener('click', e => {
            const chip = e.target.closest('.chip');
            if (!chip) return;
            $$('#angkatanChips .chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            selectedAngkatan = chip.dataset.angkatan;
            renderDashboard();
        });

        // Semester chip buttons (dashboard level)
        $('#semesterChips').addEventListener('click', e => {
            const chip = e.target.closest('.chip');
            if (!chip) return;
            $$('#semesterChips .chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            selectedSemester = chip.dataset.semester;
            renderDashboard();
        });

        // Kelas chip buttons (dashboard level)
        $('#kelasChips').addEventListener('click', e => {
            const chip = e.target.closest('.chip');
            if (!chip) return;
            $$('#kelasChips .chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            selectedKelas = chip.dataset.kelas;
            renderDashboard();
        });

        // Mahasiswa table filters
        $('#filterMagang').addEventListener('change', applyFilters);
        $('#sortBy').addEventListener('change', applyFilters);
        $('#filterAngkatan').addEventListener('change', applyFilters);
        $('#filterSemester').addEventListener('change', applyFilters);
        $('#filterKelas').addEventListener('change', applyFilters);
        $('#filterJumlahApply').addEventListener('change', applyFilters);

        // Magang table filters
        let magangSearchTimeout;
        $('#filterMagangSearch').addEventListener('input', () => {
            clearTimeout(magangSearchTimeout);
            magangSearchTimeout = setTimeout(applyMagangFilters, 300);
        });
        $('#filterMagangSemester').addEventListener('change', applyMagangFilters);
        $('#filterMagangCompany').addEventListener('change', applyMagangFilters);
        $('#filterMagangPosition').addEventListener('change', applyMagangFilters);
        $('#filterMagangType').addEventListener('change', applyMagangFilters);
        $('#sortMagang').addEventListener('change', applyMagangFilters);

        // Posisi section filters
        let posSearchTimeout;
        $('#filterPosisiSearch').addEventListener('input', () => {
            clearTimeout(posSearchTimeout);
            posSearchTimeout = setTimeout(renderPositionSection, 300);
        });
        $('#sortPosisi').addEventListener('change', renderPositionSection);

        // Context-aware global search
        let searchTimeout;
        $('#searchInput').addEventListener('input', () => {
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => {
                const query = $('#searchInput').value.trim();
                if (currentSection === 'magang' || currentSection === 'dashboard') {
                    // On magang or dashboard page, search goes to magang
                    if (query) {
                        $('#filterMagangSearch').value = query;
                        switchSection('magang');
                        applyMagangFilters();
                    }
                } else {
                    // Default: search in mahasiswa
                    applyFilters();
                    if (query) {
                        switchSection('mahasiswa');
                    }
                }
            }, 300);
        });

        // Export PDF buttons
        $('#exportDashboard').addEventListener('click', exportDashboardPDF);
        $('#exportMahasiswa').addEventListener('click', exportMahasiswaPDF);
        $('#exportMagang').addEventListener('click', exportMagangPDF);
        $('#exportPosisi').addEventListener('click', exportPosisiPDF);
    }

    // ===== Show App =====
    function showApp() {
        setTimeout(() => {
            $('#loadingScreen').classList.add('hidden');
            $('#app').classList.remove('hidden');
        }, 800);
    }

})();
