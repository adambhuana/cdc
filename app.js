/* ===== CDC Dashboard - Main Application ===== */
(function () {
    'use strict';

    // ===== State =====
    let mahasiswaData = [];
    let magangData = [];
    let mergedData = [];
    let filteredData = [];
    let currentPage = 1;
    let selectedAngkatan = 'all';
    const PAGE_SIZE = 15;

    // Chart instances for destroy/recreate
    let barChartInstance = null;
    let doughnutChartInstance = null;
    let companyChartInstance = null;

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
            renderDashboard();
            renderTable();
            renderMagangTable();
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

    function getFilteredByAngkatan(data) {
        if (selectedAngkatan === 'all') return data;
        return data.filter(m => getAngkatan(m.nim) === selectedAngkatan);
    }

    // ===== Data Loading =====
    async function loadData() {
        const [mhsRes, mgRes] = await Promise.all([
            fetch('data_mahasiswa.xlsx').then(r => r.arrayBuffer()),
            fetch('data_magang.xlsx').then(r => r.arrayBuffer())
        ]);

        const mhsWb = XLSX.read(mhsRes, { type: 'array' });
        const mgWb = XLSX.read(mgRes, { type: 'array' });

        const mhsSheet = mhsWb.Sheets[mhsWb.SheetNames[0]];
        const mgSheet = mgWb.Sheets[mgWb.SheetNames[0]];

        mahasiswaData = XLSX.utils.sheet_to_json(mhsSheet).map(row => ({
            nim: String(row['NIM'] || '').replace(/^'/, '').trim(),
            nama: String(row['Nama'] || '').trim()
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
            return { name, major, company, position, semester, compensation, type, feedback, intake };
        }).filter(r => r.name && r.company);
    }

    // ===== Data Processing =====
    function normalizeName(n) {
        return n.toUpperCase()
            .replace(/[^A-Z\s]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
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

            return {
                nim: mhs.nim,
                nama: mhs.nama,
                angkatan: getAngkatan(mhs.nim),
                jumlahMagang: internships.length,
                internships: internships
            };
        });

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

    // ===== Dashboard Rendering =====
    function renderDashboard() {
        const data = getFilteredByAngkatan(mergedData);
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

    // ===== Table Rendering =====
    function renderTable() {
        const tbody = $('#mahasiswaBody');
        const start = (currentPage - 1) * PAGE_SIZE;
        const pageData = filteredData.slice(start, start + PAGE_SIZE);

        tbody.innerHTML = pageData.map((m, i) => {
            const badgeClass = m.jumlahMagang === 0 ? 'badge-red' :
                               m.jumlahMagang === 1 ? 'badge-orange' :
                               m.jumlahMagang === 2 ? 'badge-blue' : 'badge-green';
            return `<tr>
                <td>${start + i + 1}</td>
                <td style="font-family:monospace;color:var(--accent-cyan)">${m.nim}</td>
                <td style="font-weight:600;color:var(--text-primary)">${m.nama}</td>
                <td><span class="badge badge-purple">${m.angkatan}</span></td>
                <td><span class="badge ${badgeClass}">${m.jumlahMagang} magang</span></td>
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

    function renderMagangTable() {
        const dsMagang = magangData.filter(m => m.major.toLowerCase().includes('data science'));
        const tbody = $('#magangBody');
        tbody.innerHTML = dsMagang.map((m, i) => `<tr>
            <td>${i + 1}</td>
            <td style="font-weight:600;color:var(--text-primary)">${m.name}</td>
            <td>${m.company}</td>
            <td>${m.position}</td>
            <td><span class="badge badge-blue">Semester ${m.semester}</span></td>
            <td>${m.type}</td>
            <td>${m.compensation}</td>
        </tr>`).join('');
    }

    // ===== Filtering & Sorting =====
    function applyFilters() {
        const filterVal = $('#filterMagang').value;
        const sortVal = $('#sortBy').value;
        const search = ($('#searchInput').value || '').toLowerCase().trim();
        const angkatanFilter = $('#filterAngkatan').value;

        filteredData = mergedData.filter(m => {
            // Angkatan filter (table-level)
            if (angkatanFilter !== 'all' && m.angkatan !== angkatanFilter) return false;
            // Filter by magang count
            if (filterVal === '0' && m.jumlahMagang !== 0) return false;
            if (filterVal === '1' && m.jumlahMagang !== 1) return false;
            if (filterVal === '2+' && m.jumlahMagang < 2) return false;
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

        const titles = {
            dashboard: ['Dashboard Statistik', 'Ringkasan data magang mahasiswa Sains Data'],
            mahasiswa: ['Data Mahasiswa', 'Daftar lengkap mahasiswa dan pengalaman magang'],
            magang: ['Data Magang', 'Seluruh data magang mahasiswa Sains Data']
        };
        $('#pageTitle').textContent = titles[section][0];
        $('#pageSubtitle').textContent = titles[section][1];
    }

    // ===== Pagination Global =====
    window.goToPage = function (page) {
        currentPage = page;
        renderTable();
        $('.table-container').scrollIntoView({ behavior: 'smooth', block: 'start' });
    };

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

        // Table filters
        $('#filterMagang').addEventListener('change', applyFilters);
        $('#sortBy').addEventListener('change', applyFilters);
        $('#filterAngkatan').addEventListener('change', applyFilters);

        // Search
        let searchTimeout;
        $('#searchInput').addEventListener('input', () => {
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => {
                applyFilters();
                if ($('#searchInput').value.trim()) {
                    switchSection('mahasiswa');
                }
            }, 300);
        });
    }

    // ===== Show App =====
    function showApp() {
        setTimeout(() => {
            $('#loadingScreen').classList.add('hidden');
            $('#app').classList.remove('hidden');
        }, 800);
    }

})();
