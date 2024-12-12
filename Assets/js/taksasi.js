document.getElementById('taksasiForm').addEventListener('submit', function(e) {
    e.preventDefault();

    const blokData = {
        "S62": { luas: 32.14, sph: 81 },
        "S61": { luas: 28.29, sph: 128 },
        "S60": { luas: 32.03, sph: 142 },
        "S59": { luas: 14.53, sph: 185 },
        "R61": { luas: 14.74, sph: 52 },
        "Q61": { luas: 24.13, sph: 159 },
        "Q60": { luas: 7.73, sph: 142 },
        "Q59": { luas: 23.19, sph: 132 },
        "Q58": { luas: 30.05, sph: 140 },
        "Q57": { luas: 13.29, sph: 133 },
        "R60": { luas: 34.26, sph: 145 },
        "R59": { luas: 39.58, sph: 131 },
        "R58": { luas: 38.79, sph: 128 },
        "R57": { luas: 35.33, sph: 134 },
        "R56": { luas: 27.53, sph: 131 },
        "R55": { luas: 18.62, sph: 129 },
        "R54": { luas: 22.7, sph: 153 },
        "S55": { luas: 41.54, sph: 147 },
        "S54": { luas: 36.88, sph: 140 },
        "S53": { luas: 37.05, sph: 140 },
        "S52": { luas: 30.27, sph: 134 },
        "S51": { luas: 35.4, sph: 135 },
        "S50": { luas: 17.69, sph: 136 }
    };

    const selectedBlocks = Array.from(document.querySelectorAll('#blok .form-check-input:checked')).map(cb => cb.value);
    const akp = parseFloat(document.getElementById('akp').value);
    const bjr = parseFloat(document.getElementById('bjr').value);
    const tenagaKerja = parseInt(document.getElementById('tenagaKerja').value);

    // Validasi jika tidak ada blok yang dipilih
    if (!selectedBlocks.length) {
        alert('Pilih setidaknya satu blok.');
        return;
    }

    let totalLuas = 0, totalPokok = 0;

    selectedBlocks.forEach(blok => {
        totalLuas += blokData[blok].luas;
        totalPokok += Math.round(blokData[blok].luas * blokData[blok].sph);
    });

    const estimasiJjg = Math.round(totalPokok * akp / 100);
    const estimasiBeratKg = Math.round(estimasiJjg * bjr);
    const estimasiBeratTon = (estimasiBeratKg / 1000).toFixed(2);
    const rit = Math.ceil(estimasiBeratKg / 6500);

    const now = new Date();
    const tanggal = `${now.getDate().toString().padStart(2, '0')}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getFullYear()}`;

    const hasil = `
        <h2>Taksasi Panen DMRE Divisi 4</h2>
        <h4>Tanggal: ${tanggal}</h4>
        <hr>
        <strong>Hasil Perhitungan:</strong><br>
        Blok: ${selectedBlocks.join(', ')}<br>
        Luas: ${totalLuas.toFixed(2)} ha<br>
        Total Pokok: ${totalPokok}<br>
        AKP: ${akp}%<br>
        Janjang Panen: ${estimasiJjg}<br>
        BJR: ${bjr} kg<br>
        Estimasi Berat: ${estimasiBeratKg} kg (${estimasiBeratTon} ton)<br>
        Estimasi Pengiriman: ${rit} rit<br>
        Tenaga Kerja Hadir: ${tenagaKerja}
    `;

    // Menampilkan hasil perhitungan di halaman
    const hasilElement = document.getElementById('hasilTaksasi');
    hasilElement.innerHTML = hasil;
    hasilElement.classList.remove('d-none');
});

// Menyalin hasil perhitungan ke clipboard
document.getElementById('copyToClipboard').addEventListener('click', function() {
    const text = document.getElementById('hasilTaksasi').innerText;
    navigator.clipboard.writeText(text).then(() => alert('Disalin ke clipboard!'));
});

// Menyimpan hasil perhitungan ke file Excel
document.getElementById('saveToExcel').addEventListener('click', function() {
    const selectedBlocks = Array.from(document.querySelectorAll('#blok .form-check-input:checked')).map(cb => cb.value);
    const akp = parseFloat(document.getElementById('akp').value);
    const bjr = parseFloat(document.getElementById('bjr').value);
    const tenagaKerja = parseInt(document.getElementById('tenagaKerja').value);
    let totalLuas = 0, totalPokok = 0;

    selectedBlocks.forEach(blok => {
        totalLuas += blokData[blok].luas;
        totalPokok += Math.round(blokData[blok].luas * blokData[blok].sph);
    });

    const estimasiJjg = Math.round(totalPokok * akp / 100);
    const estimasiBeratKg = Math.round(estimasiJjg * bjr);
    const estimasiBeratTon = (estimasiBeratKg / 1000).toFixed(2);
    const rit = Math.ceil(estimasiBeratKg / 6500);

    const now = new Date();
    const fileName = `Taksasi_DMRE_4_${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}.xlsx`;

    const data = [
        ["Blok", "Luas (ha)", "Total Pokok", "AKP", "Janjang Panen", "BJR (kg)", "Berat (kg)", "Berat (ton)", "Rit", "Tenaga Kerja"],
        [
            selectedBlocks.join(', '),
            totalLuas.toFixed(2), totalPokok, akp, estimasiJjg, bjr,
            estimasiBeratKg, estimasiBeratTon, rit, tenagaKerja
        ]
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, fileName);
});
