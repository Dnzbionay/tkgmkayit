const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;
const SIFRE = process.env.SIFRE || 'ega2024';

app.use(express.json({ limit: '50mb' }));

function sifreKontrol(req, res, next) {
    const sifre = req.headers['x-sifre'] || req.query.sifre;
    if (sifre !== SIFRE) return res.status(401).json({ hata: 'Gecersiz sifre' });
    next();
}

app.use(express.static('public'));

// ============================================================
// BASLIK ESLESTIRME - Farkli Excel basliklarini otomatik tanir
// ============================================================
function normalizeTurkce(s) {
    return s.replace(/İ/g, 'i').replace(/I/g, 'i')
        .replace(/ğ/g, 'g').replace(/ü/g, 'u')
        .replace(/ş/g, 's').replace(/ı/g, 'i')
        .replace(/ö/g, 'o').replace(/ç/g, 'c')
        .replace(/â/g, 'a').replace(/î/g, 'i')
        .replace(/û/g, 'u').replace(/ô/g, 'o')
        .toLowerCase().trim();
}

const KOLON_MAP = {
    'kec sn': 'kecSn', 'seri no': 'kecSn', 'cihaz seri no': 'kecSn',
    'cihaz seri numarasi': 'kecSn', 'cihaz serino': 'kecSn',
    'is yeri no': 'isYeriNo', 'isyeri no': 'isYeriNo', 'is yeri numarasi': 'isYeriNo',
    'yer no': 'isYeriNo', 'is no': 'isYeriNo',
    'is yeri adi': 'isYeriAdi', 'isyeri adi': 'isYeriAdi', 'tapu mudurlugu': 'isYeriAdi',
    'tapu mudur': 'isYeriAdi', 'is yeri': 'isYeriAdi', 'isyeri': 'isYeriAdi',
    'ad': 'isYeriAdi', 'islem yeri': 'isYeriAdi',
    'adres': 'adres', 'is yeri adresi': 'adres', 'isyeri adresi': 'adres',
    'ilce': 'ilce', 'semt/koy': 'ilce', 'semt/koy': 'ilce',
    'il': 'il', 'sehir': 'il',
    'tel1': 'tel1', 'tel 1': 'tel1', 'telefon': 'tel1', 'telefon 1': 'tel1',
    'ceptel': 'tel1', 'cep tel': 'tel1', 'tel no': 'tel1', 'tel': 'tel1',
    'terminal no': 'terminalNo', 'terminal': 'terminalNo', 'term no': 'terminalNo',
    'konu': 'konu', 'islem konusu': 'konu',
    'aciklama': 'aciklama', 'not': 'aciklama', 'notlar': 'aciklama',
    'ariza aciklamasi': 'aciklama', 'ariza': 'aciklama',
    'firma id': 'firmaId', 'firma no': 'firmaId', 'firma numarasi': 'firmaId',
    'anahtar': 'anahtar', 'aktivasyon anahtari': 'anahtar', 'key': 'anahtar',
};

function baslikEslestir(baslikSatiri) {
    const eslesme = {};
    for (let i = 0; i < baslikSatiri.length; i++) {
        const raw = normalizeTurkce(String(baslikSatiri[i] || ''));
        if (!raw) continue;
        const eslesen = KOLON_MAP[raw];
        if (eslesen) {
            eslesme[eslesen] = i;
        }
    }
    return eslesme;
}

function htmlTableParseBuffer(buf) {
    if (buf[0] === 0xFF && buf[1] === 0xFE) {
        const html = buf.toString('utf16le');
        if (!html.includes('<table')) return null;
        const rows = [];
        let i = 0;
        while (i < html.length) {
            const trS = html.indexOf('<tr', i);
            if (trS === -1) break;
            const trE = html.indexOf('</tr>', trS);
            if (trE === -1) break;
            const tr = html.substring(trS, trE);
            const cells = [];
            let j = 0;
            while (j < tr.length) {
                const tdS = tr.indexOf('<td', j);
                if (tdS === -1) break;
                const tdOpenE = tr.indexOf('>', tdS);
                if (tdOpenE === -1) break;
                const tdC = tr.indexOf('</td>', tdOpenE);
                if (tdC === -1) break;
                cells.push(tr.substring(tdOpenE + 1, tdC).replace(/<[^>]*>/g, '').trim());
                j = tdC + 5;
            }
            if (cells.length > 0) rows.push(cells);
            i = trE + 5;
        }
        return rows;
    }
    return null;
}

function htmlTableParse(filePath) {
    const buf = fs.readFileSync(filePath);
    return htmlTableParseBuffer(buf);
}

function cihazlariOku(filePath) {
    const cihazlar = [];
    if (!fs.existsSync(filePath)) return cihazlar;

    let data;
    const htmlRows = htmlTableParse(filePath);
    if (htmlRows && htmlRows.length > 1) {
        data = htmlRows;
    } else {
        const wb = XLSX.readFile(filePath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    }

    if (data.length < 2) return cihazlar;

    // EGATAPU formati: ilk satir baslik degil, baslik 2. satirda
    let baslikSatir = 0;
    const ilkBaslik = String(data[0][0] || '').trim();
    if (ilkBaslik.includes('EGATAPU') || ilkBaslik.includes('Aktif Terminal')) {
        baslikSatir = 1;
    }

    const eslesme = baslikEslestir(data[baslikSatir]);

    for (let i = baslikSatir + 1; i < data.length; i++) {
        const satir = data[i];
        if (!satir || satir.length === 0) continue;

        const kecSn = eslesme.kecSn !== undefined ? String(satir[eslesme.kecSn] || '').trim() : '';
        if (!kecSn) continue;

        cihazlar.push({
            konu: eslesme.konu !== undefined ? (satir[eslesme.konu] || 'Ariza') : 'Ariza',
            isYeriNo: eslesme.isYeriNo !== undefined ? String(satir[eslesme.isYeriNo] || '') : '',
            isYeriAdi: eslesme.isYeriAdi !== undefined ? (satir[eslesme.isYeriAdi] || '') : '',
            adres: eslesme.adres !== undefined ? (satir[eslesme.adres] || '') : '',
            ilce: eslesme.ilce !== undefined ? (satir[eslesme.ilce] || '') : '',
            il: eslesme.il !== undefined ? (satir[eslesme.il] || '') : '',
            tel1: eslesme.tel1 !== undefined ? String(satir[eslesme.tel1] || '') : '',
            terminalNo: eslesme.terminalNo !== undefined ? (satir[eslesme.terminalNo] || '') : '',
            kecSn: kecSn,
            aciklama: eslesme.aciklama !== undefined ? (satir[eslesme.aciklama] || '') : ''
        });
    }
    return cihazlar;
}

function anahtarlariOku(filePath) {
    const anahtarlar = [];
    if (!fs.existsSync(filePath)) return anahtarlar;

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (data.length < 3) return anahtarlar;

    for (let i = 2; i < data.length; i++) {
        if (data[i] && data[i][0] && data[i][1]) {
            const eslesme = baslikEslestir(data[1] || []);
            anahtarlar.push({
                firmaId: data[i][0],
                tapuMudurlugu: String(data[i][1]),
                anahtar: data[i][2] || ''
            });
        }
    }
    return anahtarlar;
}

function loadData() {
    let cihazDosya = path.join(__dirname, 'TKGM Aktif cihazlar.xlsx');
    const xlsDosyalar = fs.readdirSync(__dirname).filter(f => f.startsWith('report') && (f.endsWith('.xls') || f.endsWith('.xlsx')));
    if (xlsDosyalar.length > 0) {
        const enYeni = xlsDosyalar.sort().pop();
        cihazDosya = path.join(__dirname, enYeni);
    }
    return {
        cihazlar: cihazlariOku(cihazDosya),
        anahtarlar: anahtarlariOku(path.join(__dirname, 'TKGM Aktivasyon Anahtarları.xlsx')),
        cihazKaynagi: path.basename(cihazDosya)
    };
}

// ============================================================
// API'LER
// ============================================================

app.get('/api/tapular', sifreKontrol, (req, res) => {
    const { cihazlar } = loadData();
    res.json([...new Set(cihazlar.map(c => c.isYeriAdi))].sort());
});

app.get('/api/tapu-bilgi/:tapuAdi', sifreKontrol, (req, res) => {
    const { cihazlar, anahtarlar } = loadData();
    const tapuAdi = decodeURIComponent(req.params.tapuAdi);
    const tapuCihazlari = cihazlar.filter(c => c.isYeriAdi === tapuAdi);
    const anahtar = anahtarlar.find(a => a.tapuMudurlugu === tapuAdi);
    if (tapuCihazlari.length === 0) return res.status(404).json({ hata: 'Tapu bulunamadi' });
    res.json({ tapu: tapuCihazlari[0], cihazlar: tapuCihazlari, anahtar: anahtar || null });
});

// YENI: Excel yukle ve basliklari analiz et
app.post('/api/excel-yukle', sifreKontrol, (req, res) => {
    const { dosyaIcerigi, dosyaAdi } = req.body;
    if (!dosyaIcerigi) return res.status(400).json({ hata: 'Dosya icerigi gerekli' });

    try {
        const buffer = Buffer.from(dosyaIcerigi, 'base64');

        let data;
        const htmlRows = htmlTableParseBuffer(buffer);
        if (htmlRows && htmlRows.length > 1) {
            data = htmlRows;
        } else {
            const wb = XLSX.read(buffer, { type: 'buffer' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        }

        if (data.length < 1) return res.status(400).json({ hata: 'Dosya bos' });

        // EGATAPU formati: ilk satir baslik degil, baslik 2. satirda
        let baslikSatir = 0;
        const ilkBaslik = String(data[0][0] || '').trim();
        if (ilkBaslik.includes('EGATAPU') || ilkBaslik.includes('Aktif Terminal')) {
            baslikSatir = 1;
        }

        const basliklar = data[baslikSatir] || [];
        const eslesme = baslikEslestir(basliklar);

        const sonuc = {
            dosyaAdi: dosyaAdi,
            toplamSatir: data.length - 1,
            basliklar: basliklar.map((b, i) => {
                const normalized = normalizeTurkce(String(b || ''));
                return {
                    orijinal: b,
                    eslesen: KOLON_MAP[normalized] || null,
                    kolonIndex: i
                };
            }),
            eslesenKolonlar: Object.keys(eslesme),
            ornekVeri: data.length > 1 ? data[1] : null
        };

        res.json(sonuc);
    } catch (e) {
        res.status(500).json({ hata: 'Dosya okunamadi: ' + e.message });
    }
});

// Excel'i kaydet - dosyayi orijinal adiyla dizine kaydeder
app.post('/api/excel-kaydet', sifreKontrol, (req, res) => {
    const { dosyaIcerigi, dosyaAdi } = req.body;
    if (!dosyaIcerigi || !dosyaAdi) return res.status(400).json({ hata: 'Eksik bilgi' });

    try {
        const buffer = Buffer.from(dosyaIcerigi, 'base64');

        const hedefDosya = path.join(__dirname, dosyaAdi);
        fs.writeFileSync(hedefDosya, buffer);

        const guncel = loadData();
        res.json({
            durum: 'ok',
            cihazSayisi: guncel.cihazlar.length,
            kaynak: guncel.cihazKaynagi,
            mesaj: guncel.cihazlar.length + ' cihaz yuklendi (' + guncel.cihazKaynagi + ')'
        });
    } catch (e) {
        res.status(500).json({ hata: 'Kaydedilemedi: ' + e.message });
    }
});

function buildMailHTML(tabloSatirlari) {
    let tablo = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:12pt;">';
    tablo += '<tr style="background:#2c3e50;color:white;font-weight:bold;">';
    ['Konu','Is Yeri No','Is Yeri Adi','Adres','Ilce','Il','TEL1','Terminal No','KEC SN','Aciklama'].forEach(h => {
        tablo += '<th>' + h + '</th>';
    });
    tablo += '</tr>';
    for (const s of tabloSatirlari) {
        tablo += '<tr>';
        tablo += '<td>' + s.konu + '</td><td>' + s.isYeriNo + '</td><td>' + s.isYeriAdi + '</td>';
        tablo += '<td>' + s.adres + '</td><td>' + s.ilce + '</td><td>' + s.il + '</td>';
        tablo += '<td>' + s.tel1 + '</td><td>' + s.terminalNo + '</td><td>' + s.kecSn + '</td>';
        tablo += '<td>' + s.aciklama + '</td></tr>';
    }
    tablo += '</table>';
    return tablo;
}

app.post('/api/ariza-mail', sifreKontrol, (req, res) => {
    const { kayitlar } = req.body;
    if (!kayitlar || !Array.isArray(kayitlar) || kayitlar.length === 0) {
        return res.status(400).json({ hata: 'En az bir kayit gerekli' });
    }

    const { cihazlar, anahtarlar } = loadData();
    const tabloSatirlari = [];
    const hatalar = [];

    for (const kayit of kayitlar) {
        const { tapuAdi, kecSn, aciklama } = kayit;
        if (!tapuAdi || !kecSn || !aciklama) { hatalar.push('Eksik: ' + (kecSn || '?')); continue; }
        const cihaz = cihazlar.find(c => c.kecSn === kecSn && c.isYeriAdi === tapuAdi);
        if (!cihaz) { hatalar.push(kecSn + ' bulunamadi'); continue; }
        const anahtar = anahtarlar.find(a => a.tapuMudurlugu === tapuAdi);
        if (!anahtar) { hatalar.push(tapuAdi + ' anahtari bulunamadi'); continue; }

        tabloSatirlari.push({
            konu: cihaz.konu, isYeriNo: cihaz.isYeriNo, isYeriAdi: cihaz.isYeriAdi,
            adres: cihaz.adres, ilce: cihaz.ilce, il: cihaz.il, tel1: cihaz.tel1,
            terminalNo: cihaz.terminalNo, kecSn: cihaz.kecSn,
            aciklama: aciklama + '. Arizali KEC pasif edilirken Firma no: 1 Anahtar: 4a1a0f9b Kurulacak KEC Aktif ederken Firma No:' + anahtar.firmaId + ' Anahtar:' + anahtar.anahtar
        });
    }

    if (tabloSatirlari.length === 0) {
        return res.status(400).json({ hata: 'Gecerli kayit bulunamadi', detay: hatalar });
    }

    const htmlTablo = buildMailHTML(tabloSatirlari);
    const mailHTML = 'Merhaba,<br><br>Asagida bilgileri yer alan tapu mudurlukleri icin kayit acilmasi konusunda yardimlariniz rica olunur.<br><br>Tesekkurler,<br><br>' + htmlTablo;
    const plainText = 'Merhaba,\n\nAsagida bilgileri yer alan tapu mudurlukleri icin kayit acilmasi konusunda yardimlariniz rica olunur.\n\nTesekkurler,\n\n' + tabloSatirlari.map(s => s.isYeriAdi + ' | ' + s.kecSn + ' | ' + s.terminalNo + ' | ' + s.ilce + ' | ' + s.tel1).join('\n');

    res.json({ konu: 'TKGM ARIZA KAYIT', html: mailHTML, plainText: plainText, kayitSayisi: tabloSatirlari.length, hatalar: hatalar.length > 0 ? hatalar : null });
});

app.post('/api/outlook-ac', sifreKontrol, (req, res) => {
    const { html, konu } = req.body;
    if (!html) return res.status(400).json({ hata: 'HTML gerekli' });

    if (process.platform !== 'win32') {
        return res.json({ durum: 'ok', html: html, konu: konu, bulut: true });
    }

    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

    const htmPath = path.join(tempDir, 'ariza_mail.htm');
    fs.writeFileSync(htmPath, html, 'utf8');

    const psScript = `$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = '${konu || 'TKGM ARIZA KAYIT'}'
$mail.HTMLBody = [System.IO.File]::ReadAllText('${htmPath.replace(/\\/g, '\\\\')}', [System.Text.Encoding]::UTF8)
$mail.Display()`;

    const psPath = path.join(tempDir, 'outlook_ac.ps1');
    fs.writeFileSync(psPath, psScript, 'utf8');

    try {
        execSync('powershell -ExecutionPolicy Bypass -File "' + psPath + '"', { windowsHide: true });
        res.json({ durum: 'ok' });
    } catch (e) {
        res.status(500).json({ hata: 'Outlook acilamadi: ' + e.message });
    }
});

app.listen(PORT, () => {
    console.log('TKGM Ariza Kayit Sistemi calisiyor: http://localhost:' + PORT);
});
