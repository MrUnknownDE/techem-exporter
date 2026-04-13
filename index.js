require('dotenv').config();
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const dayjs = require('dayjs');

async function runExport() {
    const isDebug = process.env.DEBUG_MODE === 'true';

    // Konfiguration aus .env laden
    const config = {
        email: process.env.TECHEM_EMAIL,
        password: process.env.TECHEM_PASSWORD,
        unitId: process.env.TECHEM_UNIT_ID,
        startDate: process.env.START_DATE,
        endDate: process.env.END_DATE,
        pageTimeout: parseInt(process.env.TIMEOUT_PAGE, 10) || 60000,
        selectorTimeout: parseInt(process.env.TIMEOUT_SELECTOR, 10) || 15000,
    };

    console.log(`🚀 Techem Export gestartet... (Debug: ${isDebug ? 'AN' : 'AUS'})`);

    const browser = await chromium.launch({ 
        headless: !isDebug,    
        slowMo: isDebug ? 250 : 0 
    });
    
    const context = await browser.newContext();
    const page = await context.newPage();

    try {
        // ==========================================
        // 1. LOGIN & AUTHENTIFIZIERUNG
        // ==========================================
        console.log('🔑 Login-Prozess startet...');
        await page.goto('https://mieter.techem.de/', { waitUntil: 'networkidle', timeout: config.pageTimeout });
        
        // Cookie-Banner killen
        try {
            const cookieBtn = await page.waitForSelector('#CybotCookiebotDialogBodyButtonDecline', { timeout: 5000 });
            if (cookieBtn) {
                await cookieBtn.click();
                await page.waitForTimeout(1000); // Kurz warten bis Animation weg ist
            }
        } catch (e) { /* Kein Banner da */ }

        // Login-Bereich öffnen & Formular ausfüllen
        const loginTrigger = page.locator('text=/login|anmelden/i').first();
        await loginTrigger.click();
        
        await page.waitForSelector('input[type="email"]', { timeout: config.selectorTimeout });
        await page.fill('input[type="email"]', config.email);
        await page.fill('input[type="password"]', config.password);
        await page.click('button[type="submit"]');

        console.log('⏳ Warte auf Dashboard-Weiterleitung...');
        await page.waitForURL(/.*consumptions.*/, { timeout: config.pageTimeout });
        console.log('✅ Login erfolgreich.');

        // ==========================================
        // 2. DATEN EXTRAKTION
        // ==========================================
        const months = [];
        let current = dayjs(config.startDate);
        const end = dayjs(config.endDate);
        while (current.isBefore(end) || current.isSame(end)) {
            months.push(current.format('YYYY-MM'));
            current = current.add(1, 'month');
        }

        const results = [];

        for (const date of months) {
            console.log(`📡 Lade Daten für ${date}...`);
            
            let hWert = 0, hUnit = 'kWh';
            let wWert = 0, wUnit = 'm³';

            try {
                // Heizung
                const urlHeating = `https://mieter.techem.de/en/${config.unitId}/consumptions/heating/${date}`;
                await page.goto(urlHeating, { waitUntil: 'domcontentloaded', timeout: config.pageTimeout });
                await page.waitForTimeout(1000); // React kurz rendern lassen
                
                const bodyHeating = await page.innerText('body');
                const matchH = bodyHeating.match(/([\d,.]+)\s*(kWh|units)/i);
                if (matchH) {
                    hWert = parseFloat(matchH[1].replace(',', '.'));
                    hUnit = matchH[2];
                }

                // Warmwasser
                const urlWater = `https://mieter.techem.de/en/${config.unitId}/consumptions/hot-water/${date}`;
                await page.goto(urlWater, { waitUntil: 'domcontentloaded', timeout: config.pageTimeout });
                await page.waitForTimeout(1000);
                
                const bodyWater = await page.innerText('body');
                const matchW = bodyWater.match(/([\d,.]+)\s*(m³|kWh)/i);
                if (matchW) {
                    wWert = parseFloat(matchW[1].replace(',', '.'));
                    wUnit = matchW[2];
                }

                results.push({ monat: date, water: wWert, waterUnit: wUnit, heating: hWert, heatingUnit: hUnit });
                console.log(`   ✅ HZ: ${hWert} ${hUnit} | WW: ${wWert} ${wUnit}`);

            } catch (err) {
                console.error(`⚠️ Fehler bei ${date}: ${err.message}`);
            }
        }

        // ==========================================
        // 3. EXCEL EXPORT (Monatlich, Horizontal)
        // ==========================================
        console.log('📊 Erstelle Excel-Datei...');
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Verbrauchsübersicht');

        sheet.columns = [
            { header: 'Monat', key: 'monat', width: 15 },
            { header: 'Verbrauch Warmwasser', key: 'water', width: 25 },
            { header: 'Einheit WW', key: 'waterUnit', width: 12 },
            { header: 'Verbrauch Heizung', key: 'heating', width: 25 },
            { header: 'Einheit HZ', key: 'heatingUnit', width: 12 }
        ];

        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00549F' } };
        sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

        results.forEach((data, index) => {
            const row = sheet.addRow(data);
            row.alignment = { vertical: 'middle', horizontal: 'center' };
            if (index % 2 !== 0) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        });

        sheet.autoFilter = 'A1:E1';
        sheet.views = [{ state: 'frozen', ySplit: 1 }];

        const excelName = `Techem_Export_${dayjs().format('YYYY-MM-DD')}.xlsx`;
        await workbook.xlsx.writeFile(excelName);
        console.log(`✨ Excel gespeichert: ${excelName}`);

        // ==========================================
        // 4. PDF EXPORT (Jährlich mit Graphen)
        // ==========================================
        if (results.length > 0) {
            console.log('📄 Erstelle PDF-Report...');
            
            // Daten nach Jahren aggregieren
            const yearlyStats = results.reduce((acc, curr) => {
                const year = curr.monat.split('-')[0];
                if (!acc[year]) acc[year] = { year, heating: 0, water: 0, heatingUnit: curr.heatingUnit, waterUnit: curr.waterUnit };
                acc[year].heating += curr.heating;
                acc[year].water += curr.water;
                return acc;
            }, {});
            const statsArray = Object.values(yearlyStats);

            const htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
                <style>
                    body { font-family: Arial, sans-serif; padding: 40px; color: #333; }
                    h1 { color: #00549F; border-bottom: 2px solid #00549F; padding-bottom: 10px; }
                    .info { margin-bottom: 30px; font-size: 0.9em; color: #666; }
                    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                    th, td { border: 1px solid #ddd; padding: 12px; text-align: center; }
                    th { background-color: #f4f4f4; }
                    .chart-container { width: 100%; height: 400px; margin: 40px 0; }
                </style>
            </head>
            <body>
                <h1>Techem Jahres-Report</h1>
                <div class="info">
                    Einheit-ID: ${config.unitId}<br>
                    Zeitraum: ${config.startDate} bis ${config.endDate}<br>
                    Erstellt am: ${dayjs().format('DD.MM.YYYY HH:mm')}
                </div>

                <div class="chart-container">
                    <canvas id="myChart"></canvas>
                </div>

                <table>
                    <thead>
                        <tr>
                            <th>Jahr</th>
                            <th>Heizung Gesamt (${results[0].heatingUnit})</th>
                            <th>Warmwasser Gesamt (${results[0].waterUnit})</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${statsArray.map(s => `
                            <tr>
                                <td><b>${s.year}</b></td>
                                <td>${s.heating.toFixed(2)}</td>
                                <td>${s.water.toFixed(2)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>

                <script>
                    const ctx = document.getElementById('myChart').getContext('2d');
                    new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: ${JSON.stringify(statsArray.map(s => s.year))},
                            datasets: [{
                                label: 'Heizung',
                                data: ${JSON.stringify(statsArray.map(s => s.heating))},
                                backgroundColor: '#00549F',
                                yAxisID: 'y'
                            }, {
                                label: 'Warmwasser',
                                data: ${JSON.stringify(statsArray.map(s => s.water))},
                                backgroundColor: '#5CB85C',
                                yAxisID: 'y1'
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            animation: false, // WICHTIG FÜR PDF: Keine Animationen beim Rendern!
                            scales: {
                                y: { type: 'linear', position: 'left', title: { display: true, text: 'Heizung' } },
                                y1: { type: 'linear', position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Wasser' } }
                            }
                        }
                    });
                </script>
            </body>
            </html>`;

            const reportPage = await context.newPage();
            // networkidle stellt sicher, dass das externe Chart.js CDN geladen wurde
            await reportPage.setContent(htmlContent, { waitUntil: 'networkidle' });
            
            const pdfName = `Techem_Report_${dayjs().format('YYYY-MM-DD')}.pdf`;
            await reportPage.pdf({
                path: pdfName,
                format: 'A4',
                printBackground: true,
                margin: { top: '20px', bottom: '20px', left: '20px', right: '20px' }
            });
            console.log(`✨ PDF gespeichert: ${pdfName}`);
        }

    } catch (err) {
        console.error('❌ Kritischer Fehler im Skript:', err.message);
    } finally {
        console.log('🧹 Räume auf und schließe Browser...');
        await browser.close();
    }
}

runExport();