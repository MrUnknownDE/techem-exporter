require('dotenv').config();
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const dayjs = require('dayjs');

async function runExport() {
    const isDebug = process.env.DEBUG_MODE === 'true';

    const config = {
        email: process.env.TECHEM_EMAIL,
        password: process.env.TECHEM_PASSWORD,
        unitId: process.env.TECHEM_UNIT_ID,
        startDate: process.env.START_DATE,
        endDate: process.env.END_DATE,
        pageTimeout: parseInt(process.env.TIMEOUT_PAGE, 10) || 60000,
        selectorTimeout: parseInt(process.env.TIMEOUT_SELECTOR, 10) || 15000,
    };

    console.log(`🚀 Techem Export (Horizontal-Layout) gestartet...`);

    const browser = await chromium.launch({ 
        headless: !isDebug,    
        slowMo: isDebug ? 250 : 0 
    });
    
    const context = await browser.newContext();
    const page = await context.newPage();

    try {
        // --- LOGIN FLOW ---
        console.log('🔑 Login-Prozess...');
        await page.goto('https://mieter.techem.de/', { waitUntil: 'networkidle', timeout: config.pageTimeout });
        
        // Cookie-Banner weg
        try {
            const cookieBtn = await page.waitForSelector('#CybotCookiebotDialogBodyButtonDecline', { timeout: 5000 });
            if (cookieBtn) await cookieBtn.click();
        } catch (e) {}

        // Login-Bereich öffnen & Formular ausfüllen
        const loginTrigger = page.locator('text=/login|anmelden/i').first();
        await loginTrigger.click();
        
        await page.waitForSelector('input[type="email"]', { timeout: config.selectorTimeout });
        await page.fill('input[type="email"]', config.email);
        await page.fill('input[type="password"]', config.password);
        await page.click('button[type="submit"]');

        await page.waitForURL(/.*consumptions.*/, { timeout: config.pageTimeout });
        console.log('✅ Login erfolgreich.');

        // --- DATEN EXTRAKTION ---
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
                // 1. Heizung abrufen
                const urlHeating = `https://mieter.techem.de/en/${config.unitId}/consumptions/heating/${date}`;
                // Turbo an: Wir warten nur auf das DOM, nicht auf Tracking-Skripte
                await page.goto(urlHeating, { waitUntil: 'domcontentloaded', timeout: config.pageTimeout });
                await page.waitForTimeout(1000); // Kurze 1-Sekunden-Atempause für React
                
                const bodyHeating = await page.innerText('body');
                const matchH = bodyHeating.match(/([\d,.]+)\s*(kWh|units)/i);
                if (matchH) {
                    hWert = parseFloat(matchH[1].replace(',', '.'));
                    hUnit = matchH[2];
                }

                // 2. Warmwasser abrufen
                const urlWater = `https://mieter.techem.de/en/${config.unitId}/consumptions/hot-water/${date}`;
                await page.goto(urlWater, { waitUntil: 'domcontentloaded', timeout: config.pageTimeout });
                await page.waitForTimeout(1000);
                
                const bodyWater = await page.innerText('body');
                const matchW = bodyWater.match(/([\d,.]+)\s*(m³|kWh)/i);
                if (matchW) {
                    wWert = parseFloat(matchW[1].replace(',', '.'));
                    wUnit = matchW[2];
                }

                // Daten zusammenführen
                results.push({
                    monat: date,
                    water: wWert,
                    waterUnit: wUnit,
                    heating: hWert,
                    heatingUnit: hUnit
                });
                
                console.log(`   ✅ HZ: ${hWert} ${hUnit} | WW: ${wWert} ${wUnit}`);

            } catch (err) {
                console.error(`⚠️ Fehler bei ${date}: ${err.message}`);
            }
        }

        // --- EXCEL DESIGN ---
        console.log('📊 Erstelle übersichtliche Excel-Datei...');
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Verbrauchsübersicht');

        sheet.columns = [
            { header: 'Monat', key: 'monat', width: 15 },
            { header: 'Verbrauch Warmwasser', key: 'water', width: 25 },
            { header: 'Einheit WW', key: 'waterUnit', width: 12 },
            { header: 'Verbrauch Heizung', key: 'heating', width: 25 },
            { header: 'Einheit HZ', key: 'heatingUnit', width: 12 }
        ];

        // Styling (Blauer Header, alternierende Zeilen)
        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00549F' } };
        sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

        results.forEach((data, index) => {
            const row = sheet.addRow(data);
            row.alignment = { vertical: 'middle', horizontal: 'center' };
            if (index % 2 !== 0) {
                row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
            }
        });

        sheet.autoFilter = 'A1:E1';
        sheet.views = [{ state: 'frozen', ySplit: 1 }];

        const fileName = `Techem_Smart_Export_${dayjs().format('YYYY-MM-DD')}.xlsx`;
        await workbook.xlsx.writeFile(fileName);

        console.log(`✨ Datei wurde generiert: ${fileName}`);

    } catch (err) {
        console.error('❌ Kritischer Fehler:', err.message);
    } finally {
        await browser.close();
    }
}

runExport();