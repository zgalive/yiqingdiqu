const puppeteer = require('puppeteer');
const fs = require('fs');
const XLSX = require('xlsx-js-style');
// const xlsx = require('node-xlsx').default;
// const xlsxstyle = require('xlsx-style');

(async () => {
    const targetUrl = 'http://sh.bendibao.com/news/gelizhengce/fengxianmingdan.php';

    const browser = await puppeteer.launch({
        headless: true,
        executablePath: '/Program Files/Google/Chrome/Application/chrome.exe',
        // args: [
        //     '--start-fullscreen'
        // ],
        ignoreDefaultArgs: ['--enable-automation'],
    });
    // const context = await browser.createIncognitoBrowserContext();
    const page = await browser.newPage();
    // const page = pages[0];

    await page.setViewport({
        width: 0,
        height: 0,
    });

    await page.goto(targetUrl);
    await page.waitForTimeout(2000);

    const result = await page.evaluate(() => {
        let amount = document.querySelector('.num.gao').innerText;
        amount = `${amount}`.substring(0, amount.length - 1);
        const cityCollection = document.querySelectorAll('.height .detail-message .city');
        const cityAry = [];
        // const areaCollection = document.querySelectorAll('.height .detail-message-show .ditu');
        // const areaAry = [];

        Array.prototype.forEach.call(cityCollection, function (city) {
            const textAry = city.innerText.split(' ');
            const childrenElements = city.parentElement.parentElement.parentElement.querySelectorAll('.ditu');
            const children = [];
            Array.prototype.forEach.call(childrenElements, function (ch) {
                children.push(ch.innerText);
            });
            const cityObj = {
                first: textAry[0],
                second: textAry[1] ? textAry[1] : textAry[0],
                children
            }
            cityAry.push(cityObj);
        })

        const headStyle = {
            fill: {
                fgColor: { rgb: 'F79646' }
            }
        }
        const xlsxData = [
            [
                {
                    v: '??????',
                    s: { font: { bold: true } }
                },
                {
                    v: `${amount}`,
                    s: {
                        font: {
                            color: { rgb: 'FF0000' }
                        }
                    }
                },
                {
                    v: ''
                },
            ],
            [
                { v: '??????', s: headStyle },
                { v: '??????', s: headStyle },
                { v: '??????', s: headStyle }
            ]
        ];
        const sheetOptions = { '!merges': [] };
        let lastRow = 2;
        cityAry.forEach((x, i) => {
            const range1 = {
                s: { c: 0, r: lastRow },
                e: { c: 0, r: lastRow + x.children.length - 1 }
            };
            const range2 = {
                s: { c: 1, r: lastRow },
                e: { c: 1, r: lastRow + x.children.length - 1 }
            };
            const range3 = {
                s: { c: 2, r: lastRow },
                e: { c: 2, r: lastRow + x.children.length - 1 }
            }
            x.children.forEach((ch, idx) => {
                const firstV = idx == 0 ? x.first : '';
                const secondV = idx == 0 ? x.second : '';
                const borderSetting = {
                    style: 'medium',
                    color: { rgb: '000000' }
                };
                const commonSetting = {
                    alignment: { vertical: 'top', horizontal: 'top' },
                    border: {
                        top: {
                            ...borderSetting
                        },
                        left: {
                            ...borderSetting
                        },
                        right: {
                            ...borderSetting
                        },
                        bottom: {
                            ...borderSetting
                        }

                    }
                };
                const s = i % 2 === 0 ? {
                    fill: {
                        fgColor: { rgb: 'FDE9D9' },
                    },
                    ...commonSetting  
                } : {
                    ...commonSetting
                };
                const tmp = [
                    { v: firstV, s },
                    { v: secondV, s },
                    { v: ch, s }
                ];
                xlsxData.push(tmp);
            })
            sheetOptions['!merges'].push(range1);
            sheetOptions['!merges'].push(range2);
            // sheetOptions['!merges'].push(range3);

            lastRow += (x.children.length);
            //??????
            // xlsxData.push([]);
        })

        return { xlsxData, sheetOptions };
    });

    const worksheet = XLSX.utils.aoa_to_sheet(result.xlsxData);
    worksheet['!merges'] = result.sheetOptions['!merges'];
    
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "?????????");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([]), "?????????");

    const currentDay = new Date();
    const fileName = `????????????????????????????????????????????????${currentDay.getFullYear()}.${currentDay.getMonth()+1}.${currentDay.getDate()}.xlsx`;
    XLSX.writeFile(workbook, fileName, { compression: true });

    browser.close();
})();