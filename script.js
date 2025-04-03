async function generateReport() {
    const year = document.getElementById("year").value;
    const month = document.getElementById("month").value;
    const name = document.getElementById("FIO").value;



    const daysInMonth = new Date(year, month, 0).getDate();
    const workbook = new ExcelJS.Workbook();

    
    for (let day = 1; day <= daysInMonth; day++) {

        
        const currentDate = new Date(year, month - 1, day); 
        const formattedDate = currentDate.toLocaleDateString('en-GB').replace(/\//g, '.'); 
        const dayOfWeek = currentDate.getDay();
        
        const sheet = workbook.addWorksheet(`${formattedDate}`);

        sheet.getCell("G3").value = name;


        if (dayOfWeek === 0) {
            workbook.removeWorksheet(sheet.id);
            continue;
        }

        if (dayOfWeek === 6) {

            for (let row = 4; row <= 24; row++) {
                const cell = sheet.getCell(`B${row}`);
                cell.value = "ОФИСНЫЙ ДЕНЬ";
                cell.font = { bold: true, size: 14 };
                cell.alignment = { vertical: "middle", horizontal: "center" };
            }
    
            for (let row = 27; row <= 41; row++) {
                const cell = sheet.getCell(`B${row}`);
                cell.value = "ОФИСНЫЙ ДЕНЬ";
                cell.font = { bold: true, size: 14 };
                cell.alignment = { vertical: "middle", horizontal: "center" };
            }
        }


        

        const columnWidths = [7, 42, 30, 30, 30, 41, 60, 30, 30];
        sheet.columns = columnWidths.map(width => ({ width }));


        sheet.getRow(1).height = 22;
        sheet.getRow(2).height = 22;
        sheet.getRow(3).height = 40;
        sheet.getRow(4).height = 60;
        sheet.getRow(25).height = 35;
        sheet.getRow(26).height = 55;



        sheet.mergeCells("A1:I2");
        const titleCell = sheet.getCell("A1");
        titleCell.value = "DAILY VISITACTIVITY REPORT (TASHRIF FAOLIYATI BO'YICHA KUNLIK HISOBOT)";
        titleCell.font = { bold: true, size: 20 };
        titleCell.alignment = { vertical: "middle", horizontal: "center" };


        sheet.mergeCells("A25:A26");
        sheet.mergeCells("A3:A4");
        const secondSheet = sheet.getCell("A3");
        secondSheet.value = "S. NO";
        secondSheet.font = { bold: true, size: 11 };
        secondSheet.alignment = { vertical: "middle", horizontal: "center" };

        sheet.mergeCells("B3:D3");
        sheet.getCell("B3:D3").value = "HOSPITALS (SHIFOXONALAR)"
        sheet.getCell("B3:D3").font = { bold: true, size: 14 };
        sheet.getCell("B3:D3").alignment = { vertical: 'middle', horizontal: 'center' };

        sheet.mergeCells("E3:F3");
        sheet.getCell("E3").value = "MR NAME  ( TIBBIY VAKIL TO'LIQ ISM)";
        sheet.getCell("E3").font = { bold: true, size: 14 };
        sheet.getCell("E3").alignment = { vertical: "middle", horizontal: "center" };

        sheet.getCell("H3").value = "DATE (SANA):";
        sheet.getCell("H3").font = { bold: true, size: 14 };
        sheet.getCell("H3").alignment = { vertical: "middle", horizontal: "center" };

        sheet.getCell("G3").font = { bold: true, size: 14, color: { argb: "FFFF0000" } };
        sheet.getCell("G3").alignment = { vertical: "middle", horizontal: "center" };

        sheet.getCell("I3").value = `${day.toString().padStart(2, "0")}.${month.padStart(2, "0")}.${year}`;
        sheet.getCell("I3").font = { bold: true, size: 14 }; 
        sheet.getCell("I3").alignment = { vertical: "middle", horizontal: "center" }; 


        const row4Content = [
            {
                richText: [
                    { text: "Full name ", font: { bold: true } },
                    { text: "\n (To'liq ism)", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "Specialty", font: { bold: true } },
                    { text: "\n(Mutaxassisligi)", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "How many visits", font: { bold: true } },
                    { text: "\n have been made", font: { bold: true } },
                    { text: "\n(Qancha tashriflar", font: { bold: false } },
                    { text: "\n amalga oshirildi)", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "Place of work ", font: { bold: true } },
                    { text: "\n( Ish joyi )", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "Promoted product on visit", font: { bold: true } },
                    { text: "\n(Tashrifda promotsiya qilingan mahsulot)", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "Agreement with the doctor at the end of the visit?", font: { bold: true } },
                    { text: "\n(Tashrif oxirida shifokor bilan kelishuv?)", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "City, district ", font: { bold: true } },
                    { text: "\n( Shahar, tuman )", font: { bold: false } }
                ]
            },
            {
                richText: [
                    { text: "Phone - Email", font: { bold: true } },
                    { text: "\n(Tel nomer - Email)", font: { bold: false } }
                ]
            }
        ];
        

        row4Content.forEach((text, index) => {
            const cell = sheet.getCell(4, index + 2); 
            cell.value = text;
            cell.font = { size: 11 };
            cell.alignment = { vertical: "middle", horizontal: "center" };
        });


        for (let i = 5; i <= 24; i++) {
            sheet.getCell(`A${i}`).value = i - 4;
            sheet.getCell(`A${i}`).font = { size: 11, bold: true };
            sheet.getCell(`A${i}`).alignment = { vertical: "middle", horizontal: "center" };
            sheet.getRow(i).height = 39;
        }




        row4Content.forEach((text, index) => {
            const cell = sheet.getCell(26, index + 2);
            cell.value = text;
            cell.font = { size: 12 };
            cell.alignment = { vertical: "middle", horizontal: "center" };
        });

        for (let i = 27; i <= 41; i++) {
            sheet.getCell(`A${i}`).value = i - 26;
            sheet.getCell(`A${i}`).font = { bold: true, size: 11 };
            sheet.getCell(`A${i}`).alignment = { vertical: "middle", horizontal: "center" };
            sheet.getRow(i).height = 39;
        }
        const blueFill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'ff93cddd' } 
        };
    
        const greenFill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFd9f1cf' } 
        };

        const border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
            vertical: { style: 'thin' },
            horizontal: { style: 'thin' }
        };
    

        for (let row = 1; row <= 4; row++) {
            for (let col = 1; col <= 9; col++) {
                const cell = sheet.getCell(`${String.fromCharCode(65 + col - 1)}${row}`);
                if (!name && row == 3 && col == 7) {
                    cell.fill = greenFill; 
                    cell.protection = { locked: false };
                }
                else {
                    cell.fill = blueFill; 
                    cell.protection = { locked: true }; 
                }

            }
        }
    

        for (let row = 25; row <= 26; row++) {
            for (let col = 1; col <= 9; col++) {
                const cell = sheet.getCell(`${String.fromCharCode(65 + col - 1)}${row}`);
                cell.fill = blueFill;
                cell.protection = { locked: true };
                cell.font = { bold: true, size: 11 }; 
            }
        }

        sheet.mergeCells("B25:D25");
        sheet.getCell("B25").value = "PHARMACIES (DORIXONALAR):";
        sheet.getCell("B25").font = { bold: true, size: 14 };
        sheet.getCell("B25").alignment = { vertical: "middle", horizontal: "center" };

        for (let row = 5; row <= 24; row++) {
            for (let col = 1; col <= 9; col++) {
                const cell = sheet.getCell(`${String.fromCharCode(65 + col - 1)}${row}`);
                cell.fill = greenFill;
                cell.protection = { locked: false };
                cell.font = { bold: true, size: 12 }; 
                cell.alignment = { horizontal: 'center', vertical: 'middle' };  
            }
        }
    
        for (let row = 27; row <= 41; row++) {
            for (let col = 1; col <= 9; col++) {
                const cell = sheet.getCell(`${String.fromCharCode(65 + col - 1)}${row}`);
                cell.fill = greenFill;
                cell.protection = { locked: false };
                cell.font = { bold: true, size: 12 }; 
                cell.alignment = { horizontal: 'center', vertical: 'middle' };  
            }
        }
        for (let row = 1; row <= 41; row++) {
            for (let col = 1; col <= 9; col++) {
                const cell = sheet.getCell(`${String.fromCharCode(65 + col - 1)}${row}`);

                cell.border = border;
            }
        }

        workbook.eachSheet((sheet) => {
            sheet.eachRow((row) => {
                row.eachCell((cell) => {

                    const currentFont = cell.font || {};
                    const currentColor = currentFont.color || {}; 
        

                    cell.font = { 
                        name: 'Arial', 
                        bold: currentFont.bold,   
                        size: currentFont.size,   
                        italic: currentFont.italic, 
                        color: currentColor.argb 
                    };
                });
            });
        });
        
        const pass = "12345"
        sheet.protect({
            password: pass, 
            selectLockedCells: false,
            selectUnlockedCells: true,
            formatCells: false,
            formatColumns: true,
            formatRows: true,
            insertColumns: true,
            insertRows: true,
            deleteColumns: false,
            deleteRows: false
        });
        
    }

    

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    if (!name) {
        link.download = `ОТЧЁТ_${month}_${year}_ШАБЛОН.xlsx`;
    }
    else {
        link.download = `ОТЧЁТ_${month}_${year}_${name}.xlsx`;
    }
    
    link.click();
}
