function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getRange("J4:J4587");
    let values = range.getValues();
    
    let ukrainianNamesAndSurnames: string[] = [
        "Oksana", "Taras", "Mykola", "Serhii", "Andrii",
        "Kostiantyn", "Svitlana", "Valentyn", "Petrenko",
        "Danylo", "Kovalenko", "Melnyk",
    ];
    
    let matchCount = 0;
    
    for (let i = 0; i < values.length; i++) {
        let cellValue = values[i][0];
        if (typeof cellValue === 'string') {
            for (let name of ukrainianNamesAndSurnames) {
                if (cellValue.toLowerCase().includes(name.toLowerCase())) {
                    let cell = sheet.getRange(`J${i + 4}`);
                    cell.getFormat().getFill().setColor("yellow");
                    matchCount++;
                    break;
                }
            }
        }
    }
    
    console.log(`Found and highlighted ${matchCount} matches.`);
}
