/**
 * Вспомогательные методы работы с ячейками ClosedXML / SheetJS
 */
export class ExcelHelpers {
  /**
   * Получение ячейки по 1-based координатам (row, col)
   */
  static getCell(sheet, row, col) {
    const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    return sheet[address] || null;
  }

  static getValue(sheet, row, col) {
    const cell = this.getCell(sheet, row, col);
    if (!cell || cell.v === undefined || cell.v === null) return "";
    return cell.v;
  }

  static isEmpty(sheet, row, col) {
    const val = this.getValue(sheet, row, col);
    return val === "" || val === null || val === undefined;
  }

  /**
   * Запись значения в ячейку (1-based)
   */
  static setValue(sheet, row, col, value, format = null) {
    const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
    const cell = { v: value };

    if (typeof value === "number") {
      cell.t = "n";
    } else if (typeof value === "boolean") {
      cell.t = "b";
    } else if (value instanceof Date) {
      cell.t = "d";
    } else {
      cell.t = "s";
    }

    if (format) {
      cell.z = format;
    }

    sheet[address] = cell;

    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    if (row - 1 < range.s.r) range.s.r = row - 1;
    if (col - 1 < range.s.c) range.s.c = col - 1;
    if (row - 1 > range.e.r) range.e.r = row - 1;
    if (col - 1 > range.e.c) range.e.c = col - 1;
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }

  /**
   * Преобразование ячейки в объект Date
   */
  static getDateValue(sheet, row, col) {
    const cell = this.getCell(sheet, row, col);
    if (!cell || cell.v === undefined || cell.v === "") return null;

    if (cell.v instanceof Date) return cell.v;

    if (typeof cell.v === "number") {
      return this.excelSerialToDate(cell.v);
    }

    if (typeof cell.v === "string") {
      if (cell.v.includes(":")) {
        const parts = cell.v.split(":");
        return new Date(1975, 0, 1, parseInt(parts[0], 10), parseInt(parts[1], 10), 0);
      }
      const parsed = Date.parse(cell.v);
      if (!isNaN(parsed)) return new Date(parsed);
    }

    return null;
  }

  static excelSerialToDate(serial) {
    if (serial < 1) {
      const totalSeconds = Math.round(serial * 86400);
      const h = Math.floor(totalSeconds / 3600);
      const m = Math.floor((totalSeconds % 3600) / 60);
      const s = totalSeconds % 60;
      return new Date(1975, 0, 1, h, m, s);
    }
    const utcDays = Math.floor(serial - 25569);
    const remainder = serial - Math.floor(serial);
    const date = new Date(utcDays * 86400 * 1000);
    const totalSeconds = Math.round(remainder * 86400);
    date.setHours(Math.floor(totalSeconds / 3600), Math.floor((totalSeconds % 3600) / 60), totalSeconds % 60);
    return date;
  }

  static timeToTotalDays(date) {
    if (!date || isNaN(date.getTime())) return 0;
    const h = date.getHours();
    const m = date.getMinutes();
    const s = date.getSeconds();
    return (h * 3600 + m * 60 + s) / 86400;
  }

  static getMaxRowInColumn(sheet, colIndex) {
    if (!sheet['!ref']) return 0;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    let max = 0;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: colIndex - 1 })];
      if (cell && cell.v !== undefined && cell.v !== "") {
        max = r + 1;
      }
    }
    return max;
  }
}