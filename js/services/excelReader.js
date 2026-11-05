import { Schedule } from '../models/Schedule.js';
import { ExcelHelpers } from '../utils/excelHelpers.js';
import { DictionaryService } from './dictionaryService.js';

export class ExcelReader {
  /**
   * Быстрое чтение без тормозящих пауз в цикле
   */
  static async readWorkbook(file, dictionaries) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    let schedules = [];
    let source = "";
    let sourceLabel = "";
    let sourceColor = "";

    if (firstSheetName === "Data") {
      schedules = this.readSchedulesFromProlog(sheet);
      if (dictionaries) {
        DictionaryService.setCities(schedules, dictionaries);
      }
      schedules = this.sortSchedules(schedules);
      source = "Prolog";
      sourceLabel = "ПроКТОлог";
      sourceColor = "#0000cd";
    } else {
      schedules = this.readSchedulesFromNewsLetter(sheet);
      if (dictionaries) {
        DictionaryService.setAXCities(schedules, dictionaries);
        if (dictionaries.loadFormats) {
          DictionaryService.translateLoadFormatsNewsIntoProlog(dictionaries.loadFormats, schedules);
        }
      }
      schedules.forEach(s => s.setRoute());
      source = "Графіки з розсилки";
      sourceLabel = "Файл з розсилки";
      sourceColor = "#008000";
    }

    return { schedules, source, sourceLabel, sourceColor };
  }

  static readSchedulesFromProlog(sheet) {
    const list = [];
    const maxRow = ExcelHelpers.getMaxRowInColumn(sheet, 5); // колонка E
    for (let r = 3; r <= maxRow; r++) {
      const routeVal = ExcelHelpers.getValue(sheet, r, 5);
      const grfVal = ExcelHelpers.getValue(sheet, r, 192);
      const startDateVal = ExcelHelpers.getValue(sheet, r, 2);
      if (!routeVal && !grfVal && !startDateVal) continue;

      const sched = new Schedule();
      sched.readFromProlog(sheet, r);
      if (sched.points.length > 0 || sched.route || sched.grf) {
        list.push(sched);
      }
    }
    return list;
  }

  static readSchedulesFromNewsLetter(sheet) {
    const list = [];
    const maxRow = ExcelHelpers.getMaxRowInColumn(sheet, 1); // колонка A
    for (let r = 5; r <= maxRow; r++) {
      const routePravoslVal = ExcelHelpers.getValue(sheet, r, 11);
      const grfVal = ExcelHelpers.getValue(sheet, r, 3);
      const startDateVal = ExcelHelpers.getValue(sheet, r, 1);
      if (!routePravoslVal && !grfVal && !startDateVal) continue;

      const sched = new Schedule();
      sched.readFromNewsLetter(sheet, r);
      if (sched.points.length > 0 || sched.routePravosl || sched.grf) {
        list.push(sched);
      }
    }
    return list;
  }

  static sortSchedules(schedules) {
    return [...schedules].sort((a, b) => {
      if (b.logMon !== a.logMon) return b.logMon ? 1 : -1;
      if (b.logTue !== a.logTue) return b.logTue ? 1 : -1;
      if (b.logWed !== a.logWed) return b.logWed ? 1 : -1;
      if (b.logThu !== a.logThu) return b.logThu ? 1 : -1;
      if (b.logFri !== a.logFri) return b.logFri ? 1 : -1;
      if (b.logSat !== a.logSat) return b.logSat ? 1 : -1;
      if (b.logSun !== a.logSun) return b.logSun ? 1 : -1;
      if (a.routePravosl !== b.routePravosl) return (a.routePravosl || "").localeCompare(b.routePravosl || "");
      return (a.t2?.getTime() || 0) - (b.t2?.getTime() || 0);
    });
  }
}