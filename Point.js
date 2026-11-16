import { ExcelHelpers } from '../utils/excelHelpers.js';

export class Point {
  constructor() {
    this.name = "";
    this.namePravosl = "";
    this.city = "";
    this.cityRing = "";
    this.number = 0;
    this.arrival = null;
    this.departure = null;
    this.waitingTime = null;
    this.movingType = "";
    this.loadRamp1 = 0;
    this.loadRamp2 = 0;
    this.unloadRamp1 = 0;
    this.unloadRamp2 = 0;
    this.startLoadDate1 = null;
    this.startLoadDate2 = null;
    this.startUnLoadDate1 = null;
    this.startUnLoadDate2 = null;
    this.transitTime = null;
    this.distance = 0.0;
    this.endPoint = false;
    this.nodeCat = "";
  }

  read(sheet, n, isEndPoint, row, col) {
    this.arrival = ExcelHelpers.getDateValue(sheet, row, col);
    this.departure = ExcelHelpers.getDateValue(sheet, row, col + 1);
    this.waitingTime = ExcelHelpers.getDateValue(sheet, row, col + 2);
    this.number = n;
    this.endPoint = isEndPoint;

    if (!isEndPoint) {
      if (n === 1) {
        if (!ExcelHelpers.isEmpty(sheet, row, col + 3)) this.loadRamp1 = parseInt(ExcelHelpers.getValue(sheet, row, col + 3), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 4)) this.startLoadDate1 = ExcelHelpers.getDateValue(sheet, row, col + 4);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 5)) this.loadRamp2 = parseInt(ExcelHelpers.getValue(sheet, row, col + 5), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 6)) this.startLoadDate2 = ExcelHelpers.getDateValue(sheet, row, col + 6);
      } else {
        if (!ExcelHelpers.isEmpty(sheet, row, col + 3)) this.unloadRamp1 = parseInt(ExcelHelpers.getValue(sheet, row, col + 3), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 4)) this.startUnLoadDate1 = ExcelHelpers.getDateValue(sheet, row, col + 4);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 5)) this.unloadRamp2 = parseInt(ExcelHelpers.getValue(sheet, row, col + 5), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 6)) this.startUnLoadDate2 = ExcelHelpers.getDateValue(sheet, row, col + 6);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 7)) this.loadRamp1 = parseInt(ExcelHelpers.getValue(sheet, row, col + 7), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 8)) this.startLoadDate1 = ExcelHelpers.getDateValue(sheet, row, col + 8);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 9)) this.loadRamp2 = parseInt(ExcelHelpers.getValue(sheet, row, col + 9), 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 10)) this.startLoadDate2 = ExcelHelpers.getDateValue(sheet, row, col + 10);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 11)) this.movingType = String(ExcelHelpers.getValue(sheet, row, col + 11));
        if (!ExcelHelpers.isEmpty(sheet, row, col + 12)) this.transitTime = ExcelHelpers.getDateValue(sheet, row, col + 12);
        if (!ExcelHelpers.isEmpty(sheet, row, col + 13)) this.distance = parseFloat(ExcelHelpers.getValue(sheet, row, col + 13)) || 0;
      }
    } else {
      if (!ExcelHelpers.isEmpty(sheet, row, col + 3)) this.unloadRamp1 = parseInt(ExcelHelpers.getValue(sheet, row, col + 3), 10);
      if (!ExcelHelpers.isEmpty(sheet, row, col + 4)) this.startUnLoadDate1 = ExcelHelpers.getDateValue(sheet, row, col + 4);
      if (!ExcelHelpers.isEmpty(sheet, row, col + 5)) this.unloadRamp2 = parseInt(ExcelHelpers.getValue(sheet, row, col + 5), 10);
      if (!ExcelHelpers.isEmpty(sheet, row, col + 6)) this.startUnLoadDate2 = ExcelHelpers.getDateValue(sheet, row, col + 6);
      if (!ExcelHelpers.isEmpty(sheet, row, col + 7)) this.movingType = String(ExcelHelpers.getValue(sheet, row, col + 7));
      if (!ExcelHelpers.isEmpty(sheet, row, col + 8)) this.transitTime = ExcelHelpers.getDateValue(sheet, row, col + 8);
      if (!ExcelHelpers.isEmpty(sheet, row, col + 9)) this.distance = parseFloat(ExcelHelpers.getValue(sheet, row, col + 9)) || 0;
    }
  }

  getCopy() {
    const copy = new Point();
    Object.assign(copy, this);
    if (this.arrival) copy.arrival = new Date(this.arrival.getTime());
    if (this.departure) copy.departure = new Date(this.departure.getTime());
    return copy;
  }
}