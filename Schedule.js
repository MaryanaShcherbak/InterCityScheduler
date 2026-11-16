import { Point } from './Point.js';
import { ExcelHelpers } from '../utils/excelHelpers.js';

export class Schedule {
  constructor() {
    this.isActive = false;
    this.startDate = null;
    this.endDate = null;
    this.routeGroup = "";
    this.route = "";
    this.additional = false;
    this.number = "";
    this.deadline = false;
    this.plannedDate = false;
    this.mon = false;
    this.tue = false;
    this.wed = false;
    this.thu = false;
    this.fri = false;
    this.sat = false;
    this.sun = false;

    this.points = [];
    this.distance = 0.0;
    this.duration = 0.0;
    this.region = "";
    this.deliveryType = "";
    this.vehicleType = "";
    this.loadFormat = "";
    this.grf = "";
    this.name = "";
    this.comment = "";
    this.order = "";
    this.costCalculationType = "";
    this.movingType = "";
    this.movingWithoutCar = false;
    this.deadlineSegments = "";
    this.deadlineSegmentsSP = "";
    this.deadlineSegmentsP = "";
    this.deadlineSegmentsC = "";

    this.routePravosl = "";
    this.changedFields = [];
    this.changed = false;
    this.isNew = false;
    this.logMon = false;
    this.logTue = false;
    this.logWed = false;
    this.logThu = false;
    this.logFri = false;
    this.logSat = false;
    this.logSun = false;
    this.loadFormatPravosl = "";
    this.astroDay = 0;
    this.logisticDay = 0;

    for (let i = 1; i <= 24; i++) {
      this[`t${i}`] = null;
    }
  }

  setRoute() {
    if (!this.points || this.points.length === 0) return;
    
    this.route = this.points[0].name || "";
    this.routePravosl = this.points[0].namePravosl || this.points[0].name || "";

    for (let i = 1; i < this.points.length; i++) {
      const ax = this.points[i].name || "";
      const pr = this.points[i].namePravosl || ax;
      this.route += "_" + ax;
      this.routePravosl += " - " + pr;
    }
  }

  getTimeDiff(t1, t2) {
    if (!t1 || !t2) return 0;
    let ms1 = (t1.getHours() * 3600 + t1.getMinutes() * 60 + t1.getSeconds()) * 1000;
    let ms2 = (t2.getHours() * 3600 + t2.getMinutes() * 60 + t2.getSeconds()) * 1000;
    if (ms2 < ms1) {
      ms2 += 24 * 3600 * 1000;
    }
    return ms2 - ms1;
  }

  calculateDurations() {
    let totalMs = 0;
    let totalDist = 0;
    for (let i = 1; i < this.points.length; i++) {
      const diff = this.getTimeDiff(this.points[i - 1].departure, this.points[i].arrival);
      totalDist += this.points[i].distance || 0;
      totalMs += diff;
    }
    this.distance = totalDist;
    this.duration = totalMs;
  }

  calculateTimes() {
    if (this.points.length === 0) return;
    this.t23 = this.points[this.points.length - 1].arrival;
    this.t24 = this.points[this.points.length - 1].departure;

    for (let j = 0; j <= this.points.length - 2; j++) {
      const pt = this.points[j];
      const i = (j + 1) * 2 - 1;
      this[`t${i}`] = pt.arrival;
      this[`t${i + 1}`] = pt.departure;
    }
    this.calculateDurations();
  }

  normalizeTimes() {
    for (let i = 1; i <= 24; i++) {
      const t = this[`t${i}`];
      if (t) {
        this[`t${i}`] = new Date(1975, 0, 1, t.getHours(), t.getMinutes(), t.getSeconds());
      }
    }
  }

  readFromProlog(sheet, row) {
    this.isActive = !ExcelHelpers.isEmpty(sheet, row, 1);
    this.startDate = ExcelHelpers.getDateValue(sheet, row, 2);
    this.endDate = ExcelHelpers.getDateValue(sheet, row, 3);
    this.routeGroup = String(ExcelHelpers.getValue(sheet, row, 4));
    this.route = String(ExcelHelpers.getValue(sheet, row, 5));
    this.additional = !ExcelHelpers.isEmpty(sheet, row, 6);
    this.number = String(ExcelHelpers.getValue(sheet, row, 7));
    this.deadline = !ExcelHelpers.isEmpty(sheet, row, 8);
    this.plannedDate = !ExcelHelpers.isEmpty(sheet, row, 9);
    this.mon = !ExcelHelpers.isEmpty(sheet, row, 10);
    this.tue = !ExcelHelpers.isEmpty(sheet, row, 11);
    this.wed = !ExcelHelpers.isEmpty(sheet, row, 12);
    this.thu = !ExcelHelpers.isEmpty(sheet, row, 13);
    this.fri = !ExcelHelpers.isEmpty(sheet, row, 14);
    this.sat = !ExcelHelpers.isEmpty(sheet, row, 15);
    this.sun = !ExcelHelpers.isEmpty(sheet, row, 16);

    const nodes = [];
    for (let clmn = 17; clmn <= 28; clmn++) {
      const vl = ExcelHelpers.getValue(sheet, row, clmn);
      if (vl !== "") nodes.push(String(vl));
    }

    const l = nodes.length;
    const firstPointsCol = 29;
    this.points = [];

    nodes.forEach((node, idx) => {
      const i = idx + 1;
      const pt = new Point();
      pt.name = node;
      if (i === 1) {
        pt.read(sheet, i, i === l, row, firstPointsCol);
      } else if (i !== l) {
        pt.read(sheet, i, false, row, firstPointsCol + 7 + 14 * (i - 2));
      } else {
        pt.read(sheet, i, true, row, 176);
      }
      this.points.push(pt);
    });

    this.distance = parseFloat(ExcelHelpers.getValue(sheet, row, 186)) || 0;
    this.region = String(ExcelHelpers.getValue(sheet, row, 188));
    this.deliveryType = String(ExcelHelpers.getValue(sheet, row, 189));
    this.vehicleType = String(ExcelHelpers.getValue(sheet, row, 190));
    this.loadFormat = String(ExcelHelpers.getValue(sheet, row, 191));
    this.grf = String(ExcelHelpers.getValue(sheet, row, 192));
    if (this.grf === "") this.isNew = true;

    this.name = String(ExcelHelpers.getValue(sheet, row, 193));
    this.comment = String(ExcelHelpers.getValue(sheet, row, 194));
    this.order = String(ExcelHelpers.getValue(sheet, row, 195));
    this.costCalculationType = String(ExcelHelpers.getValue(sheet, row, 196));
    this.movingType = String(ExcelHelpers.getValue(sheet, row, 197));
    this.movingWithoutCar = !ExcelHelpers.isEmpty(sheet, row, 203);
    this.deadlineSegments = String(ExcelHelpers.getValue(sheet, row, 205));
    this.deadlineSegmentsSP = String(ExcelHelpers.getValue(sheet, row, 206));
    this.deadlineSegmentsP = String(ExcelHelpers.getValue(sheet, row, 207));
    this.deadlineSegmentsC = String(ExcelHelpers.getValue(sheet, row, 208));

    if (!ExcelHelpers.isEmpty(sheet, row, 220)) {
      this.changedFields = String(ExcelHelpers.getValue(sheet, row, 220)).split(",");
      this.changed = true;
    }

    if (this.mon) this.astroDay = 1;
    if (this.tue) this.astroDay = 2;
    if (this.wed) this.astroDay = 3;
    if (this.thu) this.astroDay = 4;
    if (this.fri) this.astroDay = 5;
    if (this.sat) this.astroDay = 6;
    if (this.sun) this.astroDay = 7;

    this.calculateTimes();

    this.logisticDay = this.astroDay;
    if (this.t2 && this.t2.getHours() < 12) this.logisticDay = this.astroDay - 1;
    if (this.logisticDay === 0) this.logisticDay = 7;

    this.logMon = this.logisticDay === 1;
    this.logTue = this.logisticDay === 2;
    this.logWed = this.logisticDay === 3;
    this.logThu = this.logisticDay === 4;
    this.logFri = this.logisticDay === 5;
    this.logSat = this.logisticDay === 6;
    this.logSun = this.logisticDay === 7;

    this.normalizeTimes();
  }

  readFromNewsLetter(sheet, row) {
    this.startDate = ExcelHelpers.getDateValue(sheet, row, 1);
    this.endDate = ExcelHelpers.getDateValue(sheet, row, 2);
    this.routeGroup = String(ExcelHelpers.getValue(sheet, row, 6));
    this.route = String(ExcelHelpers.getValue(sheet, row, 93));
    this.routePravosl = String(ExcelHelpers.getValue(sheet, row, 11));
    this.number = String(ExcelHelpers.getValue(sheet, row, 4));
    this.deadline = String(ExcelHelpers.getValue(sheet, row, 39)).trim() === "Так";

    this.logMon = !ExcelHelpers.isEmpty(sheet, row, 31);
    this.logTue = !ExcelHelpers.isEmpty(sheet, row, 32);
    this.logWed = !ExcelHelpers.isEmpty(sheet, row, 33);
    this.logThu = !ExcelHelpers.isEmpty(sheet, row, 34);
    this.logFri = !ExcelHelpers.isEmpty(sheet, row, 35);
    this.logSat = !ExcelHelpers.isEmpty(sheet, row, 36);
    this.logSun = !ExcelHelpers.isEmpty(sheet, row, 37);

    const nodes = this.routePravosl.split(" - ");
    const times = [];
    for (let i = 40; i <= 63; i++) {
      if (!ExcelHelpers.isEmpty(sheet, row, i)) {
        times.push(ExcelHelpers.getDateValue(sheet, row, i));
      }
    }

    this.points = [];
    for (let i = 0; i < nodes.length; i++) {
      const pt = new Point();
      pt.namePravosl = nodes[i];
      if (i * 2 + 1 < times.length) {
        pt.arrival = times[i * 2];
        pt.departure = times[i * 2 + 1];
      }
      pt.distance = 0;
      this.points.push(pt);
    }

    this.deliveryType = "Міжрегіональна";
    let veh = String(ExcelHelpers.getValue(sheet, row, 9));
    this.vehicleType = veh === "БДФ" ? "Шасі BDF" : veh;
    this.loadFormat = String(ExcelHelpers.getValue(sheet, row, 10));
    this.grf = String(ExcelHelpers.getValue(sheet, row, 3));
    this.name = String(ExcelHelpers.getValue(sheet, row, 5));
    this.comment = String(ExcelHelpers.getValue(sheet, row, 90));

    this.deadlineSegmentsSP = String(ExcelHelpers.getValue(sheet, row, 94)).split(" = ")[0];
    this.deadlineSegmentsP = String(ExcelHelpers.getValue(sheet, row, 95)).split(" = ")[0];
    this.deadlineSegmentsC = String(ExcelHelpers.getValue(sheet, row, 96)).split(" = ")[0];
    this.order = String(ExcelHelpers.getValue(sheet, row, 7));
    this.costCalculationType = "Усі";
    this.movingType = String(ExcelHelpers.getValue(sheet, row, 8));
    this.deadlineSegments = String(ExcelHelpers.getValue(sheet, row, 92)).split(" = ")[0];

    if (this.logMon) this.logisticDay = 1;
    if (this.logTue) this.logisticDay = 2;
    if (this.logWed) this.logisticDay = 3;
    if (this.logThu) this.logisticDay = 4;
    if (this.logFri) this.logisticDay = 5;
    if (this.logSat) this.logisticDay = 6;
    if (this.logSun) this.logisticDay = 7;

    this.calculateTimes();
    this.astroDay = this.logisticDay;
    if (this.points[0] && this.points[0].departure && this.points[0].departure.getHours() < 12) {
      this.astroDay = this.logisticDay + 1;
    }
    if (this.astroDay === 8) this.astroDay = 1;

    this.mon = this.astroDay === 1;
    this.tue = this.astroDay === 2;
    this.wed = this.astroDay === 3;
    this.thu = this.astroDay === 4;
    this.fri = this.astroDay === 5;
    this.sat = this.astroDay === 6;
    this.sun = this.astroDay === 7;

    this.isActive = true;
  }
}