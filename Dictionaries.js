export class Warehouse {
  constructor() {
    this.ax = "";
    this.pravosl = "";
    this.city = "";
    this.cityRing = "";
    this.nodeCat = "";
  }

  readFromArray(row) {
    this.ax = String(row[0] || "");
    this.pravosl = String(row[1] || "");
    this.city = String(row[2] || "");
    this.cityRing = String(row[3] || "");
    if (row[4] !== undefined && row[4] !== "") {
      this.nodeCat = String(row[4]);
    }
  }
}

export class LoadFormat {
  constructor() {
    this.nameProlog = "";
    this.namePravosl = "";
  }

  readFromArray(row) {
    this.nameProlog = String(row[0] || "");
    this.namePravosl = String(row[1] || "");
  }
}

export class Matrix {
  constructor() {
    this.nodeA = "";
    this.nodeB = "";
    this.route = "";
    this.distance = 0;
    this.duration = null;
  }

  readFromArray(row) {
    this.nodeA = String(row[0] || "");
    this.nodeB = String(row[1] || "");
    this.route = `${this.nodeA}_${this.nodeB}`;
    this.distance = parseFloat(row[2]) || 0;
    this.duration = row[3] ? new Date(row[3]) : null;
  }
}

export class Ramp {
  constructor() {
    this.nodeA = "";
    this.nodeB = "";
    this.ramp = 0;
    this.direction = "";
  }

  readFromArray(row) {
    this.nodeA = String(row[0] || "");
    this.nodeB = String(row[1] || "");
    if (row[2] !== undefined && row[2] !== "") {
      this.ramp = parseInt(row[2], 10);
    }
    this.direction = `${this.nodeA}${this.nodeB}`;
  }
}