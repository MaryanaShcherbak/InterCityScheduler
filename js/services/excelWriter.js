/**
 * Экспорт графиков на базе ExcelJS с сохранением шапок и стилей шаблонов
 */
export class ExcelWriter {

  static async exportToProlog(schedules, filename = "Графіки ММ.xlsx", saveChanges = false) {
    try {
      // Относительный путь для локальной работы и GitHub Pages
      const response = await fetch('./Templates/PrologTemplate.xlsx');
      if (!response.ok) throw new Error("Не вдалося завантажити Templates/PrologTemplate.xlsx");
      const buffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const sheet = workbook.getWorksheet("Data");

      const sorted = [...schedules].sort((a, b) => {
        const pA1 = a.points[0]?.namePravosl || "";
        const pB1 = b.points[0]?.namePravosl || "";
        if (pA1 !== pB1) return pA1.localeCompare(pB1);

        if (b.logMon !== a.logMon) return b.logMon ? 1 : -1;
        if (b.logTue !== a.logTue) return b.logTue ? 1 : -1;
        if (b.logWed !== a.logWed) return b.logWed ? 1 : -1;
        if (b.logThu !== a.logThu) return b.logThu ? 1 : -1;
        if (b.logFri !== a.logFri) return b.logFri ? 1 : -1;
        if (b.logSat !== a.logSat) return b.logSat ? 1 : -1;
        return b.logSun ? 1 : -1;
      });

      let rowNum = 3;
      for (const sched of sorted) {
        const row = sheet.getRow(rowNum);

        if (sched.isActive) row.getCell(1).value = 1;
        if (sched.startDate) {
          row.getCell(2).value = sched.startDate;
          row.getCell(2).numFmt = "dd.mm.yyyy";
        }
        if (sched.endDate) {
          row.getCell(3).value = sched.endDate;
          row.getCell(3).numFmt = "dd.mm.yyyy";
        }

        row.getCell(4).value = sched.routeGroup;
        row.getCell(5).value = sched.route;
        row.getCell(7).value = sched.number;
        if (sched.deadline) row.getCell(8).value = 1;
        if (sched.plannedDate) row.getCell(9).value = 1;
        if (sched.mon) row.getCell(10).value = 1;
        if (sched.tue) row.getCell(11).value = 1;
        if (sched.wed) row.getCell(12).value = 1;
        if (sched.thu) row.getCell(13).value = 1;
        if (sched.fri) row.getCell(14).value = 1;
        if (sched.sat) row.getCell(15).value = 1;
        if (sched.sun) row.getCell(16).value = 1;

        if (sched.points.length > 0) {
          row.getCell(17).value = sched.points[0].name;
          for (let i = 1; i < sched.points.length - 1; i++) {
            row.getCell(17 + i).value = sched.points[i].name;
          }
          row.getCell(28).value = sched.points[sched.points.length - 1].name;

          row.getCell(29).value = this.formatTimeOnly(sched.points[0].arrival);
          row.getCell(30).value = this.formatTimeOnly(sched.points[0].departure);
          row.getCell(32).value = sched.points[0].loadRamp1 || "";
          row.getCell(34).value = sched.points[0].loadRamp2 || "";

          for (let i = 1; i < sched.points.length - 1; i++) {
            const nextCol = 29 + 7 + 14 * (i - 1);
            row.getCell(nextCol).value = this.formatTimeOnly(sched.points[i].arrival);
            row.getCell(nextCol + 1).value = this.formatTimeOnly(sched.points[i].departure);
            row.getCell(nextCol + 11).value = sched.points[i].movingType;
            row.getCell(nextCol + 13).value = sched.points[i].distance;
            row.getCell(nextCol + 7).value = sched.points[i].loadRamp1 || "";
            row.getCell(nextCol + 9).value = sched.points[i].loadRamp2 || "";
          }

          const last = sched.points[sched.points.length - 1];
          row.getCell(176).value = this.formatTimeOnly(last.arrival);
          row.getCell(177).value = this.formatTimeOnly(last.departure);
          row.getCell(183).value = last.movingType;
          row.getCell(185).value = last.distance;
        }

        row.getCell(186).value = sched.distance;
        row.getCell(188).value = sched.region;
        row.getCell(189).value = sched.deliveryType;
        row.getCell(190).value = sched.vehicleType;
        row.getCell(191).value = sched.loadFormat;
        row.getCell(192).value = sched.grf;
        row.getCell(193).value = sched.name;
        row.getCell(194).value = sched.comment;
        row.getCell(195).value = sched.order;
        row.getCell(196).value = "Усі";
        row.getCell(197).value = sched.movingType;
        if (sched.movingWithoutCar) row.getCell(203).value = 1;

        row.getCell(205).value = sched.deadlineSegments;
        row.getCell(206).value = sched.deadlineSegmentsSP;
        row.getCell(207).value = sched.deadlineSegmentsP;
        row.getCell(208).value = sched.deadlineSegmentsC;

        // --- ЦВЕТОВАЯ РАЗМЕТКА "СВІТЛОФОР" ---
        if (saveChanges) {
          if (sched.changedFields && sched.changedFields.length > 0) {
            row.getCell(220).value = sched.changedFields.join(",");
          }

          let argbColor = null;
          if (sched.endDate && !isNaN(sched.endDate.getTime())) {
            argbColor = "FFFA8072"; // Salmon (закрыт)
          } else if (sched.isNew) {
            argbColor = "FF90EE90"; // LightGreen (новый)
          } else if (sched.changed) {
            argbColor = "FFFFFF00"; // Yellow (изменен)
          }

          if (argbColor) {
            for (let c = 1; c <= 208; c++) {
              row.getCell(c).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: argbColor }
              };
            }
          }
        }

        row.commit();
        rowNum++;
      }

      await this.saveWorkbookFile(workbook, filename);
    } catch (e) {
      console.error(e);
      alert("Помилка формування файлу: " + e.message);
    }
  }

  static async exportToNewsLetter(schedules, filename = "Графіки ММ.xlsx") {
    try {
      const response = await fetch('./Templates/SchedulesTemplate.xlsx');
      if (!response.ok) throw new Error("Не вдалося завантажити Templates/SchedulesTemplate.xlsx");
      const buffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const mainSheet = workbook.getWorksheet("Графіки (Регіони)");
      const relaysSheet = workbook.getWorksheet("Естафети");

      let r = 5;
      for (const s of schedules) {
        this.fillNewsLetterRow(mainSheet.getRow(r), s);
        r++;
      }

      let rEst = 5;
      const relays = schedules.filter(s => s.name && s.name.toLowerCase().includes("s"));
      for (const s of relays) {
        this.fillNewsLetterRow(relaysSheet.getRow(rEst), s);
        rEst++;
      }

      await this.saveWorkbookFile(workbook, filename);
    } catch (e) {
      console.error(e);
      alert("Помилка формування файлу розсилки: " + e.message);
    }
  }

  static fillNewsLetterRow(row, sched) {
    if (sched.startDate) {
      row.getCell(1).value = sched.startDate;
      row.getCell(1).numFmt = "dd.mm.yyyy";
    }
    if (sched.endDate) {
      row.getCell(2).value = sched.endDate;
      row.getCell(2).numFmt = "dd.mm.yyyy";
    }
    row.getCell(3).value = sched.grf;
    row.getCell(4).value = sched.number;
    row.getCell(5).value = sched.name;
    row.getCell(6).value = sched.routeGroup;
    row.getCell(7).value = sched.order;
    row.getCell(8).value = sched.movingType;
    row.getCell(9).value = sched.vehicleType === "Шасі BDF" ? "БДФ" : sched.vehicleType;
    row.getCell(10).value = sched.loadFormatPravosl || sched.loadFormat;

    let totalTimeMs = 0;

    // --- ЗАПИСЬ ТОЧЕК, МАРШРУТА И РАСЧЕТ ТРАНЗИТА ---
    if (sched.points && sched.points.length > 0) {
      const firstPtNameB = sched.points[0].namePravosl || sched.points[0].name || "";
      row.getCell(12).value = firstPtNameB;

      const routePartsB = [firstPtNameB];

      for (let i = 1; i < sched.points.length; i++) {
        const ptNameB = sched.points[i].namePravosl || sched.points[i].name || "";
        const pt = sched.points[i];
        const prevPt = sched.points[i - 1];

        // Запись промежуточных названий и времени
        if (i < sched.points.length - 1) {
          routePartsB.push(ptNameB);
          row.getCell(12 + i).value = ptNameB;
          row.getCell(41 + i * 2 - 1).value = this.formatTimeOnly(pt.arrival);
          row.getCell(41 + i * 2).value = this.formatTimeOnly(pt.departure);
        }

        // Дистанция участков (колонка 64+)
        row.getCell(64 + i - 1).value = pt.distance || 0;

        // Транзитное время участков (колонка 76+)
        let diffMs = 0;
        if (prevPt.departure && pt.arrival) {
          diffMs = pt.arrival.getTime() - prevPt.departure.getTime();
          if (diffMs < 0) diffMs += 24 * 3600 * 1000;
        }
        totalTimeMs += diffMs;
        row.getCell(76 + i - 1).value = diffMs / (86400 * 1000); // Excel time
        row.getCell(76 + i - 1).numFmt = "hh:mm";
      }

      if (sched.points.length > 1) {
        const lastIdx = sched.points.length - 1;
        const lastPtNameB = sched.points[lastIdx].namePravosl || sched.points[lastIdx].name || "";
        routePartsB.push(lastPtNameB);
        row.getCell(23).value = lastPtNameB;
        row.getCell(62).value = this.formatTimeOnly(sched.points[lastIdx].arrival);
        row.getCell(63).value = this.formatTimeOnly(sched.points[lastIdx].departure);
      } else {
        row.getCell(23).value = firstPtNameB;
      }

      row.getCell(11).value = sched.routePravosl || routePartsB.join(" - ");
      row.getCell(40).value = this.formatTimeOnly(sched.points[0].arrival);
      row.getCell(41).value = this.formatTimeOnly(sched.points[0].departure);
    } else {
      row.getCell(11).value = sched.routePravosl || sched.route || "";
    }

    if (sched.astroDay >= 1 && sched.astroDay <= 7) row.getCell(23 + sched.astroDay).value = "+";
    if (sched.logisticDay >= 1 && sched.logisticDay <= 7) row.getCell(30 + sched.logisticDay).value = "+";

    row.getCell(39).value = sched.deadline ? "Так" : "Ні";
    
    // --- РАСЧЕТ ИТОГОВ ---
    const dist = sched.distance || 0;
    const totalHours = totalTimeMs / (3600 * 1000);

    // Колонка 75: Суммарная дистанция
    row.getCell(75).value = dist;
    
    // Колонка 87: Суммарное время
    row.getCell(87).value = totalTimeMs / (86400 * 1000);
    row.getCell(87).numFmt = "hh:mm";

    // Колонка 38 (AL): Кількість водіїв (если дистанция > 629 км ИЛИ время в пути > 11 часов)
    let driversCount = 1;
    if (dist > 629) {
      driversCount = 2;
    } else if (totalHours > 11) {
      driversCount = 2;
    }
    row.getCell(38).value = driversCount;

    // Колонка 88: Зупинки (по 15 мин на каждые 3 часа пути)
    const stopsMinutes = Math.ceil(totalHours / 3) * 15;
    row.getCell(88).value = stopsMinutes / (24 * 60);
    row.getCell(88).numFmt = "hh:mm";

    // Колонка 89: Швидкість
    const stopsHours = stopsMinutes / 60;
    const netHours = totalHours - stopsHours;
    const speed = netHours > 0 ? dist / netHours : 0;
    row.getCell(89).value = speed;

    row.getCell(90).value = sched.comment;
    row.getCell(91).value = sched.grf;
    row.getCell(93).value = sched.route || "";

    row.commit();
  }

  static formatTimeOnly(date) {
    if (!date || isNaN(date.getTime())) return "";
    const h = String(date.getHours()).padStart(2, '0');
    const m = String(date.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }

  static async saveWorkbookFile(workbook, filename) {
    const uint8 = await workbook.xlsx.writeBuffer();
    const blob = new Blob([uint8], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  }
}