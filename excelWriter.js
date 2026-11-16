/**
 * Модуль экспорта графиков на базе ExcelJS со 100% сохранением логики SaveForUpload и WriteForNewsLetter
 */
export class ExcelWriter {

  /**
   * Полная реализация SaveToProlog / SaveForUpload из Schedule.vb
   */
  static async exportToProlog(schedules, filename = "Графіки ММ.xlsx", saveChanges = false) {
    try {
      const response = await fetch('./Templates/PrologTemplate.xlsx');
      if (!response.ok) throw new Error("Не знайдено файл Templates/PrologTemplate.xlsx");
      const buffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const sheet = workbook.getWorksheet("Data");

      // Сортировка точно как в SchedulesWorker.vb: SaveToProlog
      const sorted = [...schedules].sort((a, b) => {
        const pA1 = a.points[0]?.namePravosl || a.points[0]?.name || "";
        const pB1 = b.points[0]?.namePravosl || b.points[0]?.name || "";
        if (pA1 !== pB1) return pA1.localeCompare(pB1);

        const pA_last = a.points[a.points.length - 1]?.namePravosl || a.points[a.points.length - 1]?.name || "";
        const pB_last = b.points[b.points.length - 1]?.namePravosl || b.points[b.points.length - 1]?.name || "";
        if (pA_last !== pB_last) return pA_last.localeCompare(pB_last);

        if (b.logMon !== a.logMon) return b.logMon ? 1 : -1;
        if (b.logTue !== a.logTue) return b.logTue ? 1 : -1;
        if (b.logWed !== a.logWed) return b.logWed ? 1 : -1;
        if (b.logThu !== a.logThu) return b.logThu ? 1 : -1;
        if (b.logFri !== a.logFri) return b.logFri ? 1 : -1;
        if (b.logSat !== a.logSat) return b.logSat ? 1 : -1;
        return b.logSun ? 1 : -1;
      });

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      let rowNum = 3;
      for (const sched of sorted) {
        const row = sheet.getRow(rowNum);

        // Колонка 1: Активный
        if (sched.isActive) row.getCell(1).value = 1;
        if (!sched.grf) {
          if (sched.startDate && sched.startDate > today) {
            row.getCell(1).value = "";
          } else {
            row.getCell(1).value = 1;
          }
        }

        // Колонки 2, 3: Даты
        if (sched.startDate) {
          row.getCell(2).value = sched.startDate;
          row.getCell(2).numFmt = "dd.MM.yyyy";
        }
        if (sched.endDate && !isNaN(sched.endDate.getTime())) {
          row.getCell(3).value = sched.endDate;
          row.getCell(3).numFmt = "dd.MM.yyyy";
        }

        row.getCell(4).value = sched.routeGroup || "";
        row.getCell(5).value = sched.route || "";
        // Колонка 6 закомментирована в оригинале (Additional)
        row.getCell(7).value = sched.number || "";
        if (sched.deadline) row.getCell(8).value = 1;
        if (sched.plannedDate) row.getCell(9).value = 1;
        if (sched.mon) row.getCell(10).value = 1;
        if (sched.tue) row.getCell(11).value = 1;
        if (sched.wed) row.getCell(12).value = 1;
        if (sched.thu) row.getCell(13).value = 1;
        if (sched.fri) row.getCell(14).value = 1;
        if (sched.sat) row.getCell(15).value = 1;
        if (sched.sun) row.getCell(16).value = 1;

        if (sched.points && sched.points.length > 0) {
          const l = sched.points.length;

          // Имена узлов в колонках 17 .. 28
          row.getCell(17).value = sched.points[0].name;
          for (let i = 1; i <= l - 2; i++) {
            row.getCell(17 + i).value = sched.points[i].name;
          }
          row.getCell(28).value = sched.points[l - 1].name;

          // --- ТОЧКА 0 (Старт) ---
          const p0 = sched.points[0];
          this.setTimeOfDayCell(row.getCell(29), p0.arrival);
          this.setTimeOfDayCell(row.getCell(30), p0.departure);
          this.setTransitTimeCell(row.getCell(31), p0.arrival, p0.departure);
          row.getCell(32).value = p0.loadRamp1 || 0;
          row.getCell(34).value = p0.loadRamp2 || 0;

          // --- ПРОМЕЖУТОЧНЫЕ ТОЧКИ (1 .. l - 2) ---
          const firstPointsCol = 29;
          for (let i = 1; i <= l - 2; i++) {
            const nextPointCol = firstPointsCol + 7 + 14 * (i - 1);
            const pt = sched.points[i];
            const prevPt = sched.points[i - 1];

            this.setTimeOfDayCell(row.getCell(nextPointCol), pt.arrival);
            this.setTimeOfDayCell(row.getCell(nextPointCol + 1), pt.departure);
            this.setTransitTimeCell(row.getCell(nextPointCol + 2), pt.arrival, pt.departure); // Ожидание
            row.getCell(nextPointCol + 7).value = pt.loadRamp1 || 0;
            row.getCell(nextPointCol + 9).value = pt.loadRamp2 || 0;
            row.getCell(nextPointCol + 11).value = pt.movingType || "";
            this.setTransitTimeCell(row.getCell(nextPointCol + 12), prevPt.departure, pt.arrival); // Транзит
            row.getCell(nextPointCol + 13).value = pt.distance || 0;
          }

          // --- ТОЧКА ФИНАЛЬНАЯ (Last) ---
          const pLast = sched.points[l - 1];
          const penultPt = l >= 2 ? sched.points[l - 2] : sched.points[0];

          this.setTimeOfDayCell(row.getCell(176), pLast.arrival);
          this.setTimeOfDayCell(row.getCell(177), pLast.departure);
          this.setTransitTimeCell(row.getCell(178), pLast.arrival, pLast.departure); // Ожидание финиша
          row.getCell(183).value = pLast.movingType || "";
          this.setTransitTimeCell(row.getCell(184), penultPt.departure, pLast.arrival); // Транзит финиша
          row.getCell(185).value = pLast.distance || 0;
        }

        // Общие итоги дистанции и времени
        row.getCell(186).value = sched.distance || 0;
        if (sched.duration) {
          // Duration в долях суток (TimeSpan)
          const durationDays = typeof sched.duration === 'number' 
            ? sched.duration / (86400 * 1000) 
            : 0;
          row.getCell(187).value = durationDays;
          row.getCell(187).numFmt = "[h]:mm:ss";
        }

        row.getCell(188).value = sched.region || "";
        row.getCell(189).value = sched.deliveryType || "";
        row.getCell(190).value = sched.vehicleType || "";
        row.getCell(191).value = sched.loadFormat || "";
        row.getCell(192).value = sched.grf || "";
        row.getCell(193).value = sched.name || "";
        row.getCell(194).value = sched.comment || "";
        row.getCell(195).value = sched.order || "";
        row.getCell(196).value = "Усі";
        row.getCell(197).value = sched.movingType || "";
        if (sched.movingWithoutCar) row.getCell(203).value = 1;

        // Точное условие из Schedule.vb для дедлайнов
        if (!sched.deadlineSegmentsSP && !sched.deadlineSegmentsP && !sched.deadlineSegmentsC) {
          row.getCell(205).value = sched.deadlineSegments || "";
        }
        row.getCell(206).value = sched.deadlineSegmentsSP || "";
        row.getCell(207).value = sched.deadlineSegmentsP || "";
        row.getCell(208).value = sched.deadlineSegmentsC || "";

        // Режим сохранения изменений
        if (saveChanges) {
          if (sched.changedFields && sched.changedFields.length > 0) {
            row.getCell(220).value = sched.changedFields.join(",");
          }

          // Цветовая подсветка строк «Світлофор»
          let argbColor = null;
          if (sched.endDate && !isNaN(sched.endDate.getTime())) {
            argbColor = "FFFA8072"; // Salmon
          } else if (sched.isNew) {
            argbColor = "FF90EE90"; // LightGreen
          } else if (sched.changed) {
            argbColor = "FFFFFF00"; // Yellow
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

  /**
   * Полная реализация SaveToXL / WriteForNewsLetter из Schedule.vb
   */
  static async exportToNewsLetter(schedules, filename = "Графіки ММ.xlsx") {
    try {
      const response = await fetch('./Templates/SchedulesTemplate.xlsx');
      if (!response.ok) throw new Error("Не знайдено файл Templates/SchedulesTemplate.xlsx");
      const buffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const mainSheet = workbook.getWorksheet("Графіки (Регіони)");
      const relaysSheet = workbook.getWorksheet("Естафети");

      // Сортировка листа рассылки
      const sorted = [...schedules].sort((a, b) => {
        if (b.logMon !== a.logMon) return b.logMon ? 1 : -1;
        if (b.logTue !== a.logTue) return b.logTue ? 1 : -1;
        if (b.logWed !== a.logWed) return b.logWed ? 1 : -1;
        if (b.logThu !== a.logThu) return b.logThu ? 1 : -1;
        if (b.logFri !== a.logFri) return b.logFri ? 1 : -1;
        if (b.logSat !== a.logSat) return b.logSat ? 1 : -1;
        if (b.logSun !== a.logSun) return b.logSun ? 1 : -1;
        const diffT2 = (a.t2?.getTime() || 0) - (b.t2?.getTime() || 0);
        if (diffT2 !== 0) return diffT2;
        return (a.routePravosl || "").localeCompare(b.routePravosl || "");
      });

      // Заполнение листа регионов
      let r = 5;
      for (const s of sorted) {
        this.fillNewsLetterRow(mainSheet.getRow(r), s);
        r++;
      }

      // Заполнение листа эстафет (Name содержит 's')
      let rEst = 5;
      const relays = sorted.filter(s => s.name && s.name.toLowerCase().includes("s"));
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
      row.getCell(1).numFmt = "dd.MM.yyyy";
    }
    if (sched.endDate && !isNaN(sched.endDate.getTime())) {
      row.getCell(2).value = sched.endDate;
      row.getCell(2).numFmt = "dd.MM.yyyy";
    }
    row.getCell(3).value = sched.grf || "";
    row.getCell(4).value = sched.number || "";
    row.getCell(5).value = sched.name || "";
    row.getCell(6).value = sched.routeGroup || "";
    row.getCell(7).value = sched.order || "";
    row.getCell(8).value = sched.movingType || "";
    row.getCell(9).value = sched.vehicleType === "Шасі BDF" ? "БДФ" : sched.vehicleType || "";
    row.getCell(10).value = sched.loadFormatPravosl || sched.loadFormat || "";

    let totalTimeDays = 0;

    if (sched.points && sched.points.length > 0) {
      const l = sched.points.length;
      let routePravosl = sched.points[0].namePravosl || sched.points[0].name || "";
      row.getCell(12).value = sched.points[0].namePravosl || sched.points[0].name || "";

      for (let i = 1; i < l; i++) {
        const pt = sched.points[i];
        const prevPt = sched.points[i - 1];
        const ptNamePr = pt.namePravosl || pt.name || "";
        routePravosl += " - " + ptNamePr;

        if (i < l - 1) {
          row.getCell(12 + i).value = ptNamePr;
          this.setTimeFractionCell(row.getCell(41 + i * 2 - 1), pt.arrival);
          this.setTimeFractionCell(row.getCell(41 + i * 2), pt.departure);
        }

        // Дистанция и транзитное время участков
        row.getCell(64 + i - 1).value = pt.distance || 0;
        const durDays = this.calcTransitDays(prevPt.departure, pt.arrival);
        row.getCell(76 + i - 1).value = durDays;
        row.getCell(76 + i - 1).numFmt = "hh:mm";
        totalTimeDays += durDays;
      }

      row.getCell(11).value = routePravosl;
      row.getCell(23).value = sched.points[l - 1].namePravosl || sched.points[l - 1].name || "";

      // Подача и выезд начальной точки
      this.setTimeFractionCell(row.getCell(40), sched.points[0].arrival);
      this.setTimeFractionCell(row.getCell(41), sched.points[0].departure);

      // Приезд и выезд конечной точки
      this.setTimeFractionCell(row.getCell(62), sched.points[l - 1].arrival);
      this.setTimeFractionCell(row.getCell(63), sched.points[l - 1].departure);
    } else {
      row.getCell(11).value = sched.routePravosl || sched.route || "";
    }

    // Дни недели
    if (sched.astroDay >= 1 && sched.astroDay <= 7) row.getCell(23 + sched.astroDay).value = "+";
    if (sched.logisticDay >= 1 && sched.logisticDay <= 7) row.getCell(30 + sched.logisticDay).value = "+";

    row.getCell(39).value = sched.deadline ? "Так" : "Ні";

    // Итоги времени и дистанции
    const dist = sched.distance || 0;
    row.getCell(75).value = dist;
    row.getCell(87).value = totalTimeDays;
    row.getCell(87).numFmt = "hh:mm";

    // Количество водителей (колонка 38)
    const totalHours = totalTimeDays * 24;
    row.getCell(38).value = (dist > 629 || totalHours >= 11) ? 2 : 1;

    // Остановки и скорость
    const stopsMinutes = Math.ceil(totalHours / 3) * 15;
    const stopsDays = stopsMinutes / (24 * 60);
    row.getCell(88).value = stopsDays;
    row.getCell(88).numFmt = "hh:mm";

    const netHours = totalHours - (stopsMinutes / 60);
    const speed = netHours > 0 ? dist / netHours : 0;
    row.getCell(89).value = speed;

    row.getCell(90).value = sched.comment || "";
    row.getCell(91).value = sched.grf || "";
    row.getCell(93).value = sched.route || "";

    row.commit();
  }

  // --- Вспомогательные методы вычисления времени ClosedXML ---

  static setTimeOfDayCell(cell, date) {
    if (!date || isNaN(date.getTime())) return;
    const sec = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
    cell.value = sec / 86400;
    cell.numFmt = "[h]:mm:ss";
  }

  static setTransitTimeCell(cell, d1, d2) {
    if (!d1 || !d2) return;
    const days = this.calcTransitDays(d1, d2);
    cell.value = days;
    cell.numFmt = "[h]:mm:ss";
  }

  static setTimeFractionCell(cell, date) {
    if (!date || isNaN(date.getTime())) return;
    const sec = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
    cell.value = sec / 86400;
  }

  static calcTransitDays(d1, d2) {
    if (!d1 || !d2) return 0;
    const sec1 = d1.getHours() * 3600 + d1.getMinutes() * 60 + d1.getSeconds();
    const sec2 = d2.getHours() * 3600 + d2.getMinutes() * 60 + d2.getSeconds();
    let diffSec = sec2 - sec1;
    if (diffSec < 0) diffSec += 86400;
    return diffSec / 86400;
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