import { Warehouse, LoadFormat, Matrix, Ramp } from '../models/Dictionaries.js';

export class DictionaryService {
  static unknownNodesProlog = new Set();
  static unknownNodesNews = new Set();
  static apiUrl = "";

  static async loadFromGoogleAPI(apiUrl) {
    this.apiUrl = apiUrl;
    const response = await fetch(apiUrl);
    if (!response.ok) throw new Error("Помилка з'єднання з Google Sheets API");
    
    const data = await response.json();
    
    const warehouses = [];
    const loadFormats = [];
    const matrixes = [];
    const ramps = [];

    const whByAx = new Map();
    const whByPravosl = new Map();

    if (data.warehouses) {
      for (let r = 1; r < data.warehouses.length; r++) {
        const wh = new Warehouse();
        wh.readFromArray(data.warehouses[r]);
        warehouses.push(wh);

        if (wh.ax) whByAx.set(wh.ax.trim().toLowerCase(), wh);
        if (wh.pravosl) whByPravosl.set(wh.pravosl.trim().toLowerCase(), wh);
      }
    }

    if (data.dict) {
      for (let r = 1; r < data.dict.length; r++) {
        const lf = new LoadFormat();
        lf.readFromArray(data.dict[r]);
        loadFormats.push(lf);
      }
    }

    if (data.matrixes) {
      for (let r = 1; r < data.matrixes.length; r++) {
        const mtr = new Matrix();
        mtr.readFromArray(data.matrixes[r]);
        matrixes.push(mtr);
      }
    }

    if (data.ramps) {
      for (let r = 1; r < data.ramps.length; r++) {
        const rmp = new Ramp();
        rmp.readFromArray(data.ramps[r]);
        if (rmp.ramp > 0) ramps.push(rmp);
      }
    }

    // Очищаем списки неизвестных при каждой загрузке словарей
    this.unknownNodesProlog.clear();
    this.unknownNodesNews.clear();

    return { 
      warehouses, 
      loadFormats, 
      matrixes, 
      ramps,
      whByAx,
      whByPravosl
    };
  }

  static searchWH(dictContainer, targetAX) {
    if (!targetAX || !dictContainer) return null;
    const key = targetAX.trim().toLowerCase();
    if (dictContainer.whByAx && dictContainer.whByAx.has(key)) return dictContainer.whByAx.get(key);
    return null;
  }

  static searchWHByPravosl(dictContainer, namePravosl) {
    if (!namePravosl || !dictContainer) return null;
    const key = namePravosl.trim().toLowerCase();
    if (dictContainer.whByPravosl && dictContainer.whByPravosl.has(key)) return dictContainer.whByPravosl.get(key);
    return null;
  }

  static setCities(schedules, dictContainer) {
    if (!dictContainer) return;
    for (const sched of schedules) {
      for (const pt of sched.points) {
        const wh = this.searchWH(dictContainer, pt.name);
        if (wh && wh.pravosl) {
          pt.namePravosl = wh.pravosl;
          pt.city = wh.city;
          pt.cityRing = wh.cityRing;
          pt.nodeCat = wh.nodeCat;
        } else {
          pt.namePravosl = pt.name;
          // Фиксируем неизвестный AX узел
          this.unknownNodesProlog.add(pt.name);
        }
      }
      sched.setRoute();
    }
  }

  static setAXCities(schedules, dictContainer) {
    if (!dictContainer) return;
    for (const sched of schedules) {
      for (const pt of sched.points) {
        const wh = this.searchWHByPravosl(dictContainer, pt.namePravosl);
        if (wh && wh.ax) {
          pt.name = wh.ax;
          pt.city = wh.city;
          pt.cityRing = wh.cityRing;
        } else if (!pt.name) {
          pt.name = pt.namePravosl;
          // Фиксируем неизвестный православный узел
          this.unknownNodesNews.add(pt.namePravosl);
        }
      }
      sched.setRoute();
    }
  }

  static translateLoadFormatsNewsIntoProlog(loadFormats, schedules) {
    if (!loadFormats) return;
    for (const sched of schedules) {
      for (const lf of loadFormats) {
        if (lf.namePravosl === sched.loadFormat) {
          sched.loadFormatPravosl = sched.loadFormat;
          sched.loadFormat = lf.nameProlog;
          break;
        }
      }
    }
  }

  /**
   * Отправка новых узлов в Google Sheets
   */
  static async saveNewNodes(newNodesList) {
    if (!this.apiUrl) throw new Error("API URL не встановлено");
    
    const response = await fetch(this.apiUrl, {
      method: "POST",
      body: JSON.stringify({
        action: "addWarehouses",
        newNodes: newNodesList
      })
    });
    
    const result = await response.json();
    if (result.status !== "success") {
      throw new Error(result.message || "Помилка збереження на сервері");
    }
    return true;
  }
}