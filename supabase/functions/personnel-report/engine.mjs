function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeFio(text) {
  return normalize(text).replace(/[^a-zа-я0-9 ]/gi, "");
}

function hasRibsMarker(value) {
  const text = String(value || "").toLowerCase().replace(/ё/g, "е");
  return text.includes("ribs") || text.includes("рибс") || /[rр][iи][bб][sс]/i.test(text);
}

function canonicalRestaurantName(value) {
  const original = String(value || "").trim();
  const key = revenueNameKey(original);
  if (!key) return "";
  if (key.includes("фабрик") && (key.includes("кухн") || key.includes("белинского"))) {
    return "Фабрика кухни";
  }
  if (key.includes("жюль верн") || key.includes("жюльверн") || key.includes("ударник")
    || (key.includes("ленина") && key.includes("64"))) {
    return "Ленина, 64 Ударник";
  }
  if (key.includes("октября") && key.includes("2")) {
    return "Октября, 2 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("59")) {
    return "Б. Покровская, 59 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("63")) {
    return "Б. Покровская, 63 Самурай";
  }
  if ((key.includes("верхнепечер") || key.includes("верхне печер") || key.includes("вп")) && key.includes("14")) {
    return "Верхне-Печерская, 14Б Самурай";
  }
  if (key.includes("циолковского") && key.includes("19")) {
    return "Циолковского, 19А Самурай";
  }
  if (key.includes("геологов") && key.includes("7")) {
    return "Геологов 7А Самурай";
  }
  if (key.includes("ошар") || key.includes("ресторан 19") || key.includes("ресторан xix")) {
    return "Ошарская, 8А 19";
  }
  if ((key.includes("гагарина") && key.includes("35")) || key.includes("парк швейцария")) {
    return "Парк Швейцария Самурай";
  }
  if (key.includes("моторн") && (key.includes("2к1") || key.includes("2 1"))) {
    return "Моторный, 2/1 доставка Самурай";
  }
  if (key.includes("коминтерн") && key.includes("166")) {
    return "Коминтерна 166 CALL CENTRE";
  }
  if (key.includes("коминтерн") && key.includes("115")) {
    return "Коминтерна, 115 Самурай";
  }
  if (key.includes("волжская") && key.includes("13")) {
    return "Волжская, 13 Самурай";
  }
  if (key.includes("веденяпина") && key.includes("1а")) {
    return "Веденяпина, 1А Самурай";
  }
  if (key.includes("ленина") && key.includes("36")) {
    return "Ленина, 36 Самурай";
  }
  if (key.includes("белинского") && key.includes("61") && hasRibsMarker(original)) {
    return "Белинского, 61 Рибс";
  }
  if (key.includes("белинского") && key.includes("61") && key.includes("достав")) {
    return "Белинского, 61 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("63")) {
    return "Б. Покровская, 63 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("59")) {
    return "Б. Покровская, 59 Самурай";
  }
  if (key.includes("винедо") || key.includes("vinedo")) {
    return "Октябрьская, 1 Винедо";
  }
  if (key.includes("белинского") && key.includes("61") && !key.includes("фабрик")) {
    return "Белинского, 61 Самурай";
  }
  return original;
}

function isNonRestaurantDepartment(value) {
  const key = revenueNameKey(value);
  return !key
    || key === "не определен в списке сотрудников"
    || key === "подразделение не определено"
    || key === "основной зал"
    || key === "технический персонал"
    || (key.includes("фабрика") && key !== "фабрика кухни")
    || key.includes("сотрудники сторонних организаций")
    || key.includes("сдача отчетности");
}

function classifyRole(roleText) {
  const role = normalize(roleText);
  if (/повар|шеф/.test(role)) return "Кухня";
  if (/официант|менеджер зала|мойщ|мойк|администратор|кассир|оператор/.test(role)) return "Зал";
  if (/логист|курьер|водител/.test(role)) return "Доставка";
  if (/барменедж|барбэк|барбек|бармен/.test(role)) return "Бар";
  return null;
}

function splitRevenueWarehouseName(name) {
  const original = String(name || "").trim();
  const key = revenueNameKey(original);
  let restaurant = original;

  if ((key.includes("бп") || key.includes("б покровская") || key.includes("большая покровская")) && key.includes("59")) {
    restaurant = "Б. Покровская, 59 Самурай";
  } else if ((key.includes("бп") || key.includes("б покровская") || key.includes("большая покровская")) && key.includes("63")) {
    restaurant = "Б. Покровская, 63 Самурай";
  } else if (key.includes("вп 14")) {
    restaurant = "Верхне-Печерская, 14Б Самурай";
  } else if (key.includes("каспарус") || (key.includes("циолковского") && key.includes("19"))) {
    restaurant = "Циолковского, 19А Самурай";
  } else if (key.includes("геологов")) {
    restaurant = "Геологов 7А Самурай";
  } else if ((key.includes("ленина") && key.includes("64")) || key.includes("ударник")) {
    restaurant = "Ленина, 64 Ударник";
  } else if (key.includes("ошар") || key.includes("ресторан xix")) {
    restaurant = "Ошарская, 8А 19";
  } else if (key.includes("винедо") || key.includes("vinedo")) {
    restaurant = "Октябрьская, 1 Винедо";
  } else if (key.includes("белинского") && key.includes("61") && key.includes("достав")) {
    restaurant = "Белинского, 61 доставка Самурай";
  } else if ((key.includes("белинского") && key.includes("61") && hasRibsMarker(original)) || /^RIBS\b/i.test(original) || hasRibsMarker(original)) {
    restaurant = "Белинского, 61 Рибс";
  } else if (key.includes("белинского") && key.includes("61")) {
    restaurant = "Белинского, 61 Самурай";
  } else if ((key.includes("гагарина") && key.includes("35")) || key.includes("парк швейцария")) {
    restaurant = "Парк Швейцария Самурай";
  } else if (key.includes("моторн") && (key.includes("пер") || key.includes("2к1") || key.includes("2 1"))) {
    restaurant = "Моторный, 2/1 доставка Самурай";
  } else if (key.includes("коминтерн") && key.includes("166")) {
    restaurant = "Коминтерна 166 CALL CENTRE";
  } else if (key.includes("коминтерн") && key.includes("115")) {
    restaurant = "Коминтерна, 115 Самурай";
  } else if (key.includes("волжская") && key.includes("13")) {
    restaurant = "Волжская, 13 Самурай";
  } else if (key.includes("веденяпина") && key.includes("1а")) {
    restaurant = "Веденяпина, 1А Самурай";
  } else if (key.includes("октября") && key.includes("2")) {
    restaurant = "Октября, 2 Самурай";
  } else if (key.includes("ленина") && key.includes("36")) {
    restaurant = "Ленина, 36 Самурай";
  } else if (key.includes("детский центр жюль верн")) {
    restaurant = "Ленина, 64 Ударник";
  } else if (/^Самурай,\s*/i.test(original)) {
    restaurant = original.replace(/\s*\([^)]*\)\s*$/g, "").replace(/^Самурай,\s*/i, "").trim();
    restaurant = restaurant + " Самурай";
  } else {
    restaurant = original.replace(/\s*\([^)]*\)\s*$/g, "").trim();
  }

  return { restaurant: canonicalRestaurantName(restaurant || original), warehouse: original };
}

function revenueNameKey(value) {
  return String(value || "")
    .toLowerCase()
    .replaceAll("ё", "е")
    .replace(/\u00a0/g, " ")
    .replace(/[.,;:()\-_/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function restaurantMatchKey(value) {
  let key = revenueNameKey(value);
  key = key
    .replace(/\bсамурай\b/g, "")
    .replace(/\bcall\b|\bcentre\b|\bcenter\b/g, "")
    .replace(/\bд\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
  if (key.includes("коминтерна") && key.includes("166")) key = key.replace(/\bдоставка\b/g, "").trim();
  return key;
}

function isBlockedRevenueName(value) {
  const key = revenueNameKey(value);
  return DEFAULT_REVENUE_EXCLUSIONS.some((rule) => rule && key.includes(rule));
}

function getRevenueWarehouseKind(name) {
  const key = revenueNameKey(name);
  if (key.includes("кухня")) return "kitchen";
  if (key.includes("бар")) return "bar";
  return "single";
}

const DEFAULT_REVENUE_EXCLUSIONS = [
  "Онлайн оплата, СБП",
  "Основной склад",
  "БУРГЕР БИК ООО Чайка",
  "Бургер Бик",
  "Фабрика разделка",
  "Шале №15",
  "Совнаркомовская 13",
  "НТО ООО Приспех пр-кт Гагарина, д. 35",
  "Юность ул. Зеленский Съезд, д. 8/10",
  "ИП Амельченко Евгений Андреевич",
  "Фудтрак Амельченко пл. Маркина, д. 12А",
  "Фабрика кондитерка",
  "ул. Большая Покровская, д. 13",
  "Швейцария БИК \"ПРИСПЕХ\"",
  "Фабрика пекарня",
  "Рождественская",
  "ВЕНУСТО"
].map(revenueNameKey);

function calcWorkedSeconds(events) {
  const sorted = [...events].sort((a, b) => a.absSec - b.absSec);
  let total = 0;
  let inWork = false;
  let startSec = 0;

  sorted.forEach((e) => {
    if (e.direction === "Вход") {
      inWork = true;
      startSec = e.absSec;
      return;
    }
    if (e.direction === "Выход" && inWork && e.absSec >= startSec) {
      total += e.absSec - startSec;
      inWork = false;
    }
  });

  if (total === 0 && sorted.length >= 2) {
    const fallback = sorted[sorted.length - 1].absSec - sorted[0].absSec;
    if (fallback > 0) total = fallback;
  }

  return total;
}

function workedSecondsToShift(workedSeconds) {
  if (workedSeconds <= 0) return 0;
  return workedSeconds > 7 * 3600 ? 1 : 0.5;
}


export const roleKeys = ["kitchen","hall","delivery","bar"];
const groups = {"Кухня":"kitchen","Зал":"hall","Доставка":"delivery","Бар":"bar"};
export function summarizeAttendance(days, revenues) {
  const result = new Map();
  const get = (dateIso,restaurant) => {
    const key = dateIso+"|"+restaurant;
    if (!result.has(key)) result.set(key,{dateIso,restaurant,kitchen:0,hall:0,delivery:0,bar:0,total:0,revenue:null,kitchenRevenue:null,barRevenue:null,hasConflict:false,details:{kitchen:[],hall:[],delivery:[],bar:[]}});
    return result.get(key);
  };
  for (const day of days) {
    const people = new Map();
    for (const event of day.rows || []) {
      const group = groups[classifyRole(event.position)];
      if (!group) continue;
      const key = String(event.personId || normalizeFio(event.person));
      if (!key) continue;
      if (!people.has(key)) people.set(key,[]);
      people.get(key).push({...event,group});
    }
    for (const events of people.values()) {
      events.sort((a,b)=>a.dateTime.localeCompare(b.dateTime));
      const exits = events.filter(e=>Number(e.actionType)===0);
      const place = e=>canonicalRestaurantName([e.department,e.address,e.accessPoint].filter(Boolean).join(" "));
      const locations = new Set(events.map(place));
      const exit = exits.at(-1);
      const location = exit ? place(exit) : locations.size===1 ? [...locations][0] : "";
      const restaurant = location && !isNonRestaurantDepartment(location) ? location : "Подразделение не определено";
      const event = exit || events.at(-1);
      const worked = calcWorkedSeconds(events.map(e=>({direction:Number(e.actionType)===1?"Вход":"Выход",absSec:Date.parse(e.dateTime.replace(" ","T")+"+03:00")/1000})));
      const shift = workedSecondsToShift(worked);
      if (!shift) continue;
      const row = get(day.work_date,restaurant);
      row[event.group]+=shift; row.total+=shift;
      const reviewReason = !exit ? "нет отметки выхода" : restaurant==="Подразделение не определено" ? "точка не определена" : "";
      row.hasConflict ||= !!reviewReason;
      row.details[event.group].push({person:event.person,shift,hasConflict:!!reviewReason,reviewReason});
    }
  }
  for (const sale of revenues) {
    if (isBlockedRevenueName(sale.restaurant)) continue;
    const restaurant = splitRevenueWarehouseName(sale.restaurant).restaurant;
    const row = get(sale.sale_date,restaurant);
    const value = Number(sale.revenue);
    if (!Number.isFinite(value)) continue;
    row.revenue = (row.revenue ?? 0)+value;
    const kind = getRevenueWarehouseKind(sale.restaurant);
    if (kind==="kitchen") row.kitchenRevenue=(row.kitchenRevenue??0)+value;
    if (kind==="bar") row.barRevenue=(row.barRevenue??0)+value;
  }
  return [...result.values()];
}
export function restrictRows(rows, grants, admin) {
  return rows.filter(row=>admin || grants.some(g=>g.restaurant===row.restaurant)).map(row=>{
    const allowedRoles=admin?roleKeys:grants.filter(g=>g.restaurant===row.restaurant).flatMap(g=>g.roles);
    const copy={dateIso:row.dateIso,restaurant:row.restaurant,allowedRoles,details:{},total:0,hasConflict:false};
    for(const key of roleKeys) if(allowedRoles.includes(key)) {
      copy[key]=row[key]; copy.total+=row[key]; copy.details[key]=row.details[key];
      copy.hasConflict ||= row.details[key].some(p=>p.hasConflict);
    }
    if(allowedRoles.includes("kitchen")) copy.kitchenRevenue=row.kitchenRevenue;
    if(allowedRoles.includes("bar")) copy.barRevenue=row.barRevenue;
    if(allowedRoles.includes("hall")||allowedRoles.includes("delivery")) copy.revenue=row.revenue;
    return copy;
  });
}

