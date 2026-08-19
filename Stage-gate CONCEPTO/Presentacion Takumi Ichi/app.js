/**
 * TAKUMI ICHI (匠一) - Commercial Presentation & Sales Projection Application
 * Integrated with Shared Google Sheets Live Database
 * Tecnofood / Mr. Wings Innovation 2026
 */

const GOOGLE_SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbztMhmeVl3x5VNuJSMSJzooz0MKXeZGPMRNRZ5qpyfzo2MYaLXrIvoNgUvg7DHigm1M/exec";

// ==========================================================================
// 1. DATASET: 9 FLAVORS, 29 SKUS (FROM 260324 COP preliminar.xlsx - Tab 260729 costeo)
// Column E (desarrollar): 1 = Fase 1 Lanzamiento Prioritario (14 SKUs), 0 = Fase 2 (15 SKUs)
// ==========================================================================
const PRODUCTS_DATA = {
  "teriyaki": {
    id: "teriyaki",
    name: "Salsa Teriyaki",
    japanese: "照り焼き",
    kanji: "照焼",
    category: "Glaseados & Cocina Caliente",
    heat: 0,
    viscosity: "Espesa caramelizada",
    description: "Glaseado dulce y salado con notas de mirin, jengibre y caramelo. Logra un acabado brillante 'teri' que se adhiere perfectamente a proteínas calientes sobre parrilla, plancha o sartén.",
    pairings: "Brochetas Yakitori de pollo, Salmón teriyaki a la plancha, Bowls de pollo y res, Hamburguesas estilo oriental.",
    dishImage: "assets/dishes/Teriyaki.jpg",
    dishCaption: "Brochetas Yakitori glaseadas al carbón servidas con Salsa Teriyaki Takumi Ichi.",
    formats: [
      {
        sizeLabel: "3.8 L",
        code: "Teriyaki-3.78",
        netKg: 4.72,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Restaurantes / Planchas Teppanyaki",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 316.00,
        pricePerKg: 66.99,
        mockupImage: "assets/mockups/Teriyaki-3.78.jpg",
        description: "Garrafa de 3.78 L con asa ergonómica de alto rendimiento para cocina."
      },
      {
        sizeLabel: "946 ml",
        code: "Teriyaki-0.946",
        netKg: 1.18,
        desarrollar: 1,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Food Service",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.00,
        pricePerKg: 88.95,
        mockupImage: "assets/mockups/Teriyaki-0.946.jpg",
        description: "Squeeze 946 ml ideal para línea de armado y baño de rollos/bowls."
      },
      {
        sizeLabel: "414 ml",
        code: "Teriyaki-0.414",
        netKg: 0.52,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Servicio en Mesa/Barra)",
        channelType: "Distribuidor Food Service",
        priceMxn: 59.21,
        pricePerKg: 114.61,
        mockupImage: "assets/mockups/Teriyaki-0.414.jpg",
        description: "Squeeze de 414 ml para dipping individual y terminación."
      }
    ]
  },

  "sriracha": {
    id: "sriracha",
    name: "Salsa Sriracha",
    japanese: "シラチャー",
    kanji: "赤唐辛子",
    category: "Picantes Tradicionales",
    heat: 4,
    viscosity: "Media pastosa con tropezones de chile",
    description: "Salsa picante a base de jalapeños rojos maduros y chiles tailandeses madurados al sol con ajo y vinagre. Textura suave con gran brillo y acidez brillante.",
    pairings: "Sopas Pho vietnamitas, Rollos de sushi spicy tuna, Huevos y bowls de desayuno, Mayonesas preparadas, Alitas.",
    dishImage: "assets/dishes/Sriracha.jpg",
    dishCaption: "Alitas y tenders crujientes bañados en Salsa Sriracha Takumi Ichi.",
    formats: [
      {
        sizeLabel: "3.8 L",
        code: "Sriracha-3.78",
        netKg: 4.50,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Restaurantes / Producción de Aderezos",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 316.00,
        pricePerKg: 70.15,
        mockupImage: "assets/mockups/Sriracha-3.78.jpg",
        description: "Garrafa de 3.78 L para bases de cocina y aderezos de la casa."
      },
      {
        sizeLabel: "946 ml",
        code: "Sriracha-0.946",
        netKg: 1.13,
        desarrollar: 1,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Barras de Sushi",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.00,
        pricePerKg: 93.14,
        mockupImage: "assets/mockups/Sriracha-0.946.jpg",
        description: "Squeeze 946 ml ergonómico para uso intensivo en cocina."
      },
      {
        sizeLabel: "414 ml",
        code: "Sriracha-0.414",
        netKg: 0.49,
        desarrollar: 1,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Servicio de Mesa)",
        channelType: "Distribuidor Food Service",
        priceMxn: 59.00,
        pricePerKg: 119.59,
        mockupImage: "assets/mockups/Sriracha-0.414.jpg",
        description: "Squeeze 414 ml estándar para colocar al centro de cada mesa."
      }
    ]
  },

  "soya": {
    id: "soya",
    name: "Salsa de Soya",
    japanese: "しょうゆ",
    kanji: "醤油",
    category: "Básicos & Sazonadores",
    heat: 0,
    viscosity: "Líquida tradicional",
    description: "Salsa de soya tradicional japonesa elaborada mediante fermentación controlada. Perfil umami profundo y balance salino limpio, indispensable para sazonar caldos, salteados al wok y servicio en mesa de sushi.",
    pairings: "Ramen, Arroz Frito Yakimeshi, Sushi & Sashimi, Marinados de res, cerdo y aves, Dumplings y Gyozas.",
    dishImage: "assets/dishes/Soya.jpg",
    dishCaption: "Tazón de Ramen Tonkotsu humeante con ajitama y chashu maridado con Salsa de Soya Takumi Ichi.",
    formats: [
      {
        sizeLabel: "20 L",
        code: "Soya-20",
        netKg: 21.16,
        desarrollar: 1,
        unitsPerBox: 1,
        channel: "Transformadores de Alimentos / Cadenas",
        channelType: "Transformadores de Alimentos",
        priceMxn: 505.00,
        pricePerKg: 23.86,
        mockupImage: "assets/mockups/Soya-20.jpg",
        description: "Porrón industrial de 20 L en HDPE natural semitranslúcido con asa superior."
      },
      {
        sizeLabel: "3.8 L",
        code: "Soya-3.78",
        netKg: 4.00,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Restaurantes / Catering / Comedores",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 127.00,
        pricePerKg: 31.75,
        mockupImage: "assets/mockups/Soya-3.78.jpg",
        description: "Garrafa ergonómica de 1 Galón (3.78 L) con asa lateral para cocina caliente."
      },
      {
        sizeLabel: "946 ml",
        code: "Soya-0.946",
        netKg: 1.00,
        desarrollar: 0,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Barras de Sushi",
        channelType: "Distribuidor Food Service",
        priceMxn: 59.21,
        pricePerKg: 59.15,
        mockupImage: "assets/mockups/Soya-0.946.jpg",
        description: "Botella squeeze de 946 ml (32 oz) con dosificador cónico estriado."
      },
      {
        sizeLabel: "414 ml",
        code: "Soya-0.414",
        netKg: 0.44,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Servicio en Mesa)",
        channelType: "Distribuidor Food Service",
        priceMxn: 39.47,
        pricePerKg: 90.10,
        mockupImage: "assets/mockups/Soya-0.414.jpg",
        description: "Botella squeeze compacta de 414 ml para servicio directo al comensal."
      }
    ]
  },

  "ostion": {
    id: "ostion",
    name: "Salsa de Ostión",
    japanese: "オイスターソース",
    kanji: "蠔油",
    category: "Glaseados & Wok",
    heat: 0,
    viscosity: "Espesa brillante",
    description: "Salsa aterciopelada y rica en extracto de ostión natural con tono marrón oscuro. Imparte el auténtico sabor 'Wok Hei' a vegetales salteados, carnes de res y arroces.",
    pairings: "Res con Brócoli, Chow Mein, Verduras salteadas al wok, Marinados de cerdo Char Siu, Costillitas orientales.",
    dishImage: "assets/dishes/Ostion.jpg",
    dishCaption: "Salteado de res y brócoli al wok con Salsa de Ostión Takumi Ichi.",
    formats: [
      {
        sizeLabel: "946 ml",
        code: "Ostión-0.946",
        netKg: 1.10,
        desarrollar: 1,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Food Service",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.00,
        pricePerKg: 95.24,
        mockupImage: "assets/mockups/Ostión-0.946.jpg",
        description: "Squeeze 946 ml de fácil flujo para dosificación exacta por porción."
      },
      {
        sizeLabel: "3.8 L",
        code: "Ostión-3.78",
        netKg: 4.41,
        desarrollar: 0,
        unitsPerBox: 4,
        channel: "Cadenas de Comida China / Wok Stations",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 263.16,
        pricePerKg: 59.74,
        mockupImage: "assets/mockups/Ostión-3.78.jpg",
        description: "Garrafa de 3.78 L para uso continuo en estaciones de wok."
      },
      {
        sizeLabel: "414 ml",
        code: "Ostión-0.414",
        netKg: 0.48,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes / Cocina Caliente",
        channelType: "Distribuidor Food Service",
        priceMxn: 52.63,
        pricePerKg: 109.08,
        mockupImage: "assets/mockups/Ostión-0.414.jpg",
        description: "Squeeze 414 ml compacto para barras y restaurantes pequeños."
      }
    ]
  },

  "mayo_chipotle": {
    id: "mayo_chipotle",
    name: "Aderezo Mayo - Chipotle",
    japanese: "マヨチポトレ",
    kanji: "辛マヨ",
    category: "Emulsiones & Aderezos Fusión",
    heat: 2,
    viscosity: "Cremosa untable",
    description: "Emulsión sedosa estilo mayonesa japonesa Kewpie combinada con el ahumado y picor cálido del chipotle mexicano. Textura ultra cremosa que no se separa sobre platillos calientes.",
    pairings: "Rollos de sushi empanizados, Camarones roca tempura, Tacos gobernador y baja de pescado, Sandwiches de pollo crujiente.",
    dishImage: "assets/dishes/Mayo Chipotle.jpg",
    dishCaption: "Rollos de sushi empanizados y camarones roca bañados en Aderezo Mayo Chipotle Takumi Ichi.",
    formats: [
      {
        sizeLabel: "3.8 L",
        code: "Mayo Chipotle-3.78",
        netKg: 4.12,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Sushi / Comida Casual",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 316.00,
        pricePerKg: 76.79,
        mockupImage: "assets/mockups/Mayo Chipotle-3.78.jpg",
        description: "Garrafa de 3.78 L de alto rendimiento para barras frías."
      },
      {
        sizeLabel: "946 ml",
        code: "Mayo Chipotle-0.946",
        netKg: 1.03,
        desarrollar: 0,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Food Service",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.26,
        pricePerKg: 102.21,
        mockupImage: "assets/mockups/Mayo Chipotle-0.946.jpg",
        description: "Squeeze 946 ml para aplicación rápida en línea de sushi."
      },
      {
        sizeLabel: "414 ml",
        code: "Mayo Chipotle-0.414",
        netKg: 0.45,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Mesa y Dipping)",
        channelType: "Distribuidor Food Service",
        priceMxn: 52.63,
        pricePerKg: 116.77,
        mockupImage: "assets/mockups/Mayo Chipotle-0.414.jpg",
        description: "Squeeze 414 ml para servicio directo al cliente."
      }
    ]
  },

  "anguila": {
    id: "anguila",
    name: "Salsa de Anguila",
    japanese: "うなぎのたれ",
    kanji: "鰻蒲焼",
    category: "Sushi & Especialidades",
    heat: 0,
    viscosity: "Muy espesa / Almíbar brillante",
    description: "Salsa Kabayaki dulce y untuosa, con cuerpo denso que permanece sobre el sushi y la proteína sin escurrir. Clásica para coronar rollos empanizados, nigiris de anguila y bowls de arroz.",
    pairings: "Rollos de sushi horneados y fritos, Unadon (bowl de anguila), Brochetas de camarón, Poke bowls fusión.",
    dishImage: "assets/dishes/Anguila.jpg",
    dishCaption: "Tazón Unadon tradicional con anguila glaseada y Salsa de Anguila Takumi Ichi.",
    formats: [
      {
        sizeLabel: "946 ml",
        code: "Anguila-0.946",
        netKg: 1.18,
        desarrollar: 1,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Barras de Sushi",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.00,
        pricePerKg: 88.94,
        mockupImage: "assets/mockups/Anguila-0.946.jpg",
        description: "Squeeze 946 ml con boquilla fina para decoración artística de rollos."
      },
      {
        sizeLabel: "3.8 L",
        code: "Anguila-3.78",
        netKg: 4.72,
        desarrollar: 0,
        unitsPerBox: 4,
        channel: "Cadenas de Sushi / Dark Kitchens",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 315.79,
        pricePerKg: 66.94,
        mockupImage: "assets/mockups/Anguila-3.78.jpg",
        description: "Garrafa de 3.78 L para recarga de dispensadores de barra."
      },
      {
        sizeLabel: "414 ml",
        code: "Anguila-0.414",
        netKg: 0.52,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes / Servicio a Mesa",
        channelType: "Distribuidor Food Service",
        priceMxn: 52.63,
        pricePerKg: 101.87,
        mockupImage: "assets/mockups/Anguila-0.414.jpg",
        description: "Squeeze 414 ml para autoservicio y terminación en barra."
      }
    ]
  },

  "ponzu": {
    id: "ponzu",
    name: "Salsa Ponzu",
    japanese: "ポン酢しょうゆ",
    kanji: "柚子ポン酢",
    category: "Cítricos & Aderezos",
    heat: 0,
    viscosity: "Líquida cítrica",
    description: "Salsa a base de soya sazonada con jugo de cítricos orientales (tipo Yuzu) y notas avinagradas. Aporta frescura instantánea, corta la grasa en cortes marmoleados y resalta mariscos frescos.",
    pairings: "Shabu-Shabu, Hot Pot, Tiraditos y Sashimi de atún/salmón, Ostras frescas, Ensaladas orientales, Gyozas al vapor.",
    dishImage: "assets/dishes/Ponzu.jpg",
    dishCaption: "Mesa tradicional de Shabu-Shabu japonés con láminas de res wagyu y salsa Ponzu cítrica.",
    formats: [
      {
        sizeLabel: "20 L",
        code: "Ponzu-20",
        netKg: 21.16,
        desarrollar: 1,
        unitsPerBox: 1,
        channel: "Cadenas de Sushi / Transformadores",
        channelType: "Transformadores de Alimentos",
        priceMxn: 734.00,
        pricePerKg: 34.69,
        mockupImage: "assets/mockups/Ponzu-20.jpg",
        description: "Porrón industrial de 20 L en HDPE natural semitranslúcido."
      },
      {
        sizeLabel: "3.8 L",
        code: "Ponzu-3.78",
        netKg: 4.00,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Restaurantes / Marisquerías",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 178.00,
        pricePerKg: 44.51,
        mockupImage: "assets/mockups/Ponzu-3.78.jpg",
        description: "Garrafa de 1 Galón (3.78 L) para mise en place en barra de sushi."
      },
      {
        sizeLabel: "946 ml",
        code: "Ponzu-0.946",
        netKg: 1.00,
        desarrollar: 0,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Food Service",
        channelType: "Distribuidor Food Service",
        priceMxn: 65.79,
        pricePerKg: 65.73,
        mockupImage: "assets/mockups/Ponzu-0.946.jpg",
        description: "Squeeze de 946 ml para preparación rápida y porcionado."
      },
      {
        sizeLabel: "414 ml",
        code: "Ponzu-0.414",
        netKg: 0.44,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Servicio en Mesa)",
        channelType: "Distribuidor Food Service",
        priceMxn: 39.47,
        pricePerKg: 90.12,
        mockupImage: "assets/mockups/Ponzu-0.414.jpg",
        description: "Squeeze de 414 ml para autoservicio de comensales en mesa."
      }
    ]
  },

  "samyang": {
    id: "samyang",
    name: "Salsa Samyang",
    japanese: "激辛サムヤン",
    kanji: "火鶏激辛",
    category: "Picantes Extremos / K-Food",
    heat: 5,
    viscosity: "Media espesa con aceite de chile",
    description: "Inspirada en la tendencia coreana 'Buldak' de fuego extremo. Una salsa de color carmesí profundo con notas tostadas de sésamo, ajo y un picor ardiente de alta persistencia.",
    pairings: "Fideos coreanos picantes, Alitas y boneless ardientes, Tacos coreanos de pork belly, Hamburguesas picantes.",
    dishImage: "assets/dishes/Samyang.jpg",
    dishCaption: "Fideos coreanos picantes al fuego con ajonjolí y chiles rojos maridados con Salsa Samyang.",
    formats: [
      {
        sizeLabel: "946 ml",
        code: "Samyang-0.946",
        netKg: 1.13,
        desarrollar: 1,
        unitsPerBox: 12,
        channel: "Restaurantes Casual Dining / Alitas y Burgers",
        channelType: "Distribuidor Food Service",
        priceMxn: 105.00,
        pricePerKg: 93.12,
        mockupImage: "assets/mockups/Samyang-0.946.jpg",
        description: "Squeeze 946 ml para línea de aderezado y salsas de reto de picor."
      },
      {
        sizeLabel: "3.8 L",
        code: "Samyang-3.78",
        netKg: 4.51,
        desarrollar: 0,
        unitsPerBox: 4,
        channel: "Cadenas de Alitas / Dark Kitchens / Restaurantes Coreanos",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 315.79,
        pricePerKg: 70.09,
        mockupImage: "assets/mockups/Samyang-3.78.jpg",
        description: "Garrafa de 3.78 L para salteo de alitas y fideos en volumen."
      },
      {
        sizeLabel: "414 ml",
        code: "Samyang-0.414",
        netKg: 0.49,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Mesa)",
        channelType: "Distribuidor Food Service",
        priceMxn: 52.63,
        pricePerKg: 106.65,
        mockupImage: "assets/mockups/Samyang-0.414.jpg",
        description: "Squeeze 414 ml para comensales amantes del picor extremo."
      }
    ]
  },

  "agridulce": {
    id: "agridulce",
    name: "Salsa Agridulce",
    japanese: "甘酢あん",
    kanji: "糖醋",
    category: "Cocina China & Frituras",
    heat: 0,
    viscosity: "Media gelatinosa traslúcida",
    description: "Salsa color rojo-anaranjado traslúcido con balance perfecto entre dulzor frutal y acidez refrescante. Diseñada para cubrir frituras crujientes conservando su textura crocante.",
    pairings: "Pollo y cerdo agridulce, Rollos primavera (Spring rolls), Wontons fritos, Dedos de queso y tempura.",
    dishImage: "assets/dishes/Agridulce.jpg",
    dishCaption: "Pollo agridulce crujiente con piña y pimientos en caja takeout oriental.",
    formats: [
      {
        sizeLabel: "3.8 L",
        code: "Agridulce-3.78",
        netKg: 4.72,
        desarrollar: 1,
        unitsPerBox: 4,
        channel: "Cadenas de Comida Rápida Oriental / Buffets",
        channelType: "Cadenas de Restaurantes",
        priceMxn: 316.00,
        pricePerKg: 66.99,
        mockupImage: "assets/mockups/Agridulce-3.78.jpg",
        description: "Garrafa de 3.78 L de alto volumen para preparación masiva."
      },
      {
        sizeLabel: "946 ml",
        code: "Agridulce-0.946",
        netKg: 1.18,
        desarrollar: 0,
        unitsPerBox: 12,
        channel: "Restaurantes Independientes / Food Service",
        channelType: "Distribuidor Food Service",
        priceMxn: 98.68,
        pricePerKg: 83.59,
        mockupImage: "assets/mockups/Agridulce-0.946.jpg",
        description: "Squeeze 946 ml para ensamble en línea de servicio y delivery."
      },
      {
        sizeLabel: "414 ml",
        code: "Agridulce-0.414",
        netKg: 0.52,
        desarrollar: 0,
        unitsPerBox: 24,
        channel: "Restaurantes Independientes (Dipping de Mesa)",
        channelType: "Distribuidor Food Service",
        priceMxn: 52.63,
        pricePerKg: 101.87,
        mockupImage: "assets/mockups/Agridulce-0.414.jpg",
        description: "Squeeze de 414 ml para dip individual en mesa."
      }
    ]
  }
};

// ==========================================================================
// 2. STATE MANAGEMENT & GOOGLE SHEETS LIVE SYNC
// ==========================================================================
const STORAGE_KEY = "takumi_ichi_projections_v3";

let portfolioFilter = "develop"; // "develop" (14 SKUs) or "all" (29 SKUs)
let currentSelectedFlavor = "teriyaki";
let currentSelectedFormatIndex = 0;
let currentViewMode = "mockup"; // "mockup" or "dish"

let projections = [];
let charts = {};
let isSyncing = false;

// Initial fallback demo projections
const FALLBACK_DEMO_PROJECTIONS = [
  {
    id: "proj-001",
    kamName: "KAM Food Service",
    customerName: "Grupo Sushi Roll (50 Sucursales)",
    channel: "Cadenas de Restaurantes",
    flavorId: "teriyaki",
    flavorName: "Salsa Teriyaki",
    formatLabel: "3.8 L",
    skuCode: "Teriyaki-3.78",
    piecesAnnual: 1800,
    netKgPerPiece: 4.72,
    pricePerPiece: 316.00,
    totalKg: 8496.0,
    totalRevenue: 568800.0,
    probability: 0.8,
    targetDate: "2026-10",
    notes: "Abastecimiento mensual centralizado para cocina caliente."
  },
  {
    id: "proj-002",
    kamName: "KAM Food Service",
    customerName: "Cadena Alitas & Chelas",
    channel: "Cadenas de Restaurantes",
    flavorId: "sriracha",
    flavorName: "Salsa Sriracha",
    formatLabel: "3.8 L",
    skuCode: "Sriracha-3.78",
    piecesAnnual: 1200,
    netKgPerPiece: 4.50,
    pricePerPiece: 316.00,
    totalKg: 5400.0,
    totalRevenue: 379200.0,
    probability: 0.8,
    targetDate: "2026-10",
    notes: "Base para marinado y glaseado de alitas spicy."
  },
  {
    id: "proj-003",
    kamName: "KAM Clientes Especiales",
    customerName: "Alimentos Procesados de Occidente",
    channel: "Transformadores de Alimentos",
    flavorId: "soya",
    flavorName: "Salsa de Soya",
    formatLabel: "20 L",
    skuCode: "Soya-20",
    piecesAnnual: 480,
    netKgPerPiece: 21.16,
    pricePerPiece: 505.00,
    totalKg: 10156.8,
    totalRevenue: 242400.0,
    probability: 1.0,
    targetDate: "2026-09",
    notes: "Marinado industrial de proteína cárnica."
  },
  {
    id: "proj-004",
    kamName: "KAM Mayoristas/Retail",
    customerName: "Distribuidora Gastronómica Nacional",
    channel: "Distribuidor Food Service",
    flavorId: "anguila",
    flavorName: "Salsa de Anguila",
    formatLabel: "946 ml",
    skuCode: "Anguila-0.946",
    piecesAnnual: 2400,
    netKgPerPiece: 1.18,
    pricePerPiece: 105.00,
    totalKg: 2832.0,
    totalRevenue: 252000.0,
    probability: 0.8,
    targetDate: "2026-10",
    notes: "Distribución en restaurantes japoneses independientes."
  },
  {
    id: "proj-005",
    kamName: "KAM Sureste",
    customerName: "Catering Cancún & Riviera Maya",
    channel: "Distribuidor Mayorista",
    flavorId: "ponzu",
    flavorName: "Salsa Ponzu",
    formatLabel: "3.8 L",
    skuCode: "Ponzu-3.78",
    piecesAnnual: 600,
    netKgPerPiece: 4.00,
    pricePerPiece: 178.00,
    totalKg: 2400.0,
    totalRevenue: 106800.0,
    probability: 0.8,
    targetDate: "2026-10",
    notes: "Barras de mariscos y ceviches en hotelería."
  },
  {
    id: "proj-006",
    kamName: "Ecommerce",
    customerName: "Canal Digital MrWings.com / Mercado Libre",
    channel: "Mercado Libre",
    flavorId: "samyang",
    flavorName: "Salsa Samyang",
    formatLabel: "946 ml",
    skuCode: "Samyang-0.946",
    piecesAnnual: 1500,
    netKgPerPiece: 1.13,
    pricePerPiece: 105.00,
    totalKg: 1695.0,
    totalRevenue: 157500.0,
    probability: 0.8,
    targetDate: "2026-09",
    notes: "Venta directa a consumidores amantes del picor extremo coreano."
  }
];

function updateSyncStatus(msg, type = "success") {
  const el = document.getElementById("syncStatusBadge");
  if (!el) return;
  if (type === "syncing") {
    el.innerHTML = `<i data-lucide="refresh-cw" class="spin"></i> <span>Sincronizando Google Sheets...</span>`;
    el.style.background = "rgba(59, 130, 246, 0.15)";
    el.style.color = "#60A5FA";
  } else if (type === "success") {
    el.innerHTML = `<i data-lucide="cloud-check"></i> <span>Google Sheets Conectado (${projections.length} filas)</span>`;
    el.style.background = "rgba(16, 185, 129, 0.15)";
    el.style.color = "#34D399";
  } else {
    el.innerHTML = `<i data-lucide="cloud-off"></i> <span>Modo Offline / Local</span>`;
    el.style.background = "rgba(245, 158, 11, 0.15)";
    el.style.color = "#FBBF24";
  }
  lucide.createIcons();
}

async function loadStoredProjections() {
  updateSyncStatus("Sincronizando...", "syncing");
  
  // 1. First try loading cached local data immediately for fast load
  try {
    const cached = localStorage.getItem(STORAGE_KEY);
    if (cached) {
      projections = JSON.parse(cached);
      renderProjectionsTable();
      updateDashboard();
    }
  } catch (e) {
    console.warn("Local storage cache unavailable:", e);
  }

  // 2. Fetch live data from Google Sheets Webhook
  try {
    const res = await fetch(GOOGLE_SHEETS_WEBHOOK_URL, { method: "GET" });
    if (res.ok) {
      const liveData = await res.json();
      if (Array.isArray(liveData) && liveData.length > 0) {
        projections = liveData;
        saveLocalCache();
        renderProjectionsTable();
        renderProductShowcase();
        updateDashboard();
        updateSyncStatus("Conectado", "success");
        return;
      }
    }
  } catch (err) {
    console.warn("Google Sheets fetch error (falling back to cache/demo):", err);
  }

  // 3. Fallback if empty or failed
  if (projections.length === 0) {
    projections = [...FALLBACK_DEMO_PROJECTIONS];
  }
  saveLocalCache();
  renderProjectionsTable();
  renderProductShowcase();
  updateDashboard();
  updateSyncStatus("Offline", "offline");
}

function saveLocalCache() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(projections));
  } catch (e) {
    console.error("Local storage save error:", e);
  }
}

async function pushToGoogleSheets(item) {
  try {
    updateSyncStatus("Guardando en Google Sheets...", "syncing");
    // Use text/plain payload to avoid CORS preflight issues with Google Apps Script
    await fetch(GOOGLE_SHEETS_WEBHOOK_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(item)
    });
    updateSyncStatus("Guardado en Google Sheets", "success");
  } catch (err) {
    console.error("Error sending to Google Sheets:", err);
    updateSyncStatus("Guardado solo localmente", "offline");
  }
}

// ==========================================================================
// 3. VIEW FILTER SWITCHER (FASE 1 DESARROLLAR vs TODOS)
// ==========================================================================
window.setPortfolioFilter = function(filter) {
  portfolioFilter = filter;

  // Update Navbar buttons
  document.getElementById("navFilterDevelop").classList.toggle("active", filter === "develop");
  document.getElementById("navFilterAll").classList.toggle("active", filter === "all");

  // Update In-page buttons
  document.getElementById("catFilterDevelop").classList.toggle("active", filter === "develop");
  document.getElementById("catFilterAll").classList.toggle("active", filter === "all");

  // Update Banner text
  const bannerText = document.getElementById("activeFilterBannerText");
  const skuMetricsTag = document.getElementById("skuMetricsViewTag");
  if (filter === "develop") {
    bannerText.innerHTML = `Mostrando <strong>14 productos prioritarios</strong> para desarrollo (Columna E = 1 / Fase 1).`;
    if (skuMetricsTag) skuMetricsTag.innerHTML = `14 SKUs Fase 1 (Desarrollar)`;
  } else {
    bannerText.innerHTML = `Mostrando <strong>Portafolio Completo (29 SKUs)</strong> de lanzamiento extendido.`;
    if (skuMetricsTag) skuMetricsTag.innerHTML = `29 SKUs Portafolio Completo`;
  }

  currentSelectedFormatIndex = 0;
  renderFlavorTabs();
  renderProductShowcase();
  populateFormSelects();
  updateDashboard();
};

function getVisibleFormats(prod) {
  if (portfolioFilter === "develop") {
    return prod.formats.filter(f => f.desarrollar === 1);
  }
  return prod.formats;
}

function getAllFlatSkus() {
  let skus = [];
  Object.values(PRODUCTS_DATA).forEach(prod => {
    prod.formats.forEach(f => {
      skus.push({
        flavorId: prod.id,
        flavorName: prod.name,
        sizeLabel: f.sizeLabel,
        code: f.code,
        netKg: f.netKg,
        priceMxn: f.priceMxn,
        pricePerKg: f.pricePerKg,
        desarrollar: f.desarrollar,
        unitsPerBox: f.unitsPerBox || 6,
        channel: f.channel
      });
    });
  });
  return skus;
}

// ==========================================================================
// 4. UI RENDERING: TABS, PRODUCT SHOWCASE, FORMAT SWITCHER
// ==========================================================================
function renderFlavorTabs() {
  const container = document.getElementById("flavorTabs");
  container.innerHTML = "";

  Object.values(PRODUCTS_DATA).forEach(prod => {
    const visibleFormats = getVisibleFormats(prod);
    if (visibleFormats.length === 0) return;

    const btn = document.createElement("button");
    btn.className = `flavor-tab-btn ${prod.id === currentSelectedFlavor ? 'active' : ''}`;
    btn.innerHTML = `
      <span class="flavor-tab-kanji">${prod.kanji}</span>
      <span>${prod.name}</span>
      <span class="tag ${portfolioFilter === 'develop' ? 'develop-tag' : ''}" style="font-size:0.7rem; padding: 2px 6px;">${visibleFormats.length} SKU${visibleFormats.length > 1 ? 's' : ''}</span>
    `;
    btn.addEventListener("click", () => {
      currentSelectedFlavor = prod.id;
      currentSelectedFormatIndex = 0;
      currentViewMode = "mockup";
      renderFlavorTabs();
      renderProductShowcase();
      syncFormFlavorSelect();
    });
    container.appendChild(btn);
  });
}

function renderProductShowcase() {
  const prod = PRODUCTS_DATA[currentSelectedFlavor];
  const visibleFormats = getVisibleFormats(prod);

  if (currentSelectedFormatIndex >= visibleFormats.length) {
    currentSelectedFormatIndex = 0;
  }
  const format = visibleFormats[currentSelectedFormatIndex] || prod.formats[0];
  const container = document.getElementById("productShowcase");

  // Heat stars/flame rendering
  let heatHtml = "";
  if (prod.heat > 0) {
    heatHtml = `<span class="spec-badge heat"><i data-lucide="flame"></i> Picor: Nivel ${prod.heat}/5</span>`;
  } else {
    heatHtml = `<span class="spec-badge"><i data-lucide="sparkle"></i> Sin picor</span>`;
  }

  // Format pills HTML
  let formatPillsHtml = visibleFormats.map((f, idx) => `
    <button class="format-pill-btn ${idx === currentSelectedFormatIndex ? 'active' : ''}" onclick="switchFormat(${idx})">
      ${f.desarrollar === 1 ? '<span class="pill-dot" title="Prioritario para Desarrollar"></span>' : ''}
      ${f.sizeLabel}
    </button>
  `).join("");

  // Calculate projected volume for this current SKU
  let currentSkuPieces = 0;
  let currentSkuKg = 0;
  let currentSkuRevenue = 0;
  projections.forEach(p => {
    if (p.skuCode === format.code) {
      currentSkuPieces += p.piecesAnnual;
      currentSkuKg += p.totalKg;
      currentSkuRevenue += p.totalRevenue;
    }
  });

  // Format table rows HTML
  let tableRowsHtml = visibleFormats.map((f, idx) => `
    <tr class="${idx === currentSelectedFormatIndex ? 'active-row' : ''}">
      <td>
        <strong>${f.sizeLabel}</strong>
        ${f.desarrollar === 1 ? '<span class="tag develop-tag" style="margin-left: 6px; font-size: 0.65rem;">🚀 Fase 1</span>' : '<span class="tag" style="margin-left: 6px; font-size: 0.65rem; color: var(--text-muted);">Fase 2</span>'}
      </td>
      <td>${f.netKg.toFixed(2)} kg</td>
      <td><small>${f.channel}</small></td>
      <td class="price-val">$${f.priceMxn.toFixed(2)} MXN</td>
      <td class="price-kg">$${f.pricePerKg.toFixed(2)}/kg</td>
    </tr>
  `).join("");

  container.innerHTML = `
    <!-- Left: Interactive Gallery -->
    <div class="showcase-visuals">
      <div class="visual-view-selector">
        <button class="view-btn ${currentViewMode === 'mockup' ? 'active' : ''}" onclick="switchView('mockup')">
          <i data-lucide="package"></i> Envase de Estudio
        </button>
        <button class="view-btn ${currentViewMode === 'dish' ? 'active' : ''}" onclick="switchView('dish')">
          <i data-lucide="utensils"></i> En Restaurante con Platillo
        </button>
      </div>

      <div class="visual-display-stage">
        ${currentViewMode === 'mockup' ? `
          <img src="${format.mockupImage}" alt="${prod.name} ${format.sizeLabel}" class="mockup-img" id="mainMockupImg" />
        ` : `
          <img src="${prod.dishImage}" alt="${prod.name} con platillo" class="dish-img" id="mainDishImg" />
        `}
      </div>

      ${currentViewMode === 'mockup' ? `
        <div class="format-pills-row">
          <span style="font-size: 0.8rem; color: var(--text-muted); align-self: center; margin-right: 4px;">Formato:</span>
          ${formatPillsHtml}
        </div>
      ` : `
        <div style="text-align: center; font-size: 0.85rem; color: var(--text-secondary); font-style: italic;">
          <i data-lucide="info" style="width: 14px; height: 14px; vertical-align: middle;"></i> ${prod.dishCaption}
        </div>
      `}

      <!-- Live SKU Commercial Performance Widget -->
      <div style="background: rgba(0,0,0,0.35); border: 1px solid var(--border-subtle); border-radius: var(--radius-sm); padding: 0.85rem 1rem; display: flex; justify-content: space-between; align-items: center; margin-top: 0.5rem;">
        <div>
          <span style="font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; display: block;">Métricas en Pipeline (${format.code}):</span>
          <span style="font-size: 0.95rem; font-weight: 700; color: #fff;">${currentSkuPieces.toLocaleString('es-MX')} pzas <small style="color:var(--text-muted); font-weight: normal;">(${currentSkuKg.toLocaleString('es-MX', {maximumFractionDigits:1})} kg)</small></span>
        </div>
        <div style="text-align: right;">
          <span style="font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; display: block;">Venta Anual Proyectada:</span>
          <span style="font-size: 1.05rem; font-weight: 800; color: var(--color-gold);">$${currentSkuRevenue.toLocaleString('es-MX', {maximumFractionDigits:0})} MXN</span>
        </div>
      </div>
    </div>

    <!-- Right: Specs & Commercial Data -->
    <div class="showcase-info">
      <div class="showcase-header">
        <span class="product-japanese-title">${prod.japanese} (${prod.kanji})</span>
        <h2 class="product-main-title">${prod.name}</h2>
        <div class="product-badge-row">
          <span class="spec-badge"><i data-lucide="tag"></i> ${prod.category}</span>
          <span class="spec-badge"><i data-lucide="droplet"></i> ${prod.viscosity}</span>
          ${heatHtml}
        </div>
      </div>

      <p class="product-profile-text">${prod.description}</p>

      <div class="product-pairings">
        <strong><i data-lucide="chef-hat" style="width: 14px; height: 14px; vertical-align: middle;"></i> Aplicaciones Gastronómicas & Maridaje:</strong>
        <p>${prod.pairings}</p>
      </div>

      <div class="format-pricing-table-wrap">
        <h4><i data-lucide="table"></i> Especificaciones y Precios (${portfolioFilter === 'develop' ? 'Fase 1: Desarrollar' : 'Portafolio Completo'}):</h4>
        <table class="format-table">
          <thead>
            <tr>
              <th>Formato</th>
              <th>Cont. Neto</th>
              <th>Canal Objetivo</th>
              <th>Precio Sugerido</th>
              <th>Precio / Kg</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>
      </div>
    </div>
  `;

  lucide.createIcons();
}

window.switchFormat = function(idx) {
  currentSelectedFormatIndex = idx;
  currentViewMode = "mockup";
  renderProductShowcase();
};

window.switchView = function(mode) {
  currentViewMode = mode;
  renderProductShowcase();
};

// ==========================================================================
// 5. PROJECTIONS FORM LOGIC & PREVIEW
// ==========================================================================
function populateFormSelects() {
  const flavorSelect = document.getElementById("flavorSelect");
  flavorSelect.innerHTML = "";

  Object.values(PRODUCTS_DATA).forEach(prod => {
    const visibleFormats = getVisibleFormats(prod);
    if (visibleFormats.length === 0) return;

    const opt = document.createElement("option");
    opt.value = prod.id;
    opt.textContent = `${prod.name} (${prod.japanese})`;
    flavorSelect.appendChild(opt);
  });

  flavorSelect.addEventListener("change", (e) => {
    updateSkuSelect(e.target.value);
  });

  document.getElementById("skuSelect").addEventListener("change", updateFormPreview);
  document.getElementById("annualPieces").addEventListener("input", updateFormPreview);
  document.getElementById("channelSelect").addEventListener("change", updateFormPreview);

  updateSkuSelect(flavorSelect.value || currentSelectedFlavor);
}

function syncFormFlavorSelect() {
  const flavorSelect = document.getElementById("flavorSelect");
  if (flavorSelect && flavorSelect.value !== currentSelectedFlavor) {
    flavorSelect.value = currentSelectedFlavor;
    updateSkuSelect(currentSelectedFlavor);
  }
}

function updateSkuSelect(flavorId) {
  const skuSelect = document.getElementById("skuSelect");
  skuSelect.innerHTML = "";

  const prod = PRODUCTS_DATA[flavorId];
  if (!prod) return;

  const visibleFormats = getVisibleFormats(prod);
  visibleFormats.forEach((f, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.textContent = `${f.sizeLabel} — $${f.priceMxn.toFixed(2)} MXN ($${f.pricePerKg.toFixed(2)}/kg) ${f.desarrollar === 1 ? '⭐' : ''}`;
    skuSelect.appendChild(opt);
  });

  updateFormPreview();
}

function updateFormPreview() {
  const flavorId = document.getElementById("flavorSelect").value;
  const skuIdx = parseInt(document.getElementById("skuSelect").value || 0, 10);
  const pieces = parseInt(document.getElementById("annualPieces").value || 0, 10);

  const prod = PRODUCTS_DATA[flavorId];
  if (!prod) return;

  const visibleFormats = getVisibleFormats(prod);
  const format = visibleFormats[skuIdx] || prod.formats[0];

  if (!format) return;

  const totalKg = pieces * format.netKg;
  const totalMxn = pieces * format.priceMxn;

  document.getElementById("prevKg").textContent = `${totalKg.toLocaleString('es-MX', { minimumFractionDigits: 1, maximumFractionDigits: 1 })} kg`;
  document.getElementById("prevMxn").textContent = `$${totalMxn.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} MXN`;
}

async function handleFormSubmit(e) {
  e.preventDefault();

  const kamName = document.getElementById("kamSelect").value;
  const customerName = document.getElementById("customerName").value.trim();
  const channel = document.getElementById("channelSelect").value;
  const probability = parseFloat(document.getElementById("probabilitySelect").value);
  const flavorId = document.getElementById("flavorSelect").value;
  const skuIdx = parseInt(document.getElementById("skuSelect").value, 10);
  const pieces = parseInt(document.getElementById("annualPieces").value, 10);
  const targetDate = document.getElementById("targetStartDate").value;
  const notes = document.getElementById("notesText").value.trim();

  const prod = PRODUCTS_DATA[flavorId];
  const visibleFormats = getVisibleFormats(prod);
  const format = visibleFormats[skuIdx];

  const newProj = {
    id: `proj-${Date.now()}`,
    kamName,
    customerName,
    channel,
    flavorId,
    flavorName: prod.name,
    formatLabel: format.sizeLabel,
    skuCode: format.code,
    desarrollar: format.desarrollar,
    piecesAnnual: pieces,
    netKgPerPiece: format.netKg,
    pricePerPiece: format.priceMxn,
    totalKg: pieces * format.netKg,
    totalRevenue: pieces * format.priceMxn,
    probability,
    targetDate,
    notes
  };

  // Add immediately to local state and update UI
  projections.unshift(newProj);
  saveLocalCache();
  renderProjectionsTable();
  renderProductShowcase();
  updateDashboard();

  // Send to shared Google Sheet in background
  pushToGoogleSheets(newProj);

  // Reset form inputs
  document.getElementById("customerName").value = "";
  document.getElementById("annualPieces").value = "";
  document.getElementById("notesText").value = "";
  updateFormPreview();
}

window.deleteProjection = function(id) {
  if (confirm("¿Deseas eliminar este registro de proyección?")) {
    projections = projections.filter(p => p.id !== id);
    saveLocalCache();
    renderProjectionsTable();
    renderProductShowcase();
    updateDashboard();
  }
};

window.refreshFromGoogleSheets = async function() {
  await loadStoredProjections();
  alert("Datos actualizados desde Google Sheets.");
};

// ==========================================================================
// 6. PROJECTIONS TABLE RENDERING
// ==========================================================================
function renderProjectionsTable() {
  const tbody = document.getElementById("projectionsTableBody");
  tbody.innerHTML = "";

  document.getElementById("tableCounterText").textContent = `${projections.length} cuentas registradas`;

  if (projections.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: var(--text-muted); padding: 2rem;">No hay proyecciones registradas aún. Ingresa prospectos en el formulario superior.</td></tr>`;
    return;
  }

  projections.forEach(p => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${p.kamName || p.sellerName || 'KAM Food Service'}</strong></td>
      <td>${p.customerName}</td>
      <td><span class="tag">${p.channel}</span></td>
      <td>${p.flavorName}</td>
      <td><span class="tag" style="background: var(--bg-card);">${p.formatLabel}</span></td>
      <td><strong>${p.piecesAnnual.toLocaleString('es-MX')}</strong> pz</td>
      <td>${p.totalKg.toLocaleString('es-MX', { maximumFractionDigits: 1 })} kg</td>
      <td style="color: var(--color-gold); font-weight: 700;">$${p.totalRevenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })}</td>
      <td><span class="tag" style="color: ${p.probability >= 0.8 ? '#4ADE80' : '#FBBF24'};">${Math.round(p.probability * 100)}%</span></td>
      <td>
        <button class="btn-icon-danger" onclick="deleteProjection('${p.id}')" title="Eliminar registro">
          <i data-lucide="trash-2"></i>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  lucide.createIcons();
}

// ==========================================================================
// 7. DASHBOARD & KPIS ENGINE (INCLUDING SKU METRICS BREAKDOWN)
// ==========================================================================
function updateDashboard() {
  let totalPieces = 0;
  let totalKg = 0;
  let totalRevenue = 0;
  let weightedRevenue = 0;

  let formatPieces = {};
  let channelRevenue = {};
  let productVolume = {};
  let customerMap = {};
  let skuAggregates = {};

  projections.forEach(p => {
    totalPieces += p.piecesAnnual;
    totalKg += p.totalKg;
    totalRevenue += p.totalRevenue;
    weightedRevenue += (p.totalRevenue * p.probability);

    // Format Mix
    formatPieces[p.formatLabel] = (formatPieces[p.formatLabel] || 0) + p.totalKg;

    // Channel Revenue
    channelRevenue[p.channel] = (channelRevenue[p.channel] || 0) + p.totalRevenue;

    // Product Volume
    productVolume[p.flavorName] = (productVolume[p.flavorName] || 0) + p.totalKg;

    // SKU Aggregation
    if (!skuAggregates[p.skuCode]) {
      skuAggregates[p.skuCode] = { pieces: 0, kg: 0, revenue: 0, weighted: 0 };
    }
    skuAggregates[p.skuCode].pieces += p.piecesAnnual;
    skuAggregates[p.skuCode].kg += p.totalKg;
    skuAggregates[p.skuCode].revenue += p.totalRevenue;
    skuAggregates[p.skuCode].weighted += (p.totalRevenue * p.probability);

    // Customer aggregation
    if (!customerMap[p.customerName]) {
      customerMap[p.customerName] = {
        name: p.customerName,
        channel: p.channel,
        kam: p.kamName || p.sellerName || 'KAM Food Service',
        pieces: 0,
        kg: 0,
        revenue: 0,
        weighted: 0
      };
    }
    customerMap[p.customerName].pieces += p.piecesAnnual;
    customerMap[p.customerName].kg += p.totalKg;
    customerMap[p.customerName].revenue += p.totalRevenue;
    customerMap[p.customerName].weighted += (p.totalRevenue * p.probability);
  });

  const avgPriceKg = totalKg > 0 ? (totalRevenue / totalKg) : 0;
  const approxBoxes = Math.round(totalPieces / 6);

  // Update KPI cards
  document.getElementById("kpiTotalPieces").innerHTML = `${totalPieces.toLocaleString('es-MX')} <small>pzas</small>`;
  document.getElementById("kpiTotalBoxes").textContent = `~${approxBoxes.toLocaleString('es-MX')} cajas maestras`;

  document.getElementById("kpiTotalKg").innerHTML = `${totalKg.toLocaleString('es-MX', { maximumFractionDigits: 0 })} <small>kg</small>`;
  document.getElementById("kpiTotalLiters").textContent = `~${(totalKg * 0.9).toLocaleString('es-MX', { maximumFractionDigits: 0 })} L equivalentes`;

  document.getElementById("kpiTotalRevenue").innerHTML = `$${totalRevenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })} <small>MXN</small>`;
  document.getElementById("kpiWeightedRevenue").textContent = `Ponderado: $${weightedRevenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })} MXN`;

  document.getElementById("kpiAvgPriceKg").innerHTML = `$${avgPriceKg.toFixed(2)} <small>/kg</small>`;

  // Update Charts
  updateCharts(formatPieces, channelRevenue, productVolume);

  // Update SKU Metrics Breakdown Table
  renderSkuMetricsTable(skuAggregates, totalRevenue);

  // Update Top Clients Table
  renderTopClientsTable(Object.values(customerMap));
}

function renderSkuMetricsTable(skuAggregates, totalCompanyRevenue) {
  const tbody = document.getElementById("skuMetricsTableBody");
  const tfoot = document.getElementById("skuMetricsTableFoot");
  if (!tbody) return;

  tbody.innerHTML = "";
  tfoot.innerHTML = "";

  const allSkus = getAllFlatSkus();
  const visibleSkus = portfolioFilter === "develop" 
    ? allSkus.filter(s => s.desarrollar === 1) 
    : allSkus;

  let sumPieces = 0;
  let sumBoxes = 0;
  let sumKg = 0;
  let sumRevenue = 0;
  let sumWeighted = 0;

  visibleSkus.forEach(sku => {
    const agg = skuAggregates[sku.code] || { pieces: 0, kg: 0, revenue: 0, weighted: 0 };
    const approxBoxes = Math.ceil(agg.pieces / (sku.unitsPerBox || 6));
    const sharePercent = totalCompanyRevenue > 0 ? ((agg.revenue / totalCompanyRevenue) * 100).toFixed(1) : "0.0";

    sumPieces += agg.pieces;
    sumBoxes += approxBoxes;
    sumKg += agg.kg;
    sumRevenue += agg.revenue;
    sumWeighted += agg.weighted;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><code>${sku.code}</code></td>
      <td><strong>${sku.flavorName}</strong></td>
      <td><span class="tag" style="background: var(--bg-card);">${sku.sizeLabel}</span></td>
      <td>${sku.desarrollar === 1 ? '<span class="tag develop-tag">🚀 Fase 1</span>' : '<span class="tag" style="color:var(--text-muted);">Fase 2</span>'}</td>
      <td>$${sku.priceMxn.toFixed(2)}</td>
      <td>${sku.netKg.toFixed(2)} kg</td>
      <td><strong>${agg.pieces.toLocaleString('es-MX')}</strong> pz</td>
      <td>${approxBoxes.toLocaleString('es-MX')} cjs</td>
      <td>${agg.kg.toLocaleString('es-MX', { maximumFractionDigits: 1 })} kg</td>
      <td style="color: var(--color-gold); font-weight: 700;">$${agg.revenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })}</td>
      <td style="color: #4ADE80;">$${agg.weighted.toLocaleString('es-MX', { maximumFractionDigits: 0 })}</td>
      <td>
        <div style="display: flex; align-items: center; gap: 6px;">
          <span style="font-weight: 600; width: 40px;">${sharePercent}%</span>
          <div style="flex:1; height: 6px; background: rgba(255,255,255,0.06); border-radius: 3px; overflow: hidden; min-width: 50px;">
            <div style="width: ${Math.min(parseFloat(sharePercent), 100)}%; height: 100%; background: ${sku.desarrollar === 1 ? '#10B981' : '#D4AF37'};"></div>
          </div>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="6">TOTAL CONSOLIDADO (${visibleSkus.length} SKUs)</td>
      <td><strong>${sumPieces.toLocaleString('es-MX')} pz</strong></td>
      <td><strong>${sumBoxes.toLocaleString('es-MX')} cjs</strong></td>
      <td><strong>${sumKg.toLocaleString('es-MX', { maximumFractionDigits: 1 })} kg</strong></td>
      <td style="color: var(--color-gold);"><strong>$${sumRevenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })} MXN</strong></td>
      <td style="color: #4ADE80;"><strong>$${sumWeighted.toLocaleString('es-MX', { maximumFractionDigits: 0 })} MXN</strong></td>
      <td><strong>100%</strong></td>
    </tr>
  `;
}

function updateCharts(formatPieces, channelRevenue, productVolume) {
  // 1. Format Mix Chart
  const ctxFormat = document.getElementById("chartFormatMix").getContext("2d");
  const formatLabels = Object.keys(formatPieces);
  const formatData = Object.values(formatPieces);

  if (charts.formatMix) charts.formatMix.destroy();
  charts.formatMix = new Chart(ctxFormat, {
    type: "doughnut",
    data: {
      labels: formatLabels.length ? formatLabels : ["Sin datos"],
      datasets: [{
        data: formatData.length ? formatData : [1],
        backgroundColor: ["#E53935", "#F59E0B", "#3B82F6", "#10B981"],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "right", labels: { color: "#CBD5E1", font: { family: "Outfit" } } }
      }
    }
  });

  // 2. Channel Revenue Chart
  const ctxChannel = document.getElementById("chartChannelRevenue").getContext("2d");
  const channelLabels = Object.keys(channelRevenue);
  const channelData = Object.values(channelRevenue);

  if (charts.channelRevenue) charts.channelRevenue.destroy();
  charts.channelRevenue = new Chart(ctxChannel, {
    type: "bar",
    data: {
      labels: channelLabels.length ? channelLabels.map(c => c.length > 15 ? c.slice(0, 15) + '...' : c) : ["Sin datos"],
      datasets: [{
        label: "Venta ($ MXN)",
        data: channelData.length ? channelData : [0],
        backgroundColor: "#D4AF37",
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#94A3B8" }, grid: { display: false } },
        y: { ticks: { color: "#94A3B8" }, grid: { color: "rgba(255,255,255,0.05)" } }
      }
    }
  });

  // 3. Product Volume Chart
  const ctxProduct = document.getElementById("chartProductVolume").getContext("2d");
  const prodLabels = Object.keys(productVolume);
  const prodData = Object.values(productVolume);

  if (charts.productVolume) charts.productVolume.destroy();
  charts.productVolume = new Chart(ctxProduct, {
    type: "bar",
    data: {
      labels: prodLabels.length ? prodLabels : ["Sin datos"],
      datasets: [{
        label: "Volumen (Kg Anual)",
        data: prodData.length ? prodData : [0],
        backgroundColor: "#D32F2F",
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: "#94A3B8" }, grid: { display: false } },
        y: { ticks: { color: "#94A3B8" }, grid: { color: "rgba(255,255,255,0.05)" } }
      }
    }
  });
}

function renderTopClientsTable(clientsList) {
  const tbody = document.getElementById("topClientsBody");
  tbody.innerHTML = "";

  clientsList.sort((a, b) => b.revenue - a.revenue);
  const top10 = clientsList.slice(0, 10);

  if (top10.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color: var(--text-muted);">Sin datos suficientes</td></tr>`;
    return;
  }

  top10.forEach((c, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>#${idx + 1}</strong></td>
      <td><strong>${c.name}</strong></td>
      <td><span class="tag">${c.channel}</span></td>
      <td>${c.kam}</td>
      <td>${c.pieces.toLocaleString('es-MX')} pz</td>
      <td>${c.kg.toLocaleString('es-MX', { maximumFractionDigits: 1 })} kg</td>
      <td style="color: var(--color-gold); font-weight: 700;">$${c.revenue.toLocaleString('es-MX', { maximumFractionDigits: 0 })}</td>
      <td style="color: #4ADE80;">$${c.weighted.toLocaleString('es-MX', { maximumFractionDigits: 0 })}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ==========================================================================
// 8. EXPORT CSV & DEMO LOADER
// ==========================================================================
function exportCSV() {
  if (projections.length === 0) {
    alert("No hay datos para exportar.");
    return;
  }

  const headers = ["ID", "KAM", "Cliente", "Canal", "Producto", "Formato", "SKU_Codigo", "Fase_Desarrollar", "Piezas_Anual", "Kg_Unitario", "Precio_Unitario_MXN", "Volumen_Total_Kg", "Venta_Total_MXN", "Probabilidad", "Venta_Ponderada_MXN", "Fecha_Inicio", "Notas"];
  
  const rows = projections.map(p => [
    p.id,
    `"${p.kamName || p.sellerName || 'KAM Food Service'}"`,
    `"${p.customerName}"`,
    `"${p.channel}"`,
    `"${p.flavorName}"`,
    `"${p.formatLabel}"`,
    `"${p.skuCode}"`,
    p.desarrollar !== undefined ? p.desarrollar : 1,
    p.piecesAnnual,
    p.netKgPerPiece,
    p.pricePerPiece,
    p.totalKg.toFixed(2),
    p.totalRevenue.toFixed(2),
    p.probability,
    (p.totalRevenue * p.probability).toFixed(2),
    `"${p.targetDate}"`,
    `"${(p.notes || '').replace(/"/g, '""')}"`
  ]);

  const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement("a");
  a.href = url;
  a.download = `Takumi_Ichi_Proyecciones_${portfolioFilter === 'develop' ? 'Fase1_Desarrollar' : 'Portafolio_Completo'}_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function exportSkuCSV() {
  let skuAggregates = {};
  let totalCompanyRevenue = 0;

  projections.forEach(p => {
    totalCompanyRevenue += p.totalRevenue;
    if (!skuAggregates[p.skuCode]) {
      skuAggregates[p.skuCode] = { pieces: 0, kg: 0, revenue: 0, weighted: 0 };
    }
    skuAggregates[p.skuCode].pieces += p.piecesAnnual;
    skuAggregates[p.skuCode].kg += p.totalKg;
    skuAggregates[p.skuCode].revenue += p.totalRevenue;
    skuAggregates[p.skuCode].weighted += (p.totalRevenue * p.probability);
  });

  const allSkus = getAllFlatSkus();
  const visibleSkus = portfolioFilter === "develop" 
    ? allSkus.filter(s => s.desarrollar === 1) 
    : allSkus;

  const headers = ["Codigo_SKU", "Producto", "Formato", "Fase_Desarrollo", "Precio_Unitario_MXN", "Cont_Neto_Kg", "Piezas_Proyectadas", "Cajas_Estimadas", "Volumen_Total_Kg", "Monto_Venta_MXN", "Venta_Ponderada_MXN", "Porcentaje_Participacion"];
  
  const rows = visibleSkus.map(s => {
    const agg = skuAggregates[s.code] || { pieces: 0, kg: 0, revenue: 0, weighted: 0 };
    const approxBoxes = Math.ceil(agg.pieces / (s.unitsPerBox || 6));
    const sharePercent = totalCompanyRevenue > 0 ? ((agg.revenue / totalCompanyRevenue) * 100).toFixed(2) : "0.00";
    return [
      `"${s.code}"`,
      `"${s.flavorName}"`,
      `"${s.sizeLabel}"`,
      s.desarrollar === 1 ? "Fase 1 (Desarrollar)" : "Fase 2",
      s.priceMxn.toFixed(2),
      s.netKg.toFixed(2),
      agg.pieces,
      approxBoxes,
      agg.kg.toFixed(2),
      agg.revenue.toFixed(2),
      agg.weighted.toFixed(2),
      sharePercent + "%"
    ];
  });

  const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement("a");
  a.href = url;
  a.download = `Takumi_Ichi_Metricas_Por_SKU_${portfolioFilter === 'develop' ? 'Fase1' : 'Total'}_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// ==========================================================================
// 9. INITIALIZATION
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
  loadStoredProjections();
  renderFlavorTabs();
  renderProductShowcase();
  populateFormSelects();

  document.getElementById("formProjection").addEventListener("submit", handleFormSubmit);
  document.getElementById("btnExportCSV").addEventListener("click", exportCSV);
  document.getElementById("btnQuickExport").addEventListener("click", exportCSV);
  document.getElementById("btnExportSkuCSV").addEventListener("click", exportSkuCSV);

  const btnSync = document.getElementById("btnSyncSheets");
  if (btnSync) {
    btnSync.addEventListener("click", window.refreshFromGoogleSheets);
  }

  document.getElementById("btnClearData").addEventListener("click", () => {
    if (confirm("¿Estás seguro de que deseas limpiar la vista local? (Los datos permanecen seguros en tu Google Sheet)")) {
      projections = [];
      saveLocalCache();
      renderProjectionsTable();
      renderProductShowcase();
      updateDashboard();
    }
  });

  document.getElementById("btnLoadDemo").addEventListener("click", () => {
    projections = [...FALLBACK_DEMO_PROJECTIONS];
    saveLocalCache();
    renderProjectionsTable();
    renderProductShowcase();
    updateDashboard();
    alert("Datos cargados localmente.");
  });
});
