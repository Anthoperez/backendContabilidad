// src/app.service.ts
import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository, Like, In } from 'typeorm';
import { Gasto } from './gasto.entity';
import { PicMetadata, IngresoPic, GastoPrevioPic } from './pic-metadata.entity';
import * as ExcelJS from 'exceljs';

const mapaModalidadesPIC = new Map<string, string[]>();


// Esta es la estructura que enviaremos desde Angular
export interface PicMetadataDto {
  projectName: string;
  investigador?: string;
  tesista?: string;
  asesor?: string;
  duracion?: string;
  presupuestoTotal?: number | null;
  ingresos?: IngresoPic[];
  gastosAnosAnteriores?: GastoPrevioPic[];
}

interface GastoAnoAnterior {
  year: number;
  bienesCorrientes: number | null;
  bienesCapital: number | null;
  servicios: number | null;
  subvencion: number | null;
  viaticos: number | null;
  encargoInterno: number | null;
}


// ▼▼▼ INTERFAZ ReportMetadata AMPLIADA (Sin cambios desde la última vez) ▼▼▼
interface ReportMetadata {
  tituloProyecto?: string;
  codigoProyecto?: string;
  investigador?: string;
  rr_investigador?: string;
  fechaInicio?: string;
  fechaCulminacion?: Date;
  duracion?: string;
  presupuestoEntidades?: {
    nombreEntidad: string;
    aporteNoMonetario: number | null;
    aporteMonetario: number | null;
  }[];
  ingresos?: { descripcion: string; monto: number | null }[];
  gastosAnosAnteriores?: GastoAnoAnterior[];
}
// ▲▲▲ FIN ReportMetadata AMPLIADA ▲▲▲

// --- INTERFAZ PARA REPORTE MENSUAL (Sin cambios) ---
interface ReporteDataMensual {
  gastosCorrientes: Gasto[];
  subvencionesEnsenanza: Gasto[];
  subvenciones: Gasto[];
  gastosCapital: Gasto[];
}

// ▼▼▼ CORRECCIÓN ERROR 3 (Interfaz) ▼▼▼
// La interfaz ahora debe incluir los campos que faltaban
interface ReporteDataConsolidado {
  bienesCorrientes: number;
  bienesCapital: number;
  servicios: number;
  subvencion: number;
  viaticos: number; // <-- AÑADIDO
  encargoInterno: number; // <-- AÑADIDO
  totalMes: number; // <-- AÑADIDO
}
// ▲▲▲ FIN CORRECCIÓN ERROR 3 (Interfaz) ▲▲▲

// --- INTERFACES ESPECIALIZADAS PARA PIC ---
// Define las categorías de gasto del reporte PIC
type PicCategory =
  | 'BIENES CORRIENTES' // 2.3.1
  | 'BIENES CAPITAL' // 2.6
  | 'SERVICIOS' // 2.3.2 (general)
  | 'SUBVENCION' // 2.5
  | 'VIATICOS' // 2.3.2 1.1

// Para el resumen mensual del consolidado PIC
interface PicMonthSummary {
  'BIENES CORRIENTES': number;
  'BIENES CAPITAL': number;
  SERVICIOS: number;
  SUBVENCION: number;
  VIATICOS: number;
  'CUENTA ENCARGO': number;
  TOTAL: number;
}


// --- PIC 2021 ---
mapaModalidadesPIC.set('V CONVOCATORIA - 2021 MODALIDAD 1: PROYECTOS DE INVESTIGACIÓN BÁSICA Y APLICADA', [
  'PIC 01-2021 MOD. 01 - V CONV.',
  'PIC 02-2021 MOD. 01 - V CONV.',
  'PIC 03-2021 MOD. 01 - V CONV.',
  'PIC 04-2021 MOD. 01 - V CONV.',
  'PIC 05-2021 MOD. 01 - V CONV.'
]);
mapaModalidadesPIC.set('V CONVOCATORIA - 2021 MODALIDAD 2: PROYECTOS DE INVESTIGACIÓN DE TESIS DE PREGRADO', [
  'PIC 01-2021 MOD. 02 - V CONV.A',
  'PIC 02-2021 MOD. 02 - V CONV.',
  'PIC 03-2021 MOD. 02 - V CONV.',
  'PIC 04-2021 MOD. 02 - V CONV.',
  'PIC 05-2021 MOD. 02 - V CONV.',
  'PIC 06-2021 MOD. 02 - V CONV.'
]);

mapaModalidadesPIC.set(
  'V CONVOCATORIA - 2021 MODALIDAD 3: PROYECTOS DE INVESTIGACIÓN DE TESIS DE POSGRADO',[
    'PIC 01-2021 MOD. 03 - V CONV.',
    'PIC 02-2021 MOD. 03 - V CONV.',
    'PIC 03-2021 MOD. 03 - V CONV.',
    'PIC 04-2021 MOD. 03 - V CONV.',
    'PIC 05-2021 MOD. 03 - V CONV.',
    'PIC 06-2021 MOD. 03 - V CONV.',
    'PIC 07-2021 MOD. 03 - V CONV.',
    'PIC 08-2021 MOD. 03 - V CONV.',
    'PIC 09-2021 MOD. 03 - V CONV.',
    'PIC 10-2021 MOD. 03 - V CONV.',
    'PIC 11-2021 MOD. 03 - V CONV.',
    'PIC 12-2021 MOD. 03 - V CONV.',
    'PIC 13-2021 MOD. 03 - V CONV.',
    'PIC 14-2021 MOD. 03 - V CONV.',
    'PIC 15-2021 MOD. 03 - V CONV.',
]);

mapaModalidadesPIC.set(
  'V CONVOCATORIA - 2021 MODALIDAD 4: PROYECTOS DE PUBLICACIONES',
  [
    'PIC 01-2021 MOD. 04 - V CONV.',
    'PIC 02-2021 MOD. 04 - V CONV.',
    'PIC 03-2021 MOD. 04 - V CONV.',
    'PIC 04-2021 MOD. 04 - V CONV.',
    'PIC 05-2021 MOD. 04 - V CONV.',
    'PIC 06-2021 MOD. 04 - V CONV.',
    'PIC 07-2021 MOD. 04 - V CONV.',
    'PIC 08-2021 MOD. 04 - V CONV.',
    'PIC 09-2021 MOD. 04 - V CONV.',
    'PIC 10-2021 MOD. 04 - V CONV.',
    'PIC 11-2021 MOD. 04 - V CONV.',
    'PIC 12-2021 MOD. 04 - V CONV.',
    'PIC 13-2021 MOD. 04 - V CONV.',
    'PIC 14-2021 MOD. 04 - V CONV.',
    'PIC 15-2021 MOD. 04 - V CONV.',
    'PIC 16-2021 MOD. 04 - V CONV.',
    'PIC 17-2021 MOD. 04 - V CONV.',
    'PIC 18-2021 MOD. 04 - V CONV.',
    'PIC 19-2021 MOD. 04 - V CONV.',
    'PIC 20-2021 MOD. 04 - V CONV.',
    'PIC 21-2021 MOD. 04 - V CONV.',
    'PIC 22-2021 MOD. 04 - V CONV.',
  ],
);

mapaModalidadesPIC.set('V CONVOCATORIA - 2021 MODALIDAD MODALIDAD 4-II: PROYECTOS DE INVESTIGACIÓN PUBLICACIONES (LIBROS)', [
  'PIC 23-2021 MOD. 04-II - V CONV.',
  'PIC 24-2021 MOD. 04-II - V CONV.'
]);

mapaModalidadesPIC.set(
  'VI CONVOCATORIA - 2022 MODALIDAD 1: PROYECTOS DE INVESTIGACIÓN APLICADA',
  [
    'PIC 01-2022 MOD. 01 - VI CONV.',
    'PIC 02-2022 MOD. 01 - VI CONV.',
    'PIC 03-2022 MOD. 01 - VI CONV.',
    'PIC 04-2022 MOD. 01 - VI CONV.',
  ],
);

mapaModalidadesPIC.set('VI CONVOCATORIA - 2022 MODALIDAD 2: CATEGORIAS CONSOLIDADO Y POR CONSOLIDAR', [
  'PIC 01-2022 MOD. 02 - VI CONV.',
  'PIC 02-2022 MOD. 02 - VI CONV.',
]);

mapaModalidadesPIC.set('VI CONVOCATORIA - 2022 MODALIDAD 2: CATEGORIA EMERGENTE', [
  'PIC 03-2022 MOD. 02 - VI CONV.',
  'PIC 04-2022 MOD. 02 - VI CONV.',
]);

mapaModalidadesPIC.set('VI CONVOCATORIA - 2022 MODALIDAD: PROYECTOS EMBLEMATICO', [
  'PIC 01-2023 - EMBLEMATICO',
  'PIC 02-2023 - EMBLEMATICO',
  'PIC 03-2023 -EMBLEMATICO ',
]);

mapaModalidadesPIC.set(
  'VII CONVOCATORIA - 2023 MODALIDAD: 01 - PROY. INVESTIGACION CIENTIFICA',
  [
    'PIC 01-2023 MOD. 01 - VII CONV.',
    'PIC 02-2023 MOD. 01 - VII CONV.',
    'PIC 03-2023 MOD. 01 - VII CONV.',
    'PIC 04-2023 MOD. 01 - VII CONV.',
    'PIC 05-2023 MOD. 01 - VII CONV.',
    'PIC 06-2023 MOD. 01 - VII CONV.',
    'PIC 07-2023 MOD. 01 - VII CONV.',
    'PIC 08-2023 MOD. 01 - VII CONV.',
  ],
);

mapaModalidadesPIC.set('VII CONVOCATORIA - 2023 MODALIDAD: 02 - EN CIENCIAS SOCIALES', [
  'PIC 01-2023 MOD. 02 - VII CONV.',
  'PIC 02-2023 MOD. 02 - VII CONV.',
]);

mapaModalidadesPIC.set(
  'VII CONVOCATORIA - 2023 MODALIDAD: 03 - PROY. INVESTIGACION EMBLEMATICA',
  [
    'PIC 01-2023 MOD. 03 - VII CONV.',
    'PIC 02-2023 MOD. 03 - VII CONV.',
    'PIC 03-2023 MOD. 03 - VII CONV.',
    'PIC 04-2023 MOD. 03 - VII CONV.',
    'PIC 05-2023 MOD. 03 - VII CONV.',
    'PIC 06-2023 MOD. 03 - VII CONV.',
    'PIC 07-2023 MOD. 03 - VII CONV.',
  ],
);


mapaModalidadesPIC.set('VIII CONVOCATORIA - 2024 MODALIDAD: 01 - PROY. INVESTIGACION CIENTIFICA', [
  'PIC 01-2024 MOD. 01 - VIII CONV.',
  'PIC 02-2024 MOD. 01 - VIII CONV.',
]);


mapaModalidadesPIC.set(
  'VIII CONVOCATORIA - 2024 MODALIDAD: 02 - PROY. INVESTIGACION CIENTIFICA',
  [
    'PIC 01-2024 MOD. 02 - VIII CONV.',
    'PIC 02-2024 MOD. 02 - VIII CONV.',
    'PIC 03-2024 MOD. 02 - VIII CONV.',
    'PIC 04-2024 MOD. 02 - VIII CONV.',
    'PIC 05-2024 MOD. 02 - VIII CONV.',
    'PIC 06-2024 MOD. 02 - VIII CONV.',
    'PIC 07-2024 MOD. 02 - VIII CONV.',
    'PIC 08-2024 MOD. 02 - VIII CONV.',
    'PIC 09-2024 MOD. 02 - VIII CONV.',
    'PIC 10-2024 MOD. 02 - VIII CONV.',
    'PIC 11-2024 MOD. 02 - VIII CONV.',
    'PIC 12-2024 MOD. 02 - VIII CONV.',
    'PIC 13-2024 MOD. 02 - VIII CONV.',
    'PIC 14-2024 MOD. 02 - VIII CONV.',
    'PIC 15-2024 MOD. 02 - VIII CONV.',
    'PIC 16-2024 MOD. 02 - VIII CONV.',
    'PIC 17-2024 MOD. 02 - VIII CONV.',
  ],
);


mapaModalidadesPIC.set(
  'VIII CONVOCATORIA - 2024 MODALIDAD: 03 - PROY. INVESTIGACION CIENTIFICA',
  [
    'PIC 01-2024 MOD. 03 - VIII CONV.',
    'PIC 02-2024 MOD. 03 - VIII CONV.',
    'PIC 03-2024 MOD. 03 - VIII CONV.',
    'PIC 04-2024 MOD. 03 - VIII CONV.',
  ],
);

@Injectable()
export class AppService {
  constructor(
    @InjectRepository(Gasto)
    private gastoRepository: Repository<Gasto>,

    @InjectRepository(PicMetadata)
    private picMetadataRepository: Repository<PicMetadata>,

  ) {}

  // --- MÉTODOS DE GESTIÓN DE DATOS (Sin cambios) ---
  create(gasto: Partial<Gasto>): Promise<Gasto> {
    const nuevoGasto = this.gastoRepository.create(gasto);
    return this.gastoRepository.save(nuevoGasto);
  }

  createMany(gastos: Partial<Gasto>[]): Promise<Gasto[]> {
    const gastosLimpios = gastos.map((gasto) => {
      if (gasto.proyecto && typeof gasto.proyecto === 'string') {
        gasto.proyecto = gasto.proyecto.trim();
      }
      if (gasto.numeroDocumento) {
        gasto.numeroDocumento = String(gasto.numeroDocumento);
      }
      if (gasto.siaf) {
        gasto.siaf = String(gasto.siaf);
      }
      return gasto;
    });
    const nuevosGastos = this.gastoRepository.create(gastosLimpios);
    return this.gastoRepository.save(nuevosGastos);
  }

  
  findAll(): Promise<Gasto[]> {
    return this.gastoRepository.find();
  }

  async findUniqueProjects(): Promise<string[]> {
    const projects = await this.gastoRepository
      .createQueryBuilder('gasto')
      .select('gasto.proyecto', 'proyecto')
      .distinct(true)
      .orderBy('proyecto', 'ASC')
      .getRawMany();
    return projects.map((p) => p.proyecto).filter(Boolean);
  }


  /**
   * NUEVA FUNCIÓN
   * Obtiene solo los proyectos que son CONTRATO.
   */
  async findContratoProjects(): Promise<string[]> {
    const projects = await this.gastoRepository
      .createQueryBuilder('gasto')
      .select('gasto.proyecto', 'proyecto')
      .where('gasto.proyecto LIKE :query', { query: '%CONTRATO%' })
      .distinct(true)
      .orderBy('proyecto', 'ASC')
      .getRawMany();
    return projects.map((p) => p.proyecto).filter(Boolean);
  }

  /**
   * NUEVA FUNCIÓN
   * Obtiene solo los proyectos que son PIC.
   */
  async findPicProjects(): Promise<string[]> {
    const projects = await this.gastoRepository
      .createQueryBuilder('gasto')
      .select('gasto.proyecto', 'proyecto')
      .where('gasto.proyecto LIKE :query', { query: '%PIC%' })
      .distinct(true)
      .orderBy('proyecto', 'ASC')
      .getRawMany();
    return projects.map((p) => p.proyecto).filter(Boolean);
  }
  // --- FIN DE MÉTODOS DE GESTIÓN DE DATOS ---



  /**
   * NUEVA FUNCIÓN
   * Obtiene la "lista maestra" de TODOS los proyectos PIC definidos en el mapa.
   */
  async getMasterPicProjectList(): Promise<string[]> {
    const allProjectsSet = new Set<string>();

    // Itera sobre el mapa y añade cada proyecto de cada modalidad
    mapaModalidadesPIC.forEach((projectList) => {
      projectList.forEach(projectName => {
        // .trim() por si hay algún espacio extra
        allProjectsSet.add(projectName.trim()); 
      });
    });

    // Convertir a array y ordenar alfabéticamente
    const allProjectsArray = Array.from(allProjectsSet);
    allProjectsArray.sort((a, b) => a.localeCompare(b));
    
    return allProjectsArray;
  }

  /**
   * Obtiene la metadata guardada para un proyecto PIC específico.
   */
  async getPicMetadata(projectName: string): Promise<PicMetadata | null> {
    if (!projectName) {
      return null;
    }
    return this.picMetadataRepository.findOneBy({ projectName });
  }

  /**
   * Guarda o actualiza la metadata para un proyecto PIC específico.
   * Esto es un "upsert": si existe, lo actualiza; si no, lo crea.
   */
  async savePicMetadata(
    projectName: string,
    data: PicMetadataDto,
  ): Promise<PicMetadata> {
    // 1. Buscar si ya existe metadata para este proyecto
    let metadata = await this.picMetadataRepository.findOneBy({ projectName });

    if (metadata) {
      // 2. Si existe, la actualizamos
      Object.assign(metadata, data);
    } else {
      // 3. Si no existe, creamos una nueva entidad
      metadata = this.picMetadataRepository.create(data);
    }

    // 4. Guardar los cambios en la base de datos
    return this.picMetadataRepository.save(metadata);
  }

  // ▲▲▲ FIN DE LAS NUEVAS FUNCIONES ▲▲▲
  
  // --- MÉTODO PRINCIPAL DE GENERACIÓN DE REPORTES (Sin cambios) ---
  async generateContratoReport(
    projectName: string,
    metadata: ReportMetadata = {},
  ): Promise<ExcelJS.Workbook> {
    const gastosDelProyecto = await this.gastoRepository.find({
      where: { proyecto: projectName },
      order: { fechaDevengado: 'ASC' },
    });

    const gastosAgrupadosPorMes = this.groupGastosByMonth(gastosDelProyecto);
    const workbook = new ExcelJS.Workbook();

    this.crearHojaConsolidado(
      workbook,
      projectName,
      gastosAgrupadosPorMes,
      metadata,
    );

    const monthOrderShort = [
      'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN',
      'JUL', 'AGO', 'SET', 'OCT', 'NOV', 'DIC',
    ];
    const monthOrderFull = [
      'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
    ];

    monthOrderFull.forEach((monthFullName, index) => {
      const monthShort = monthOrderShort[index];
      const gastosDelMes = gastosAgrupadosPorMes.get(monthFullName) || [];
      const dataMensual = this.clasificarGastosParaMes(gastosDelMes);
      this.crearHojaDeReportePorMes(
        workbook,
        monthShort,
        dataMensual,
        projectName,
      );
    });

    return workbook;
  }

  // ▼▼▼ CORRECCIÓN ERROR 1 (Funciones auxiliares incluidas) ▼▼▼

  // --- LÓGICA DE CLASIFICACIÓN Y AGRUPACIÓN ---
  private groupGastosByMonth(gastos: Gasto[]): Map<string, Gasto[]> {
    const gastosPorMes = new Map<string, Gasto[]>();
    const monthNames = [
      'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
    ];
    for (const gasto of gastos) {
      if (gasto.fechaDevengado) {
        // Aseguramos que la fecha se interprete correctamente (como UTC)
        const date = new Date(gasto.fechaDevengado);
        const monthName = monthNames[date.getUTCMonth()]; // Usar getUTCMonth()
        if (!gastosPorMes.has(monthName)) {
          gastosPorMes.set(monthName, []);
        }
        gastosPorMes.get(monthName)!.push(gasto);
      }
    }
    return gastosPorMes;
  }

  private clasificarGastosParaMes(gastos: Gasto[]): ReporteDataMensual {
    const reporte: ReporteDataMensual = {
      gastosCorrientes: [],
      subvencionesEnsenanza: [],
      subvenciones: [],
      gastosCapital: [],
    };

    for (const gasto of gastos) {
      const especifica = gasto.especifica?.trim().replace(/\s+/g, ' ') || '' || '';

      if (especifica.startsWith('2.3.')) {
        reporte.gastosCorrientes.push(gasto);
      } else if (especifica.startsWith('2.5.')) {
        reporte.subvenciones.push(gasto);
      } else if (especifica.startsWith('2.6.')) {
        reporte.gastosCapital.push(gasto);
      } else {
        // Categoría "catch-all" (ej. 2.5.3 1.1 2 en tu CSV)
        reporte.subvencionesEnsenanza.push(gasto);
      }
    }
    return reporte;
  }

  // ▼▼▼ CORRECCIÓN ERROR 3 (Implementación) ▼▼▼
  // Actualizamos esta función para que calcule las nuevas categorías
private clasificarGastosParaConsolidado(
    gastos: Gasto[],
  ): ReporteDataConsolidado {
    const reporte: ReporteDataConsolidado = {
      bienesCorrientes: 0,
      bienesCapital: 0,
      servicios: 0,
      subvencion: 0,
      viaticos: 0,
      encargoInterno: 0,
      totalMes: 0,
    };

    // --- INICIO DE LA NUEVA LÓGICA ---
    for (const gasto of gastos) {
      // 1. Obtenemos el monto total (monto + monto2)
      // Esta era una corrección clave: sumar ambos montos.
      const montoTotal = (Number(gasto.monto) || 0) + (Number(gasto.monto2) || 0);

      // Si el monto es 0, no hay nada que clasificar
      if (montoTotal === 0) continue;

      // 2. Obtenemos los campos para las reglas
      const doc = gasto.tipoDocumento?.trim().toUpperCase() || '';
      const especifica = gasto.especifica?.trim().replace(/\s+/g, ' ') || '';

      // 3. Acumulamos el total del mes
      reporte.totalMes += montoTotal;

      // 4. Aplicamos las reglas de negocio en orden de prioridad
      
      // Regla 1: Viáticos (P/V o C/S)
      if (doc === 'P/V' || doc === 'C/S') {
        reporte.viaticos += montoTotal;
      }
      // Regla 2: Encargo Interno (R.DGA)
      else if (doc === 'R.DGA') {
        reporte.encargoInterno += montoTotal;
      }
      // Regla 3: Servicios (O/S)
      else if (doc === 'O/S') {
        reporte.servicios += montoTotal;
      }
      // Regla 4: Subvención (P/S)
      else if (doc === 'P/S') {
        reporte.subvencion += montoTotal;
      }
      // Regla 5: Bienes Corrientes (O/C + especifica 2.3.1)
      else if (doc === 'O/C' && especifica.startsWith('2.3.')) {
        reporte.bienesCorrientes += montoTotal;
      }
      // Regla 6: Bienes Capital (solo especifica 2.6)
      else if (especifica.startsWith('2.6.')) {
        reporte.bienesCapital += montoTotal;
      }
      // (Cualquier otra combinación, como O/C + 2.3.2, no se suma 
      // a ninguna categoría específica, pero ya se sumó al totalMes)
    }
    // --- FIN DE LA NUEVA LÓGICA ---

    // Ya no necesitamos calcular el totalMes al final,
    // se calculó dentro del bucle.
    return reporte;
  }
  // ▲▲▲ FIN CORRECCIÓN ERROR 3 (Implementación) ▲▲▲

  private clasificarGastosParaPicConsolidado(
    gastos: Gasto[],
  ): PicMonthSummary {
    // Inicializa la estructura de resumen que espera el reporte PIC
    const reporte: PicMonthSummary = {
      'BIENES CORRIENTES': 0,
      'BIENES CAPITAL': 0,
      SERVICIOS: 0,
      SUBVENCION: 0,
      VIATICOS: 0,
      'CUENTA ENCARGO': 0, // Equivalente a 'encargoInterno'
      TOTAL: 0,
    };

    for (const gasto of gastos) {
      // 1. Obtenemos el monto total (monto + monto2)
      const montoTotal = (Number(gasto.monto) || 0) + (Number(gasto.monto2) || 0);

      if (montoTotal === 0) continue;

      // 2. Obtenemos los campos para las reglas
      const doc = gasto.tipoDocumento?.trim().toUpperCase() || '';
      const especifica = gasto.especifica?.trim().replace(/\s+/g, ' ') || '';

      // 3. Acumulamos el total del mes
      reporte.TOTAL += montoTotal;

      // 4. Aplicamos las reglas de negocio (las mismas de Contrato)
      // y las mapeamos a las claves del reporte PIC
      
      if (doc === 'P/V' || doc === 'C/S') {
        reporte.VIATICOS += montoTotal;
      }
      else if (doc === 'R.DGA') {
        reporte['CUENTA ENCARGO'] += montoTotal; // Mapeado a 'CUENTA ENCARGO'
      }
      else if (doc === 'O/S') {
        reporte.SERVICIOS += montoTotal;
      }
      else if (doc === 'P/S') {
        reporte.SUBVENCION += montoTotal;
      }
      else if (doc === 'O/C' && especifica.startsWith('2.3.')) {
        reporte['BIENES CORRIENTES'] += montoTotal; // Mapeado a 'BIENES CORRIENTES'
      }
      else if (especifica.startsWith('2.6.')) {
        reporte['BIENES CAPITAL'] += montoTotal; // Mapeado a 'BIENES CAPITAL'
      }
    }
    
    return reporte;
  }


  // ▲▲▲ FIN CORRECCIÓN ERROR 1 ▲▲▲

  // --- CONSTRUCTOR DE HOJA CONSOLIDADO (Con corrección para Error 2) ---
  private crearHojaConsolidado(
    workbook: ExcelJS.Workbook,
    projectName: string,
    gastosPorMes: Map<string, Gasto[]>,
    metadata: ReportMetadata,
  ): void {
    const worksheet = workbook.addWorksheet('CONSOLIDADO 2025');
    const titleStyle: Partial<ExcelJS.Style> = {
      font: { bold: true, size: 12 },
      alignment: { horizontal: 'center' },
    };
    const boldStyle: Partial<ExcelJS.Style> = { font: { bold: true } };
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';
    const borderStyle: Partial<ExcelJS.Border> = {
      style: 'thin',
      color: { argb: 'FF000000' },
    };
    const allBorders = {
      top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle,
    };
    const centerAlign: Partial<ExcelJS.Alignment> = {
      horizontal: 'center',
      vertical: 'middle',
    };

    // Definir anchos de columna
    worksheet.columns = [
      { width: 20 }, // A: Descripción / Mes/Año
      { width: 15 }, // B
      { width: 15 }, // C
      { width: 15 }, // D
      { width: 15 }, // E
      { width: 15 }, // F
      { width: 15 }, // G
      { width: 15 }, // H
    ];
    // Corrección para que la descripción de ingresos sea visible
    worksheet.getColumn('A').alignment = { wrapText: true, vertical: 'middle' };


    let currentRow = 1;

    // ... (Título principal) ...
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS DEL PROYECTO DE INVESTIGACIÓN - PROCIENCIA';
    worksheet.getCell(`A${currentRow}`).style = titleStyle;
    currentRow++;
    currentRow++;

    if (metadata.tituloProyecto) {
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`); // Ocupar todo el ancho
      worksheet.getCell(`A${currentRow}`).value = metadata.tituloProyecto;
      worksheet.getCell(`A${currentRow}`).alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      worksheet.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
      // Ajustar altura de fila si el texto es largo (aproximación)
      worksheet.getRow(currentRow).height = 45; 
      currentRow++; 
    }

    // ... (Título del proyecto) ...
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = projectName;
    worksheet.getCell(`A${currentRow}`).font = { bold: true, size: 11 };
    currentRow+=2;

    if (metadata.codigoProyecto) {
      worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = `Código del Proyecto: ${metadata.codigoProyecto}`;
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      currentRow++;
    }
    // ... (Fila Investigador y Fecha Inicio) ...
    worksheet.getCell(`A${currentRow}`).value = 'INVESTIGADOR:';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`B${currentRow}:D${currentRow}`);
    worksheet.getCell(`B${currentRow}`).value = metadata.investigador || '';
    worksheet.getCell(`B${currentRow}`).border = allBorders;
    worksheet.getCell(`E${currentRow}`).value = 'Fecha de inicio:';
    worksheet.getCell(`E${currentRow}`).font = { bold: true };
    
    const fechaInicioCell = worksheet.getCell(`F${currentRow}`);
    worksheet.getCell(`F${currentRow}`).value = metadata.fechaInicio || '';
    // if (metadata.fechaInicio) {
    //   fechaInicioCell.value = new Date(metadata.fechaInicio);
    //   fechaInicioCell.numFmt = 'dd/mm/yyyy';
    // } else {
    //   fechaInicioCell.value = '';
    // }
    fechaInicioCell.border = allBorders;

    worksheet.getCell(`G${currentRow}`).value = 'Fecha de Culminacion:';
    worksheet.getCell(`G${currentRow}`).font = { bold: true };
    
    const fechaCulminacionCell = worksheet.getCell(`H${currentRow}`);
    if (metadata.fechaCulminacion) {
      fechaCulminacionCell.value = new Date(metadata.fechaCulminacion);
      fechaCulminacionCell.numFmt = 'dd/mm/yyyy';
    } else {
      fechaCulminacionCell.value = '';
    }
    fechaCulminacionCell.border = allBorders;
    currentRow++;

    // ... (Fila RR y Duración) ...
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = metadata.rr_investigador || '';
    worksheet.getCell(`A${currentRow}`).border = allBorders;
    worksheet.getCell(`E${currentRow}`).value = 'Duración:';
    worksheet.getCell(`E${currentRow}`).font = { bold: true };
    
    worksheet.getCell(`F${currentRow}`).value = metadata.duracion || '';
    worksheet.getCell(`F${currentRow}`).border = allBorders;
    currentRow++;
    
    currentRow++; // Salto de línea

    // ... (SECCIÓN DESCRIPCIÓN DE PRESUPUESTO - Cabeceras)
   // --- SECCIÓN DESCRIPCIÓN DE PRESUPUESTO (DINÁMICA) ---
    // Cabeceras
    worksheet.getCell(`A${currentRow}`).value = 'DESCRIPCION DE PRESUPUESTO';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = 'Aporte No Monetario (Valorizado S/)'; // CAMBIADO
    worksheet.getCell(`C${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    worksheet.getCell(`C${currentRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell(`D${currentRow}`).value = 'Aporte Monetario S/';
    worksheet.getCell(`D${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    worksheet.getCell(`D${currentRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell(`E${currentRow}`).value = 'Aporte Total S/';
    worksheet.getCell(`E${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    for (let col = 1; col <= 6; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;

    // Variables para sumar el total
    let totalAporteNoMonetario = 0;
    let totalAporteMonetario = 0;
    let totalPresupuestoGeneral = 0;

    // Bucle dinámico para las entidades de presupuesto
    if (metadata.presupuestoEntidades && metadata.presupuestoEntidades.length > 0) {
      metadata.presupuestoEntidades.forEach(entidad => {
        const nombre = entidad.nombreEntidad || 'Entidad sin nombre';
        const aporteNoMonetario = Number(entidad.aporteNoMonetario) || 0;
        const aporteMonetario = Number(entidad.aporteMonetario) || 0;
        const totalFila = aporteNoMonetario + aporteMonetario;

        worksheet.getCell(`A${currentRow}`).value = nombre;
        worksheet.getCell(`A${currentRow}`).font = { bold: true };
        worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
        
        worksheet.getCell(`C${currentRow}`).value = aporteNoMonetario || null;
        worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
        
        worksheet.getCell(`D${currentRow}`).value = aporteMonetario || null;
        worksheet.getCell(`D${currentRow}`).numFmt = moneyFormat;
        
        worksheet.getCell(`E${currentRow}`).value = totalFila || null;
        worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
        
        
        for(let col = 1; col <= 6; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;

        // Acumular totales
        totalAporteNoMonetario += aporteNoMonetario;
        totalAporteMonetario += aporteMonetario;
        totalPresupuestoGeneral += totalFila;

        currentRow++;
      });
    } else {
      // Fila vacía si no hay datos (o puedes poner 3 filas vacías)
      worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
      for(let col = 1; col <= 6; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
      currentRow++;
    }

    // Fila TOTAL PRESUPUESTO (calculada dinámicamente)
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL PRESUPUESTO';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    
    worksheet.getCell(`C${currentRow}`).value = totalAporteNoMonetario || null;
    worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
    
    
    worksheet.getCell(`D${currentRow}`).value = totalAporteMonetario || null;
    worksheet.getCell(`D${currentRow}`).numFmt = moneyFormat;
    
    worksheet.getCell(`E${currentRow}`).value = totalPresupuestoGeneral || null;
    worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
    
    for(let col = 1; col <= 6; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;
    currentRow++; // Salto de línea

    // ▲▲▲ FIN DE LA SECCIÓN DE REEMPLAZO ▲▲▲

    // ... (SECCIÓN INGRESOS) ...

    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    const ingresosHeaderCell = worksheet.getCell(`A${currentRow}`);
    ingresosHeaderCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'fff2cc' } };
    ingresosHeaderCell.value = 'INGRESOS';
    ingresosHeaderCell.font = { bold: true, size:16  };
    currentRow++;

    // ▼▼▼ MODIFICACIÓN AQUÍ ▼▼▼
    let currentTotalIngresos = 0;
    // Obtenemos el año actual (ej: "2025") para filtrar
    const currentYearString = String(new Date().getFullYear());
    // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

    if (metadata.ingresos && metadata.ingresos.length > 0) {
      metadata.ingresos.forEach(ingreso => {
        if (ingreso.descripcion || ingreso.monto !== null) {


          // ▼▼▼ MODIFICACIÓN AQUÍ ▼▼▼
          const descripcion = ingreso.descripcion || '';
          const monto = Number(ingreso.monto) || 0;

          // Lógica para sumar solo si la descripción contiene el año actual
          // (ej: en "R.R Nº 0209-2025" o en "(05/02/2025)")
          if (descripcion.includes(currentYearString)) {
            currentTotalIngresos += monto;
          }
          // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲



          worksheet.mergeCells(`A${currentRow}:G${currentRow}`);
          worksheet.getCell(`A${currentRow}`).value = descripcion;
          worksheet.getCell(`H${currentRow}`).value = ingreso.monto || null;
          worksheet.getCell(`H${currentRow}`).numFmt = moneyFormat;
          for(let col = 1; col <= 8; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
          
          currentRow++;
        }
      });
    } else {
      worksheet.mergeCells(`A${currentRow}:G${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = '';
      worksheet.getCell(`H${currentRow}`).value = null;
      for(let col = 1; col <= 8; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
      currentRow++;
    }
    worksheet.mergeCells(`A${currentRow}:G${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL INGRESOS 2025';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`H${currentRow}`).value = currentTotalIngresos || null;
    worksheet.getCell(`H${currentRow}`).numFmt = moneyFormat;
    worksheet.getCell(`H${currentRow}`).font = { bold: true };
    for(let col = 1; col <= 8; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;
    currentRow++;

    // ... (PRIMERA SECCIÓN EJECUCIÓN DE GASTOS - Cabeceras) ...
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS';
    worksheet.getCell(`A${currentRow}`).style = boldStyle;
    worksheet.getCell(`A${currentRow}`).style.font = { size: 16, bold: true };
    worksheet.getCell(`A${currentRow}`).style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'fff2cc' } };
    currentRow++;
    const headers = ['MES/AÑO', 'BIENES CORRIENTES', 'BIENES CAPITAL', 'SERVICIOS', 'SUBVENCION', 'VIATICOS', 'ENCARGO INTERNO', 'TOTAL'];
    worksheet.getCell(`A${currentRow}`).value = headers[0];
    worksheet.getCell(`B${currentRow}`).value = headers[1];
    worksheet.getCell(`C${currentRow}`).value = headers[2];
    worksheet.getCell(`D${currentRow}`).value = headers[3];
    worksheet.getCell(`E${currentRow}`).value = headers[4];
    worksheet.getCell(`F${currentRow}`).value = headers[5];
    worksheet.getCell(`G${currentRow}`).value = headers[6];
    worksheet.getCell(`H${currentRow}`).value = headers[7];
    worksheet.getRow(currentRow).font = { bold: true };
    for(let col = 1; col <= 8; col++) {
      if(col === 1) worksheet.getCell(`A${currentRow}`);
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    }
    currentRow++;


    // 1. Inicializamos los totales acumulados de AÑOS ANTERIORES
    const previousYearsGastosTotals = {
      bienesCorrientes: 0,
      bienesCapital: 0,
      servicios: 0,
      subvencion: 0,
      viaticos: 0,
      encargoInterno: 0,
      total: 0,
    };

    // 2. Verificamos si el array existe y tiene datos
    if (metadata.gastosAnosAnteriores && metadata.gastosAnosAnteriores.length > 0) {
      
      // 3. Iteramos sobre CADA año anterior proporcionado
      metadata.gastosAnosAnteriores.forEach(gastoAno => {

    // ▼▼▼ CORRECCIÓN ERROR 2 ▼▼▼
    // Proporcionamos un objeto default completo si metadata.gastosAnoAnterior es nulo
    const gastosAnoAnterior = gastoAno || {
      year: new Date().getFullYear() - 1,
      bienesCorrientes: null,
      bienesCapital: null,
      servicios: null,
      subvencion: null,
      viaticos: null,
      encargoInterno: null,
    };
    // ▲▲▲ FIN CORRECCIÓN ERROR 2 ▲▲▲

    const yearAnterior = gastosAnoAnterior.year;
    const totalAnoAnterior =
      (gastosAnoAnterior.bienesCorrientes || 0) +
      (gastosAnoAnterior.bienesCapital || 0) +
      (gastosAnoAnterior.servicios || 0) +
      (gastosAnoAnterior.subvencion || 0) +
      (gastosAnoAnterior.viaticos || 0) +
      (gastosAnoAnterior.encargoInterno || 0);

    
    worksheet.getCell(`A${currentRow}`).value = `AÑO ${yearAnterior}`;
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`B${currentRow}`).value = gastosAnoAnterior.bienesCorrientes || null;
    worksheet.getCell(`C${currentRow}`).value = gastosAnoAnterior.bienesCapital || null;
    worksheet.getCell(`D${currentRow}`).value = gastosAnoAnterior.servicios || null;
    worksheet.getCell(`E${currentRow}`).value = gastosAnoAnterior.subvencion || null;
    worksheet.getCell(`F${currentRow}`).value = gastosAnoAnterior.viaticos || null;
    worksheet.getCell(`G${currentRow}`).value = gastosAnoAnterior.encargoInterno || null;
    worksheet.getCell(`H${currentRow}`).value = totalAnoAnterior || null;
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 3 && colNumber <= 8) cell.numFmt = moneyFormat;
      cell.border = allBorders;
    });

    // 5. Acumulamos los totales de TODOS los años anteriores
        previousYearsGastosTotals.bienesCorrientes += (gastosAnoAnterior.bienesCorrientes || 0);
        previousYearsGastosTotals.bienesCapital += (gastosAnoAnterior.bienesCapital || 0);
        previousYearsGastosTotals.servicios += (gastosAnoAnterior.servicios || 0);
        previousYearsGastosTotals.subvencion += (gastosAnoAnterior.subvencion || 0);
        previousYearsGastosTotals.viaticos += (gastosAnoAnterior.viaticos || 0);
        previousYearsGastosTotals.encargoInterno += (gastosAnoAnterior.encargoInterno || 0);
        previousYearsGastosTotals.total += totalAnoAnterior;

    currentRow++;
  });
}else {
  // (Comportamiento anterior si el array está vacío: imprime una fila por defecto)
      
      worksheet.getCell(`A${currentRow}`).value = `AÑO ${new Date().getFullYear() - 1}`;
      // ... (el resto de las celdas quedan vacías/nulas)
      worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell) => {
        cell.border = allBorders;
      });
      currentRow++;
    }

    currentRow++;
    // ▼▼▼ CÓDIGO A AÑADIR (AQUÍ) ▼▼▼
    // --- Fila TOTAL GASTOS (resumen de TODOS los años anteriores) ---
    
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL GASTOS';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`B${currentRow}`).value = previousYearsGastosTotals.bienesCorrientes || null;
    worksheet.getCell(`C${currentRow}`).value = previousYearsGastosTotals.bienesCapital || null;
    worksheet.getCell(`D${currentRow}`).value = previousYearsGastosTotals.servicios || null;
    worksheet.getCell(`E${currentRow}`).value = previousYearsGastosTotals.subvencion || null;
    worksheet.getCell(`F${currentRow}`).value = previousYearsGastosTotals.viaticos || null;
    worksheet.getCell(`G${currentRow}`).value = previousYearsGastosTotals.encargoInterno || null;
    worksheet.getCell(`H${currentRow}`).value = previousYearsGastosTotals.total || null;
    
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 2 && colNumber <= 8) cell.numFmt = moneyFormat;
      if (colNumber >= 2 && !cell.value) cell.value = 0; // Poner 0 si es nulo
      cell.border = allBorders;
    });
    worksheet.getRow(currentRow).font = { bold: true };
    currentRow++;
    // ▲▲▲ FIN DEL CÓDIGO A AÑADIR ▲▲▲
    

    currentRow++;

    // ... (SEGUNDA SECCIÓN EJECUCIÓN DE GASTOS (MENSUAL AÑO ACTUAL) - Cabeceras) ...
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS';
    worksheet.getCell(`A${currentRow}`).style = boldStyle;
    worksheet.getCell(`A${currentRow}`).style.font = { size: 16, bold: true };
    worksheet.getCell(`A${currentRow}`).style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'fff2cc' } };
    currentRow++;
    worksheet.getCell(`A${currentRow}`).value = headers[0];
    worksheet.getCell(`B${currentRow}`).value = headers[1];
    worksheet.getCell(`C${currentRow}`).value = headers[2];
    worksheet.getCell(`D${currentRow}`).value = headers[3];
    worksheet.getCell(`E${currentRow}`).value = headers[4];
    worksheet.getCell(`F${currentRow}`).value = headers[5];
    worksheet.getCell(`G${currentRow}`).value = headers[6];
    worksheet.getCell(`H${currentRow}`).value = headers[7];
    worksheet.getRow(currentRow).font = { bold: true };
    for(let col = 1; col <= 8; col++) {
      if(col === 1) worksheet.getCell(`A${currentRow}`);
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    }
    currentRow++;

    const monthOrderShort = [
      'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN',
      'JUL', 'AGO', 'SET', 'OCT', 'NOV', 'DIC',
    ];
    const monthOrderFull = [
      'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
    ];
    
    let currentYearGastosTotals = {
      bienesCorrientes: 0,
      bienesCapital: 0,
      servicios: 0,
      subvencion: 0,
      viaticos: 0,
      encargoInterno: 0,
      totalMes: 0,
    };

    monthOrderFull.forEach((monthFullName, index) => {
      const monthShort = monthOrderShort[index];
      const gastosDelMes = gastosPorMes.get(monthFullName) || [];
      const data = this.clasificarGastosParaConsolidado(gastosDelMes); // <-- USA LA FUNCIÓN CORREGIDA

      const row = worksheet.getRow(currentRow);
      
      worksheet.getCell(`A${currentRow}`).value = `${monthShort} - ${new Date().getFullYear()}`;
      // ▼▼▼ CORRECCIÓN ERROR 3 (Uso) ▼▼▼
      worksheet.getCell(`B${currentRow}`).value = data.bienesCorrientes || null;
      worksheet.getCell(`C${currentRow}`).value = data.bienesCapital || null;
      worksheet.getCell(`D${currentRow}`).value = data.servicios || null;
      worksheet.getCell(`E${currentRow}`).value = data.subvencion || null;
      worksheet.getCell(`F${currentRow}`).value = data.viaticos || null; // <-- CAMPO AÑADIDO
      worksheet.getCell(`G${currentRow}`).value = data.encargoInterno || null; // <-- CAMPO AÑADIDO
      worksheet.getCell(`H${currentRow}`).value = data.totalMes || null; // <-- CAMPO AÑADIDO
      // ▲▲▲ FIN CORRECCIÓN ERROR 3 (Uso) ▲▲▲

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber >= 2 && colNumber <= 8) cell.numFmt = moneyFormat;
        cell.border = allBorders;
        if (colNumber === 1) cell.alignment = { horizontal: 'left', vertical: 'middle' };
      });

      // Sumar para los totales anuales
      currentYearGastosTotals.bienesCorrientes += data.bienesCorrientes;
      currentYearGastosTotals.bienesCapital += data.bienesCapital;
      currentYearGastosTotals.servicios += data.servicios;
      currentYearGastosTotals.subvencion += data.subvencion;
      currentYearGastosTotals.viaticos += data.viaticos;
      currentYearGastosTotals.encargoInterno += data.encargoInterno;
      currentYearGastosTotals.totalMes += data.totalMes;

      currentRow++;
    });

    // ... (Fila TOTALES DEL AÑO ACTUAL (2025)) ...
    
    worksheet.getCell(`A${currentRow}`).value = `AÑO ${new Date().getFullYear()}`;
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`B${currentRow}`).value = currentYearGastosTotals.bienesCorrientes || null;
    worksheet.getCell(`C${currentRow}`).value = currentYearGastosTotals.bienesCapital || null;
    worksheet.getCell(`D${currentRow}`).value = currentYearGastosTotals.servicios || null;
    worksheet.getCell(`E${currentRow}`).value = currentYearGastosTotals.subvencion || null;
    worksheet.getCell(`F${currentRow}`).value = currentYearGastosTotals.viaticos || null;
    worksheet.getCell(`G${currentRow}`).value = currentYearGastosTotals.encargoInterno || null;
    worksheet.getCell(`H${currentRow}`).value = currentYearGastosTotals.totalMes || null;
    worksheet.getRow(currentRow).font = { bold: true };
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 2 && colNumber <= 8) cell.numFmt = moneyFormat;
      cell.border = allBorders;
    });
    currentRow+=2;

    // ... (Fila TOTAL GASTOS (GLOBAL)) ...
    
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL GASTOS - 2025';
    worksheet.getCell(`A${currentRow}`).font = { bold: true};
    worksheet.getCell(`B${currentRow}`).value = previousYearsGastosTotals.bienesCorrientes + currentYearGastosTotals.bienesCorrientes;
    worksheet.getCell(`C${currentRow}`).value = previousYearsGastosTotals.bienesCapital + currentYearGastosTotals.bienesCapital;
    worksheet.getCell(`D${currentRow}`).value = previousYearsGastosTotals.servicios + currentYearGastosTotals.servicios;
    worksheet.getCell(`E${currentRow}`).value = previousYearsGastosTotals.subvencion + currentYearGastosTotals.subvencion;
    worksheet.getCell(`F${currentRow}`).value = previousYearsGastosTotals.viaticos + currentYearGastosTotals.viaticos;
    worksheet.getCell(`G${currentRow}`).value = previousYearsGastosTotals.encargoInterno + currentYearGastosTotals.encargoInterno;
    worksheet.getCell(`H${currentRow}`).value = previousYearsGastosTotals.total + currentYearGastosTotals.totalMes;
    worksheet.getRow(currentRow).font = { bold: true };
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 2 && colNumber <= 8) cell.numFmt = moneyFormat;
      (cell.numFmt = moneyFormat), (cell.border = allBorders);
      // Poner 0 si es nulo
      if (colNumber >= 2 && !cell.value) cell.value = 0;
    });
    currentRow++;

    // ... (SALDO AL AÑO) ...
    worksheet.mergeCells(`A${currentRow}:G${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `SALDO AL AÑO ${new Date().getFullYear()}`;
    worksheet.getCell(`A${currentRow}`).style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'c9c9c9' } };
    worksheet.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
    const saldo = currentTotalIngresos - (currentYearGastosTotals.totalMes);
    worksheet.getCell(`H${currentRow}`).value = saldo;
    worksheet.getCell(`H${currentRow}`).numFmt = moneyFormat;
    worksheet.getCell(`H${currentRow}`).font = { bold: true };
    for(let col = 1; col <= 8; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;
  }

  // --- CONSTRUCTOR DE HOJA DE REPORTE POR MES (Sin cambios) ---
  

  
  private crearHojaDeReportePorMes(
  workbook: ExcelJS.Workbook,
  sheetName: string, // 'ENE', 'FEB', etc.
  data: ReporteDataMensual, // ¡Recibe el objeto YA CLASIFICADO!
  projectName: string,
): void {
  const worksheet = workbook.addWorksheet(sheetName);

  // --- Definición de Estilos ---
  const headerStyle: Partial<ExcelJS.Style> = {
    font: { name: 'Arial', size: 10, bold: true, color: { argb: 'FF000000' } },
    alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }, // Verde claro
    border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
  };
  const sectionTitleStyle: Partial<ExcelJS.Style> = {
    font: { bold: true, color: { argb: 'FF000000'} },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }, // Azul claro
    alignment: { horizontal: 'left', vertical: 'middle' },
    border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
  };
  const totalRowStyle: Partial<ExcelJS.Style> = {
    font: { bold: true },
    alignment: { horizontal: 'right' },
    border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
  };
  const totalAmountStyle: Partial<ExcelJS.Style> = {
    ...totalRowStyle,
    numFmt: '"S/" #,##0.00;[Red]-"S/" #,##0.00',
    alignment: { horizontal: 'left' }, // Para que el S/ quede pegado al número
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }, // Verde claro
  };
  const grandTotalAmountStyle: Partial<ExcelJS.Style> = {
    ...totalAmountStyle,
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }, // Amarillo
    font: { bold: true, size: 12 },
  };
  const grandTotalLabelStyle: Partial<ExcelJS.Style> = {
    ...totalRowStyle,
    font: { bold: true, size: 12 },
  };
  const cellStyle: Partial<ExcelJS.Style> = {
    alignment: { vertical: 'middle' },
    border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
  };
  const wrapTextStyle: Partial<ExcelJS.Style> = { ...cellStyle, alignment: { ...cellStyle.alignment, wrapText: true } };
  const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';
  const dateFormat = 'dd/mm/yyyy';

  // --- Cabeceras de Columnas (las 14 columnas de tu ejemplo) ---
  const headers = [
    'DOC','N°','SIAF','A NOMBRE DE','CONCEPTO','MONTO','ESPECIFICA',
    'MONTO2','ESPECIFICA2','F.F.','MES','F. DEVENGADO','PROYECTO','META',
    'CERT. VIATICO', 'DESTINO', 'F. SALIDA', 'F. RETORNO'
  ];
  // Mapeo de claves de objeto Gasto al orden de las cabeceras
  const headerKeys = [
    'tipoDocumento', 'numeroDocumento', 'siaf', 'aNombreDe', 'concepto', 'monto', 'especifica',
    'monto2', 'especifica2', 'ff', 'mes', 'fechaDevengado', 'proyecto', 'meta',
    'certificacionViatico', 'destino', 'fechaSalida', 'fechaRetorno'
  ];

  // Indices de columnas para aplicar formatos (Excel empieza en 1)
    const montoColIndex = 6; 
    const monto2ColIndex = 8;
    const fDevengadoIndex = 12;
    const fSalidaIndex = 17;  
    const fRetornoIndex = 18;

  // --- Títulos de la Hoja ---
  let currentRow = 1;
  worksheet.mergeCells(currentRow, 1, currentRow, headers.length);
  worksheet.getCell(`A${currentRow}`).value = `EJECUCIÓN DE GASTOS - ${sheetName}`;
  worksheet.getCell(`A${currentRow}`).font = { name: 'Arial', size: 14, bold: true };
  worksheet.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
  currentRow++;
  
  worksheet.mergeCells(currentRow, 1, currentRow, headers.length);
  worksheet.getCell(`A${currentRow}`).value = projectName;
  worksheet.getCell(`A${currentRow}`).font = { name: 'Arial', size: 11, italic: true };
  worksheet.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
  currentRow++;
  currentRow++; // Dejar fila en blanco


  // --- 1. Escribir las Cabeceras de Columna UNA SOLA VEZ ---
  const headerRow = worksheet.getRow(currentRow);
  headerRow.values = headers;
  headerRow.eachCell((cell) => cell.style = headerStyle);
  headerRow.height = 30;
  currentRow++;

  // --- Definir las secciones y los datos que usarán (vienen de 'data') ---
  const sections = [
    { title: 'Gastos Corrientes', data: data.gastosCorrientes },
    { title: 'Subvenciones por enseñanza', data: data.subvencionesEnsenanza },
    { title: 'Subvenciones', data: data.subvenciones },
    { title: 'Gastos de Capital', data: data.gastosCapital }
  ];

  const filasVaciasPorDefecto = 4;
  let totalGeneral = 0;
 

  // --- 2. Iterar sobre las secciones de datos ---
  for (const section of sections) {
    // Escribir el título de la sección
    worksheet.mergeCells(currentRow, 1, currentRow, headers.length);
    const titleCell = worksheet.getCell(`A${currentRow}`);
    titleCell.value = section.title;
    titleCell.style = sectionTitleStyle;
    currentRow++;

    let totalSection = 0;
    let rowsAdded = 0;
    const startDataRow = currentRow; // Fila donde empiezan los datos/filas vacías

    // Escribir los datos de la sección si existen
    if (section.data.length > 0) {
      section.data.forEach((gasto) => {
        // Mapear el objeto Gasto a un array en el orden de headerKeys
        const rowData = headerKeys.map(key => (gasto as any)[key]);
        worksheet.addRow(rowData);
        
        const monto = Number(gasto.monto) || 0;
        const monto2 = Number(gasto.monto2) || 0;
        totalSection += monto + monto2; // Sumar ambos montos al total
        rowsAdded++;
        currentRow++;
      });
    }

    // --- 3. Añadir filas vacías por defecto ---
    const rowsToAdd = Math.max(0, filasVaciasPorDefecto - rowsAdded);
    for (let i = 0; i < rowsToAdd; i++) {
      // Añadir una fila vacía con el número correcto de columnas
      worksheet.addRow(Array(headers.length).fill(null));
      currentRow++;
    }
    
    // --- Aplicar estilos a las filas de datos y vacías ---
    for (let i = startDataRow; i < currentRow; i++) {
      const row = worksheet.getRow(i);
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // Aplicar estilo de celda base
        cell.style = cellStyle;
        
        // Aplicar formatos especiales
        if (colNumber === montoColIndex || colNumber === monto2ColIndex) {
          cell.numFmt = moneyFormat;
        } else if (colNumber === fDevengadoIndex || colNumber === fSalidaIndex || colNumber === fRetornoIndex) { // F. DEVENGADO
          cell.numFmt = dateFormat;
          // Corregir la fecha si es una cadena (viene de la BD como Fecha)
          if (cell.value && typeof cell.value === 'string') {
            cell.value = new Date(cell.value);
          }
        } else if ([4, 5, 13, 16].includes(colNumber)) { // A NOMBRE DE, CONCEPTO, PROYECTO
          cell.alignment = { ...cellStyle.alignment, wrapText: false };
        } else if (colNumber === 2 || colNumber === 3) { // N° y SIAF
          cell.numFmt = '@'; // Forzar texto
        }
      });
    }

    // --- 4. Escribir la fila de Total de la sección ---
    const totalRow = worksheet.getRow(currentRow);
    // Combinar celdas ANTES de la columna de monto
    worksheet.mergeCells(currentRow, 1, currentRow, montoColIndex - 1); // Merge A:E
    totalRow.getCell(1).value = `Total ${section.title}`;
    
    // Aplicar estilo de borde y alineación a las celdas mergeadas
    for(let i = 1; i < montoColIndex; i++) {
      totalRow.getCell(i).style = totalRowStyle;
    }

    // Poner el total en la columna 'MONTO'
    totalRow.getCell(montoColIndex).value = totalSection || null; // Columna F (MONTO)
    totalRow.getCell(montoColIndex).style = totalAmountStyle;
    
    // Aplicar bordes al resto de la fila de total (para que no quede vacía)
    for (let i = montoColIndex + 1; i <= headers.length; i++) {
      totalRow.getCell(i).style = cellStyle;
    }

    currentRow++;
    currentRow++; // Dejar una fila en blanco entre secciones
    totalGeneral += totalSection;
  }

  // --- 5. Escribir la fila de TOTAL GENERAL ---
  const grandTotalRow = worksheet.getRow(currentRow);
  worksheet.mergeCells(currentRow, 1, currentRow, montoColIndex - 1); // Merge A:E
  grandTotalRow.getCell(1).value = `TOTAL GENERAL`;
  
  // Aplicar estilo a celdas mergeadas
  for(let i = 1; i < montoColIndex; i++) {
    grandTotalRow.getCell(i).style = grandTotalLabelStyle;
  }

  grandTotalRow.getCell(montoColIndex).value = totalGeneral || null; // Columna F (MONTO)
  grandTotalRow.getCell(montoColIndex).style = grandTotalAmountStyle;
  
  // Aplicar bordes al resto de la fila
  for (let i = montoColIndex + 1; i <= headers.length; i++) {
    grandTotalRow.getCell(i).style = cellStyle;
  }

    // --- Ajustar Anchos de Columna (al final) ---
    worksheet.columns = [
      { width: 10 }, // DOC
      { width: 12 }, // N°
      { width: 12 }, // SIAF
      { width: 35 }, // A NOMBRE DE
      { width: 35 }, // CONCEPTO
      { width: 15 }, // MONTO
      { width: 15 }, // ESPECIFICA
      { width: 15 }, // MONTO2
      { width: 15 }, // ESPECIFICA2
      { width: 10 }, // F.F.
      { width: 10 }, // MES
      { width: 15 }, // F. DEVENGADO
      { width: 30 }, // PROYECTO
      { width: 10 },  // META
      { width: 15 }, // CERT. VIATICO (Nuevo)
      { width: 20 }, // DESTINO (Nuevo)
      { width: 14 }, // F. SALIDA (Nuevo)
      { width: 14 }  // F. RETORNO (Nuevo)
    ];
}


// =============================================================
  // --- SECCIÓN 2: NUEVA LÓGICA DE REPORTES "PIC" ---
  // =============================================================

  // ▼▼▼ REEMPLAZA ESTA FUNCIÓN COMPLETA ▼▼▼
  async generatePicReportGroup(
    modalityName: string, 
  ): Promise<ExcelJS.Workbook> {
    
    // 1. Obtener la lista de proyectos de la modalidad (sin cambios)
    const projectList = mapaModalidadesPIC.get(modalityName);
    if (!projectList || projectList.length === 0) {
      throw new Error(
        `No se pudo encontrar la lista de proyectos para la modalidad: ${modalityName}`,
      );
    }

    // 2. Obtener TODOS los gastos de la base de datos para esta lista (sin cambios)
    const allGastosInGroup = await this.gastoRepository.find({
      where: projectList.map(proyecto => ({ proyecto })), 
      order: { proyecto: 'ASC', fechaDevengado: 'ASC' },
    });

    // 3. Agrupar los gastos encontrados (sin cambios)
    const gastosPorProyecto = new Map<string, Gasto[]>();
    for (const gasto of allGastosInGroup) {
      const projectName = gasto.proyecto;
      let lista = gastosPorProyecto.get(projectName);
      if (!lista) {
        lista = [];
        gastosPorProyecto.set(projectName, lista);
      }
      lista.push(gasto);
    }

    // --- NUEVO PASO 4: Obtener toda la Metadata guardada ---
    const allMetadata = await this.picMetadataRepository.find({
      where: { projectName: In(projectList) },
    });
    
    // Convertir el array de metadata en un Mapa para búsqueda rápida
    const metadataMap = new Map<string, PicMetadata>();
    for (const metadata of allMetadata) {
      metadataMap.set(metadata.projectName, metadata);
    }
    // --- FIN DEL NUEVO PASO ---

    // 5. Crear el libro de Excel
    const workbook = new ExcelJS.Workbook();

    // 6. Crear la Hoja "CONSOLIDADO" (Formato PIC)
    this.crearHojaPicConsolidado(
      workbook,
      gastosPorProyecto,
      projectList, 
      modalityName,
      metadataMap, // <-- ¡Ahora pasamos la metadata!
    );

    // 7. Crear las Hojas de Detalle (sin cambios)
    for (const projectName of projectList) {
      const projectGastos = gastosPorProyecto.get(projectName) || [];
      
      const shortNameMatch = projectName.match(/PIC\s*\d+-\d{4}/i);
      let sheetName = shortNameMatch ? shortNameMatch[0] : projectName.substring(0, 10);
      sheetName = sheetName.replace(/[\/\?\*\[\]\:]/g, '-').substring(0, 31);

      this.crearHojaPicDetalle(
        workbook,
        sheetName,
        projectName, 
        projectGastos,
      );
    }

    return workbook;
  }
// ▲▲▲ FIN DE LA FUNCIÓN REEMPLAZADA ▲▲▲
  /**
   * NUEVO HELPER (Solo para PIC)
   * Clasifica un gasto según las categorías del reporte PIC.
   */
  private clasificarGastoPic(especifica: string): PicCategory | null {
    if (!especifica) return null;
    
    // Normalizar la 'especifica' para eliminar espacios duplicados
    const spec = especifica.trim().replace(/\s+/g, ' ');

    // El orden es importante: de más específico a más general.
    // if (spec.startsWith('2.3.2 1.1')) return 'VIATICOS';
    // if (spec.startsWith('2.3.2 1.2')) return 'CUENTA ENCARGO';
    // if (spec.startsWith('2.3.2')) return 'SERVICIOS';
    if (spec.startsWith('2.3.')) return 'BIENES CORRIENTES';
    if (spec.startsWith('2.6.')) return 'BIENES CAPITAL';
    if (spec.startsWith('2.5.')) return 'SUBVENCION';

    return null;
  }

  /**
   * NUEVO HELPER (Solo para PIC)
   * Devuelve un objeto de resumen PIC vacío.
   */
  private getEmptyPicSummary(): PicMonthSummary {
    return {
      'BIENES CORRIENTES': 0,
      'BIENES CAPITAL': 0,
      SERVICIOS: 0,
      SUBVENCION: 0,
      VIATICOS: 0,
      'CUENTA ENCARGO': 0,
      TOTAL: 0,
    };
  }



  // ▼▼▼ REEMPLAZA ESTA FUNCIÓN COMPLETA ▼▼▼
  private crearHojaPicConsolidado(
    workbook: ExcelJS.Workbook,
    gastosPorProyecto: Map<string, Gasto[]>,
    projectList: string[], 
    groupIdentifier: string,
    metadataMap: Map<string, PicMetadata>, // <-- AHORA RECIBE LA METADATA
  ): void {
    const worksheet = workbook.addWorksheet('CONSOLIDADO');
    let currentRow = 1;

    // --- Estilos (sin cambios) ---
    const titleStyle: Partial<ExcelJS.Style> = { font: { bold: true, size: 12 } };
    const boldStyle: Partial<ExcelJS.Style> = { font: { bold: true } };
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';
    const allBorders: Partial<ExcelJS.Borders> = {
      top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
    };
    const headerStyle: Partial<ExcelJS.Style> = { 
      font: { bold: true }, 
      border: allBorders, 
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } 
    };
    const centerAlign: Partial<ExcelJS.Alignment> = {
      horizontal: 'center',
      vertical: 'middle',
    };
    
    // --- Título global (sin cambios) ---
    worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `EJECUCIÓN DE GASTOS DE PROYECTOS - ${groupIdentifier}`;
    worksheet.getCell(`A${currentRow}`).style = titleStyle;
    worksheet.getCell(`A${currentRow}`).alignment = centerAlign;
    currentRow += 2; // Espacio

    const monthNames = [
      'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
    ];
    
    const categories: (keyof PicMonthSummary)[] = [
      'BIENES CORRIENTES', 'BIENES CAPITAL', 'SERVICIOS',
      'SUBVENCION', 'VIATICOS', 'CUENTA ENCARGO', 'TOTAL',
    ];
    
    // Anchos de columna (sin cambios)
    worksheet.columns = [
      { width: 35 }, { width: 14 }, { width: 14 }, { width: 14 },
      { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 15 },
    ];
    worksheet.getColumn('A').alignment = { wrapText: true, vertical: 'middle' };

    // --- Inicio del Bucle por Proyecto ---
    for (const projectName of projectList) {
      // Obtenemos los gastos (o un array vacío si no hay)
      const projectGastos = gastosPorProyecto.get(projectName) || [];
      // ¡Obtenemos la metadata guardada para este proyecto!
      const metadata = metadataMap.get(projectName) || {} as PicMetadata;

      // --- Título del Proyecto ---
      worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = projectName;
      worksheet.getCell(`A${currentRow}`).style = { ...titleStyle, alignment: centerAlign };
      currentRow++;

      // --- Metadata del Proyecto (Rellenada desde la BD) ---
      worksheet.getCell(`A${currentRow}`).value = `Investigador: ${metadata.investigador || ''}`;
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      currentRow++;
      // 2. Tesista (Solo si existe en metadata)
      if (metadata.tesista) {
        worksheet.getCell(`A${currentRow}`).value = `Tesista: ${metadata.tesista}`;
        worksheet.getCell(`A${currentRow}`).font = { bold: true };
        currentRow++;
      }

      // 3. Asesor (Solo si existe en metadata)
      if (metadata.asesor) {
        worksheet.getCell(`A${currentRow}`).value = `Asesor: ${metadata.asesor}`;
        worksheet.getCell(`A${currentRow}`).font = { bold: true };
        currentRow++;
      }
      worksheet.getCell(`A${currentRow}`).value = `Duración: ${metadata.duracion || ''}`;
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      currentRow++;
      currentRow++; // Espacio

      // --- Sección INGRESOS (Rellenada desde la BD) ---
      worksheet.getCell(`A${currentRow}`).value = 'INGRESOS';
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      worksheet.getCell(`I${currentRow}`).value = 'TOTAL';
      worksheet.getCell(`I${currentRow}`).font = { bold: true };
      currentRow++;

      let totalIngresos = 0;
      if (metadata.ingresos && metadata.ingresos.length > 0) {
        metadata.ingresos.forEach(ingreso => {
          const monto = Number(ingreso.monto) || 0;
          worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
          worksheet.getCell(`A${currentRow}`).value = ingreso.descripcion || '';
          worksheet.getCell(`I${currentRow}`).value = monto || null;
          worksheet.getCell(`I${currentRow}`).numFmt = moneyFormat;
          totalIngresos += monto;
          currentRow++;
        });
      } else {
        // Dejar una fila vacía si no hay ingresos
        worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
        currentRow++;
      }
      
      // Fila de TOTAL INGRESOS
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = 'TOTAL INGRESOS';
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      worksheet.getCell(`I${currentRow}`).value = totalIngresos || null;
      worksheet.getCell(`I${currentRow}`).numFmt = moneyFormat;
      worksheet.getCell(`I${currentRow}`).font = { bold: true };
      currentRow++;
      currentRow++; // Espacio

      // --- Cabeceras de la tabla resumen (sin cambios) ---
      const headers = ['MES', 'PRESUPUESTO', 'EJECUCIÓN DE GASTOS', '', '', '', '', '', 'TOTAL'];
      const subHeaders = ['', '', ...categories.slice(0, -1), '']; 
      let headerRow = worksheet.getRow(currentRow);
      headerRow.values = headers;
      worksheet.mergeCells(currentRow, 3, currentRow, 8);
      headerRow.eachCell((cell) => cell.style = headerStyle);
      currentRow++;
      headerRow = worksheet.getRow(currentRow);
      headerRow.values = subHeaders;
      worksheet.getRow(currentRow - 1).getCell(9).value = 'TOTAL'; 
      worksheet.mergeCells(currentRow - 1, 9, currentRow, 9);
      headerRow.getCell(1).value = 'MES';
      worksheet.mergeCells(currentRow - 1, 1, currentRow, 1);
      headerRow.getCell(2).value = 'PRESUPUESTO';
      worksheet.mergeCells(currentRow - 1, 2, currentRow, 2);
      headerRow.eachCell((cell) => cell.style = headerStyle);

      // Escribir el Presupuesto Total (del modal)
      const presupuestoTotal = Number(metadata.presupuestoTotal) || 0;
      const presupuestoCell = worksheet.getCell(`B${currentRow}`);
      presupuestoCell.value = presupuestoTotal || null;
      presupuestoCell.numFmt = moneyFormat;
      currentRow++;
      
      // Totales para el "cuadro verde" (gastos actuales)
      const projectTotals: PicMonthSummary = this.getEmptyPicSummary();
      // Totales para el "cuadro rojo" (gastos previos)
      const previousGastosTotals: PicMonthSummary = this.getEmptyPicSummary();

      // --- Sección GASTOS AÑOS ANTERIORES (Rellenada desde la BD) ---
      // ¡¡¡ BUG ARREGLADO: Ahora usa un bucle forEach !!!
      if (metadata.gastosAnosAnteriores && metadata.gastosAnosAnteriores.length > 0) {
        metadata.gastosAnosAnteriores.forEach(gastoPrevio => {
          const totalGastoPrevio = 
              (gastoPrevio.bienesCorrientes || 0) + (gastoPrevio.bienesCapital || 0) +
              (gastoPrevio.servicios || 0) + (gastoPrevio.subvencion || 0) +
              (gastoPrevio.viaticos || 0) + (gastoPrevio.encargoInterno || 0);

          const row = worksheet.addRow([
            `AÑO ${gastoPrevio.year || '...'}`,
            null, // Presupuesto
            gastoPrevio.bienesCorrientes || 0,
            gastoPrevio.bienesCapital || 0,
            gastoPrevio.servicios || 0,
            gastoPrevio.subvencion || 0,
            gastoPrevio.viaticos || 0,
            gastoPrevio.encargoInterno || 0,
            totalGastoPrevio || 0
          ]);
          
          row.font = { bold: true };
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            cell.border = allBorders;
            if (colNumber > 2) {
              cell.numFmt = moneyFormat;
              if (!cell.value) cell.value = 0;
            }
          });

          // Acumular en los totales de AÑOS PREVIOS
          previousGastosTotals['BIENES CORRIENTES'] += (gastoPrevio.bienesCorrientes || 0);
          previousGastosTotals['BIENES CAPITAL'] += (gastoPrevio.bienesCapital || 0);
          previousGastosTotals.SERVICIOS += (gastoPrevio.servicios || 0);
          previousGastosTotals.SUBVENCION += (gastoPrevio.subvencion || 0);
          previousGastosTotals.VIATICOS += (gastoPrevio.viaticos || 0);
          previousGastosTotals['CUENTA ENCARGO'] += (gastoPrevio.encargoInterno || 0);
          previousGastosTotals.TOTAL += totalGastoPrevio;
          
          currentRow++;
        });
      }
      
      // Fila de GASTO TOTAL AL AÑO ANTERIOR
      const saldoAnterior = presupuestoTotal - previousGastosTotals.TOTAL;
      const gastoTotalPrevioRow = worksheet.addRow([
        `GASTO TOTAL AL ${new Date().getFullYear() - 1}`,
        null,
        ...categories.map(cat => (previousGastosTotals as any)[cat] || 0)
      ]);
      gastoTotalPrevioRow.font = { bold: true };
      gastoTotalPrevioRow.eachCell({ includeEmpty: true }, (cell, col) => {
        cell.border = allBorders;
        if (col > 2) { cell.numFmt = moneyFormat; if (!cell.value) cell.value = 0; }
      });
      currentRow++;

      // Fila de SALDO (ej. SALDO AL 2024)
      const saldoRow = worksheet.addRow([
        `SALDO AL ${new Date().getFullYear() - 1}`,
         null, null, null, null, null, null, null,
         saldoAnterior
      ]);
      saldoRow.font = { bold: true };
      saldoRow.getCell(9).numFmt = moneyFormat;
      currentRow++;


      // --- Resumen por Mes (Año Actual) ---
      for (const month of monthNames) {
        const monthShort = month.substring(0, 3);
        const gastosDelMes = projectGastos.filter(
          (g) => g.mes && g.mes.toUpperCase() === month,
        );
        const monthSummary = this.clasificarGastosParaPicConsolidado(gastosDelMes);
        
        const rowData = [
          `${monthShort} - ${new Date().getFullYear()}`,
          null, // Presupuesto
          ...categories.map(cat => monthSummary[cat] || null)
        ];
        const row = worksheet.addRow(rowData);
        
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = allBorders;
          if (colNumber > 2) {
            cell.numFmt = moneyFormat;
            if (!cell.value) cell.value = 0; 
          }
        });
        
        // Acumular totales del AÑO ACTUAL
        categories.forEach(cat => (projectTotals as any)[cat] += monthSummary[cat]);
        currentRow++;
      }
      
      const totalAnoActualRow = worksheet.addRow([
        `AÑO ${new Date().getFullYear()}`,
        null, // Presupuesto
        ...categories.map(cat => (projectTotals as any)[cat] || 0)
      ]);
      totalAnoActualRow.font = { bold: true };
      totalAnoActualRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = allBorders;
        if (colNumber > 2) {
          cell.numFmt = moneyFormat;
          if (!cell.value) cell.value = 0;
        }
      });
      currentRow++;

      // --- Fila de Total del Proyecto (GASTO TOTAL AÑOS PREVIOS + AÑO ACTUAL) ---
      const granTotal: PicMonthSummary = this.getEmptyPicSummary();
      categories.forEach(cat => {
        (granTotal as any)[cat] = (previousGastosTotals as any)[cat] + (projectTotals as any)[cat];
      });

      const totalRowData = [
        'TOTAL GASTOS',
        presupuestoTotal || null,
        ...categories.map(cat => (granTotal as any)[cat] || 0) // Usamos granTotal
      ];
      const totalRow = worksheet.addRow(totalRowData);
      totalRow.font = { bold: true };
      totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          cell.border = allBorders;
          if (colNumber > 1) { 
            cell.numFmt = moneyFormat;
            if (!cell.value) cell.value = 0;
          }
      });
      currentRow++;

      // --- Fila de SALDO AL AÑO (Año Actual) ---
      const saldoActual = saldoAnterior - projectTotals.TOTAL;
      const saldoActualRow = worksheet.addRow([
        `SALDO AL AÑO ${new Date().getFullYear()}`,
         null, null, null, null, null, null, null,
         saldoActual
      ]);
      saldoActualRow.font = { bold: true };
      saldoActualRow.getCell(9).numFmt = moneyFormat;
      currentRow++;
      
      currentRow += 2; // Espacio entre proyectos
    }
  }
// ▲▲▲ FIN DE LA FUNCIÓN REEMPLAZADA ▲▲▲

  /**
   * NUEVO HELPER (Solo para PIC)
   * Crea la hoja de DETALLE para un proyecto PIC individual.
   * Esta lógica coincide con tu descripción.
   */
  private crearHojaPicDetalle(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    fullProjectName: string,
    projectGastos: Gasto[],
  ): void {
    const worksheet = workbook.addWorksheet(sheetName);
    let currentRow = 1;

    // Estilos
    const titleStyle: Partial<ExcelJS.Style> = { font: { bold: true, size: 12 } };
    const categoryTitleStyle: Partial<ExcelJS.Style> = { font: { bold: true, size: 11 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } } }; // Azul claro
    const headerStyle: Partial<ExcelJS.Style> = { font: { bold: true }, border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }, alignment: { horizontal: 'center' } };
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';
    const textFormat = '@';
    const monthTotalStyle: Partial<ExcelJS.Font> = { bold: true };

    // --- Títulos ---
    worksheet.mergeCells(`A${currentRow}:K${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = fullProjectName;
    worksheet.getCell(`A${currentRow}`).style = titleStyle;
    currentRow += 2;
    // Opción A: Dejar metadatos específicos en blanco
    worksheet.getCell(`B${currentRow}`).value = `Investigador:`;
    currentRow += 2;

    // --- Definir Cabeceras de Detalle y Categorías ---
    const detailHeaders = [
      'DOC', 'N°', 'SIAF', 'A NOMBRE DE', 'CONCEPTO', 'MES', 
      'IMPORTE', // monto
      'ESPECIFICA', // especifica
      'IMPORTE1', // monto2
      'ESPECIFICA' // especifica2 (col 10)
    ];
    // Mapeo de claves de objeto Gasto al orden de las cabeceras
    const detailKeys = [
      'tipoDocumento', 'numeroDocumento', 'siaf', 'aNombreDe', 'concepto', 'mes',
      'monto', 'especifica', 'monto2', 'especifica2'
    ];
    // Categorías en el orden del detalle PIC
    const categories: PicCategory[] = [
      'BIENES CORRIENTES', 
      'BIENES CAPITAL', 
      'SERVICIOS', 
      'SUBVENCION',
      'VIATICOS',

    ];
    const monthNames = [
      'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
    ];
    
    // Columnas de Importe (basado en los 'detailHeaders')
    const importeColIndex = 7;
    const importe1ColIndex = 9;

    // --- Iterar por Categoría ---
    for (const category of categories) {
      worksheet.mergeCells(`A${currentRow}:J${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = category;
      worksheet.getCell(`A${currentRow}`).style = categoryTitleStyle;
      currentRow++;

      // Escribir cabeceras de detalle
      const headerRow = worksheet.getRow(currentRow);
      headerRow.values = detailHeaders;
      headerRow.eachCell(cell => cell.style = headerStyle);
      currentRow++;

      let totalCategory = 0;
      let totalCategoryMonto2 = 0;

      // --- Iterar por Mes ---
      for (const month of monthNames) {
        let totalMonth = 0;
        let totalMonthMonto2 = 0;

        const gastosDelMes = projectGastos.filter(
          (g) =>
            g.mes &&
            g.mes.toUpperCase() === month &&
            this.clasificarGastoPic(g.especifica) === category,
        );

        if (gastosDelMes.length > 0) {
          for (const gasto of gastosDelMes) {
            const rowData = detailKeys.map(key => (gasto as any)[key]);
            const row = worksheet.addRow(rowData);
            
            // Aplicar formatos
            row.getCell(importeColIndex).numFmt = moneyFormat;  // IMPORTE
            row.getCell(importe1ColIndex).numFmt = moneyFormat; // IMPORTE1
            row.getCell(2).numFmt = textFormat; // N°
            row.getCell(3).numFmt = textFormat; // SIAF
            
            const monto = (Number(gasto.monto) || 0);
            const monto2 = (Number(gasto.monto2) || 0);
            totalMonth += monto;
            totalMonthMonto2 += monto2;
            currentRow++;
          }
        }
        
        // Fila de Total del Mes (siempre se muestra)
         const totalMonthRow = worksheet.getRow(currentRow);
         worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
         totalMonthRow.getCell(1).value = `TOTAL ${month}`;
         totalMonthRow.getCell(importeColIndex).value = totalMonth || 0;
         totalMonthRow.getCell(importeColIndex).numFmt = moneyFormat;
         totalMonthRow.getCell(importe1ColIndex).value = totalMonthMonto2 || 0;
         totalMonthRow.getCell(importe1ColIndex).numFmt = moneyFormat;
         totalMonthRow.font = monthTotalStyle;

         // Poner 0 si es null para que se muestre
         if (totalMonth === 0) totalMonthRow.getCell(importeColIndex).value = 0;
         if (totalMonthMonto2 === 0) totalMonthRow.getCell(importe1ColIndex).value = 0;


         totalMonthRow.alignment = { horizontal: 'center' };
         currentRow++;
         
         totalCategory += totalMonth;
         totalCategoryMonto2 += totalMonthMonto2;
      }
      
      currentRow++; // Espacio
      // Fila de Total de Categoría
      const totalCategoryRow = worksheet.getRow(currentRow);
      worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
      totalCategoryRow.getCell(1).value = `TOTAL ${category}`;
      totalCategoryRow.getCell(importeColIndex).value = totalCategory || 0;
      totalCategoryRow.getCell(importeColIndex).numFmt = moneyFormat;
      totalCategoryRow.getCell(importe1ColIndex).value = totalCategoryMonto2 || 0;
      totalCategoryRow.getCell(importe1ColIndex).numFmt = moneyFormat;
      totalCategoryRow.font = { bold: true, size: 11 };
      totalCategoryRow.alignment = { horizontal: 'center' };
      currentRow += 2; // Espacio
    }
    
    // Ajustar anchos de columna
    worksheet.columns = [
      { width: 10 }, // DOC
      { width: 10 }, // N°
      { width: 10 }, // SIAF
      { width: 30 }, // A NOMBRE DE
      { width: 45 }, // CONCEPTO
      { width: 10 }, // MES
      { width: 14 }, // IMPORTE
      { width: 14 }, // ESPECIFICA
      { width: 14 }, // IMPORTE1
      { width: 14 }, // ESPECIFICA2 (col 10)
    ];
  }





  // ▼▼▼ AÑADIR ESTOS DOS NUEVOS MÉTODOS ▼▼▼

  // 1. MÉTODO PRINCIPAL PARA EL REPORTE POR META
  async generateReportByMeta(): Promise<ExcelJS.Workbook> {
    // 1. Obtener TODOS los gastos
    const allGastos = await this.gastoRepository.find({
      order: { fechaDevengado: 'ASC' },
    });

    // 2. Agruparlos por 'meta'
    const gastosPorMeta = new Map<string, Gasto[]>();
    for (const gasto of allGastos) {
      const meta = gasto.meta?.trim() || 'SIN_META';
      if (!gastosPorMeta.has(meta)) {
        gastosPorMeta.set(meta, []);
      }
      gastosPorMeta.get(meta)!.push(gasto);
    }

    // 3. Crear el libro de Excel
    const workbook = new ExcelJS.Workbook();

    // 4. Crear una hoja por cada meta
    for (const [meta, gastos] of gastosPorMeta.entries()) {
      // Sanitizar el nombre de la hoja (máx 31 chars, sin caracteres inválidos)
      const safeSheetName = meta.replace(/[\\/*?[\]:]/g, '_').substring(0, 31);
      this.crearHojaMeta(workbook, safeSheetName, gastos);
    }

    return workbook;
  }

  // 2. FUNCIÓN AUXILIAR PARA CREAR UNA HOJA DE "META"
  /**
   * Crea una hoja de cálculo con un volcado de datos de gastos.
   */
  private crearHojaMeta(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    gastos: Gasto[],
  ): void {
    const worksheet = workbook.addWorksheet(sheetName);

    // Definir Estilos
    const headerStyle: Partial<ExcelJS.Style> = {
      font: { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } },
      alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    const cellStyle: Partial<ExcelJS.Style> = {
      alignment: { vertical: 'middle' },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    const wrapTextStyle: Partial<ExcelJS.Style> = {
      ...cellStyle,
      alignment: { ...cellStyle.alignment, wrapText: true }
    };
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';

    // Definir Cabeceras (las mismas 18 columnas que en Gasto-List)
    const headers = [
      'DOC','N°','SIAF','A NOMBRE DE','CONCEPTO','MONTO','ESPECIFICA',
      'MONTO2','ESPECIFICA2','F.F.','MES','F. DEVENGADO','PROYECTO','META',
      'CERTIFICACION VIATICO','DESTINO','F. SALIDA','F. RETORNO'
    ];
    
    const headerRow = worksheet.getRow(1);
    headerRow.values = headers;
    headerRow.eachCell(cell => cell.style = headerStyle);
    headerRow.height = 30;

    // Añadir Filas de Datos
    let currentRow = 2;
    gastos.forEach(gasto => {
      worksheet.addRow([
        gasto.tipoDocumento, gasto.numeroDocumento, gasto.siaf,
        gasto.aNombreDe, gasto.concepto, gasto.monto,
        gasto.especifica, gasto.monto2, gasto.especifica2,
        gasto.ff, gasto.mes, gasto.fechaDevengado,
        gasto.proyecto, gasto.meta,
        gasto.certificacionViatico, gasto.destino,
        gasto.fechaSalida, gasto.fechaRetorno
      ]);
      const addedRow = worksheet.getRow(currentRow);
      
      // Aplicar formatos
      addedRow.getCell(6).numFmt = moneyFormat; // Monto
      if(gasto.monto2) {
        addedRow.getCell(8).numFmt = moneyFormat; // Monto2
      }
      if (gasto.fechaDevengado) {
        addedRow.getCell(12).value = new Date(gasto.fechaDevengado);
        addedRow.getCell(12).numFmt = 'dd/mm/yyyy';
      }
      if (gasto.fechaSalida) {
        addedRow.getCell(17).value = new Date(gasto.fechaSalida);
        addedRow.getCell(17).numFmt = 'dd/mm/yyyy';
      }
      if (gasto.fechaRetorno) {
        addedRow.getCell(18).value = new Date(gasto.fechaRetorno);
        addedRow.getCell(18).numFmt = 'dd/mm/yyyy';
      }

      // Aplicar estilos de celda
      addedRow.eachCell((cell, colNumber) => {
        if (colNumber === 4 || colNumber === 5 || colNumber === 13) {
          cell.style = wrapTextStyle; // 'A NOMBRE DE', 'CONCEPTO', 'PROYECTO'
        } else {
          cell.style = cellStyle;
        }
      });
      
      currentRow++;
    });

    // Definir Anchos de Columna
    const columnWidths = [
      { key: 'DOC', width: 10 },
      { key: 'N°', width: 12 },
      { key: 'SIAF', width: 12 },
      { key: 'A NOMBRE DE', width: 35 },
      { key: 'CONCEPTO', width: 50 },
      { key: 'MONTO', width: 15 },
      { key: 'ESPECIFICA', width: 15 },
      { key: 'MONTO2', width: 15 },
      { key: 'ESPECIFICA2', width: 15 },
      { key: 'F.F.', width: 10 },
      { key: 'MES', width: 10 },
      { key: 'F. DEVENGADO', width: 15 },
      { key: 'PROYECTO', width: 30 },
      { key: 'META', width: 10 },
      { key: 'CERT. VIATICO', width: 15 },
      { key: 'DESTINO', width: 20 },
      { key: 'F. SALIDA', width: 15 },
      { key: 'F. RETORNO', width: 15 }
    ];

    columnWidths.forEach((col, index) => {
      const colNum = index + 1;
      const column = worksheet.getColumn(colNum);
      column.width = col.width;
      if (colNum === 2 || colNum === 3) {
        column.numFmt = '@'; // Forzar texto para N° y SIAF
      }
    });
  }
  // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

  async update(id: number, gastoData: Partial<Gasto>): Promise<Gasto> {
    // Primero busca el gasto para asegurarse de que existe
    const gasto = await this.gastoRepository.findOneBy({ id });
    if (!gasto) {
      throw new Error(`Gasto con ID ${id} no encontrado.`);
    }
    // Asigna los nuevos datos al gasto encontrado
    Object.assign(gasto, gastoData);
    
    // Guarda y devuelve el gasto actualizado
    return this.gastoRepository.save(gasto);
  }
  
  /**
   * Elimina un gasto por su ID.
   */
  async delete(id: number): Promise<void> {
    const result = await this.gastoRepository.delete(id);
    if (result.affected === 0) {
      throw new Error(`Gasto con ID ${id} no encontrado.`);
    }
  }

  /**
   * Elimina TODOS los gastos de la base de datos.
   */
  async deleteAll(): Promise<void> {
    // Esto borra todos los registros de la tabla 'Gasto'
    await this.gastoRepository.clear();
  }

}