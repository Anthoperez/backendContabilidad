// src/app.service.ts
import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { Gasto } from './gasto.entity';
import * as ExcelJS from 'exceljs';

// ▼▼▼ INTERFAZ ReportMetadata AMPLIADA (Sin cambios desde la última vez) ▼▼▼
interface ReportMetadata {
  investigador?: string;
  rr_investigador?: string;
  fechaInicio?: string | Date;
  fechaCulminacion?: Date;
  duracion?: string;
  presupuestoProcienciaAporteMonetario?: number | null;
  presupuestoProcienciaAporteNoMonetario?: number | null;
  presupuestoEntidadEjecutoraAporteMonetario?: number | null;
  presupuestoEntidadEjecutoraAporteNoMonetario?: number | null;
  presupuestoEntidadAsociadaAporteMonetario?: number | null;
  presupuestoEntidadAsociadaAporteNoMonetario?: number | null;
  ingresos?: { descripcion: string; monto: number | null }[];
  gastosAnoAnterior?: {
    year: number;
    bienesCorrientes: number | null;
    bienesCapital: number | null;
    servicios: number | null;
    subvencion: number | null;
    viaticos: number | null;
    encargoInterno: number | null;
  };
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

@Injectable()
export class AppService {
  constructor(
    @InjectRepository(Gasto)
    private gastoRepository: Repository<Gasto>,
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

  
  // --- MÉTODO PRINCIPAL DE GENERACIÓN DE REPORTES (Sin cambios) ---
  async generateProjectReport(
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
      const especifica = gasto.especifica?.trim() || '';

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
      viaticos: 0, // <-- Inicializado
      encargoInterno: 0, // <-- Inicializado
      totalMes: 0, // <-- Inicializado
    };

    for (const gasto of gastos) {
      // Normalizar espacios (ej. "2.3. 2 1. 1 2" -> "2.3.2 1.1 2")
      const especifica = gasto.especifica?.trim().replace(/\s+/g, ' ') || '';
      const monto = Number(gasto.monto) || 0;

      if (especifica.startsWith('2.3.1')) {
        reporte.bienesCorrientes += monto;
      } else if (especifica.startsWith('2.6.')) {
        reporte.bienesCapital += monto;
      } else if (especifica.startsWith('2.5.')) {
        reporte.subvencion += monto;
      } //else if (especifica.startsWith('2.3.2 1.1')) {
      //   // Lógica específica para Viáticos (basada en tus CSV)
      //   reporte.viaticos += monto;
      // } else if (especifica.startsWith('2.3.2')) {
      //   // Todo lo demás que sea 2.3.2 va a Servicios
      //   reporte.servicios += monto;
      // }
      // Nota: Aún no tengo la lógica para 'encargoInterno', así que se quedará en 0.
      // Si "Encargo Interno" tiene un código (ej. '2.3.2 1.2'), añádelo aquí.
    }

    // Calculamos el total
    reporte.totalMes =
      reporte.bienesCorrientes +
      reporte.bienesCapital +
      reporte.servicios +
      reporte.subvencion +
      reporte.viaticos +
      reporte.encargoInterno;

    return reporte;
  }
  // ▲▲▲ FIN CORRECCIÓN ERROR 3 (Implementación) ▲▲▲

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
      { width: 20 }, // C
      { width: 15 }, // D
      { width: 15 }, // E
      { width: 15 }, // F
      { width: 15 }, // G
      { width: 20 }, // H
      { width: 15 }, // I: TOTAL
    ];
    // Corrección para que la descripción de ingresos sea visible
    worksheet.getColumn('A').alignment = { wrapText: true, vertical: 'middle' };


    let currentRow = 1;

    // ... (Título principal) ...
    worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS DEL PROYECTO DE INVESTIGACIÓN - PROCIENCIA';
    worksheet.getCell(`A${currentRow}`).style = titleStyle;
    currentRow++;
    currentRow++;

    // ... (Título del proyecto) ...
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = projectName;
    worksheet.getCell(`A${currentRow}`).font = { bold: true, size: 11 };
    currentRow++;

    // ... (Fila Investigador y Fecha Inicio) ...
    worksheet.getCell(`A${currentRow}`).value = 'INVESTIGADOR:';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`C${currentRow}:E${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = metadata.investigador || '';
    worksheet.getCell(`C${currentRow}`).border = allBorders;
    worksheet.getCell(`F${currentRow}`).value = 'Fecha de inicio:';
    worksheet.getCell(`F${currentRow}`).font = { bold: true };
    
    const fechaInicioCell = worksheet.getCell(`G${currentRow}`);
    if (metadata.fechaInicio) {
      fechaInicioCell.value = new Date(metadata.fechaInicio);
      fechaInicioCell.numFmt = 'dd/mm/yyyy';
    } else {
      fechaInicioCell.value = '';
    }
    fechaInicioCell.border = allBorders;

    worksheet.getCell(`H${currentRow}`).value = 'Fecha de Culminacion:';
    worksheet.getCell(`H${currentRow}`).font = { bold: true };
    
    const fechaCulminacionCell = worksheet.getCell(`I${currentRow}`);
    if (metadata.fechaCulminacion) {
      fechaCulminacionCell.value = new Date(metadata.fechaCulminacion);
      fechaCulminacionCell.numFmt = 'dd/mm/yyyy';
    } else {
      fechaCulminacionCell.value = '';
    }
    fechaCulminacionCell.border = allBorders;
    currentRow++;

    // ... (Fila RR y Duración) ...
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = metadata.rr_investigador || '';
    worksheet.getCell(`A${currentRow}`).border = allBorders;
    worksheet.getCell(`F${currentRow}`).value = 'Duración:';
    worksheet.getCell(`F${currentRow}`).font = { bold: true };
    
    worksheet.getCell(`G${currentRow}`).value = metadata.duracion || '';
    worksheet.getCell(`G${currentRow}`).border = allBorders;
    currentRow++;
    
    currentRow++; // Salto de línea

    // ... (SECCIÓN DESCRIPCIÓN DE PRESUPUESTO - Cabeceras)
    // Colocamos la DESCRIPCION DE PRESUPUESTO en la misma fila que los headers de aporte
    worksheet.getCell(`A${currentRow}`).value = 'DESCRIPCION DE PRESUPUESTO';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    // Reservar A:B para la descripción
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = 'Aporte Monetaria o no monetaria (Valorizado S/)';
    worksheet.mergeCells(`C${currentRow}:D${currentRow}`);
    worksheet.getCell(`C${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    worksheet.getCell(`C${currentRow}`).alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    worksheet.getCell(`E${currentRow}`).value = 'Aporte Monetario S/';
    worksheet.getCell(`E${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    worksheet.getCell(`E${currentRow}`).alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    worksheet.getCell(`F${currentRow}`).value = 'Aporte Total S/';
    worksheet.mergeCells(`F${currentRow}:G${currentRow}`);
    worksheet.getCell(`F${currentRow}`).style = { font: { bold: true }, alignment: centerAlign, border: allBorders };
    // Aplicar bordes a toda la fila de cabecera de presupuesto
    for (let col = 1; col <= 7; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;

    // ... (Fila PROCIENCIA) ...
    worksheet.getCell(`A${currentRow}`).value = 'PROCIENCIA';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = metadata.presupuestoProcienciaAporteNoMonetario || null;
    worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`C${currentRow}:D${currentRow}`);
    worksheet.getCell(`E${currentRow}`).value = metadata.presupuestoProcienciaAporteMonetario || null;
    worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
    const procienciaTotal = (metadata.presupuestoProcienciaAporteNoMonetario || 0) + (metadata.presupuestoProcienciaAporteMonetario || 0);
    worksheet.getCell(`F${currentRow}`).value = procienciaTotal || null;
    worksheet.getCell(`F${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`F${currentRow}:G${currentRow}`);
    for(let col = 1; col <= 7; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;

    // ... (Fila Entidad Ejecutora) ...
    worksheet.getCell(`A${currentRow}`).value = 'Entidad Ejecutora';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = metadata.presupuestoEntidadEjecutoraAporteNoMonetario || null;
    worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`C${currentRow}:D${currentRow}`);
    worksheet.getCell(`E${currentRow}`).value = metadata.presupuestoEntidadEjecutoraAporteMonetario || null;
    worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
    const ejecutoraTotal = (metadata.presupuestoEntidadEjecutoraAporteNoMonetario || 0) + (metadata.presupuestoEntidadEjecutoraAporteMonetario || 0);
    worksheet.getCell(`F${currentRow}`).value = ejecutoraTotal || null;
    worksheet.getCell(`F${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`F${currentRow}:G${currentRow}`);
    for(let col = 1; col <= 7; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;

    // ... (Fila Entidad Asociada) ...
    worksheet.getCell(`A${currentRow}`).value = 'Entidad Asociada';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`C${currentRow}`).value = metadata.presupuestoEntidadAsociadaAporteNoMonetario || null;
    worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`C${currentRow}:D${currentRow}`);
    worksheet.getCell(`E${currentRow}`).value = metadata.presupuestoEntidadAsociadaAporteMonetario || null;
    worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
    const asociadaTotal = (metadata.presupuestoEntidadAsociadaAporteNoMonetario || 0) + (metadata.presupuestoEntidadAsociadaAporteMonetario || 0);
    worksheet.getCell(`F${currentRow}`).value = asociadaTotal || null;
    worksheet.getCell(`F${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`F${currentRow}:G${currentRow}`);
    for(let col = 1; col <= 7; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;

    // ... (Fila TOTAL PRESUPUESTO) ...
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL PRESUPUESTO';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    const totalAporteNoMonetario = (metadata.presupuestoProcienciaAporteNoMonetario || 0) + (metadata.presupuestoEntidadEjecutoraAporteNoMonetario || 0) + (metadata.presupuestoEntidadAsociadaAporteNoMonetario || 0);
    worksheet.getCell(`C${currentRow}`).value = totalAporteNoMonetario || null;
    worksheet.getCell(`C${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`C${currentRow}:D${currentRow}`);
    const totalAporteMonetario = (metadata.presupuestoProcienciaAporteMonetario || 0) + (metadata.presupuestoEntidadEjecutoraAporteMonetario || 0) + (metadata.presupuestoEntidadAsociadaAporteMonetario || 0);
    worksheet.getCell(`E${currentRow}`).value = totalAporteMonetario || null;
    worksheet.getCell(`E${currentRow}`).numFmt = moneyFormat;
    worksheet.getCell(`F${currentRow}`).value = (procienciaTotal + ejecutoraTotal + asociadaTotal) || null;
    worksheet.getCell(`F${currentRow}`).numFmt = moneyFormat;
    worksheet.mergeCells(`F${currentRow}:G${currentRow}`);
    for(let col = 1; col <= 7; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;
    currentRow++;

    // ... (SECCIÓN INGRESOS) ...

    worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
    const ingresosHeaderCell = worksheet.getCell(`A${currentRow}`);
    ingresosHeaderCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    ingresosHeaderCell.value = 'INGRESOS';
    ingresosHeaderCell.font = { bold: true };
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



          worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
          worksheet.getCell(`A${currentRow}`).value = descripcion;
          worksheet.getCell(`I${currentRow}`).value = ingreso.monto || null;
          worksheet.getCell(`I${currentRow}`).numFmt = moneyFormat;
          for(let col = 1; col <= 9; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
          
          currentRow++;
        }
      });
    } else {
      worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = '';
      worksheet.getCell(`I${currentRow}`).value = null;
      for(let col = 1; col <= 9; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
      currentRow++;
    }
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL INGRESOS ${currentYearString}';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`I${currentRow}`).value = currentTotalIngresos || null;
    worksheet.getCell(`I${currentRow}`).numFmt = moneyFormat;
    worksheet.getCell(`I${currentRow}`).font = { bold: true };
    for(let col = 1; col <= 9; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    currentRow++;
    currentRow++;

    // ... (PRIMERA SECCIÓN EJECUCIÓN DE GASTOS - Cabeceras) ...
    worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS';
    worksheet.getCell(`A${currentRow}`).style = boldStyle;
    currentRow++;
    const headers = ['MES/AÑO', '', 'BIENES CORRIENTES', 'BIENES CAPITAL', 'SERVICIOS', 'SUBVENCION', 'VIATICOS', 'ENCARGO INTERNO', 'TOTAL'];
    worksheet.getCell(`A${currentRow}`).value = headers[0];
    worksheet.getCell(`C${currentRow}`).value = headers[2];
    worksheet.getCell(`D${currentRow}`).value = headers[3];
    worksheet.getCell(`E${currentRow}`).value = headers[4];
    worksheet.getCell(`F${currentRow}`).value = headers[5];
    worksheet.getCell(`G${currentRow}`).value = headers[6];
    worksheet.getCell(`H${currentRow}`).value = headers[7];
    worksheet.getCell(`I${currentRow}`).value = headers[8];
    worksheet.getRow(currentRow).font = { bold: true };
    for(let col = 1; col <= 9; col++) {
      if(col === 1) worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).alignment = centerAlign;
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
    }
    currentRow++;

    // ▼▼▼ CORRECCIÓN ERROR 2 ▼▼▼
    // Proporcionamos un objeto default completo si metadata.gastosAnoAnterior es nulo
    const gastosAnoAnterior = metadata.gastosAnoAnterior || {
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

    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `AÑO ${yearAnterior}`;
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`C${currentRow}`).value = gastosAnoAnterior.bienesCorrientes || null;
    worksheet.getCell(`D${currentRow}`).value = gastosAnoAnterior.bienesCapital || null;
    worksheet.getCell(`E${currentRow}`).value = gastosAnoAnterior.servicios || null;
    worksheet.getCell(`F${currentRow}`).value = gastosAnoAnterior.subvencion || null;
    worksheet.getCell(`G${currentRow}`).value = gastosAnoAnterior.viaticos || null;
    worksheet.getCell(`H${currentRow}`).value = gastosAnoAnterior.encargoInterno || null;
    worksheet.getCell(`I${currentRow}`).value = totalAnoAnterior || null;
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 3 && colNumber <= 9) cell.numFmt = moneyFormat;
      cell.border = allBorders;
    });
    currentRow++;

    // ... (Fila TOTAL GASTOS (solo año anterior)) ...
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL GASTOS';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`C${currentRow}`).value = gastosAnoAnterior.bienesCorrientes || null;
    worksheet.getCell(`D${currentRow}`).value = gastosAnoAnterior.bienesCapital || null;
    worksheet.getCell(`E${currentRow}`).value = gastosAnoAnterior.servicios || null;
    worksheet.getCell(`F${currentRow}`).value = gastosAnoAnterior.subvencion || null;
    worksheet.getCell(`G${currentRow}`).value = gastosAnoAnterior.viaticos || null;
    worksheet.getCell(`H${currentRow}`).value = gastosAnoAnterior.encargoInterno || null;
    worksheet.getCell(`I${currentRow}`).value = totalAnoAnterior || null;
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 3 && colNumber <= 9) cell.numFmt = moneyFormat;
      cell.border = allBorders;
    });
    currentRow++;
    currentRow++;

    // ... (SEGUNDA SECCIÓN EJECUCIÓN DE GASTOS (MENSUAL AÑO ACTUAL) - Cabeceras) ...
    worksheet.mergeCells(`A${currentRow}:I${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'EJECUCIÓN DE GASTOS';
    worksheet.getCell(`A${currentRow}`).style = boldStyle;
    currentRow++;
    worksheet.getCell(`A${currentRow}`).value = headers[0];
    worksheet.getCell(`C${currentRow}`).value = headers[2];
    worksheet.getCell(`D${currentRow}`).value = headers[3];
    worksheet.getCell(`E${currentRow}`).value = headers[4];
    worksheet.getCell(`F${currentRow}`).value = headers[5];
    worksheet.getCell(`G${currentRow}`).value = headers[6];
    worksheet.getCell(`H${currentRow}`).value = headers[7];
    worksheet.getCell(`I${currentRow}`).value = headers[8];
    worksheet.getRow(currentRow).font = { bold: true };
    for(let col = 1; col <= 9; col++) {
      if(col === 1) worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
      worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).alignment = centerAlign;
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
      worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = `${monthShort} - ${new Date().getFullYear()}`;
      // ▼▼▼ CORRECCIÓN ERROR 3 (Uso) ▼▼▼
      worksheet.getCell(`C${currentRow}`).value = data.bienesCorrientes || null;
      worksheet.getCell(`D${currentRow}`).value = data.bienesCapital || null;
      worksheet.getCell(`E${currentRow}`).value = data.servicios || null;
      worksheet.getCell(`F${currentRow}`).value = data.subvencion || null;
      worksheet.getCell(`G${currentRow}`).value = data.viaticos || null; // <-- CAMPO AÑADIDO
      worksheet.getCell(`H${currentRow}`).value = data.encargoInterno || null; // <-- CAMPO AÑADIDO
      worksheet.getCell(`I${currentRow}`).value = data.totalMes || null; // <-- CAMPO AÑADIDO
      // ▲▲▲ FIN CORRECCIÓN ERROR 3 (Uso) ▲▲▲

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber >= 3 && colNumber <= 9) cell.numFmt = moneyFormat;
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
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `AÑO ${new Date().getFullYear()}`;
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`C${currentRow}`).value = currentYearGastosTotals.bienesCorrientes || null;
    worksheet.getCell(`D${currentRow}`).value = currentYearGastosTotals.bienesCapital || null;
    worksheet.getCell(`E${currentRow}`).value = currentYearGastosTotals.servicios || null;
    worksheet.getCell(`F${currentRow}`).value = currentYearGastosTotals.subvencion || null;
    worksheet.getCell(`G${currentRow}`).value = currentYearGastosTotals.viaticos || null;
    worksheet.getCell(`H${currentRow}`).value = currentYearGastosTotals.encargoInterno || null;
    worksheet.getCell(`I${currentRow}`).value = currentYearGastosTotals.totalMes || null;
    worksheet.getRow(currentRow).font = { bold: true };
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 3 && colNumber <= 9) cell.numFmt = moneyFormat;
      cell.border = allBorders;
    });
    currentRow++;

    // ... (Fila TOTAL GASTOS (GLOBAL)) ...
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = 'TOTAL GASTOS';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    worksheet.getCell(`C${currentRow}`).value = (gastosAnoAnterior.bienesCorrientes || 0) + currentYearGastosTotals.bienesCorrientes;
    worksheet.getCell(`D${currentRow}`).value = (gastosAnoAnterior.bienesCapital || 0) + currentYearGastosTotals.bienesCapital;
    worksheet.getCell(`E${currentRow}`).value = (gastosAnoAnterior.servicios || 0) + currentYearGastosTotals.servicios;
    worksheet.getCell(`F${currentRow}`).value = (gastosAnoAnterior.subvencion || 0) + currentYearGastosTotals.subvencion;
    worksheet.getCell(`G${currentRow}`).value = (gastosAnoAnterior.viaticos || 0) + currentYearGastosTotals.viaticos;
    worksheet.getCell(`H${currentRow}`).value = (gastosAnoAnterior.encargoInterno || 0) + currentYearGastosTotals.encargoInterno;
    worksheet.getCell(`I${currentRow}`).value = totalAnoAnterior + currentYearGastosTotals.totalMes;
    worksheet.getRow(currentRow).font = { bold: true };
    worksheet.getRow(currentRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber >= 3 && colNumber <= 9) cell.numFmt = moneyFormat;
      (cell.numFmt = moneyFormat), (cell.border = allBorders);
      // Poner 0 si es nulo
      if (colNumber >= 3 && !cell.value) cell.value = 0;
    });
    currentRow++;
    currentRow++;

    // ... (SALDO AL AÑO) ...
    worksheet.mergeCells(`A${currentRow}:H${currentRow}`);
    worksheet.getCell(`A${currentRow}`).value = `SALDO AL AÑO ${new Date().getFullYear()}`;
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    const saldo = currentTotalIngresos - (currentYearGastosTotals.totalMes);
    worksheet.getCell(`I${currentRow}`).value = saldo;
    worksheet.getCell(`I${currentRow}`).numFmt = moneyFormat;
    worksheet.getCell(`I${currentRow}`).font = { bold: true };
    for(let col = 1; col <= 9; col++) worksheet.getCell(`${String.fromCharCode(64 + col)}${currentRow}`).border = allBorders;
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
    'MONTO2','ESPECIFICA2','F.F.','MES','F. DEVENGADO','PROYECTO','META'
  ];
  // Mapeo de claves de objeto Gasto al orden de las cabeceras
  const headerKeys = [
    'tipoDocumento', 'numeroDocumento', 'siaf', 'aNombreDe', 'concepto', 'monto', 'especifica',
    'monto2', 'especifica2', 'ff', 'mes', 'fechaDevengado', 'proyecto', 'meta'
  ];

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
  const montoColIndex = 6; // 'MONTO' es la 6ta columna
  const monto2ColIndex = 8; // 'MONTO2' es la 8va columna

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
        } else if (colNumber === 12) { // F. DEVENGADO
          cell.numFmt = dateFormat;
          // Corregir la fecha si es una cadena (viene de la BD como Fecha)
          if (cell.value && typeof cell.value === 'string') {
            cell.value = new Date(cell.value);
          }
        } else if (colNumber === 4 || colNumber === 5 || colNumber === 13) { // A NOMBRE DE, CONCEPTO, PROYECTO
          cell.style = wrapTextStyle;
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
    { width: 50 }, // CONCEPTO
    { width: 15 }, // MONTO
    { width: 15 }, // ESPECIFICA
    { width: 15 }, // MONTO2
    { width: 15 }, // ESPECIFICA2
    { width: 10 }, // F.F.
    { width: 10 }, // MES
    { width: 15 }, // F. DEVENGADO
    { width: 30 }, // PROYECTO
    { width: 10 }  // META
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

  
}