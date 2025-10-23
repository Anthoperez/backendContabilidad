import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { Gasto } from './gasto.entity';
import * as ExcelJS from 'exceljs';

// --- INTERFACES PARA LA ESTRUCTURA DE LOS REPORTES ---
interface ReporteDataMensual {
  gastosCorrientes: Gasto[];
  subvenciones: Gasto[];
  gastosCapital: Gasto[];
}

interface ReporteDataConsolidado {
  bienesCorrientes: number;
  bienesCapital: number;
  servicios: number;
  subvencion: number;
}

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
    const gastosLimpios = gastos.map(gasto => {
      if (gasto.proyecto && typeof gasto.proyecto === 'string') {
        gasto.proyecto = gasto.proyecto.trim();
      }
      // Aseguramos que los campos que deben ser texto, lo sean
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
  async generateProjectReport(projectName: string): Promise<ExcelJS.Workbook> {
    const gastosDelProyecto = await this.gastoRepository.find({
      where: { proyecto: projectName },
      order: { fechaDevengado: 'ASC' },
    });

    const gastosAgrupadosPorMes = this.groupGastosByMonth(gastosDelProyecto);
    const workbook = new ExcelJS.Workbook();
    
    this.crearHojaConsolidado(workbook, projectName, gastosAgrupadosPorMes);

    const monthOrderShort = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"];
    const monthOrderFull = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

    monthOrderFull.forEach((monthFullName, index) => {
      const monthShortName = monthOrderShort[index];
      const gastosDelMes = gastosAgrupadosPorMes.get(monthFullName) || [];
      const reporteDataMes = this.clasificarGastosParaMes(gastosDelMes);
      this.crearHojaDeReportePorMes(workbook, monthShortName, reporteDataMes, projectName);
    });
    
    return workbook;
  }

  // --- LÓGICA DE CLASIFICACIÓN Y AGRUPACIÓN (Sin cambios) ---
  private groupGastosByMonth(gastos: Gasto[]): Map<string, Gasto[]> {
    const gastosPorMes = new Map<string, Gasto[]>();
    const monthNames = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    for (const gasto of gastos) {
      if (gasto.fechaDevengado) {
        const date = new Date(gasto.fechaDevengado);
        const monthName = monthNames[date.getMonth()];
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
      subvenciones: [],
      gastosCapital: [],
    };
    for (const gasto of gastos) {
      const especifica = gasto.especifica?.trim() || '';
      if (especifica.startsWith('2.6.')) {
        reporte.gastosCapital.push(gasto);
      } else if (especifica.startsWith('2.5.')) {
        reporte.subvenciones.push(gasto);
      } else {
        reporte.gastosCorrientes.push(gasto);
      }
    }
    return reporte;
  }

  private clasificarGastosParaConsolidado(gastos: Gasto[]): ReporteDataConsolidado {
    const reporte: ReporteDataConsolidado = {
      bienesCorrientes: 0,
      bienesCapital: 0,
      servicios: 0,
      subvencion: 0,
    };
    for (const gasto of gastos) {
      const especifica = gasto.especifica?.trim() || '';
      const monto = Number(gasto.monto) || 0;
      if (especifica.startsWith('2.3.1')) {
        reporte.bienesCorrientes += monto;
      } else if (especifica.startsWith('2.3.2')) {
        reporte.servicios += monto;
      } else if (especifica.startsWith('2.6.')) {
        reporte.bienesCapital += monto;
      } else if (especifica.startsWith('2.5.')) {
        reporte.subvencion += monto;
      }
    }
    return reporte;
  }

  // --- CONSTRUCTORES DE HOJAS DE EXCEL ---

  private crearHojaConsolidado(workbook: ExcelJS.Workbook, projectName: string, gastosPorMes: Map<string, Gasto[]>): void {
    const worksheet = workbook.addWorksheet('CONSOLIDADO 2025');
    const titleStyle: Partial<ExcelJS.Style> = { font: { bold: true, size: 12 }, alignment: { horizontal: 'center' } };
    const boldStyle: Partial<ExcelJS.Style> = { font: { bold: true } };
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';

    // --- (INICIO) SECCIONES RESTAURADAS ---
    worksheet.mergeCells('A1:I1');
    worksheet.getCell('A1').value = 'EJECUCIÓN DE GASTOS DEL PROYECTO DE INVESTIGACIÓN - PROCIENCIA';
    worksheet.getCell('A1').style = titleStyle;
    worksheet.mergeCells('A3:E3');
    worksheet.getCell('A3').value = projectName;
    worksheet.getCell('A3').font = { bold: true, size: 11 };
    worksheet.getCell('A5').value = 'INVESTIGADOR:';
    worksheet.getCell('A5').font = { bold: true };
    worksheet.getCell('F5').value = 'Fecha de inicio:';
    worksheet.getCell('F6').value = 'Duración:';
    
    worksheet.getCell('A8').value = 'DESCRIPCION DE PRESUPUESTO';
    worksheet.getCell('A8').font = { bold: true };
    const presupuestoHeaders = ['', '', 'Aporte Monetaria o no monetaria (Valorizado S/)', 'Aporte Monetario S/', 'Aporte Total S/'];
    worksheet.getRow(9).values = presupuestoHeaders;
    worksheet.getRow(9).font = { bold: true };
    worksheet.getRow(10).values = ['PROCIENCIA', '', 0, 99999, 99999];
    worksheet.getRow(11).values = ['Entidad Ejecutora', '', 99999, 0, 99999];
    worksheet.getRow(12).values = ['TOTAL PRESUPUESTO', '', 99999, 99999, 99999];
    worksheet.getRow(12).font = { bold: true };
    for (let i = 10; i <= 12; i++) {
        for (let j = 3; j <= 5; j++) {
            worksheet.getCell(i, j).numFmt = moneyFormat;
        }
    }

    worksheet.getCell('A14').value = 'INGRESOS';
    worksheet.getCell('A14').font = { bold: true };
    const incomeRow = worksheet.getRow(15);
    incomeRow.values = ['TOTAL INGRESOS 2025', '', '', '', '', '', '', '', 99999];
    incomeRow.font = { bold: true };
    incomeRow.getCell(9).numFmt = moneyFormat;
    // --- (FIN) SECCIONES RESTAURADAS ---
    
    worksheet.getCell('A17').value = 'EJECUCIÓN DE GASTOS';
    worksheet.getCell('A17').style = boldStyle;
    let currentRow = 18;
    const ejecucionHeaders = ['MES/AÑO','','BIENES CORRIENTES','BIENES CAPITAL','SERVICIOS','SUBVENCION','VIATICOS','ENCARGO INTERNO','TOTAL'];
    worksheet.getRow(currentRow).values = ejecucionHeaders;
    worksheet.getRow(currentRow).font = { bold: true };
    currentRow++;

    const monthOrderShort = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"];
    const monthOrderFull = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    const totals = { bienesCorrientes: 0, bienesCapital: 0, servicios: 0, subvencion: 0, totalGeneral: 0 };

    monthOrderFull.forEach((monthFullName, index) => {
      const gastosDelMes = gastosPorMes.get(monthFullName) || [];
      const dataMes = this.clasificarGastosParaConsolidado(gastosDelMes);
      const totalMes = dataMes.bienesCorrientes + dataMes.bienesCapital + dataMes.servicios + dataMes.subvencion;
      
      const row = worksheet.getRow(currentRow);
      row.values = [`${monthOrderShort[index]} - 2025`, '', dataMes.bienesCorrientes || 0, dataMes.bienesCapital || 0, dataMes.servicios || 0, dataMes.subvencion || 0, 0, 0, totalMes || 0];
      for(let i = 3; i <= 9; i++) {
        row.getCell(i).numFmt = moneyFormat;
      }
      
      totals.bienesCorrientes += dataMes.bienesCorrientes;
      totals.bienesCapital += dataMes.bienesCapital;
      totals.servicios += dataMes.servicios;
      totals.subvencion += dataMes.subvencion;
      totals.totalGeneral += totalMes;
      currentRow++;
    });

    const totalRow = worksheet.getRow(currentRow);
    totalRow.values = ['TOTAL', '', totals.bienesCorrientes, totals.bienesCapital, totals.servicios, totals.subvencion, 0, 0, totals.totalGeneral];
    totalRow.font = { bold: true };
      for(let i = 3; i <= 9; i++) {
        totalRow.getCell(i).numFmt = moneyFormat;
    }

    // ⭐ --- NUEVA SECCIÓN: Anchos de columna fijos ---
    // Se definen los anchos para cada columna de la hoja CONSOLIDADO.
    worksheet.getColumn('A').width = 25; // MES/AÑO y etiquetas de presupuesto
    worksheet.getColumn('B').width = 5;  // Columna vacía de separación
    worksheet.getColumn('C').width = 22; // Aporte... / BIENES CORRIENTES
    worksheet.getColumn('D').width = 22; // Aporte... / BIENES CAPITAL
    worksheet.getColumn('E').width = 22; // Aporte... / SERVICIOS
    worksheet.getColumn('F').width = 22; // SUBVENCION
    worksheet.getColumn('G').width = 22; // VIATICOS
    worksheet.getColumn('H').width = 22; // ENCARGO INTERNO
    worksheet.getColumn('I').width = 22; // TOTAL
    
  }

  private crearHojaDeReportePorMes(workbook: ExcelJS.Workbook, sheetName: string, data: ReporteDataMensual, projectName: string): void {
    const worksheet = workbook.addWorksheet(sheetName);

    const headerStyle: Partial<ExcelJS.Style> = {
        font: { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } },
        alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    const totalRowStyle: Partial<ExcelJS.Style> = {
        font: { bold: true, color: { argb: 'FF000000'} },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } },
        alignment: { horizontal: 'right' }
    };
    const cellStyle: Partial<ExcelJS.Style> = {
        alignment: { vertical: 'middle' },
        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    };
    // ⭐ NUEVO: Estilo específico para celdas con texto que debe ajustarse
    const wrapTextStyle: Partial<ExcelJS.Style> = {
        ...cellStyle,
        alignment: { ...cellStyle.alignment, wrapText: true }
    };

    // ▼▼▼ MODIFICACIÓN AQUÍ (1) ▼▼▼
    // Añadido formato de moneda con rojos para negativos
    const moneyFormat = '"S/" #,##0.00;[Red]-"S/" #,##0.00';
    // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲


    worksheet.mergeCells('A1:N1');
    worksheet.getCell('A1').value = `EJECUCIÓN DE GASTOS - ${sheetName}`;
    worksheet.getCell('A1').font = { name: 'Arial', size: 14, bold: true };
    worksheet.getCell('A1').alignment = { horizontal: 'center' };
    
    worksheet.mergeCells('A2:N2');
    worksheet.getCell('A2').value = projectName;
    worksheet.getCell('A2').font = { name: 'Arial', size: 11, italic: true };
    worksheet.getCell('A2').alignment = { horizontal: 'center' };
    
    let currentRow = 4;
    let totalGeneral = 0;

    const renderSection = (title: string, gastos: Gasto[]) => {
      if (gastos.length === 0) return;
      worksheet.mergeCells(currentRow, 1, currentRow, 14);
      worksheet.getCell(`A${currentRow}`).value = title;
      worksheet.getCell(`A${currentRow}`).font = { name: 'Arial', size: 11, bold: true };
      currentRow++;
      
      const headers = ['DOC','N°','SIAF','A NOMBRE DE','CONCEPTO','MONTO','ESPECIFICA','MONTO2','ESPECIFICA2','F.F.','MES','F. DEVENGADO','PROYECTO','META'];
      const headerRow = worksheet.getRow(currentRow);
      headerRow.values = headers;
      headerRow.eachCell(cell => cell.style = headerStyle);
      currentRow++;

      const startDataRow = currentRow;
      let totalSection = 0;
      gastos.forEach(gasto => {
        worksheet.addRow([
          gasto.tipoDocumento, gasto.numeroDocumento, gasto.siaf,
          gasto.aNombreDe, gasto.concepto, gasto.monto,
          gasto.especifica, gasto.monto2, gasto.especifica2,
          gasto.ff, gasto.mes, gasto.fechaDevengado,
          gasto.proyecto, gasto.meta
        ]);
        const addedRow = worksheet.getRow(currentRow);
        
        // ▼▼▼ MODIFICACIÓN AQUÍ (2) ▼▼▼
        addedRow.getCell(6).numFmt = moneyFormat;
        // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

        totalSection += Number(gasto.monto) || 0;
        currentRow++;
      });
      
      for(let i = startDataRow; i < currentRow; i++) {
        const row = worksheet.getRow(i);
        row.eachCell((cell, colNumber) => {
            //  MODIFICACIÓN: Aplicar estilo de ajuste de texto a columnas específicas
            if (colNumber === 4 || colNumber === 5) { // Columnas 'A NOMBRE DE' y 'CONCEPTO'
                cell.style = wrapTextStyle;
            } else {
                cell.style = cellStyle;
            }
        });
      }

      const totalRow = worksheet.getRow(currentRow);
      totalRow.getCell(5).value = `Total ${title}`;
      totalRow.getCell(6).value = totalSection;
      totalRow.getCell(5).style = totalRowStyle;

      // ▼▼▼ MODIFICACIÓN AQUÍ (3) ▼▼▼
      totalRow.getCell(6).style = { ...totalRowStyle, numFmt: moneyFormat, alignment: { horizontal: 'left' } };
      // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
      
      currentRow += 2;
      totalGeneral += totalSection;
    };

    renderSection('Gastos Corrientes', data.gastosCorrientes);
    renderSection('Subvenciones', data.subvenciones);
    renderSection('Gastos de Capital', data.gastosCapital);

    const grandTotalRow = worksheet.getRow(currentRow);
    grandTotalRow.getCell(5).value = `TOTAL GENERAL`;
    grandTotalRow.getCell(6).value = totalGeneral;
    grandTotalRow.getCell(5).style = totalRowStyle;

    // ▼▼▼ MODIFICACIÓN AQUÍ (4) ▼▼▼
    grandTotalRow.getCell(6).style = { ...totalRowStyle, numFmt: moneyFormat, alignment: { horizontal: 'left' } };
    // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
    
    // --- SECCIÓN MODIFICADA: Anchos de columna fijos ---
    // Se reemplaza el bucle de auto-ajuste por anchos definidos.
    // Puedes cambiar estos valores según tus preferencias.
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
        { key: 'META', width: 10 }
    ];

    columnWidths.forEach((col, index) => {
      // ▼▼▼ MODIFICACIÓN AQUÍ (5) ▼▼▼
      const colNum = index + 1;
      const column = worksheet.getColumn(colNum);
      column.width = col.width;

      // Forzar formato de TEXTO para N° (col 2) y SIAF (col 3)
      if (colNum === 2 || colNum === 3) {
        column.numFmt = '@';
      }
      // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

      // ▼▼▼ MODIFICACIÓN AQUÍ ▼▼▼
      // Forzar formato de FECHA para F. DEVENGADO (col 12)
      if (colNum === 12) {
        column.numFmt = 'dd/mm/yyyy';
      }
      // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

      
    });

    /*
    // Bucle original de auto-ajuste (eliminado)
    ...
    */
  }
}