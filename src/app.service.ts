import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { Gasto } from './gasto.entity';
import * as ExcelJS from 'exceljs';

// --- NUEVA INTERFAZ PARA DEFINIR LA ESTRUCTURA DE UN REPORTE ---
interface ReporteData {
  gastosCorrientes: Gasto[];
  subvenciones: Gasto[];
  gastosCapital: Gasto[];
  // Se pueden añadir más categorías aquí si es necesario
}

@Injectable()
export class AppService {
  constructor(
    @InjectRepository(Gasto)
    private gastoRepository: Repository<Gasto>,
  ) {}

  // --- MÉTODOS EXISTENTES (sin cambios) ---
  create(gasto: Partial<Gasto>): Promise<Gasto> {
    const nuevoGasto = this.gastoRepository.create(gasto);
    return this.gastoRepository.save(nuevoGasto);
  }

// --- MÉTODO A MODIFICAR ---
  createMany(gastos: Partial<Gasto>[]): Promise<Gasto[]> {
    // 1. LIMPIAR LOS DATOS ANTES DE GUARDAR
    const gastosLimpios = gastos.map(gasto => {
      // Si el campo 'proyecto' existe y es un string, le quitamos los espacios
      if (gasto.proyecto && typeof gasto.proyecto === 'string') {
        gasto.proyecto = gasto.proyecto.trim();
      }
      // Podrías añadir más limpiezas aquí para otros campos si lo necesitas
      return gasto;
    });

    // 2. CREAR Y GUARDAR LOS GASTOS YA LIMPIOS
    const nuevosGastos = this.gastoRepository.create(gastosLimpios);
    return this.gastoRepository.save(nuevosGastos);
  }

  findAll(): Promise<Gasto[]> {
    return this.gastoRepository.find();
  }

  // --- NUEVO MÉTODO: OBTENER PROYECTOS ÚNICOS ---
  // Este método consultará la base de datos y devolverá una lista
  // de nombres de proyectos sin repetir, para el dropdown del frontend.
 async findUniqueProjects(): Promise<string[]> {
    const projects = await this.gastoRepository
      .createQueryBuilder('gasto')
      // CAMBIO CLAVE: Añadimos un alias 'proyecto' a la selección.
      .select('gasto.proyecto', 'proyecto') 
      .distinct(true)
      // AHORA ORDENAMOS POR EL ALIAS
      .orderBy('proyecto', 'ASC') 
      .getRawMany();

    // Ahora, p.proyecto existirá y el resto del código funcionará perfecto.
    return projects.map((p) => p.proyecto).filter(Boolean); 
  }

  // --- NUEVO MÉTODO: GENERAR REPORTE PARA UN PROYECTO ---
  // Esta es la función principal que orquesta todo el proceso.
  async generateProjectReport(projectName: string): Promise<ExcelJS.Workbook> {
    // 1. FILTRAR: Obtenemos todos los gastos del proyecto seleccionado.
    const gastosDelProyecto = await this.gastoRepository.find({
      where: { proyecto: projectName },
    });

    // 2. CLASIFICAR: Agrupamos los gastos en categorías.
    const reporteData = this.clasificarGastos(gastosDelProyecto);

    // 3. GENERAR EXCEL: Creamos el archivo Excel en memoria.
    const workbook = this.crearLibroExcelDeReporte(reporteData, projectName);

    return workbook;
  }

  // --- NUEVO MÉTODO AUXILIAR: LÓGICA DE CLASIFICACIÓN ---
  // Aquí es donde definimos las reglas para clasificar cada gasto.
  // ¡Puedes modificar estas reglas fácilmente si cambian en el futuro!
  private clasificarGastos(gastos: Gasto[]): ReporteData {
    const reporte: ReporteData = {
      gastosCorrientes: [],
      subvenciones: [],
      gastosCapital: [],
    };

    const REGLAS = {
      // Si el código 'especifica' empieza con '2.3.', es un Gasto Corriente.
      GASTOS_CORRIENTES: '2.3.',
      // Si empieza con '2.5.', es una Subvención.
      SUBVENCIONES: '2.5.',
      // Si empieza con '2.6.', es un Gasto de Capital.
      GASTOS_CAPITAL: '2.6.',
    };

    for (const gasto of gastos) {
      const especifica = gasto.especifica?.trim() || '';

      if (especifica.startsWith(REGLAS.GASTOS_CAPITAL)) {
        reporte.gastosCapital.push(gasto);
      } else if (especifica.startsWith(REGLAS.SUBVENCIONES)) {
        reporte.subvenciones.push(gasto);
      } else if (especifica.startsWith(REGLAS.GASTOS_CORRIENTES)) {
        // Por defecto, o si cumple la regla, lo asignamos a gastos corrientes
        reporte.gastosCorrientes.push(gasto);
      } else {
        // Si no coincide con ninguna regla específica, lo asignamos por defecto a Corrientes.
        reporte.gastosCorrientes.push(gasto);
      }
    }
    return reporte;
  }
  
  // --- NUEVO MÉTODO AUXILIAR: CONSTRUCTOR DEL ARCHIVO EXCEL ---
  // Este método usa 'exceljs' para crear la estructura del reporte.
  private crearLibroExcelDeReporte(data: ReporteData, projectName: string): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    // Por ahora, creamos una sola hoja. Se puede expandir para crear una por mes.
    const worksheet = workbook.addWorksheet('Reporte General');

    // Título del reporte
    worksheet.mergeCells('A1:E1');
    worksheet.getCell('A1').value = `Reporte de Ejecución de Gastos`;
    worksheet.getCell('A1').font = { bold: true, size: 16 };
    
    worksheet.mergeCells('A2:E2');
    worksheet.getCell('A2').value = projectName;
    worksheet.getCell('A2').font = { italic: true, size: 14 };

    let currentRow = 4; // Empezamos a escribir los datos desde la fila 4

    // --- SECCIÓN DE GASTOS CORRIENTES ---
    worksheet.getCell(`A${currentRow}`).value = 'Gastos Corrientes';
    worksheet.getCell(`A${currentRow}`).font = { bold: true, color: { argb: 'FF0000FF' } };
    currentRow++;
    // Encabezados de la tabla
    worksheet.getRow(currentRow).values = ['Fecha Devengado', 'Concepto', 'Específica', 'Monto'];
    worksheet.getRow(currentRow).font = { bold: true };
    currentRow++;
    // Filas de datos
    data.gastosCorrientes.forEach(gasto => {
        worksheet.getRow(currentRow).values = [gasto.fechaDevengado, gasto.concepto, gasto.especifica, gasto.monto];
        currentRow++;
    });
    currentRow++; // Dejar una fila en blanco

    // --- SECCIÓN DE SUBVENCIONES ---
    worksheet.getCell(`A${currentRow}`).value = 'Subvenciones';
    worksheet.getCell(`A${currentRow}`).font = { bold: true, color: { argb: 'FF0000FF' } };
    currentRow++;
    worksheet.getRow(currentRow).values = ['Fecha Devengado', 'Concepto', 'Específica', 'Monto'];
    worksheet.getRow(currentRow).font = { bold: true };
    currentRow++;
    data.subvenciones.forEach(gasto => {
        worksheet.getRow(currentRow).values = [gasto.fechaDevengado, gasto.concepto, gasto.especifica, gasto.monto];
        currentRow++;
    });
    currentRow++;

    // --- SECCIÓN DE GASTOS DE CAPITAL ---
    worksheet.getCell(`A${currentRow}`).value = 'Gastos de Capital';
    worksheet.getCell(`A${currentRow}`).font = { bold: true, color: { argb: 'FF0000FF' } };
    currentRow++;
    worksheet.getRow(currentRow).values = ['Fecha Devengado', 'Concepto', 'Específica', 'Monto'];
    worksheet.getRow(currentRow).font = { bold: true };
    currentRow++;
    data.gastosCapital.forEach(gasto => {
        worksheet.getRow(currentRow).values = [gasto.fechaDevengado, gasto.concepto, gasto.especifica, gasto.monto];
        currentRow++;
    });

    // Auto-ajustar el ancho de las columnas
    worksheet.columns.forEach(column => {
        column.width = 30;
    });

    return workbook;
  }
}