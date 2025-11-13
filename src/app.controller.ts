// src/app.controller.ts
// ▼▼▼ MODIFICAR IMPORTACIONES ▼▼▼
import { Controller, Get, Post, Body, Query, Res, Param, Put, Delete } from '@nestjs/common';
// ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
import { AppService } from './app.service';
import { Gasto } from './gasto.entity';
import type { Response } from 'express';
import * as ExcelJS from 'exceljs'; // ¡Asegúrate de importar ExcelJS!
import type { PicMetadataDto } from './app.service';

@Controller('api')
export class AppController {
  constructor(private readonly appService: AppService) {}

  // --- ENDPOINTS EXISTENTES (sin cambios) ---
  @Post('gastos')
  create(@Body() gastoData: Partial<Gasto>): Promise<Gasto> {
    return this.appService.create(gastoData);
  }

  @Post('gastos/import')
  createMany(@Body() gastosData: Partial<Gasto>[]): Promise<Gasto[]> {
    return this.appService.createMany(gastosData);
  }
  

  @Get('gastos')
  findAll(): Promise<Gasto[]> {
    return this.appService.findAll();
  }

  @Get('projects')
  findUniqueProjects(): Promise<string[]> {
    return this.appService.findUniqueProjects();
  }


/**
   * NUEVO ENDPOINT
   * Devuelve solo la lista de proyectos CONTRATO
   */
  @Get('projects/contrato')
  findContratoProjects(): Promise<string[]> {
    return this.appService.findContratoProjects();
  }

  /**
   * NUEVO ENDPOINT
   * Devuelve solo la lista de proyectos PIC
   */
  @Get('projects/pic')
  findPicProjects(): Promise<string[]> {
    return this.appService.findPicProjects();
  }


  // ▼▼▼ 3. AÑADIR LOS NUEVOS ENDPOINTS PARA METADATA DE PIC ▼▼▼

  /**
   * Obtiene la metadata guardada para un proyecto PIC.
   */
  @Get('pic-metadata/:projectName')
  getPicMetadata(@Param('projectName') projectName: string) {
    // Decodificar el nombre del proyecto que viene de la URL
    const decodedProjectName = decodeURIComponent(projectName);
    return this.appService.getPicMetadata(decodedProjectName);
  }

  /**
   * Guarda o actualiza la metadata para un proyecto PIC.
   */
  @Post('pic-metadata/:projectName')
  savePicMetadata(
    @Param('projectName') projectName: string,
    @Body() data: PicMetadataDto,
  ) {
    const decodedProjectName = decodeURIComponent(projectName);
    return this.appService.savePicMetadata(decodedProjectName, data);
  }
  
  // ▲▲▲ FIN DE LOS NUEVOS ENDPOINTS ▲▲▲


// --- ENDPOINTS DE GENERACIÓN DE REPORTES (MODIFICADOS) ---

  /**
   * ENDPOINT ESPECIALIZADO
   * Genera solo reportes de CONTRATO
   */
  @Post('reports/contrato')
  async generateContratoReport(
    @Body() body: { projectName: string; metadata: any },
    @Res() res: Response,
  ) {
    const { projectName, metadata } = body;
    if (!projectName) {
      res.status(400).send('El nombre del proyecto es requerido');
      return;
    }

    try {
      const workbook = await this.appService.generateContratoReport(
        projectName,
        metadata,
      );
      
      const safeName = projectName.replace(/[^a-z0-9]/gi, '_');
      const fileName = `Reporte_Contrato_${safeName}.xlsx`;
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      await workbook.xlsx.write(res);
      res.end();
    } catch (error) {
      console.error('Error al generar reporte CONTRATO:', error);
      res.status(500).send(`Error al generar el reporte: ${error.message || error}`);
    }
  }

  /**
   * NUEVO ENDPOINT
   * Genera solo reportes de GRUPO PIC
   */
  @Post('reports/pic')
  async generatePicReport(
    @Body() body: { modalityName: string }, // <-- AHORA ESPERA 'modalityName'
    @Res() res: Response,
  ) {
    const { modalityName } = body;
    if (!modalityName) {
      res.status(400).send('El nombre del proyecto es requerido');
      return;
    }

    try {
      const workbook = await this.appService.generatePicReportGroup(
        modalityName,
      );
      
      const safeName = modalityName.replace(/[^a-z0-9]/gi, '_');
      const fileName = `Reporte_PIC_${safeName}.xlsx`;

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      await workbook.xlsx.write(res);
      res.end();
    } catch (error) {
      console.error('Error al generar reporte PIC:', error);
      res.status(500).send(`Error al generar el reporte: ${error.message || error}`);
    }
  }
  // ▼▼▼ AÑADIR ESTE NUEVO ENDPOINT ▼▼▼
  @Get('reports/by-meta')
  async generateReportByMeta(@Res() res: Response) {
    const workbook = await this.appService.generateReportByMeta();

    // --- Preparamos la respuesta para la descarga ---
    const fileName = `Reporte_Global_por_Meta.xlsx`;

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=${fileName}`,
    );

    // Escribimos el libro en la respuesta
    await workbook.xlsx.write(res);
    res.end();
  }
  // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

  
}