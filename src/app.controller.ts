// src/app.controller.ts
// ▼▼▼ MODIFICAR IMPORTACIONES ▼▼▼
import { Controller, Get, Post, Body, Query, Res, Param, Put, Delete } from '@nestjs/common';
// ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
import { AppService } from './app.service';
import { Gasto } from './gasto.entity';
import type { Response } from 'express';

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


  // --- ENDPOINT MODIFICADO: GENERAR Y DESCARGAR REPORTE ---
  // ▼▼▼ MODIFICACIÓN COMPLETA DE ESTE MÉTODO ▼▼▼
  @Post('reports/generate') // <-- CAMBIADO A @Post
  async generateReport(
    @Body() body: { projectName: string; metadata: any }, // <-- Recibimos el body
    @Res() res: Response,
  ) {
    const { projectName, metadata } = body; // <-- Extraemos los datos

    if (!projectName) {
      res.status(400).send('El nombre del proyecto es requerido');
      return;
    }

    // Pasamos los metadatos al servicio
    const workbook = await this.appService.generateProjectReport(projectName, metadata);

    // --- PREPARAMOS LA RESPUESTA (sin cambios) ---
    const fileName = `Reporte_${projectName.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=${fileName}`,
    );
    await workbook.xlsx.write(res);
    res.end();
  }
  // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

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