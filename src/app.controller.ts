import { Controller, Get, Post, Body, Query, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { Gasto } from './gasto.entity';
import type { Response } from 'express'; // Importa 'Response' usando 'import type'

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

  // --- NUEVO ENDPOINT: OBTENER LISTA DE PROYECTOS ---
  @Get('projects')
  findUniqueProjects(): Promise<string[]> {
    return this.appService.findUniqueProjects();
  }

  // --- NUEVO ENDPOINT: GENERAR Y DESCARGAR REPORTE ---
  @Get('reports/generate')
  async generateReport(
    @Query('projectName') projectName: string,
    @Res() res: Response, // Inyectamos el objeto 'Response' de Express
  ) {
    if (!projectName) {
      res.status(400).send('El nombre del proyecto es requerido');
      return;
    }

    const workbook = await this.appService.generateProjectReport(projectName);

    // --- PREPARAMOS LA RESPUESTA PARA LA DESCARGA DEL ARCHIVO ---
    
    // 1. Creamos un nombre de archivo seguro
    const fileName = `Reporte_${projectName.replace(/[^a-z0-9]/gi, '_')}.xlsx`;
    
    // 2. Configuramos las cabeceras HTTP
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=${fileName}`,
    );
    
    // 3. Escribimos el libro de Excel en la respuesta y la enviamos.
    await workbook.xlsx.write(res);
    res.end();
  }
}