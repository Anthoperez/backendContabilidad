// src/app.controller.ts
import { Body, Controller, Get, Post } from '@nestjs/common';
import { AppService } from './app.service';
import { Gasto } from './gasto.entity';

@Controller('api/gastos') // La URL base será http://localhost:3000/api/gastos
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Post()
  create(@Body() gastoData: Partial<Gasto>) {
    return this.appService.create(gastoData);
  }

    // ¡NUEVO ENDPOINT para importar en lote!
  @Post('import')
  import(@Body() gastosData: Partial<Gasto>[]) {
    return this.appService.createMany(gastosData);
  }

  @Get()
  findAll() {
    return this.appService.findAll();
  }
}