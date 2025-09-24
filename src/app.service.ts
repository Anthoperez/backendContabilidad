// src/app.service.ts
import { Injectable } from '@nestjs/common';
import { InjectRepository } from '@nestjs/typeorm';
import { Repository } from 'typeorm';
import { Gasto } from './gasto.entity';

@Injectable()
export class AppService {
  constructor(
    @InjectRepository(Gasto)
    private gastoRepository: Repository<Gasto>,
  ) {}

  // Lógica para crear un nuevo gasto
  async create(gastoData: Partial<Gasto>): Promise<Gasto> {
    const nuevoGasto = this.gastoRepository.create(gastoData);
    return this.gastoRepository.save(nuevoGasto);
  }

    // ¡NUEVA LÓGICA para guardar muchos gastos a la vez!
  async createMany(gastosData: Partial<Gasto>[]): Promise<Gasto[]> {
    const gastos = this.gastoRepository.create(gastosData);
    return this.gastoRepository.save(gastos);
  }
  

  // Lógica para obtener todos los gastos
  findAll(): Promise<Gasto[]> {
    return this.gastoRepository.find({ order: { fechaRegistro: 'DESC' } });
  }
}