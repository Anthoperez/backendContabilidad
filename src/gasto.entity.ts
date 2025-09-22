// src/gasto.entity.ts
import { Entity, Column, PrimaryGeneratedColumn, CreateDateColumn } from 'typeorm';

@Entity()
export class Gasto {
  @PrimaryGeneratedColumn()
  id: number;

  @Column()
  tipoDocumento: string;

  @Column()
  numeroDocumento: string;

  @Column({ nullable: true })
  siaf: string;

  @Column()
  aNombreDe: string;

  @Column('text')
  concepto: string;

  @Column('decimal', { precision: 10, scale: 2 })
  monto: number;

  @Column()
  especifica: string;

  @Column()
  fechaDevengado: Date;

  @Column()
  proyecto: string;

  @CreateDateColumn()
  fechaRegistro: Date; // Columna extra para saber cuándo se guardó
}