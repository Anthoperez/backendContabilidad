// src/gasto.entity.ts
import { Entity, Column, PrimaryGeneratedColumn, CreateDateColumn } from 'typeorm';

@Entity()
export class Gasto {
  @PrimaryGeneratedColumn()
  id: number;

  // --- CAMPOS OBLIGATORIOS ---
  @Column()
  tipoDocumento: string;

  @Column()
  numeroDocumento: string;

  @Column()
  aNombreDe: string;

  @Column('text')
  concepto: string;

  @Column('decimal', { precision: 10, scale: 2 })
  monto: number;

  @Column()
  especifica: string;

  @Column({nullable: true}) // <-- Mantenemos el tipo por defecto para la fecha
  fechaDevengado: Date;

  @Column()
  proyecto: string;

  @Column()
  mes: string;

  @Column()
  meta: string;

  @CreateDateColumn()
  fechaRegistro: Date;

  // --- CAMPOS OPCIONALES ---
  @Column({ nullable: true })
  siaf: string;

  @Column('decimal', { precision: 10, scale: 2, nullable: true })
  monto2: number;

  @Column({ nullable: true })
  especifica2: string;

  @Column({ nullable: true })
  ff: string;

  @Column({ nullable: true })
  certificacionViatico: string;

  @Column({ nullable: true })
  destino: string;

  @Column({ nullable: true }) // <-- CORRECCIÓN: Dejamos que TypeORM maneje el tipo de fecha
  fechaSalida: Date;

  @Column({ nullable: true }) // <-- CORRECIÓN: Dejamos que TypeORM maneje el tipo de fecha
  fechaRetorno: Date;
}