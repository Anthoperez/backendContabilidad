// src/pic-metadata.entity.ts
import { Entity, Column, PrimaryGeneratedColumn, Index } from 'typeorm';

// Esta interfaz define la estructura de los ingresos
export interface IngresoPic {
  descripcion: string;
  monto: number | null;
}

// Esta interfaz define la estructura de los gastos de años anteriores
export interface GastoPrevioPic {
  year: number;
  bienesCorrientes: number | null;
  bienesCapital: number | null;
  servicios: number | null;
  subvencion: number | null;
  viaticos: number | null;
  encargoInterno: number | null;
}

@Entity()
export class PicMetadata {
  @PrimaryGeneratedColumn()
  id: number;

  // Usamos el nombre del proyecto como un ID único
  @Index({ unique: true })
  @Column()
  projectName: string;

  // --- Campos de Metadata ---

  @Column({ nullable: true })
  investigador: string;

  @Column({ nullable: true })
  tesista: string;

  @Column({ nullable: true })
  asesor: string;

  @Column({ nullable: true })
  duracion: string ;

  // Este es el 'PRESUPUESTO' del "cuadro rojo"
  @Column('decimal', { precision: 12, scale: 2, nullable: true })
  presupuestoTotal: number | null;

  // Guardamos el array de ingresos como un JSON
  @Column('jsonb', { nullable: true })
  ingresos: IngresoPic[];

  // Guardamos el array de gastos previos como un JSON
  @Column('jsonb', { nullable: true })
  gastosAnosAnteriores: GastoPrevioPic[];
}