import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { TypeOrmModule } from '@nestjs/typeorm';
import { Gasto } from './gasto.entity';


@Module({
    imports: [
    TypeOrmModule.forRoot({
      type: 'postgres',
      host: 'localhost',
      port: 5432,
      username: 'postgres', // Este es el usuario por defecto
      password: '70707070', // <-- ¡PON AQUÍ TU CONTRASEÑA!
      database: 'contabilidad_db',
      entities: [Gasto],
      synchronize: true, // Esto crea/actualiza la tabla automáticamente
    }),
    TypeOrmModule.forFeature([Gasto]),
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
