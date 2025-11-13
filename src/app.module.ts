import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { TypeOrmModule } from '@nestjs/typeorm';
import { Gasto } from './gasto.entity';
import { PicMetadata } from './pic-metadata.entity';


@Module({
    imports: [
    TypeOrmModule.forRoot({
      type: 'postgres',
      host: '127.0.0.1',
      port: 5432,
      username: 'postgres', // Este es el usuario por defecto
      password: '70707070', // <-- ¡PON AQUÍ TU CONTRASEÑA!
      database: 'contabilidad_db',
      entities: [Gasto, PicMetadata],
      synchronize: true, // Esto crea/actualiza la tabla automáticamente
    }),
    TypeOrmModule.forFeature([Gasto, PicMetadata]),
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
