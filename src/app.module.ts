import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { ConfigModule } from '@nestjs/config';
import { AppService } from './app.service';
import { TypeOrmModule } from '@nestjs/typeorm';
import { Gasto } from './gasto.entity';
import { PicMetadata } from './pic-metadata.entity';


@Module({
    imports: [

      ConfigModule.forRoot({
      isGlobal: true, // Hace que las variables estén disponibles en toda la app
      envFilePath: '.env', // Lee el archivo .env
    }),
    
    TypeOrmModule.forRoot({
      type: 'postgres',
      url: process.env.DATABASE_URL,
      entities: [Gasto, PicMetadata],
      synchronize: true, // Esto crea/actualiza la tabla automáticamente
      ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false, // <-- IMPORTANTE PARA NEON
      logging: false,
    }),
    TypeOrmModule.forFeature([Gasto, PicMetadata]),
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
