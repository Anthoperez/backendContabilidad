import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { TypeOrmModule } from '@nestjs/typeorm';
import { Gasto } from './gasto.entity';


@Module({
    imports: [
    TypeOrmModule.forRoot({
      type: 'postgres',

      // ▼▼▼ MODIFICACIÓN AQUÍ ▼▼▼
      // Lee la URL de la base de datos desde las variables de entorno
      url: process.env.DATABASE_URL,
      // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲

      entities: [Gasto],
      
      // ▼▼▼ MUY IMPORTANTE ▼▼▼
      synchronize: false, // Poner en 'false' para producción
      ssl: true,
      extra: {
        ssl: {
          rejectUnauthorized: false
        }
      }
      // ▲▲▲ FIN DE LA MODIFICACIÓN ▲▲▲
    }),
    TypeOrmModule.forFeature([Gasto]),
  ],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
