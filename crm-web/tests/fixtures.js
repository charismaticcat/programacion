/**
 * Datos de prueba 100 % sintéticos (no son clientes reales).
 * Reproducen los casos de la base CCH: fijos de 7 dígitos, emails repetidos,
 * email mal escrito, actividad vacía (se completa con CIIU), NIT que no debe importarse.
 */
'use strict';
const HEADER = ['EST-MATRICULA', 'RAZON SOCIAL', 'NIT', 'FEC-MATRICULA', 'FEC-RENOVACION', 'ULT-ANO_REN', 'DIR-COMERCIAL',
  'BARRIO-COMERCIAL', 'MUN-COMERCIAL', 'TEL-COM-1', 'EMAIL-COMERCIAL', 'CIIU-1', 'ACTIVIDAD', 'TAM-EMPRESA'];
const ROWS = [
  ['MA', 'PANADERIA PRUEBA UNO', 900000001, 20100101, 20260301, 2026, 'CALLE 1 # 2-3', '01001 - CENTRO PRUEBA', '41001 - NEIVA', 3000000001, 'prueba.uno@example.com', 'C1081 ** Elaboracion de productos de panaderia', 'PANADERIA Y PASTELERIA', 'MICRO EMPRESA'],
  ['MA', 'FERRETERIA PRUEBA DOS SAS', 900000002, 20110101, 20250301, 2025, 'CARRERA 5 # 6-7', '01002 - BARRIO PRUEBA', '41551 - PITALITO', 8700001, 'compartido@example.com', 'G4752 ** Comercio al por menor de articulos de ferreteria', '', 'PEQUEÑA EMPRESA'],
  ['MA', 'TIENDA PRUEBA TRES', 900000003, 20120101, 20260301, 2026, 'CALLE 9 # 9-9', '01003 - OTRO BARRIO', '41524 - PALERMO', 3000000003, 'compartido@example.com', 'G4711 ** Comercio al por menor en establecimientos no especializados', 'VENTA DE VIVERES', 'MICRO EMPRESA'],
  ['MA', 'PELUQUERIA PRUEBA CUATRO', 900000004, 20130101, 20240301, 2024, 'AVENIDA 4 # 4-4', '', '41615 - RIVERA', 3000000004, 'mal@escritocom', 'S9602 ** Peluqueria y otros tratamientos de belleza', 'PELUQUERIA', 'MICRO EMPRESA'],
  ['MA', 'RESTAURANTE PRUEBA CINCO', 900000005, 20140101, 20260301, 2026, 'CALLE 5 # 5-5', '01005 - BARRIO CINCO', '41001 - NEIVA', 3000000005, 'cinco@example.com', 'I5611 ** Expendio a la mesa de comidas preparadas', 'RESTAURANTE', 'MICRO EMPRESA'],
  ['MA', 'DROGUERIA PRUEBA SEIS', 900000006, 20150101, 20260301, 2026, 'CALLE 6 # 6-6', '01006 - BARRIO SEIS', '41001 - NEIVA', 3000000005, 'seis@example.com', 'G4773 ** Comercio al por menor de productos farmaceuticos', 'DROGUERIA', 'MICRO EMPRESA'],
  ['MA', 'TALLER PRUEBA SIETE', 900000007, 20160101, 20260301, 2026, 'CALLE 7 # 7-7', '01007 - BARRIO SIETE', '41551 - PITALITO', 0, 'siete@example.com', 'G4520 ** Mantenimiento y reparacion de vehiculos', 'TALLER', 'MICRO EMPRESA']
];

/** PNG válido de 1x1. */
const PNG_1PX = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
/** Inicio de un JPEG válido (FF D8 FF E0…). */
const JPEG_HEAD = Buffer.from([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01]);
/** PDF mínimo. */
const PDF_MIN = Buffer.from('%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF').toString('base64');

module.exports = { HEADER, ROWS, PNG_1PX, JPEG_HEAD, PDF_MIN };
