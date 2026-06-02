-- Ejecutar una sola vez en Supabase SQL Editor para completar columnas de GPSRUTA Cloud
alter table facturas add column if not exists detalle text;
alter table ingresos add column if not exists factura_id bigint;
alter table egresos add column if not exists deuda_id bigint;
alter table egresos add column if not exists numero_factura_pago text;
alter table deudas add column if not exists fecha date default current_date;
alter table deudas add column if not exists categoria text default 'Compra de equipos';
alter table deudas add column if not exists email_proveedor text;
create table if not exists adjuntos (
  id bigint generated always as identity primary key,
  factura_id bigint unique,
  name text,
  size bigint,
  type text,
  path text,
  sent boolean default false,
  attached_at text,
  sent_at text,
  created_at timestamp default now()
);
