-- Tabla única de almacenamiento clave-valor para la app de finanzas.
-- Guarda todo el estado (cuentas, movimientos, montosFijos, indicadores,
-- fechasFacturacion) como un solo registro JSON por usuario.

create table app_data (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  unique (user_id)
);

-- Seguridad: cada usuario solo puede leer/escribir su propia fila.
alter table app_data enable row level security;

create policy "Los usuarios ven solo su propio registro"
  on app_data for select
  using (auth.uid() = user_id);

create policy "Los usuarios escriben solo su propio registro"
  on app_data for insert
  with check (auth.uid() = user_id);

create policy "Los usuarios actualizan solo su propio registro"
  on app_data for update
  using (auth.uid() = user_id);
