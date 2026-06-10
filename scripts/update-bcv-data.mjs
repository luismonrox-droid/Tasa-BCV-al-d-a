import fs from 'node:fs/promises';

const DATA_FILE = 'public/data.json';
const SOURCE_URL = 'https://rates.dolarvzla.com/bcv/current.json';

async function main() {
  const res = await fetch(SOURCE_URL);

  if (!res.ok) {
    throw new Error(`No se pudo consultar la fuente: ${res.status}`);
  }

  const source = await res.json();

  const fecha = source.current?.date;
  const dolar = source.current?.usd;
  const euro = source.current?.eur;

  if (!fecha || typeof dolar !== 'number' || typeof euro !== 'number') {
    throw new Error('La respuesta no trae fecha/usd/eur válidos');
  }

  let data = [];

  try {
    const raw = await fs.readFile(DATA_FILE, 'utf-8');
    data = JSON.parse(raw);
    if (!Array.isArray(data)) data = [];
  } catch {
    data = [];
  }

  const index = data.findIndex((item) => item.fecha === fecha);

  const nuevoRegistro = {
    fecha,
    dolar,
    euro
  };

  if (index >= 0) {
    data[index] = nuevoRegistro;
  } else {
    data.push(nuevoRegistro);
  }

  data.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));

  await fs.writeFile(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');

  console.log(`data.json actualizado con fecha ${fecha}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
