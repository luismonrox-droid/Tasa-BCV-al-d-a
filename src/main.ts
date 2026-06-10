import './style.css'
import Chart from 'chart.js/auto'
import * as XLSX from 'xlsx'

const app = document.getElementById('app') as HTMLDivElement

let pagina: 'inicio' | 'calculadora' | 'historico' = 'inicio'
let menuAbierto = false

let dolar = 0
let euro = 0
let pantalla = '0'
let tasa: 'dolar' | 'euro' = 'dolar'

let historico: any[] = []
let chartInstance: Chart | null = null

const MAX_RANGO_DIAS = 90
let fechaDesde = ''
let fechaHasta = ''
let rangoRapidoActivo = 30

let fechaCalc = ''

let ultimoRegistro: any = null
let anteriorRegistro: any = null

async function cargarHistorico() {
  const res = await fetch('/data.json')
  historico = await res.json()

  historico.sort(
    (a, b) => new Date(a.fecha).getTime() - new Date(b.fecha).getTime()
  )

  if (historico.length > 0) {
    ultimoRegistro = historico[historico.length - 1]
    anteriorRegistro =
      historico.length > 1
        ? historico[historico.length - 2]
        : historico[historico.length - 1]

    fechaCalc = ultimoRegistro.fecha
    aplicarRangoRapido(30, false)

    dolar = ultimoRegistro.dolar
    euro = ultimoRegistro.euro
  }
}

async function bootApp() {
  await cargarHistorico()
  render()
}

function tituloPagina() {
  if (pagina === 'inicio') return 'Inicio'
  if (pagina === 'calculadora') return 'Calculadora'
  return 'Histórico'
}

function sumarDias(fechaISO: string, dias: number) {
  const fecha = new Date(fechaISO + 'T00:00:00')
  fecha.setDate(fecha.getDate() + dias)
  return fecha.toISOString().split('T')[0]
}

function restarDias(fechaISO: string, dias: number) {
  return sumarDias(fechaISO, -dias)
}

function diferenciaDias(desdeISO: string, hastaISO: string) {
  const desde = new Date(desdeISO + 'T00:00:00').getTime()
  const hasta = new Date(hastaISO + 'T00:00:00').getTime()
  return Math.floor((hasta - desde) / (1000 * 60 * 60 * 24))
}

function normalizarRango(origen: 'desde' | 'hasta') {
  if (!fechaDesde || !fechaHasta) return

  const diff = diferenciaDias(fechaDesde, fechaHasta)

  if (diff < 0) {
    if (origen === 'desde') {
      fechaHasta = fechaDesde
    } else {
      fechaDesde = fechaHasta
    }
  }

  const diffCorregido = diferenciaDias(fechaDesde, fechaHasta)

  if (diffCorregido > MAX_RANGO_DIAS) {
    if (origen === 'desde') {
      fechaHasta = sumarDias(fechaDesde, MAX_RANGO_DIAS)
    } else {
      fechaDesde = restarDias(fechaHasta, MAX_RANGO_DIAS)
    }
  }

  if (historico.length > 0) {
    const fechaMin = historico[0].fecha
    const fechaMax = historico[historico.length - 1].fecha

    if (fechaDesde < fechaMin) fechaDesde = fechaMin
    if (fechaHasta > fechaMax) fechaHasta = fechaMax
  }
}

function aplicarRangoRapido(dias: number, rerender = true) {
  if (historico.length === 0) return

  const fechaMax = historico[historico.length - 1].fecha
  let nuevaDesde = restarDias(fechaMax, dias)

  if (nuevaDesde < historico[0].fecha) {
    nuevaDesde = historico[0].fecha
  }

  fechaDesde = nuevaDesde
  fechaHasta = fechaMax
  rangoRapidoActivo = dias

  if (rerender) render()
}

function obtenerHistoricoFiltrado() {
  if (!fechaDesde || !fechaHasta) return historico
  return historico.filter(item => item.fecha >= fechaDesde && item.fecha <= fechaHasta)
}

function obtenerRegistroCalculadora(fecha: string) {
  if (!historico.length) return null

  const exacto = historico.find(item => item.fecha === fecha)
  if (exacto) return exacto

  const anteriores = historico.filter(item => item.fecha <= fecha)
  if (anteriores.length > 0) {
    return anteriores[anteriores.length - 1]
  }

  return historico[0]
}

function formatearFecha(fechaISO: string) {
  if (!fechaISO) return ''
  const fecha = new Date(fechaISO + 'T00:00:00')
  return fecha.toLocaleDateString('es-VE', {
    day: 'numeric',
    month: 'short',
    year: 'numeric'
  })
}

function calcularVariacion(actual: number, anterior: number) {
  const delta = actual - anterior
  const porcentaje = anterior ? (delta / anterior) * 100 : 0
  const positiva = delta >= 0

  return {
    delta,
    porcentaje,
    positiva,
    icono: positiva ? '▲' : '▼'
  }
}

function infoTasa(tipo: 'dolar' | 'euro') {
  const actual = tipo === 'dolar' ? dolar : euro
  const anterior =
    tipo === 'dolar'
      ? (anteriorRegistro?.dolar ?? dolar)
      : (anteriorRegistro?.euro ?? euro)

  return calcularVariacion(actual, anterior)
}

function exportarHistoricoExcel() {
  const datosFiltrados = obtenerHistoricoFiltrado()

  if (!datosFiltrados.length) {
    alert('No hay datos para exportar en el rango seleccionado.')
    return
  }

  const filas = datosFiltrados.map(item => ({
    Fecha: item.fecha,
    Dolar: Number(item.dolar),
    Euro: Number(item.euro)
  }))

  const worksheet = XLSX.utils.json_to_sheet(filas)
  worksheet['!cols'] = [
    { wch: 14 },
    { wch: 12 },
    { wch: 12 }
  ]

  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Historico')

  const nombreArchivo = `tasas_historico_${fechaDesde}_a_${fechaHasta}.xlsx`
  XLSX.writeFile(workbook, nombreArchivo)
}

function render() {
  app.innerHTML = `
    <div class="app-shell ${menuAbierto ? 'menu-open' : ''}">
      <div class="bg-blur bg-blur-1"></div>
      <div class="bg-blur bg-blur-2"></div>

      <header class="topbar glass-panel">
        <div class="topbar-title">${tituloPagina()}</div>

        <button id="menuBtn" class="menu-btn" aria-label="Abrir menú">
          <span></span>
          <span></span>
          <span></span>
        </button>
      </header>

      <aside class="sidebar glass-panel">
        <div class="sidebar-head">
          <div class="sidebar-title">Menú</div>
          <button id="closeMenuBtn" class="sidebar-close" aria-label="Cerrar menú">✕</button>
        </div>

        <div class="sidebar-links">
          <button class="side-link ${pagina === 'inicio' ? 'active' : ''}" data-page="inicio">Inicio</button>
          <button class="side-link ${pagina === 'calculadora' ? 'active' : ''}" data-page="calculadora">Calculadora</button>
          <button class="side-link ${pagina === 'historico' ? 'active' : ''}" data-page="historico">Histórico</button>
        </div>
      </aside>

      <div id="overlay" class="overlay"></div>

      <main class="container">
        ${pagina === 'inicio' ? inicioHTML() : ''}
        ${pagina === 'calculadora' ? calculadoraHTML() : ''}
        ${pagina === 'historico' ? historicoHTML() : ''}
      </main>
    </div>
  `

  eventos()

  if (pagina === 'historico') {
    drawChart()
  }
}

function tarjetaTasa(simbolo: string, codigo: string, valor: number, tipo: 'dolar' | 'euro') {
  const variacion = infoTasa(tipo)

  return `
    <section class="rate-card glass-panel">
      <div class="rate-top">
        <div class="rate-left">
          <div class="rate-symbol">${simbolo}</div>
          <div>
            <div class="rate-code">${codigo}</div>
            <div class="rate-sub">BCV</div>
          </div>
        </div>

        <div class="delta-pill ${variacion.positiva ? 'up' : 'down'}">
          <span>${variacion.icono}</span>
          <span>${Math.abs(variacion.delta).toFixed(2)}</span>
          <span>${Math.abs(variacion.porcentaje).toFixed(2)}%</span>
        </div>
      </div>

      <div class="rate-value">${valor.toFixed(2)}</div>

      <div class="rate-bottom">
        <span>Fecha valor: ${formatearFecha(ultimoRegistro?.fecha || '')}</span>
        <span>vs. tasa anterior</span>
      </div>
    </section>
  `
}

function inicioHTML() {
  return `
    ${tarjetaTasa('$', 'USD', dolar, 'dolar')}
    ${tarjetaTasa('€', 'EUR', euro, 'euro')}

    <section class="card glass-panel converter-card">
      <h3>Convertir</h3>

      <div class="segmented">
        <button id="usd" class="btn ${tasa === 'dolar' ? 'active success' : ''}">USD</button>
        <button id="eur" class="btn ${tasa === 'euro' ? 'active primary' : ''}">EUR</button>
      </div>

      <input id="monto" placeholder="Monto en Bs" />
      <div id="resultado">Resultado: 0</div>
    </section>
  `
}

function calculadoraHTML() {
  const registroCalc = obtenerRegistroCalculadora(fechaCalc || ultimoRegistro?.fecha || '')
  const fechaMin = historico.length > 0 ? historico[0].fecha : ''
  const fechaMax = historico.length > 0 ? historico[historico.length - 1].fecha : ''

  const mensajeFecha =
    registroCalc && registroCalc.fecha !== fechaCalc
      ? `No hubo tasa exacta. Se muestra la última disponible: ${formatearFecha(registroCalc.fecha)}`
      : `Fecha valor: ${formatearFecha(registroCalc?.fecha || '')}`

  return `
    <section class="card glass-panel calc-card">
      <h3>Calculadora</h3>

      <div class="calc-date-box">
        <label for="fecha-calc">Elegir fecha</label>
        <input
          type="date"
          id="fecha-calc"
          value="${fechaCalc}"
          min="${fechaMin}"
          max="${fechaMax}"
        />
      </div>

      <div class="calc-rate-row">
        <div class="calc-rate-pill">
          <span class="calc-rate-label">$ USD</span>
          <strong>${registroCalc ? Number(registroCalc.dolar).toFixed(2) : '--'}</strong>
        </div>

        <div class="calc-rate-pill">
          <span class="calc-rate-label">€ EUR</span>
          <strong>${registroCalc ? Number(registroCalc.euro).toFixed(2) : '--'}</strong>
        </div>
      </div>

      <div class="calc-date-help">${mensajeFecha}</div>

      <div class="screen">${pantalla}</div>

      <div class="grid">
        ${[
          'C', '⌫', '/', '*',
          '7', '8', '9', '-',
          '4', '5', '6', '+',
          '1', '2', '3', '=',
          '0', '.'
        ]
          .map(b => `<button class="btn calc">${b}</button>`)
          .join('')}
      </div>
    </section>
  `
}

function historicoHTML() {
  const fechaMin = historico.length > 0 ? historico[0].fecha : ''
  const fechaMax = historico.length > 0 ? historico[historico.length - 1].fecha : ''

  return `
    <section class="card glass-panel chart-card">
      <div class="section-title-row">
        <h3>Histórico BCV</h3>
        <button id="exportarExcel" class="export-btn">Exportar Excel</button>
      </div>

      <div class="quick-range-box">
        <button class="quick-btn ${rangoRapidoActivo === 7 ? 'active' : ''}" data-range="7">7 días</button>
        <button class="quick-btn ${rangoRapidoActivo === 15 ? 'active' : ''}" data-range="15">15 días</button>
        <button class="quick-btn ${rangoRapidoActivo === 30 ? 'active' : ''}" data-range="30">30 días</button>
        <button class="quick-btn ${rangoRapidoActivo === 60 ? 'active' : ''}" data-range="60">60 días</button>
        <button class="quick-btn ${rangoRapidoActivo === 90 ? 'active' : ''}" data-range="90">90 días</button>
      </div>

      <div class="range-box">
        <div class="range-field">
          <label for="fecha-desde">Desde</label>
          <input
            type="date"
            id="fecha-desde"
            value="${fechaDesde}"
            min="${fechaMin}"
            max="${fechaMax}"
          />
        </div>

        <div class="range-field">
          <label for="fecha-hasta">Hasta</label>
          <input
            type="date"
            id="fecha-hasta"
            value="${fechaHasta}"
            min="${fechaMin}"
            max="${fechaMax}"
          />
        </div>
      </div>

      <div class="range-note">Máximo permitido: 90 días</div>

      <div class="chart-wrap">
        <canvas id="chart"></canvas>
      </div>
    </section>
  `
}

function eventos() {
  const menuBtn = document.getElementById('menuBtn')
  const closeMenuBtn = document.getElementById('closeMenuBtn')
  const overlay = document.getElementById('overlay')

  if (menuBtn) {
    menuBtn.addEventListener('click', () => {
      menuAbierto = true
      render()
    })
  }

  if (closeMenuBtn) {
    closeMenuBtn.addEventListener('click', () => {
      menuAbierto = false
      render()
    })
  }

  if (overlay) {
    overlay.addEventListener('click', () => {
      menuAbierto = false
      render()
    })
  }

  document.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', (e: Event) => {
      const target = e.currentTarget as HTMLElement
      const nuevaPagina = target.getAttribute('data-page') as 'inicio' | 'calculadora' | 'historico'

      if (nuevaPagina) {
        pagina = nuevaPagina
        menuAbierto = false
        render()
      }
    })
  })

  const input = document.getElementById('monto') as HTMLInputElement | null

  if (input) {
    input.oninput = () => {
      const valor = parseFloat(input.value) || 0
      const t = tasa === 'dolar' ? dolar : euro
      const resultado = t > 0 ? (valor / t).toFixed(2) : '0'

      const resultadoEl = document.getElementById('resultado')
      if (resultadoEl) {
        resultadoEl.innerText = 'Resultado: ' + resultado
      }
    }

    const usdBtn = document.getElementById('usd')
    const eurBtn = document.getElementById('eur')

    if (usdBtn) {
      usdBtn.onclick = () => {
        tasa = 'dolar'
        render()
      }
    }

    if (eurBtn) {
      eurBtn.onclick = () => {
        tasa = 'euro'
        render()
      }
    }
  }

  const fechaCalcInput = document.getElementById('fecha-calc') as HTMLInputElement | null
  if (fechaCalcInput) {
    fechaCalcInput.onchange = () => {
      fechaCalc = fechaCalcInput.value
      render()
    }
  }

  document.querySelectorAll('[data-range]').forEach(btn => {
    btn.addEventListener('click', (e: Event) => {
      const target = e.currentTarget as HTMLElement
      const valor = target.getAttribute('data-range')
      if (valor) aplicarRangoRapido(Number(valor))
    })
  })

  const desdeInput = document.getElementById('fecha-desde') as HTMLInputElement | null
  const hastaInput = document.getElementById('fecha-hasta') as HTMLInputElement | null

  if (desdeInput) {
    desdeInput.onchange = () => {
      fechaDesde = desdeInput.value
      rangoRapidoActivo = 0
      normalizarRango('desde')
      render()
    }
  }

  if (hastaInput) {
    hastaInput.onchange = () => {
      fechaHasta = hastaInput.value
      rangoRapidoActivo = 0
      normalizarRango('hasta')
      render()
    }
  }

  const exportBtn = document.getElementById('exportarExcel')
  if (exportBtn) {
    exportBtn.addEventListener('click', () => {
      exportarHistoricoExcel()
    })
  }

  document.querySelectorAll('.calc').forEach(btn => {
    btn.addEventListener('click', (e: Event) => {
      const target = e.currentTarget as HTMLElement
      const val = target.innerText

      if (val === 'C') {
        pantalla = '0'
      } else if (val === '⌫') {
        if (pantalla === 'Error' || pantalla.length <= 1) {
          pantalla = '0'
        } else {
          pantalla = pantalla.slice(0, -1)
        }
      } else if (val === '=') {
        try {
          pantalla = eval(pantalla).toString()
        } catch {
          pantalla = 'Error'
        }
      } else {
        if (pantalla === '0' || pantalla === 'Error') {
          pantalla = val
        } else {
          pantalla += val
        }
      }

      render()
    })
  })
}

function drawChart() {
  const canvas = document.getElementById('chart') as HTMLCanvasElement | null
  const ctx = canvas?.getContext('2d')

  if (!ctx) return

  const datosFiltrados = obtenerHistoricoFiltrado()

  if (chartInstance) {
    chartInstance.destroy()
  }

  chartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: datosFiltrados.map(d => d.fecha),
      datasets: [
        {
          label: 'Dólar',
          data: datosFiltrados.map(d => d.dolar),
          borderColor: '#38f0c3',
          backgroundColor: 'rgba(56,240,195,0.08)',
          borderWidth: 3,
          tension: 0.25,
          fill: true,
          pointRadius: 0,
          pointHoverRadius: 5,
          pointHitRadius: 12
        },
        {
          label: 'Euro',
          data: datosFiltrados.map(d => d.euro),
          borderColor: '#6da8ff',
          backgroundColor: 'rgba(109,168,255,0.06)',
          borderWidth: 3,
          tension: 0.25,
          fill: true,
          pointRadius: 0,
          pointHoverRadius: 5,
          pointHitRadius: 12
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false
      },
      plugins: {
        legend: {
          labels: {
            color: '#f5f7fa',
            boxWidth: 16,
            boxHeight: 10,
            padding: 14,
            font: {
              size: 12,
              weight: 'bold'
            }
          }
        },
        tooltip: {
          backgroundColor: 'rgba(12, 16, 24, 0.94)',
          titleColor: '#ffffff',
          bodyColor: '#ffffff',
          padding: 12,
          displayColors: true,
          callbacks: {
            title(items: any) {
              const fecha = items[0]?.label || ''
              const [y, m, d] = fecha.split('-')
              return `${d}/${m}/${y}`
            },
            label(context: any) {
              const valor = Number(context.raw).toFixed(2)
              return `${context.dataset.label}: ${valor}`
            }
          }
        }
      },
      scales: {
        x: {
          ticks: {
            color: 'rgba(255,255,255,0.78)',
            maxTicksLimit: 6,
            maxRotation: 0,
            minRotation: 0,
            callback(_value: any, index: number) {
              const fecha = datosFiltrados[index]?.fecha
              if (!fecha) return ''
              const [, m, d] = fecha.split('-')
              return `${d}/${m}`
            }
          },
          grid: { display: false },
          border: { display: false }
        },
        y: {
          ticks: {
            color: 'rgba(255,255,255,0.78)',
            callback(value: any) {
              return Number(value).toFixed(0)
            }
          },
          grid: { color: 'rgba(255,255,255,0.07)' },
          border: { display: false }
        }
      }
    }
  })
}

void bootApp()