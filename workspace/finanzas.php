<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Finanzas profesional', 'finanzas', $user);
?>
<section class="finance-hero-panel">
  <div class="finance-hero-copy">
    <span class="workspace-tag">Finanzas profesional</span>
    <h3>Control financiero de alto nivel con enfoque legal en Colombia</h3>
    <p class="workspace-muted">Consolida ingresos, gastos, inversiones, impuestos y proyecciones en un solo tablero ejecutivo.</p>
    <div class="finance-hero-kpis" id="financeHeroKpis">
      <span class="finance-kpi-pill"><i class="fas fa-piggy-bank"></i> Ahorro mensual: 0 COP</span>
      <span class="finance-kpi-pill"><i class="fas fa-wallet"></i> Flujo neto: 0 COP</span>
      <span class="finance-kpi-pill"><i class="fas fa-percent"></i> Tasa de ahorro: 0%</span>
    </div>
  </div>
  <div class="finance-hero-orbit" aria-hidden="true">
    <div class="finance-orbit-node finance-orbit-main"><i class="fas fa-chart-line"></i></div>
    <div class="finance-orbit-node finance-orbit-a"></div>
    <div class="finance-orbit-node finance-orbit-b"></div>
    <div class="finance-orbit-node finance-orbit-c"></div>
  </div>
</section>

<div class="workspace-cards finance-summary-cards" id="financeSummaryCards">
  <article class="workspace-stat summary-stat-card summary-stat-green">
    <span class="workspace-tag">Ingresos</span>
    <strong id="financeIncomeValue">0 COP</strong>
    <p class="workspace-muted">Entradas mensuales consolidadas</p>
  </article>
  <article class="workspace-stat summary-stat-card summary-stat-yellow">
    <span class="workspace-tag">Gastos</span>
    <strong id="financeExpenseValue">0 COP</strong>
    <p class="workspace-muted">Egresos y consumos del periodo</p>
  </article>
  <article class="workspace-stat summary-stat-card summary-stat-blue">
    <span class="workspace-tag">Inversiones</span>
    <strong id="financeInvestmentValue">0 COP</strong>
    <p class="workspace-muted">Capital destinado a crecimiento</p>
  </article>
  <article class="workspace-stat summary-stat-card summary-stat-yellow">
    <span class="workspace-tag">Impuestos</span>
    <strong id="financeTaxValue">0 COP</strong>
    <p class="workspace-muted">Proyección tributaria Colombia</p>
  </article>
  <article class="workspace-stat summary-stat-card summary-stat-green">
    <span class="workspace-tag">Patrimonio</span>
    <strong id="financeNetWorthValue">0 COP</strong>
    <p class="workspace-muted">Activos menos pasivos</p>
  </article>
</div>

<div class="workspace-grid finance-grid-layout">
  <article class="workspace-panel wide finance-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Flujo de caja</span>
        <h3>Ingreso y control de gastos</h3>
      </div>
      <span class="workspace-tag">Operacion diaria</span>
    </div>

    <form id="financeCashflowForm" class="workspace-form">
      <div class="workspace-form-grid two">
        <div>
          <label for="financeType" class="form-label">Tipo de movimiento</label>
          <select id="financeType" class="form-select">
            <option value="income">Ingreso</option>
            <option value="expense">Gasto</option>
            <option value="investment">Inversion</option>
            <option value="tax">Impuesto</option>
          </select>
        </div>
        <div>
          <label for="financeAmount" class="form-label">Monto (COP)</label>
          <input type="number" class="form-control" id="financeAmount" min="0" step="1000" placeholder="Ej: 2500000">
        </div>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="financeCategory" class="form-label">Categoria</label>
          <input type="text" class="form-control" id="financeCategory" placeholder="Nomina, arriendo, alimentos, transporte...">
        </div>
        <div>
          <label for="financeDate" class="form-label">Fecha</label>
          <input type="date" class="form-control" id="financeDate">
        </div>
      </div>
      <div>
        <label for="financeDescription" class="form-label">Detalle</label>
        <textarea id="financeDescription" class="form-control" rows="3" placeholder="Describe el contexto del movimiento para trazabilidad financiera."></textarea>
      </div>
      <button type="submit" class="btn btn-primary"><i class="fas fa-plus-circle"></i> Registrar movimiento</button>
    </form>

    <div class="workspace-list finance-list" id="financeMovementsList"></div>
  </article>

  <article class="workspace-panel tall finance-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Asistente</span>
        <h3>Gastos y ahorro inteligente</h3>
      </div>
      <span class="workspace-tag">Recomendaciones</span>
    </div>

    <div class="finance-assistant-box" id="financeAssistantBox">
      <h4><i class="fas fa-lightbulb"></i> Sugerencias automaticas</h4>
      <ul>
        <li>Define un presupuesto 50/30/20 para mantener disciplina financiera.</li>
        <li>Activa una meta de ahorro con transferencia automatica semanal.</li>
        <li>Prioriza pago de deudas de alta tasa antes de nuevas inversiones.</li>
        <li>Separa provisiones para seguridad social y retencion aplicable.</li>
      </ul>
    </div>

    <div class="finance-tax-guide">
      <h4><i class="fas fa-scale-balanced"></i> Checklist legal Colombia</h4>
      <p class="workspace-muted">Guia orientativa para control tributario; valida siempre con contador o asesor fiscal.</p>
      <ul>
        <li>Revisa topes de declaracion de renta (persona natural).</li>
        <li>Controla facturacion/evidencias de costos y deducciones.</li>
        <li>Programa calendario DIAN con fechas limite y anticipos.</li>
        <li>Mantener soporte de aportes seguridad social y parafiscales.</li>
      </ul>
    </div>
  </article>

  <article class="workspace-panel wide finance-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Calculadoras</span>
        <h3>Simuladores financieros</h3>
      </div>
      <span class="workspace-tag">Planeacion</span>
    </div>

    <div class="finance-calculators-grid">
      <div class="finance-calc-card">
        <h4><i class="fas fa-coins"></i> Interes compuesto</h4>
        <p class="workspace-muted">Proyecta crecimiento de capital por tasa y periodos.</p>
        <div class="finance-calc-form-grid">
          <input type="number" id="calcCompoundInitial" class="form-control" min="0" step="10000" placeholder="Capital inicial (COP)">
          <input type="number" id="calcCompoundMonthly" class="form-control" min="0" step="10000" placeholder="Aporte mensual (COP)">
          <input type="number" id="calcCompoundRate" class="form-control" min="0" step="0.1" placeholder="Tasa anual (%)">
          <input type="number" id="calcCompoundYears" class="form-control" min="1" step="1" placeholder="Anos">
        </div>
        <div class="finance-calc-actions">
          <button type="button" class="btn btn-primary finance-calc-btn" id="calcCompoundBtn">Calcular</button>
          <button type="button" class="btn finance-calc-export-btn" id="calcCompoundExportBtn"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        </div>
        <div class="finance-calc-result" id="calcCompoundResult">Completa los campos para calcular proyeccion.</div>
      </div>

      <div class="finance-calc-card">
        <h4><i class="fas fa-house"></i> Capacidad de vivienda</h4>
        <p class="workspace-muted">Estima cuota sostenible segun ingreso y endeudamiento.</p>
        <div class="finance-calc-form-grid">
          <input type="number" id="calcHomeIncome" class="form-control" min="0" step="10000" placeholder="Ingreso mensual (COP)">
          <input type="number" id="calcHomeDebts" class="form-control" min="0" step="10000" placeholder="Deudas mensuales (COP)">
          <input type="number" id="calcHomeRate" class="form-control" min="0" step="0.1" placeholder="Tasa EA (%)">
          <input type="number" id="calcHomeYears" class="form-control" min="1" step="1" placeholder="Plazo (anos)">
          <input type="number" id="calcHomeQuotaPct" class="form-control" min="5" max="60" step="1" placeholder="Tope cuota (%)">
          <input type="number" id="calcHomeDownPct" class="form-control" min="0" max="80" step="1" placeholder="Cuota inicial (%)">
        </div>
        <div class="finance-calc-actions">
          <button type="button" class="btn btn-primary finance-calc-btn" id="calcHomeBtn">Calcular</button>
          <button type="button" class="btn finance-calc-export-btn" id="calcHomeExportBtn"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        </div>
        <div class="finance-calc-result" id="calcHomeResult">Ingresa tus datos para estimar capacidad de compra.</div>
      </div>

      <div class="finance-calc-card">
        <h4><i class="fas fa-car"></i> Leasing y credito</h4>
        <p class="workspace-muted">Compara costo financiero total entre alternativas.</p>
        <div class="finance-calc-form-grid">
          <input type="number" id="calcLoanAmount" class="form-control" min="0" step="10000" placeholder="Monto financiado (COP)">
          <input type="number" id="calcLoanMonths" class="form-control" min="1" step="1" placeholder="Plazo (meses)">
          <input type="number" id="calcLoanRate" class="form-control" min="0" step="0.1" placeholder="Tasa credito EA (%)">
          <input type="number" id="calcLeaseRate" class="form-control" min="0" step="0.1" placeholder="Tasa leasing EA (%)">
          <input type="number" id="calcLeaseResidual" class="form-control" min="0" max="70" step="1" placeholder="Residual leasing (%)">
        </div>
        <div class="finance-calc-actions">
          <button type="button" class="btn btn-primary finance-calc-btn" id="calcLeaseBtn">Calcular</button>
          <button type="button" class="btn finance-calc-export-btn" id="calcLeaseExportBtn"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        </div>
        <div class="finance-calc-result" id="calcLeaseResult">Compara cuota y costo total para decidir mejor.</div>
      </div>

      <div class="finance-calc-card">
        <h4><i class="fas fa-file-invoice-dollar"></i> Carga tributaria</h4>
        <p class="workspace-muted">Aproxima impacto de impuestos y retenciones.</p>
        <div class="finance-calc-form-grid">
          <input type="number" id="calcTaxIncome" class="form-control" min="0" step="10000" placeholder="Ingreso mensual (COP)">
          <input type="number" id="calcTaxCosts" class="form-control" min="0" step="10000" placeholder="Costos deducibles (COP)">
          <input type="number" id="calcTaxRete" class="form-control" min="0" step="0.1" placeholder="Retencion (%)">
          <input type="number" id="calcTaxSecurity" class="form-control" min="0" step="0.1" placeholder="Seg. social sobre IBC (%)">
        </div>
        <div class="finance-calc-actions">
          <button type="button" class="btn btn-primary finance-calc-btn" id="calcTaxBtn">Calcular</button>
          <button type="button" class="btn finance-calc-export-btn" id="calcTaxExportBtn"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        </div>
        <div class="finance-calc-result" id="calcTaxResult">Resultado orientativo; confirma con tu contador.</div>
      </div>
    </div>

    <div class="workspace-panel-header finance-history-head">
      <div>
        <span class="eyebrow">Historial</span>
        <h3>Simulaciones guardadas</h3>
      </div>
      <span class="workspace-tag">Auditoria</span>
    </div>

    <div class="finance-history-controls">
      <select id="financeSimTypeFilter" class="form-select">
        <option value="">Todos los simuladores</option>
        <option value="compound_interest">Interes compuesto</option>
        <option value="home_capacity">Capacidad de vivienda</option>
        <option value="leasing_vs_credit">Leasing y credito</option>
        <option value="tax_burden">Carga tributaria</option>
      </select>
      <input type="date" id="financeSimDateFrom" class="form-control" placeholder="Desde">
      <input type="date" id="financeSimDateTo" class="form-control" placeholder="Hasta">
      <button type="button" class="btn" id="financeSimApplyFilterBtn">Filtrar</button>
      <button type="button" class="btn" id="financeSimClearFilterBtn">Limpiar</button>
    </div>

    <div class="finance-history-purge-row">
      <select id="financeSimPurgeDays" class="form-select">
        <option value="30">Eliminar mayores a 30 dias</option>
        <option value="90" selected>Eliminar mayores a 90 dias</option>
        <option value="180">Eliminar mayores a 180 dias</option>
        <option value="365">Eliminar mayores a 365 dias</option>
      </select>
      <button type="button" class="btn finance-danger-btn" id="financeSimPurgeBtn"><i class="fas fa-trash"></i> Limpiar antiguas</button>
    </div>

    <div class="workspace-list finance-list" id="financeSimulationHistoryList">
      <div class="workspace-empty">Aun no hay simulaciones guardadas. Ejecuta una calculadora para crear historial.</div>
    </div>

    <div class="finance-history-pagination" id="financeHistoryPagination">
      <select id="financeSimPageSize" class="form-select">
        <option value="10">10 por pagina</option>
        <option value="20" selected>20 por pagina</option>
        <option value="50">50 por pagina</option>
      </select>
      <button type="button" class="btn" id="financeSimPrevPageBtn">Anterior</button>
      <span id="financeSimPageInfo" class="workspace-muted">Pagina 1 de 1</span>
      <button type="button" class="btn" id="financeSimNextPageBtn">Siguiente</button>
    </div>
  </article>

  <article class="workspace-panel tall finance-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Inversiones</span>
        <h3>Portafolio y riesgo</h3>
      </div>
      <span class="workspace-tag">Estrategia</span>
    </div>

    <div class="workspace-list finance-investment-list">
      <div class="workspace-list-item">
        <strong>Distribucion sugerida</strong>
        <p>Liquidez 20% · Renta fija 35% · Renta variable 25% · Alternativos 20%</p>
      </div>
      <div class="workspace-list-item">
        <strong>Control de riesgo</strong>
        <p>Limite maximo 15% por activo para evitar concentracion excesiva.</p>
      </div>
      <div class="workspace-list-item">
        <strong>Rebalanceo</strong>
        <p>Revisa desviaciones y ajusta cada trimestre segun objetivos.</p>
      </div>
    </div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>
