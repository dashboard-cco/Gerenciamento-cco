/*
========================================================
INTRANET EXECUTIVA KPI CCO • SLU
- Importação de planilhas
- Dados permanentes no Supabase
- Histórico de importações
- Controle de duplicidade
- Filtro diário e mensal
- Espelho da planilha
- Data e hora no padrão brasileiro
========================================================
*/

/* =====================================================
   CONFIGURAÇÃO DO SUPABASE
===================================================== */

const SUPABASE_URL =
  "https://rotzysfvtcnezvsguanv.supabase.co";

const SUPABASE_KEY =
  "sb_publishable_qQh8LPJbvvFextBIpEbqAA_FfeaYGb0";

const banco =
  supabase.createClient(
    SUPABASE_URL,
    SUPABASE_KEY
  );

/* =====================================================
   VALORES FIXOS DOS SERVIÇOS
===================================================== */

const VALORES_FIXOS = {
  "P1": 296.00,
  "P2.1": 1027.42,
  "P2.2": 1027.42,
  "P3": 41992.93,
  "P4": 68.80,
  "P5": 160.94,
  "P6": 76.24,
  "P7": 49811.72,
  "P8": 81001.04,
  "P9": 122039.23,
  "P10": 346660.01,
  "P11": 272459.08,
  "P12": 0.83
};

const SERVICOS_FIXOS =
  ["P3", "P7", "P8", "P9", "P10"];

/* =====================================================
   VARIÁVEIS GLOBAIS
===================================================== */

let painelExecutivo = [];
let painelExecutivoOriginal = [];

let operacoes = [];
let operacoesOriginal = [];

let sheetsOriginais = {};
let todasAsAbas = [];

let graficoExecucao = null;
let graficoPizza = null;

/* =====================================================
   INICIALIZAÇÃO DO SISTEMA
===================================================== */

document.addEventListener("DOMContentLoaded", async () => {
  atualizarData();

  await carregarBanco();

  await carregarHistorico();

  const input =
    document.getElementById("arquivoExcel");

  if (input) {
    input.addEventListener(
      "change",
      importarPlanilhas
    );
  }

  mostrarKPI("geral");
});

/* =====================================================
   CONTROLE DE ACESSO
===================================================== */

document.addEventListener("DOMContentLoaded", () => {
  const usuarioLogado =
    JSON.parse(
      localStorage.getItem("usuarioLogado")
    );

  if (!usuarioLogado) {
    window.location.href = "login.html";
    return;
  }

  if (usuarioLogado.perfil === "Diretoria") {
    esconderElemento("btnImportar");
    esconderElemento("btnLimpar");
    esconderElemento("btnExportar");
    esconderElemento("btnImplantar");
    esconderElemento("menuKpi");
    esconderElemento("paginaKpi");

    bloquearBotoesDiretoria();
  }
});

function esconderElemento(id) {
  const elemento =
    document.getElementById(id);

  if (elemento) {
    elemento.style.display = "none";
  }
}

function bloquearBotoesDiretoria() {
  [
    "btnImportar",
    "btnLimpar",
    "btnExportar",
    "btnImplantar"
  ].forEach(id => {
    const botao =
      document.getElementById(id);

    if (botao) {
      botao.disabled = true;
      botao.style.display = "none";
    }
  });
}

/* =====================================================
   DATA DO CABEÇALHO
===================================================== */

function atualizarData() {
  const dataAtual =
    document.getElementById("dataAtual");

  if (!dataAtual) return;

  dataAtual.innerText =
    new Date().toLocaleDateString("pt-BR", {
      weekday: "long",
      day: "2-digit",
      month: "long",
      year: "numeric"
    });
}

/* =====================================================
   IMPORTAÇÃO DE PLANILHAS
===================================================== */

async function importarPlanilhas(evento) {
  const arquivos =
    Array.from(evento.target.files || []);

  if (!arquivos.length) return;

  mostrarLoading(true);

  try {
    painelExecutivo = [];
    painelExecutivoOriginal = [];
    operacoes = [];
    operacoesOriginal = [];
    sheetsOriginais = {};
    todasAsAbas = [];

    for (const arquivo of arquivos) {
      const buffer =
        await arquivo.arrayBuffer();

      const workbook =
        XLSX.read(buffer, {
          type: "array",
          cellDates: false
        });

      workbook.SheetNames.forEach(nomeAba => {
        const sheet =
          workbook.Sheets[nomeAba];

        const dados =
          XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: true,
            dateNF: "dd/mm/yyyy hh:mm"
          });

        const dadosFormatados =
          dados.map(linha =>
            formatarLinhaEspelho(linha)
          );

        const normalizados =
          dadosFormatados.map(item =>
            normalizarObjeto(item)
          );

        const nomeNormalizado =
          normalizar(nomeAba);

        sheetsOriginais[nomeNormalizado] = {
          nomeOriginal: nomeAba,
          codigoServico: extrairCodigo(nomeAba),
          dadosOriginais: dadosFormatados,
          dadosNormalizados: normalizados
        };

        todasAsAbas.push({
          arquivo: arquivo.name,
          aba: nomeAba
        });
      });
    }
    gerarPainelExecutivo();
    gerarOperacoes();

    painelExecutivoOriginal =
      [...painelExecutivo];

    operacoesOriginal =
      [...operacoes];
    atualizarDashboard();
    renderResumoAutomaticoDiretoria();

    const salvou =
      await salvarPlanilhaNoBanco(
        arquivos.map(a => a.name).join(", ")
      );

    await carregarHistorico();

    preencherTexto(
      "nomeArquivo",
      `${arquivos.length} arquivo(s) importado(s) | ${todasAsAbas.length} aba(s) lida(s)`
    );

    if (salvou) {
      alert("Planilha importada e salva com sucesso!");
    }

  } catch (erro) {
    console.error(erro);
    alert("Erro ao importar planilha.");
  } finally {
    mostrarLoading(false);
    evento.target.value = "";
  }
}

/* =====================================================
   FORMATAÇÃO DE LINHAS E DATAS
===================================================== */

function formatarLinhaEspelho(linha) {
  const novaLinha = {};

  Object.keys(linha).forEach(coluna => {
    const valor =
      linha[coluna];

    if (ehCampoDataHora(coluna)) {
      novaLinha[coluna] =
        formatarDataHoraBR(valor);

    } else if (ehCampoHora(coluna)) {
      novaLinha[coluna] =
        formatarHoraBR(valor);

    } else {
      novaLinha[coluna] =
        valor;
    }
  });

  return novaLinha;
}

function ehCampoDataHora(coluna) {
  const nome =
    normalizar(coluna);

  return (
    nome.includes("data") ||
    nome.includes("inicio") ||
    nome.includes("fim") ||
    nome.includes("termino")
  );
}

function ehCampoHora(coluna) {
  const nome =
    normalizar(coluna);

  return (
    nome.includes("hora") ||
    nome.includes("tempo")
  );
}

function formatarDataHoraBR(valor) {
  if (!valor) return "";

  if (typeof valor === "number") {
    const data =
      XLSX.SSF.parse_date_code(valor);

    if (!data) return valor;

    const dia =
      String(data.d).padStart(2, "0");

    const mes =
      String(data.m).padStart(2, "0");

    const ano =
      data.y;

    const hora =
      String(data.H || 0).padStart(2, "0");

    const minuto =
      String(data.M || 0).padStart(2, "0");

    if (hora === "00" && minuto === "00") {
      return `${dia}/${mes}/${ano}`;
    }

    return `${dia}/${mes}/${ano} ${hora}:${minuto}`;
  }

  if (
    /^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(
      String(valor)
    )
  ) {
    return valor;
  }

  const data =
    new Date(valor);

  if (isNaN(data)) return valor;

  return data.toLocaleString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatarHoraBR(valor) {
  if (!valor) return "";

  if (typeof valor === "number") {
    const totalMinutos =
      Math.round(valor * 24 * 60);

    const horas =
      Math.floor(totalMinutos / 60);

    const minutos =
      totalMinutos % 60;

    return `${String(horas).padStart(2, "0")}:${String(minutos).padStart(2, "0")}`;
  }

  return valor;
}

/* =====================================================
   PAINEL EXECUTIVO
===================================================== */

function gerarPainelExecutivo() {
  const painel =
    sheetsOriginais["painel executivo"];

  if (!painel) {
    alert("Aba Painel Executivo não encontrada.");
    return;
  }

  painelExecutivo =
    painel.dadosNormalizados.map(item => {
      const servico =
        String(
          item.servico ||
          item.programa ||
          item.codigo ||
          ""
        ).toUpperCase();

      const abaServico =
        buscarAbaServico(servico);

      const acumulado =
        abaServico
          ? calcularAcumulado(
              item.medicao,
              abaServico.dadosNormalizados
            )
          : numero(
              item.acumulado_mes ||
              item.acumulado_no_mes ||
              item.executado
            );

      const previsto =
        numero(
          item.previsto_mes ||
          item.meta_mes ||
          item.previsto ||
          item.meta
        );

      const valorUnitario =
        VALORES_FIXOS[servico] || 0;

      const valorFinal =
        SERVICOS_FIXOS.includes(servico)
          ? valorUnitario
          : valorUnitario * acumulado;

      return {
        servico,
        nome_servico:
          item.nome_servico ||
          item.nome_do_servico ||
          item.descricao ||
          "",

        acumulado_mes:
          acumulado,

        semana_1:
          numero(
            item.semana_1 ||
            item["1a_semana_8d"] ||
            item["1o_semana_8d"]
          ),

        semana_2:
          numero(
            item.semana_2 ||
            item["2a_semana_6d"] ||
            item["2o_semana_6d"]
          ),

        semana_3:
          numero(
            item.semana_3 ||
            item["3a_semana_6d"] ||
            item["3o_semana_6d"]
          ),

        semana_4:
          numero(
            item.semana_4 ||
            item["4a_semana_6d"] ||
            item["4o_semana_6d"]
          ),

        medicao:
          item.medicao || "",

        previsto_mes:
          previsto,

        porcentagem_execucao:
          calcularPercentual(
            acumulado,
            previsto
          ),

        valor:
          valorFinal,

        status:
          acumulado > 0
            ? "Com dados"
            : "Sem dados"
      };
    });
}

/* =====================================================
   BASE OPERACIONAL
===================================================== */

function gerarOperacoes() {
  /*
    Monta a base operacional do sistema.

    Aqui o sistema:
    - lê todas as abas P1 a P12
    - identifica datas em várias colunas possíveis
    - cria data_normalizada no formato yyyy-mm-dd
    - prepara os filtros diário e mensal
  */

  operacoes = [];

  Object.keys(sheetsOriginais).forEach(nome => {
    if (nome === "painel executivo") return;

    const sheet =
      sheetsOriginais[nome];

    if (!sheet.codigoServico) return;

    sheet.dadosNormalizados.forEach(item => {
      const dataEncontrada =
        item.data ||
        item.data_operacao ||
        item.inicio_operacao ||
        item.inicio_da_operacao ||
        item.incio_operacao ||
        item.inicio_de_operacao ||
        item.inicio_da_operação ||
        item.início_operação ||
        item.fim_operacao ||
        item.fim_da_operacao ||
        item.data_inicio ||
        item.data_fim ||
        item.inicio ||
        item.dia ||
        item.dt ||
        "";

      operacoes.push({
        servico:
          sheet.codigoServico,

        origem:
          sheet.nomeOriginal,

        data:
          dataEncontrada,

        data_normalizada:
          normalizarData(
            dataEncontrada
          ),

        turno:
          item.turno || "",

        ra:
          item.ra ||
          item.regiao_administrativa ||
          item.regiao ||
          "Por demanda",

        peso:
          numero(
            item.peso ||
            item.peso_total ||
            item.peso_t ||
            item.tonelada ||
            item.toneladas
          ),

        viagens:
          numero(
            item.viagens ||
            item.qtd_viagem ||
            item.qtd_viagens ||
            item.quantidade_de_viagens
          ),

        km:
          numero(
            item.km ||
            item.km_total ||
            item.km_executado ||
            item.quilometragem
          ),

        equipe:
          numero(
            item.equipe ||
            item.qtd_equipe ||
            item.qdt_equipe ||
            item.equipes
          ),

        status:
          "Com dados"
      });
    });
  });

  console.log("Operações carregadas:", operacoes);
}

/* =====================================================
   ATUALIZAÇÃO VISUAL
===================================================== */
function atualizarDashboard() {
  renderCards();
  renderTabelaExecutiva();
  renderTabelaContratual();
  renderResumo();
  renderResumoAutomaticoDiretoria();
  renderRankingOperacional();
  renderKpiSemanal();
  renderComparativoOperacional();
  renderAlertas();
  renderTabelaDados();
  renderFiltros();
  renderGraficos();
}

function renderCards() {
  const servicosComDados =
    painelExecutivo.filter(
      item => item.status === "Com dados"
    ).length;

  const mediaExecucao =
    painelExecutivo.length
      ? painelExecutivo.reduce(
          (soma, item) =>
            soma + numero(item.porcentagem_execucao),
          0
        ) / painelExecutivo.length
      : 0;

  preencherTexto(
    "kpiServicosDados",
    servicosComDados
  );

  preencherTexto(
    "kpiExecucaoMedia",
    `${formatarNumero(mediaExecucao)}%`
  );

  preencherTexto(
    "kpiAbas",
    todasAsAbas.length
  );
}

function renderTabelaExecutiva() {
  const tabela =
    document.getElementById("tabelaPainelExecutivo");

  if (!tabela) return;

  if (!painelExecutivo.length) {
    tabela.innerHTML =
      `<tr><td colspan="12">Importe uma planilha para visualizar os dados.</td></tr>`;
    return;
  }

  tabela.innerHTML =
    painelExecutivo.map(item => `
      <tr>
        <td><strong>${item.servico}</strong></td>
        <td>${item.nome_servico}</td>
        <td>${formatarNumero(item.acumulado_mes)}</td>
        <td>${formatarNumero(item.semana_1)}</td>
        <td>${formatarNumero(item.semana_2)}</td>
        <td>${formatarNumero(item.semana_3)}</td>
        <td>${formatarNumero(item.semana_4)}</td>
        <td>${item.medicao}</td>
        <td>${formatarNumero(item.previsto_mes)}</td>
        <td>${formatarNumero(item.porcentagem_execucao)}%</td>
        <td>${formatarMoeda(item.valor)}</td>
        <td>
          <span class="badge ${item.status === "Com dados" ? "ok" : "info"}">
            ${item.status}
          </span>
        </td>
      </tr>
    `).join("");
}

function renderTabelaContratual() {
  const tabela =
    document.getElementById("tabelaContratual");

  if (!tabela) return;

  if (!painelExecutivo.length) {
    tabela.innerHTML =
      `<tr><td colspan="6">Nenhum dado contratual importado.</td></tr>`;
    return;
  }

  tabela.innerHTML =
    painelExecutivo.map(item => `
      <tr>
        <td>${item.servico} - ${item.nome_servico}</td>
        <td>${item.medicao}</td>
        <td>${formatarNumero(item.previsto_mes)}</td>
        <td>${formatarNumero(item.acumulado_mes)}</td>
        <td>${formatarNumero(item.porcentagem_execucao)}%</td>
        <td>
          <span class="badge ${item.status === "Com dados" ? "ok" : "info"}">
            ${item.status}
          </span>
        </td>
      </tr>
    `).join("");
}

function renderTabelaDados() {

  const tabela =
    document.getElementById("tabelaDados");

  if (!tabela) return;

  // BUSCA TEXTO
  const busca =
    normalizar(
      document.getElementById("busca")?.value || ""
    );

  // FILTRO SERVIÇO
  const filtroPrograma =
    document.getElementById("filtroPrograma")?.value || "Todos";

  // FILTRO STATUS
  const filtroStatus =
    document.getElementById("filtroStatus")?.value || "Todos";

  // FILTRO DATA
  const filtroData =
    document.getElementById("filtroDataBase")?.value || "";

  /*
  =====================================================
  FILTRAR DADOS
  =====================================================
  */

  const filtrados =
    operacoes.filter(item => {

      const texto =
        normalizar(
          Object.values(item).join(" ")
        );

      // FILTRO POR DATA
      const passouData =
        !filtroData ||
        item.data_normalizada === filtroData;

      return (
        texto.includes(busca) &&

        (
          filtroPrograma === "Todos" ||
          item.servico === filtroPrograma
        ) &&

        (
          filtroStatus === "Todos" ||
          item.status === filtroStatus
        ) &&

        passouData
      );
    });

  /*
  =====================================================
  MONTAR TABELA
  =====================================================
  */

  tabela.innerHTML =
    filtrados.map(item => `
      <tr>
        <td>${item.servico}</td>
        <td>${item.origem}</td>
        <td>${item.data}</td>
        <td>${item.turno}</td>
        <td>${item.ra}</td>
        <td>${formatarNumero(item.peso)}</td>
        <td>${formatarNumero(item.viagens)}</td>
        <td>${formatarNumero(item.km)}</td>
        <td>${formatarNumero(item.equipe)}</td>
      </tr>
    `).join("") ||

    `
      <tr>
        <td colspan="9">
          Nenhum dado encontrado.
        </td>
      </tr>
    `;
}

function renderDetalheServico(codigo) {
  const detalhe =
    document.getElementById("detalheServico");

  const aba =
    buscarAbaServico(codigo);

  if (!aba) {
    detalhe.innerHTML =
      `<div class="not-found">Serviço não encontrado.</div>`;
    return;
  }

  const dadosPainel =
    painelExecutivo.find(
      item => item.servico === codigo
    );

  const previsto =
    dadosPainel ? numero(dadosPainel.previsto_mes) : 0;

  const executado =
    dadosPainel ? numero(dadosPainel.acumulado_mes) : 0;

  const percentual =
    dadosPainel ? numero(dadosPainel.porcentagem_execucao) : 0;

  const totalKm =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.km ||
          item.km_total ||
          item.km_executado
        ),
      0
    );

  const totalEquipes =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.equipe ||
          item.qtd_equipe ||
          item.qdt_equipe
        ),
      0
    );

  const totalViagens =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.viagens ||
          item.qtd_viagem ||
          item.qtd_viagens
        ),
      0
    );

  const totalPeso =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.peso ||
          item.peso_total ||
          item.peso_t
        ),
      0
    );

  const colunas =
    aba.dadosOriginais.length
      ? Object.keys(aba.dadosOriginais[0])
      : [];

  const thead =
    colunas.map(col =>
      `<th>${col}</th>`
    ).join("");

  const tbody =
    aba.dadosOriginais.map(linha => `
      <tr>
        ${colunas.map(col => `
          <td>
            ${
              linha[col] !== undefined &&
              linha[col] !== null
                ? linha[col]
                : ""
            }
          </td>
        `).join("")}
      </tr>
    `).join("");

  detalhe.innerHTML = `
    <section class="cards">
      <div class="card">
        <span>Previsto</span>
        <strong>${formatarNumero(previsto)}</strong>
        <small>meta operacional</small>
      </div>

      <div class="card">
        <span>Executado</span>
        <strong>${formatarNumero(executado)}</strong>
        <small>acumulado realizado</small>
      </div>

      <div class="card">
        <span>% Execução</span>
        <strong>${formatarNumero(percentual)}%</strong>
        <small>índice operacional</small>
      </div>

      <div class="card">
        <span>Peso</span>
        <strong>${formatarNumero(totalPeso)}</strong>
        <small>toneladas</small>
      </div>

      <div class="card">
        <span>Viagens</span>
        <strong>${formatarNumero(totalViagens)}</strong>
        <small>operações</small>
      </div>

      <div class="card">
        <span>KM Executado</span>
        <strong>${formatarNumero(totalKm)}</strong>
        <small>quilometragem</small>
      </div>

      <div class="card">
        <span>Equipes</span>
        <strong>${formatarNumero(totalEquipes)}</strong>
        <small>equipes operacionais</small>
      </div>
    </section>

    <section class="section">
      <div class="section-title">
        <span>Espelho da planilha</span>
        <h2>${codigo}</h2>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>${thead}</tr>
          </thead>
          <tbody>
            ${
              tbody ||
              `<tr>
                <td colspan="20">
                  Nenhuma informação encontrada.
                </td>
              </tr>`
            }
          </tbody>
        </table>
      </div>
    </section>
  `;
}
function limparFiltroBase() {

  // Limpa busca
  const busca =
    document.getElementById("busca");

  if (busca) {
    busca.value = "";
  }

  // Limpa data
  const data =
    document.getElementById("filtroDataBase");

  if (data) {
    data.value = "";
  }

  // Limpa serviço
  const programa =
    document.getElementById("filtroPrograma");

  if (programa) {
    programa.value = "Todos";
  }

  // Limpa status
  const status =
    document.getElementById("filtroStatus");

  if (status) {
    status.value = "Todos";
  }

  // Atualiza tabela
  renderTabelaDados();
}

function renderResumo() {
  const resumoTexto =
  gerarResumoExecutivoAutomatico();
  const resumo =
    document.getElementById("resumoExecutivo");

  if (!resumo) return;

  const top =
    [...painelExecutivo]
      .filter(item => item.acumulado_mes > 0)
      .sort((a, b) => b.acumulado_mes - a.acumulado_mes)
      .slice(0, 6);

  if (!top.length) {
    resumo.innerHTML =
      "<p>Importe dados para gerar o resumo executivo.</p>";
    return;
  }

  const maior =
    Math.max(
      ...top.map(item => item.acumulado_mes),
      1
    );

  resumo.innerHTML =
    top.map(item => `
      <div class="metric-row">
        <strong>${item.servico}</strong>
        <div class="bar-bg">
          <div
            class="bar"
            style="width:${Math.min(
              100,
              item.acumulado_mes / maior * 100
            )}%">
          </div>
        </div>
        <b>${formatarNumero(item.acumulado_mes)}</b>
      </div>
    `).join("");
}

function renderAlertas() {
  const alertas =
    document.getElementById("alertasExecutivos");

  if (!alertas) return;

  if (!painelExecutivo.length) {
    alertas.innerHTML =
      `<p><span class="badge info">Aguardando</span> Importe uma planilha.</p>`;
    return;
  }

  const semDados =
    painelExecutivo.filter(
      item => item.status === "Sem dados"
    );

  if (!semDados.length) {
    alertas.innerHTML =
      `<p><span class="badge ok">Normal</span> Sem alertas críticos.</p>`;
    return;
  }

  alertas.innerHTML =
    `<p><span class="badge critico">Atenção</span> ${semDados.length} serviço(s) sem dados.</p>`;
}


function renderFiltros() {
  /*
    Atualiza os filtros da tela.
    Agora o filtro de mês mostra somente meses
    que realmente existem na base importada.
  */

  const filtroPrograma =
    document.getElementById("filtroPrograma");

  const filtroStatus =
    document.getElementById("filtroStatus");

  const filtroMes =
    document.getElementById("filtroMes");

  const filtroAno =
    document.getElementById("filtroAno");

  /*
    Filtros da Base Importada
  */
  if (filtroPrograma && filtroStatus) {
    const programaSelecionado =
      filtroPrograma.value || "Todos";

    const statusSelecionado =
      filtroStatus.value || "Todos";

    const programas =
      [
        "Todos",
        ...new Set(
          operacoesOriginal
            .map(item => item.servico)
            .filter(Boolean)
        )
      ];

    filtroPrograma.innerHTML =
      programas.map(item =>
        `<option ${item === programaSelecionado ? "selected" : ""}>${item}</option>`
      ).join("");

    filtroStatus.innerHTML = `
      <option ${statusSelecionado === "Todos" ? "selected" : ""}>Todos</option>
      <option ${statusSelecionado === "Com dados" ? "selected" : ""}>Com dados</option>
      <option ${statusSelecionado === "Sem dados" ? "selected" : ""}>Sem dados</option>
    `;
  }

  /*
    Filtro Mensal:
    mostra apenas meses que possuem dados.
  */
  if (filtroMes) {
    const mesSelecionado =
      filtroMes.value || "";

    const mesesComDados =
      [...new Set(
        operacoesOriginal
          .map(item => item.data_normalizada)
          .filter(Boolean)
          .map(data => data.slice(5, 7))
      )].sort();

    const nomesMeses = {
      "01": "Janeiro",
      "02": "Fevereiro",
      "03": "Março",
      "04": "Abril",
      "05": "Maio",
      "06": "Junho",
      "07": "Julho",
      "08": "Agosto",
      "09": "Setembro",
      "10": "Outubro",
      "11": "Novembro",
      "12": "Dezembro"
    };

    filtroMes.innerHTML =
      `<option value="">Mês</option>` +
      mesesComDados.map(mes => `
        <option
          value="${mes}"
          ${mes === mesSelecionado ? "selected" : ""}
        >
          ${nomesMeses[mes]}
        </option>
      `).join("");
  }

  /*
    Filtro Ano:
    mostra apenas anos que possuem dados.
  */
  if (filtroAno) {
    const anoSelecionado =
      filtroAno.value || "";

    const anosComDados =
      [...new Set(
        operacoesOriginal
          .map(item => item.data_normalizada)
          .filter(Boolean)
          .map(data => data.slice(0, 4))
      )].sort();

    filtroAno.innerHTML =
      `<option value="">Ano</option>` +
      anosComDados.map(ano => `
        <option
          value="${ano}"
          ${ano === anoSelecionado ? "selected" : ""}
        >
          ${ano}
        </option>
      `).join("");
  }
}

function renderGraficos() {
  const ctxExecucao =
    document.getElementById("graficoExecucao");

  const ctxPizza =
    document.getElementById("graficoAcumulado");

  if (!ctxExecucao || !ctxPizza) return;

  if (graficoExecucao) {
    graficoExecucao.destroy();
  }

  if (graficoPizza) {
    graficoPizza.destroy();
  }

  const labels =
    painelExecutivo.map(item => item.servico);

  const dadosExecucao =
    painelExecutivo.map(item => item.porcentagem_execucao);

  const dadosAcumulado =
    painelExecutivo.map(item => item.acumulado_mes);

  graficoExecucao =
    new Chart(ctxExecucao, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: "% Execução",
          data: dadosExecucao,
          borderRadius: 10,
          backgroundColor: "rgba(12,107,63,.7)"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });

  graficoPizza =
    new Chart(ctxPizza, {
      type: "doughnut",
      data: {
        labels,
        datasets: [{
          data: dadosAcumulado
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
}

/* =====================================================
   FILTRO DIÁRIO
===================================================== */

function aplicarFiltroDiario() {
  const campoData =
    document.getElementById("filtroDia");

  const data =
    campoData ? campoData.value : "";

  if (!data) {
    limparFiltroPeriodo();
    return;
  }

  const filtroMes =
    document.getElementById("filtroMes");

  const filtroAno =
    document.getElementById("filtroAno");

  if (filtroMes) filtroMes.value = "";
  if (filtroAno) filtroAno.value = "";

  operacoes =
    operacoesOriginal.filter(item =>
      item.data_normalizada === data
    );

  recalcularPainelPorFiltro();

  atualizarDashboard();
}

/* Compatibilidade caso seu HTML ainda chame aplicarFiltroPeriodo */
function aplicarFiltroPeriodo() {
  aplicarFiltroDiario();
}

/* =====================================================
   FILTRO MENSAL
===================================================== */

function aplicarFiltroMensal() {
  const filtroMes =
    document.getElementById("filtroMes");

  const filtroAno =
    document.getElementById("filtroAno");

  const mes =
    filtroMes ? filtroMes.value : "";

  const ano =
    filtroAno ? filtroAno.value : "";

  if (!mes || !ano) return;

  const filtroDia =
    document.getElementById("filtroDia");

  if (filtroDia) filtroDia.value = "";

  operacoes =
    operacoesOriginal.filter(item => {
      if (!item.data_normalizada) return false;

      const [anoItem, mesItem] =
        item.data_normalizada.split("-");

      return (
        anoItem === ano &&
        mesItem === mes
      );
    });

  recalcularPainelPorFiltro();

  atualizarDashboard();
}

/* =====================================================
   RECALCULAR PAINEL PELO FILTRO
===================================================== */

function recalcularPainelPorFiltro() {
  painelExecutivo =
    painelExecutivoOriginal.map(item => {
      const servico =
        item.servico;

      const dadosFiltrados =
        operacoes.filter(op =>
          op.servico === servico
        );

      const acumulado =
        dadosFiltrados.reduce((soma, op) => {
          const medicao =
            normalizar(item.medicao);

          if (medicao.includes("km")) {
            return soma + numero(op.km);
          }

          if (medicao.includes("t")) {
            return soma + numero(op.peso);
          }

          if (
            medicao.includes("vg") ||
            medicao.includes("viagem")
          ) {
            return soma + numero(op.viagens);
          }

          if (medicao.includes("equipe")) {
            return soma + numero(op.equipe);
          }

          return soma;
        }, 0);

      const valorUnitario =
        VALORES_FIXOS[servico] || 0;

      const valorFinal =
        SERVICOS_FIXOS.includes(servico)
          ? valorUnitario
          : valorUnitario * acumulado;

      return {
        ...item,
        acumulado_mes:
          acumulado,

        porcentagem_execucao:
          calcularPercentual(
            acumulado,
            item.previsto_mes
          ),

        valor:
          valorFinal,

        status:
          acumulado > 0
            ? "Com dados"
            : "Sem dados"
      };
    });
}

/* =====================================================
   LIMPAR FILTROS
===================================================== */

function limparFiltroPeriodo() {
  const filtroDia =
    document.getElementById("filtroDia");

  const filtroMes =
    document.getElementById("filtroMes");

  const filtroAno =
    document.getElementById("filtroAno");

  if (filtroDia) filtroDia.value = "";
  if (filtroMes) filtroMes.value = "";
  if (filtroAno) filtroAno.value = "";

  painelExecutivo =
    [...painelExecutivoOriginal];

  operacoes =
    [...operacoesOriginal];

  atualizarDashboard();
}

/* =====================================================
   TELAS
===================================================== */

function mostrarTela(nome, botao) {
  document
    .querySelectorAll(".tela")
    .forEach(item =>
      item.classList.remove("ativa")
    );

  const tela =
    document.getElementById(`tela-${nome}`);

  if (tela) {
    tela.classList.add("ativa");
  }

  document
    .querySelectorAll("nav button")
    .forEach(btn =>
      btn.classList.remove("active")
    );

  if (botao) {
    botao.classList.add("active");
  }

  if (nome === "historico") {
    carregarHistorico();
  }
}

function mostrarServico(codigo, botao) {
  document
    .querySelectorAll("#tela-contrato .servico-btn")
    .forEach(btn =>
      btn.classList.remove("active")
    );

  if (botao) {
    botao.classList.add("active");
  }

  const geral =
    document.getElementById("servico-geral");

  const detalhe =
    document.getElementById("servico-detalhe");

  if (codigo === "geral") {
    geral.classList.add("ativa");
    detalhe.classList.remove("ativa");
    return;
  }

  geral.classList.remove("ativa");
  detalhe.classList.add("ativa");

  renderDetalheServico(codigo);
}

/* =====================================================
   KPI'S
===================================================== */

const KPIS_POR_SERVICO = {
  geral: {
    titulo:
      "Indicadores Gerais de Operação",

    descricao:
      "Indicadores estratégicos para acompanhamento de volume, frota, distância, produtividade, tempo, velocidade e eficiência operacional.",

    blocos: [
      {
        titulo:
          "Volume e Demanda",

        itens: [
          "Dias de operação efetivos por mês ou por ano.",
          "Quantidade coletada por turno.",
          "Quantidade total coletada por dia.",
          "Quantidade coletada por veículo por turno."
        ]
      }
    ]
  }
};

function mostrarKPI(codigo, botao) {
  document
    .querySelectorAll("#tela-kpi .servico-btn")
    .forEach(btn => btn.classList.remove("active"));

  if (botao) botao.classList.add("active");

  const area = document.getElementById("conteudoKPI");
  if (!area) return;

  const servicoPainel =
    painelExecutivo.find(item => item.servico === codigo);

  const dadosServico =
    operacoes.filter(item => item.servico === codigo);

  if (codigo === "geral") {
    area.innerHTML = gerarKpiGeral();
    return;
  }

  if (!servicoPainel && !dadosServico.length) {
    area.innerHTML = `
      <div class="not-found">
        Nenhum dado encontrado para ${codigo}.
      </div>
    `;
    return;
  }

  const totalPeso =
    dadosServico.reduce((soma, item) => soma + numero(item.peso), 0);

  const totalViagens =
    dadosServico.reduce((soma, item) => soma + numero(item.viagens), 0);

  const totalKm =
    dadosServico.reduce((soma, item) => soma + numero(item.km), 0);

  const totalEquipe =
    dadosServico.reduce((soma, item) => soma + numero(item.equipe), 0);

  const produtividadeViagem =
    totalViagens > 0 ? totalPeso / totalViagens : 0;

  const produtividadeKm =
    totalKm > 0 ? totalPeso / totalKm : 0;

  const execucao =
    servicoPainel ? servicoPainel.porcentagem_execucao : 0;

  const acumulado =
    servicoPainel ? servicoPainel.acumulado_mes : 0;

  const previsto =
    servicoPainel ? servicoPainel.previsto_mes : 0;

  const valor =
    servicoPainel ? servicoPainel.valor : 0;

  area.innerHTML = `
    <div class="kpi-header">
      <span>KPI Operacional</span>
      <h2>${codigo} - ${servicoPainel?.nome_servico || "Serviço operacional"}</h2>
      <p>
        Resumo automático gerado com base nos dados importados da planilha.
      </p>
    </div>

    <div class="kpi-grid">

      <div class="kpi-box">
        <h3>Resumo Executivo</h3>
        <ul>
          <li>Previsto: <strong>${formatarNumero(previsto)}</strong></li>
          <li>Executado: <strong>${formatarNumero(acumulado)}</strong></li>
          <li>Índice de execução: <strong>${formatarNumero(execucao)}%</strong></li>
          <li>Valor estimado: <strong>${formatarMoeda(valor)}</strong></li>
        </ul>
      </div>

      <div class="kpi-box">
        <h3>Volume e Demanda</h3>
        <ul>
          <li>Peso total coletado: <strong>${formatarNumero(totalPeso)}</strong></li>
          <li>Viagens realizadas: <strong>${formatarNumero(totalViagens)}</strong></li>
          <li>KM executado: <strong>${formatarNumero(totalKm)}</strong></li>
          <li>Equipes em operação: <strong>${formatarNumero(totalEquipe)}</strong></li>
        </ul>
      </div>

      <div class="kpi-box">
        <h3>Produtividade</h3>
        <ul>
          <li>Toneladas por viagem: <strong>${formatarNumero(produtividadeViagem)}</strong></li>
          <li>Toneladas por KM: <strong>${formatarNumero(produtividadeKm)}</strong></li>
          <li>Total de registros analisados: <strong>${dadosServico.length}</strong></li>
        </ul>
      </div>

      <div class="kpi-box">
        <h3>Resposta para Diretoria</h3>
        <ul>
          <li>
            O serviço ${codigo} executou 
            <strong>${formatarNumero(acumulado)}</strong>
            de um previsto de 
            <strong>${formatarNumero(previsto)}</strong>,
            atingindo 
            <strong>${formatarNumero(execucao)}%</strong>.
          </li>
          <li>
            Foram identificados 
            <strong>${formatarNumero(totalViagens)}</strong> viagens,
            <strong>${formatarNumero(totalKm)}</strong> km
            e <strong>${formatarNumero(totalPeso)}</strong> de peso/volume operacional.
          </li>
        </ul>
      </div>

    </div>
  `;
}
function gerarKpiGeral() {
  const totalPeso = operacoes.reduce((soma, item) => soma + numero(item.peso), 0);
  const totalViagens = operacoes.reduce((soma, item) => soma + numero(item.viagens), 0);
  const totalKm = operacoes.reduce((soma, item) => soma + numero(item.km), 0);

  return `
    <div class="kpi-header">
      <span>KPI Geral</span>
      <h2>Indicadores Gerais da Operação</h2>
      <p>Resumo consolidado de todos os serviços importados.</p>
    </div>

    <div class="kpi-grid">
      <div class="kpi-box">
        <h3>Consolidado</h3>
        <ul>
          <li>Total de serviços: <strong>${painelExecutivo.length}</strong></li>
          <li>Total de registros: <strong>${operacoes.length}</strong></li>
          <li>Peso total: <strong>${formatarNumero(totalPeso)}</strong></li>
          <li>Viagens totais: <strong>${formatarNumero(totalViagens)}</strong></li>
          <li>KM total: <strong>${formatarNumero(totalKm)}</strong></li>
        </ul>
      </div>
    </div>
  `;
}
/* =====================================================
   SUPABASE: SALVAR, CARREGAR, HISTÓRICO
===================================================== */

async function salvarPlanilhaNoBanco(
  nomeArquivo = "Planilha Importada"
) {
  const usuarioLogado =
    JSON.parse(
      localStorage.getItem("usuarioLogado")
    ) || {};

  const totalAbas =
    Object.keys(sheetsOriginais).length;

  const {
    data: importacaoExistente,
    error: erroConsulta
  } = await banco
    .from("importacoes")
    .select("*")
    .eq("nome_arquivo", nomeArquivo)
    .eq("total_abas", totalAbas)
    .order("id", { ascending: false })
    .limit(1);

  if (erroConsulta) {
    console.error("Erro ao verificar duplicidade:", erroConsulta);
    alert("Erro ao verificar duplicidade.");
    return false;
  }

  if (
    importacaoExistente &&
    importacaoExistente.length > 0
  ) {
    const continuar =
      confirm(
        "Esta planilha já foi importada anteriormente.\n\nDeseja importar novamente?"
      );

    if (!continuar) {
      alert("Importação cancelada.");
      return false;
    }
  }

  const {
    error: erroDesativar
  } = await banco
    .from("importacoes")
    .update({ ativo: false })
    .eq("ativo", true);

  if (erroDesativar) {
    console.error(erroDesativar);
    alert("Erro ao desativar importação anterior.");
    return false;
  }

  const {
    data: importacao,
    error: erroImportacao
  } = await banco
    .from("importacoes")
    .insert({
      nome_arquivo:
        nomeArquivo,

      usuario:
        usuarioLogado.usuario ||
        "Não identificado",

      perfil:
        usuarioLogado.perfil ||
        "Sem perfil",

      total_abas:
        totalAbas,

      ativo:
        true
    })
    .select()
    .single();

  if (erroImportacao) {
    console.error(erroImportacao);
    alert("Erro ao registrar histórico.");
    return false;
  }

  const registros = [];

  Object.keys(sheetsOriginais).forEach(nomeAba => {
    const aba =
      sheetsOriginais[nomeAba];

    registros.push({
      nome_arquivo:
        nomeArquivo,

      aba:
        aba.nomeOriginal,

      codigo_servico:
        aba.codigoServico || "GERAL",

      dados:
        aba.dadosOriginais,

      importacao_id:
        importacao.id
    });
  });

  const {
    error: erroInsert
  } = await banco
    .from("planilhas_importadas")
    .insert(registros);

  if (erroInsert) {
    console.error(erroInsert);
    alert("Erro ao salvar planilha.");
    return false;
  }

  return true;
}

async function carregarBanco() {
  const {
    data,
    error
  } = await banco
    .from("planilhas_importadas")
    .select(`
      *,
      importacoes!inner(
        id,
        ativo
      )
    `)
    .eq("importacoes.ativo", true)
    .order("id", { ascending: true });

  if (error) {
    console.error(error);

    preencherTexto(
      "nomeArquivo",
      "Erro ao carregar Supabase"
    );

    return;
  }

  if (!data || !data.length) {
    preencherTexto(
      "nomeArquivo",
      "Nenhuma planilha ativa no banco"
    );
    return;
  }

  sheetsOriginais = {};
  todasAsAbas = [];

  data.forEach(item => {
    const nome =
      normalizar(item.aba);

    sheetsOriginais[nome] = {
      nomeOriginal:
        item.aba,

      codigoServico:
        item.codigo_servico,

      dadosOriginais:
        item.dados,

      dadosNormalizados:
        item.dados.map(linha =>
          normalizarObjeto(linha)
        )
    };

    todasAsAbas.push({
      arquivo:
        item.nome_arquivo,

      aba:
        item.aba
    });
  });

  gerarPainelExecutivo();
  gerarOperacoes();

  painelExecutivoOriginal =
    [...painelExecutivo];

  operacoesOriginal =
    [...operacoes];

  atualizarDashboard();

  preencherTexto(
    "nomeArquivo",
    `${todasAsAbas.length} aba(s) carregada(s) do Supabase`
  );
}

async function carregarHistorico() {
  const tabela =
    document.getElementById("tabelaHistorico");

  if (!tabela) return;

  const {
    data,
    error
  } = await banco
    .from("importacoes")
    .select("*")
    .order("criado_em", { ascending: false });

  if (error) {
    console.error(error);

    tabela.innerHTML = `
      <tr>
        <td colspan="7">
          Erro ao carregar histórico.
        </td>
      </tr>
    `;

    return;
  }

  if (!data || !data.length) {
    tabela.innerHTML = `
      <tr>
        <td colspan="7">
          Nenhuma importação encontrada.
        </td>
      </tr>
    `;

    return;
  }

  tabela.innerHTML =
    data.map(item => {
      const dataBR =
        new Date(item.criado_em)
          .toLocaleString("pt-BR");

      return `
        <tr>
          <td>${dataBR}</td>
          <td>${item.usuario}</td>
          <td>${item.perfil}</td>
          <td>${item.nome_arquivo}</td>
          <td>${item.total_abas}</td>
          <td>
            <span class="badge ${item.ativo ? "ok" : "info"}">
              ${item.ativo ? "ATIVA" : "HISTÓRICO"}
            </span>
          </td>
          <td>
            <button
              class="btn-mini"
              onclick="restaurarImportacao(${item.id})"
            >
              Restaurar
            </button>
          </td>
        </tr>
      `;
    }).join("");
}

async function restaurarImportacao(importacaoId) {
  if (
    !confirm(
      "Deseja restaurar esta importação?"
    )
  ) {
    return;
  }

  await banco
    .from("importacoes")
    .update({ ativo: false })
    .eq("ativo", true);

  const {
    error
  } = await banco
    .from("importacoes")
    .update({ ativo: true })
    .eq("id", importacaoId);

  if (error) {
    console.error(error);
    alert("Erro ao restaurar importação.");
    return;
  }

  await carregarBanco();
  await carregarHistorico();

  alert("Importação restaurada com sucesso!");
}

async function limparBanco() {
  if (
    !confirm(
      "Deseja desativar a base atual?"
    )
  ) {
    return;
  }

  const {
    error
  } = await banco
    .from("importacoes")
    .update({ ativo: false })
    .eq("ativo", true);

  if (error) {
    console.error(error);
    alert("Erro ao desativar base atual.");
    return;
  }

  painelExecutivo = [];
  painelExecutivoOriginal = [];
  operacoes = [];
  operacoesOriginal = [];
  sheetsOriginais = {};
  todasAsAbas = [];

  atualizarDashboard();
  await carregarHistorico();

  preencherTexto(
    "nomeArquivo",
    "Nenhuma planilha ativa"
  );

  alert("Base atual desativada.");
}

/* =====================================================
   FUNÇÕES AUXILIARES
===================================================== */

function buscarAbaServico(codigo) {
  for (const nome in sheetsOriginais) {
    const aba =
      sheetsOriginais[nome];

    if (aba.codigoServico === codigo) {
      return aba;
    }
  }

  return null;
}

function calcularAcumulado(medicao, dados) {
  const tipo =
    normalizar(medicao);

  return dados.reduce((soma, item) => {
    if (tipo.includes("km")) {
      return soma + numero(
        item.km ||
        item.km_total ||
        item.km_executado
      );
    }

    if (tipo.includes("t")) {
      return soma + numero(
        item.peso ||
        item.peso_total ||
        item.peso_t
      );
    }

    if (
      tipo.includes("vg") ||
      tipo.includes("viagem")
    ) {
      return soma + numero(
        item.viagens ||
        item.qtd_viagem ||
        item.qtd_viagens
      );
    }

    if (tipo.includes("equipe")) {
      return soma + numero(
        item.equipe ||
        item.qtd_equipe ||
        item.qdt_equipe
      );
    }

    return soma + numero(
      item.total ||
      item.quantidade ||
      item.valor
    );
  }, 0);
}

function calcularPercentual(realizado, previsto) {
  if (!previsto) return 0;

  return Math.min(
    (realizado / previsto) * 100,
    100
  );
}

function mostrarLoading(ativo) {
  const loading =
    document.getElementById("loadingOverlay");

  if (!loading) return;

  loading.classList.toggle(
    "ativo",
    ativo
  );
}

function sair() {
  window.location.href =
    "login.html";
}

function extrairCodigo(texto) {
  const match =
    String(texto)
      .toUpperCase()
      .match(/P\d+(\.\d+)?/);

  return match ? match[0] : "";
}

function normalizarObjeto(obj) {
  const novo = {};

  Object.keys(obj).forEach(chave => {
    const novaChave =
      normalizar(chave)
        .replace(/ª/g, "a")
        .replace(/º/g, "o")
        .replace(/%/g, "porcentagem")
        .replace(/\$/g, "valor")
        .replace(/\./g, "")
        .replace(/\s+/g, "_")
        .replace(/[^\w]/g, "");

    novo[novaChave] =
      obj[chave];
  });

  return novo;
}

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function normalizarData(valor) {
  /*
    Converte datas para yyyy-mm-dd.
    Corrige:
    - 26/05/2026
    - 26/05/2026 08:30
    - 2026-05-26
    - número serial do Excel
  */

  if (!valor) return "";

  // Número serial do Excel
  if (typeof valor === "number") {

    const data =
      XLSX.SSF.parse_date_code(valor);

    if (!data) return "";

    const ano =
      data.y;

    const mes =
      String(data.m).padStart(2, "0");

    const dia =
      String(data.d).padStart(2, "0");

    return `${ano}-${mes}-${dia}`;
  }

  const texto =
    String(valor).trim();

  // Formato brasileiro
  const br =
    texto.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/
    );

  if (br) {

    const dia =
      br[1].padStart(2, "0");

    const mes =
      br[2].padStart(2, "0");

    let ano =
      br[3];

    if (ano.length === 2) {
      ano = "20" + ano;
    }

    return `${ano}-${mes}-${dia}`;
  }

  // Formato ISO
  const iso =
    texto.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );

  if (iso) {
    return `${iso[1]}-${iso[2]}-${iso[3]}`;
  }

  return "";
}

function numero(valor) {
  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {
    return 0;
  }

  if (typeof valor === "number") {
    return valor;
  }

  let texto =
    String(valor)
      .trim()
      .replace(/[^\d,.-]/g, "");

  const temVirgula =
    texto.includes(",");

  const temPonto =
    texto.includes(".");

  if (temVirgula && temPonto) {
    texto =
      texto.replace(/\./g, "")
           .replace(",", ".");

  } else if (temVirgula && !temPonto) {
    texto =
      texto.replace(",", ".");
  }

  const convertido =
    Number(texto);

  return isNaN(convertido)
    ? 0
    : convertido;
}

function formatarNumero(valor) {
  return numero(valor)
    .toLocaleString("pt-BR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    });
}

function formatarMoeda(valor) {
  return numero(valor)
    .toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
}

function preencherTexto(id, texto) {
  const elemento =
    document.getElementById(id);

  if (elemento) {
    elemento.innerText = texto;
  }
}


function gerarResumoExecutivoAutomatico() {
  const totalServicos =
    painelExecutivo.length;

  const servicosComDados =
    painelExecutivo.filter(item =>
      item.status === "Com dados"
    ).length;

  const mediaExecucao =
    painelExecutivo.length
      ? painelExecutivo.reduce(
          (soma, item) =>
            soma + numero(item.porcentagem_execucao),
          0
        ) / painelExecutivo.length
      : 0;

  const maiorExecucao =
    [...painelExecutivo].sort(
      (a, b) =>
        numero(b.porcentagem_execucao) -
        numero(a.porcentagem_execucao)
    )[0];

  const menorExecucao =
    [...painelExecutivo].sort(
      (a, b) =>
        numero(a.porcentagem_execucao) -
        numero(b.porcentagem_execucao)
    )[0];

  return `
    Foram analisados ${totalServicos} serviços operacionais,
    sendo ${servicosComDados} com dados registrados na base importada.

    A média geral de execução foi de ${formatarNumero(mediaExecucao)}%.

    O serviço com melhor desempenho foi ${maiorExecucao?.servico || "-"},
    com ${formatarNumero(maiorExecucao?.porcentagem_execucao || 0)}%.

    O serviço com menor execução foi ${menorExecucao?.servico || "-"},
    com ${formatarNumero(menorExecucao?.porcentagem_execucao || 0)}%.

    Esse resumo foi gerado automaticamente a partir dos dados importados.
  `;
}
function renderResumoAutomaticoDiretoria() {
  const area =
    document.getElementById("resumoAutomaticoDiretoria");

  if (!area) return;

  area.innerHTML =
    gerarResumoExecutivoAutomatico();
}
function renderRankingOperacional() {

  const area =
    document.getElementById("rankingOperacional");

  if (!area) return;

  if (!painelExecutivo.length) {

    area.innerHTML =
      "Importe uma planilha para gerar o ranking.";

    return;
  }

  /*
  =====================================================
  TOP 3 SERVIÇOS
  =====================================================
  */

  const ranking =
    [...painelExecutivo]

      .filter(item =>
        numero(item.acumulado_mes) > 0
      )

      .sort(
        (a, b) =>
          numero(b.porcentagem_execucao) -
          numero(a.porcentagem_execucao)
      )

      .slice(0, 3);

  /*
  =====================================================
  HTML
  =====================================================
  */

  area.innerHTML =
    ranking.map((item, index) => {

      let medalha = "📌";
      let classe = "ranking-normal";
      let titulo = "";

      /*
      =====================================================
      POSIÇÕES
      =====================================================
      */

      if (index === 0) {
        medalha = "🥇";
        classe = "ranking-ouro";
        titulo = "Melhor desempenho operacional";
      }

      if (index === 1) {
        medalha = "🥈";
        classe = "ranking-prata";
        titulo = "Segundo maior desempenho";
      }

      if (index === 2) {
        medalha = "🥉";
        classe = "ranking-bronze";
        titulo = "Terceiro maior desempenho";
      }

      return `
        <div class="ranking-item ${classe}">

          <div class="ranking-left">

            <div class="ranking-medalha">
              ${medalha}
            </div>

            <div>
              <strong>
                ${item.servico}
              </strong>

              <small>
                ${item.nome_servico}
              </small>

              <div class="ranking-desc">
                ${titulo}
              </div>
            </div>

          </div>

          <div class="ranking-right">

            <strong>
              ${formatarNumero(item.porcentagem_execucao)}%
            </strong>

            <span>
              ${formatarNumero(item.acumulado_mes)}
            </span>

          </div>

        </div>
      `;
    }).join("");
}
function renderKpiSemanal() {

  const area =
    document.getElementById("kpiSemanal");

  if (!area) return;

  if (!painelExecutivo.length) {

    area.innerHTML =
      "Importe uma planilha para gerar os KPI’s semanais.";

    return;
  }

  /*
  =====================================================
  SOMATÓRIOS
  =====================================================
  */

  const semana1 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana1),
      0
    );

  const semana2 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana2),
      0
    );

  const semana3 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana3),
      0
    );

  const semana4 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana4),
      0
    );

  /*
  =====================================================
  MELHOR SEMANA
  =====================================================
  */

  const semanas = [
    { nome: "1ª Semana", valor: semana1 },
    { nome: "2ª Semana", valor: semana2 },
    { nome: "3ª Semana", valor: semana3 },
    { nome: "4ª Semana", valor: semana4 }
  ];

  const melhorSemana =
    [...semanas]
      .sort((a, b) => b.valor - a.valor)[0];

  /*
  =====================================================
  HTML
  =====================================================
  */

  area.innerHTML = `

    <div class="semana-grid">

      <div class="semana-card">
        <span>1ª Semana</span>
        <strong>${formatarNumero(semana1)}</strong>
      </div>

      <div class="semana-card">
        <span>2ª Semana</span>
        <strong>${formatarNumero(semana2)}</strong>
      </div>

      <div class="semana-card">
        <span>3ª Semana</span>
        <strong>${formatarNumero(semana3)}</strong>
      </div>

      <div class="semana-card">
        <span>4ª Semana</span>
        <strong>${formatarNumero(semana4)}</strong>
      </div>

    </div>

    <div class="semana-resumo">

      🏆 Melhor desempenho:
      <strong>${melhorSemana.nome}</strong>

      com volume de
      <strong>${formatarNumero(melhorSemana.valor)}</strong>

    </div>
  `;
}
function renderComparativoOperacional() {

  const area =
    document.getElementById("comparativoOperacional");

  if (!area) return;

  if (!painelExecutivo.length) {

    area.innerHTML =
      "Importe uma planilha para gerar o comparativo.";

    return;
  }

  /*
  =====================================================
  SOMA DAS SEMANAS
  =====================================================
  */

  const semana1 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana1),
      0
    );

  const semana2 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana2),
      0
    );

  const semana3 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana3),
      0
    );

  const semana4 =
    painelExecutivo.reduce(
      (soma, item) =>
        soma + numero(item.semana4),
      0
    );

  /*
  =====================================================
  COMPARAÇÃO
  =====================================================
  */

  const ultimaSemana =
    semana4 || semana3 || semana2 || semana1;

  const semanaAnterior =
    semana3 || semana2 || semana1;

  let percentual = 0;

  if (semanaAnterior > 0) {

    percentual =
      (
        (
          ultimaSemana -
          semanaAnterior
        ) / semanaAnterior
      ) * 100;
  }

  /*
  =====================================================
  STATUS
  =====================================================
  */

  let status = "Estabilidade operacional";
  let emoji = "📊";
  let classe = "comparativo-neutro";

  if (percentual > 5) {
    status = "Crescimento operacional";
    emoji = "📈";
    classe = "comparativo-positivo";
  }

  if (percentual < -5) {
    status = "Queda operacional";
    emoji = "📉";
    classe = "comparativo-negativo";
  }

  /*
  =====================================================
  HTML
  =====================================================
  */

  area.innerHTML = `

    <div class="comparativo-card ${classe}">

      <div class="comparativo-topo">

        <div>
          <span>Tendência operacional</span>

          <h3>
            ${emoji} ${status}
          </h3>
        </div>

        <strong>
          ${formatarNumero(percentual)}%
        </strong>

      </div>

      <p>
        Comparando a última semana operacional
        com a semana anterior,
        foi identificado um comportamento de
        <strong>${status.toLowerCase()}</strong>
        na execução dos serviços.
      </p>

    </div>
  `;
}