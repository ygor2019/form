// ============================================
// FUNÇÃO MELHORADA - Gerar Excel BONITO (2 colunas) e PADRONIZADO
// Layout: Coluna A (descrição) + Coluna B (resposta)
// Subcontratados: vertical, hierarquizado, sem “apagar” campos
// ============================================
async function gerarExcel() {
  showToast("📊 Gerando Excel formatado...", "success");

  try {
    const wb = XLSX.utils.book_new();
    const ws = {};
    let maxRow = 0;

    // =========================
    // HELPERS BÁSICOS
    // =========================
    function safeStr(v) {
      // mantém vazio como vazio (não apaga linha/estrutura)
      if (v === null || v === undefined) return "";
      return String(v);
    }

    function setCell(cell, value, style = {}) {
      if (!ws[cell]) ws[cell] = {};
      ws[cell].v = safeStr(value);
      ws[cell].t = "s";

      const defaultStyle = {
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } },
        },
        alignment: { vertical: "center", wrapText: true },
      };

      ws[cell].s = { ...(ws[cell].s || {}), ...defaultStyle, ...style };

      const row = parseInt(cell.match(/\d+/)[0], 10);
      if (row > maxRow) maxRow = row;
    }

    ws["!merges"] = ws["!merges"] || [];

    // Mescla A:B na linha
    function mergeRowAB(row) {
      ws["!merges"].push({
        s: { r: row - 1, c: 0 }, // A
        e: { r: row - 1, c: 1 }, // B
      });
    }

    // =========================
    // ESTILOS
    // =========================
    const estilos = {
      titulo: {
        font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "1A365D" } },
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
      },
      subtitulo: {
        font: { bold: true, sz: 12, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "2C5282" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
      },
      secao: {
        font: { bold: true, sz: 11, color: { rgb: "1A202C" } },
        fill: { fgColor: { rgb: "E2E8F0" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
      },
      label: {
        font: { bold: true, sz: 10, color: { rgb: "1A202C" } },
        fill: { fgColor: { rgb: "EDF2F7" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
      },
      item: {
        font: { bold: true, sz: 10, color: { rgb: "1A202C" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
      },
      itemIndent1: {
        font: { sz: 10, color: { rgb: "1A202C" } },
        alignment: {
          horizontal: "left",
          vertical: "center",
          wrapText: true,
          indent: 1,
        },
      },
      itemIndent2: {
        font: { sz: 10, color: { rgb: "1A202C" } },
        alignment: {
          horizontal: "left",
          vertical: "center",
          wrapText: true,
          indent: 2,
        },
      },
      valor: {
        font: { sz: 10, color: { rgb: "1A202C" } },
        alignment: { horizontal: "left", vertical: "center", wrapText: true },
      },
    };

    // =========================
    // COMPONENTES DE LINHA (2 COLUNAS)
    // =========================
    function addTitle(text, row) {
      setCell(`A${row}`, text, estilos.titulo);
      setCell(`B${row}`, "", estilos.titulo);
      mergeRowAB(row);
      return row + 2;
    }

    function addSubtitle(text, row) {
      setCell(`A${row}`, text, estilos.subtitulo);
      setCell(`B${row}`, "", estilos.subtitulo);
      mergeRowAB(row);
      return row + 1;
    }

    function addSection(text, row) {
      setCell(`A${row}`, text, estilos.secao);
      setCell(`B${row}`, "", estilos.secao);
      mergeRowAB(row);
      return row + 1;
    }

    function addRow(row, descricao, valor) {
      setCell(`A${row}`, descricao, estilos.label);
      setCell(`B${row}`, valor, estilos.valor);
      return row + 1;
    }

    function addItem(row, descricao, valor, indent = 0) {
      const styleDesc =
        indent === 0
          ? estilos.item
          : indent === 1
          ? estilos.itemIndent1
          : estilos.itemIndent2;
      setCell(`A${row}`, descricao, styleDesc);
      setCell(`B${row}`, valor, estilos.valor);
      return row + 1;
    }

    function addBlank(row, lines = 1) {
      return row + lines;
    }

    // =========================
    // MAPEAMENTOS
    // =========================
    const statusLabel = {
      nao_iniciada: "Não iniciada",
      em_andamento: "Em andamento",
      concluida: "Concluída",
    };

    // Títulos dos grupos (9–24) (pra ficar bonito no Excel)
    const titulosSub = {
      9: "📐 Subcontratados - Grupo 1 (Topografia, Sondagem, Movimento de Terra e Fundações)",
      10: "🏗️ Subcontratados - Grupo 2 (Estrutura, Concreto, Premoldados, Formas e Segurança)",
      11: "🏭 Subcontratados - Grupo 3 (Equipamentos, Geradores, Compressores e Resíduos)",
      12: "⚡ Subcontratados - Grupo 4 (Instalações Elétricas e Hidráulicas)",
      13: "🧯 Subcontratados - Grupo 5 (Segurança, AVCB, Laudos e Inspeções)",
      14: "🔥 Subcontratados - Grupo 6 (Gás, Incêndio, Energia Solar e SPDA)",
      15: "📞 Subcontratados - Grupo 7 (Telefonia, Segurança Eletrônica e Central de Concreto)",
      16: "🏊 Subcontratados - Grupo 8 (Elevadores de Carga, Piscinas e Sistemas de Água)",
      17: "🪟 Subcontratados - Grupo 9 (Fachada, Esquadrias Metálicas e Poço Artesiano)",
      18: "🧱 Subcontratados - Grupo 10 (Alvenaria, Revestimentos e Jardinagem)",
      19: "📋 Subcontratados - Grupo 11 (PGR, PGR Terceirizadas, PCMAT e LTCAT)",
      20: "🎨 Subcontratados - Grupo 12 (Drenagem, Extintores, Pintura e Instalações Gerais)",
      21: "🛣️ Subcontratados - Grupo 13 (Pavimentação e Estudos Ambientais)",
      22: "🔌 Subcontratados - Grupo 14 (Rede Elétrica, Canteiro de Obras e Viabilidade)",
      23: "🚦 Subcontratados - Grupo 15 (Sinalização, Tráfego e Controle Tecnológico)",
      24: "📝 Outras Atividades Técnicas",
    };

    // Pega dados de subcontratados de forma segura
    function getSubData(pageNum, id) {
      // page8 (Gerenciamento) é especial: vem de formData.page8
      if (pageNum === 8 && id === "gerenciamento") {
        const g = formData?.page8 || {};

        return {
          opcao: g?.houveGerenciamento || "",

          profissional: g?.profissionalGerenciamento || "",
          cpf: g?.cpfGerenciamento || "",

          // REGISTRO (padrão)
          tipoReg: g?.tipoRegistroGerenciamento || "",
          numCrea: g?.numeroCreaGerenciamento || "",
          sigla: g?.siglaConselhoGerenciamento || "",
          numOutros: g?.numeroConselhoGerenciamento || "",

          // ART (padrão)
          tipoArt: g?.tipoArtGerenciamento || "",
          numArt: g?.artNumeroGerenciamento || "",
          nomeArt: g?.nomeOutrosArtGerenciamento || "",
          valArt: g?.numeroOutrosArtGerenciamento || "",

          // EMPRESA
          empresa: g?.empresaGerenciamento || "",
          cnpj: g?.cnpjGerenciamento || "",
          regEmpresa:
            g?.tipoRegistroEmpresaGer === "CREA"
              ? g?.numeroCreaEmpresaGer || ""
              : [g?.siglaConselhoEmpresaGer, g?.numeroConselhoEmpresaGer]
                  .filter(Boolean)
                  .join(" "),
          placa: g?.placaGerenciamento || "",
        };
      }

      // Demais páginas (9–24): vem de formData.pageN[id]
      const pageKey = "page" + pageNum;
      const pageObj = formData?.[pageKey] || {};
      return pageObj?.[id] || null;
    }

    function formatOpcao(itemData) {
      const op = itemData?.opcao || "";
      // status fica "nao_iniciada/em_andamento/concluida"
      if (statusLabel[op]) return statusLabel[op];
      return op; // Sim/Não/etc
    }

    // Escreve bloco padrão do profissional/ART/empresa (sempre embaixo, mesmo vazio)
    // Escreve bloco padrão do profissional/ART/empresa (sempre embaixo, mesmo vazio)
    // PADRÃO: Conselho = CREA OU sigla | Número = número correspondente
    function escreverDetalhesProf(row, itemData, indentBase = 1) {
      row = addItem(
        row,
        "Profissional",
        itemData?.profissional || "",
        indentBase
      );
      row = addItem(row, "CPF", itemData?.cpf || "", indentBase);

      // ---- Registro (CREA ou Outro conselho) ----
      const tipoReg = itemData?.tipoReg || "";
      const conselho = tipoReg === "CREA" ? "CREA" : itemData?.sigla || ""; // sigla do conselho quando não é CREA

      const numeroRegistro =
        tipoReg === "CREA"
          ? itemData?.numCrea || ""
          : itemData?.numOutros || "";

      row = addItem(row, "Conselho", conselho, indentBase);
      row = addItem(row, "Número do Registro", numeroRegistro, indentBase);

      // ---- ART (ou outro tipo) ----
      const tipoArt = itemData?.tipoArt || "";
      const artOuOutro = tipoArt === "ART" ? "ART" : itemData?.nomeArt || ""; // nome do "outro" tipo

      const numeroArtOuOutro =
        tipoArt === "ART" ? itemData?.numArt || "" : itemData?.valArt || "";

      row = addItem(row, "ART ou Outro Tipo?", artOuOutro, indentBase);
      row = addItem(row, "Número", numeroArtOuOutro, indentBase);

      // ---- Empresa ----
      row = addItem(row, "Empresa", itemData?.empresa || "", indentBase);
      row = addItem(row, "CNPJ", itemData?.cnpj || "", indentBase);

      // Registro empresa + placa (se existir no objeto)
      if ("regEmpresa" in (itemData || {}) || "placa" in (itemData || {})) {
        row = addItem(
          row,
          "Registro Empresa",
          itemData?.regEmpresa || "",
          indentBase
        );
        row = addItem(
          row,
          "Placa/Equipamento",
          itemData?.placa || "",
          indentBase
        );
      }

      return row;
    }

    function escreverTerceirizadas(row, itemData, indentBase = 1) {
      // Sempre padronizado: escreve pelo menos 1 linha (mesmo sem nada)
      const lista = Array.isArray(itemData?.terceirizadas)
        ? itemData.terceirizadas
        : [];

      if (lista.length === 0) {
        row = addItem(row, "Terceirizada 1", "", indentBase);
        row = addItem(row, "ART 1", "", indentBase + 1);
        return row;
      }

      lista.forEach((t, idx) => {
        const n = idx + 1;
        row = addItem(
          row,
          `Terceirizada ${n}`,
          t?.terceirizada || "",
          indentBase
        );
        row = addItem(row, `ART ${n}`, t?.art || "", indentBase + 1);
      });

      return row;
    }

    // =========================
    // COMEÇA A ESCREVER O EXCEL
    // =========================
    let row = 1;

    row = addTitle(
      "FORMULÁRIO DE FISCALIZAÇÃO EM OBRAS DE MÉDIO E GRANDE PORTE",
      row
    );

    // -------------------------
    // PÁGINAS 1–8 (FORMULÁRIO PRINCIPAL)
    // -------------------------
    row = addSubtitle("LOCAL DA OBRA", row);
    row = addRow(row, "Endereço da Obra:", formData?.page1?.localObra || "");
    row = addRow(row, "Bairro:", formData?.page1?.bairro || "");
    row = addRow(row, "Cidade:", formData?.page1?.cidade || "");
    row = addRow(row, "CEP:", formData?.page1?.cep || "");
    row = addBlank(row, 1);

    row = addSubtitle("EMPREENDIMENTO", row);
    row = addRow(row, "Tipo:", formData?.page2?.tipoNatureza || "");
    row = addRow(row, "Área (m²):", formData?.page2?.area || "");
    row = addRow(row, "N° de Pavimentos:", formData?.page2?.pavimentos || "");
    row = addRow(row, "Nº do Processo:", formData?.page2?.processoNum || "");
    row = addRow(row, "Nº do Alvará:", formData?.page2?.alvaraNum || "");
    row = addRow(row, "Ano:", formData?.page2?.alvaraDe || "");
    row = addRow(row, "Proprietário:", formData?.page2?.proprietario || "");
    row = addRow(
      row,
      "CPF ou CNPJ:",
      (formData?.page2?.tipoCpfCnpj || "").toUpperCase()
    );
    row = addRow(row, "Número:", formData?.page2?.cpfCnpj || "");
    row = addRow(
      row,
      "Endereço do Proprietário:",
      formData?.page2?.enderecoProprietario || ""
    );
    row = addBlank(row, 1);

    row = addSubtitle("AUTOR DO PROJETO", row);
    row = addRow(row, "Placa afixada?", formData?.page3?.placaAfixada || "");
    row = addRow(row, "Profissional:", formData?.page3?.profissional || "");
    row = addRow(
      row,
      "Registro:",
      formData?.page3?.tipoRegistro === "CREA"
        ? "CREA"
        : formData?.page3?.tipoRegistro === "CPF"
        ? "CPF"
        : formData?.page3?.siglaConselho || ""
    );
    row = addRow(
      row,
      "Número:",
      formData?.page3?.numeroCrea ||
        formData?.page3?.cpfProfissional ||
        formData?.page3?.numeroConselho ||
        ""
    );
    row = addRow(row, "Título:", formData?.page3?.titulo || "");
    row = addRow(row, "Nº da ART:", formData?.page3?.artNumero || "");
    row = addRow(row, "Nº do RRT:", formData?.page3?.rrtNumero || "");
    row = addRow(row, "Empresa:", formData?.page3?.empresa || "");
    row = addRow(
      row,
      "Tipo de Registro da Empresa:",
      formData?.page3?.tipoEmpresaRegistro || ""
    );
    row = addRow(row, "Número:", formData?.page3?.empresaRegistro || "");
    row = addBlank(row, 1);

    row = addSubtitle("RESPONSÁVEL TÉCNICO", row);
    row = addRow(row, "Placa afixada?", formData?.page4?.placaAfixadaRT || "");
    row = addRow(row, "Profissional:", formData?.page4?.profissionalRT || "");
    row = addRow(row, "CPF:", formData?.page4?.cpfRT || "");
    row = addRow(
      row,
      "Conselho:",
      formData?.page4?.tipoRegistroRT === "CREA"
        ? "CREA"
        : formData?.page4?.siglaConselhoRT || ""
    );
    row = addRow(
      row,
      "Número de Registro:",
      formData?.page4?.numeroCreaRT || formData?.page4?.numeroConselhoRT || ""
    );
    row = addRow(
      row,
      "ART ou Outro Tipo?",
      formData?.page4?.tipoArtOutrosRT === "ART"
        ? "ART"
        : formData?.page4?.nomeOutrosArtRT || ""
    );
    row = addRow(
      row,
      "Número:",
      formData?.page4?.artNumeroRT || formData?.page4?.numeroOutrosArtRT || ""
    );
    row = addRow(row, "Empresa:", formData?.page4?.empresaRT || "");
    row = addRow(
      row,
      "Tipo de Registro da Empresa:",
      formData?.page4?.tipoEmpresaRegistroRT || ""
    );
    row = addRow(row, "Número:", formData?.page4?.empresaRegistroRT || "");
    row = addBlank(row, 1);

    row = addSubtitle("DEMOLIÇÃO", row);
    row = addRow(
      row,
      "Houve Demolição?",
      formData?.page5?.houveDemolicao || ""
    );
    row = addSection("PROJETO / PLANO DE DEMOLIÇÃO", row);
    row = addRow(
      row,
      "Houve Projeto de Demolição?",
      formData?.page5?.projetoDemolicao || ""
    );
    row = addRow(
      row,
      "Profissional:",
      formData?.page5?.profissionalProjetoDemol || ""
    );
    row = addRow(row, "CPF:", formData?.page5?.cpfProjetoDemol || "");
    row = addRow(
      row,
      "Conselho:",
      formData?.page5?.tipoRegistroProjetoDemol === "CREA"
        ? "CREA"
        : formData?.page5?.siglaConselhoProjetoDemol || ""
    );
    row = addRow(
      row,
      "Número do Registro:",
      formData?.page5?.numeroCreaProjetoDemol ||
        formData?.page5?.numeroConselhoProjetoDemol ||
        ""
    );
    row = addRow(
      row,
      "ART ou Outro Tipo?",
      formData?.page5?.tipoArtProjetoDemol === "ART"
        ? "ART"
        : formData?.page5?.nomeOutrosArtProjetoDemol || ""
    );
    row = addRow(
      row,
      "Número:",
      formData?.page5?.artNumeroProjetoDemol ||
        formData?.page5?.numeroOutrosArtProjetoDemol ||
        ""
    );

    row = addSection("EXECUÇÃO DA DEMOLIÇÃO", row);
    row = addRow(
      row,
      "Foi Executada a Demolição?",
      formData?.page5?.execucaoDemolicao || ""
    );
    row = addRow(
      row,
      "Profissional:",
      formData?.page5?.profissionalExecDemol || ""
    );
    row = addRow(row, "CPF:", formData?.page5?.cpfExecDemol || "");
    row = addRow(
      row,
      "Conselho:",
      formData?.page5?.tipoRegistroExecDemol === "CREA"
        ? "CREA"
        : formData?.page5?.siglaConselhoExecDemol || ""
    );
    row = addRow(
      row,
      "Número do Registro:",
      formData?.page5?.numeroCreaExecDemol ||
        formData?.page5?.numeroConselhoExecDemol ||
        ""
    );
    row = addRow(
      row,
      "ART ou Outro Tipo?",
      formData?.page5?.tipoArtExecDemol === "ART"
        ? "ART"
        : formData?.page5?.nomeOutrosArtExecDemol || ""
    );
    row = addRow(
      row,
      "Número:",
      formData?.page5?.artNumeroExecDemol ||
        formData?.page5?.numeroOutrosArtExecDemol ||
        ""
    );
    row = addRow(row, "Empresa:", formData?.page5?.empresaDemol || "");
    row = addRow(row, "CNPJ:", formData?.page5?.cnpjDemol || "");
    row = addBlank(row, 1);

    row = addSubtitle("COORDENAÇÃO / ORÇAMENTO / RESIDENTE", row);
    row = addRow(
      row,
      "Engenheiro Responsável pela Coordenação?",
      formData?.page6?.houveCoordenacao || ""
    );
    row = addRow(
      row,
      "Engenheiro Responsável pelo Orçamento?",
      formData?.page6?.houveOrcamento || ""
    );
    row = addRow(
      row,
      "Houve Engenheiro Residente?",
      formData?.page6?.houveResidente || ""
    );

    function escreverEquipe(row, titulo, obj) {
      row = addSection(titulo, row);
      row = addRow(row, "Nome:", obj?.nome || "");
      row = addRow(row, "CPF:", obj?.cpf || "");
      row = addRow(
        row,
        "Conselho:",
        obj?.reg === "CREA" ? "CREA" : obj?.sigla || ""
      );
      row = addRow(
        row,
        "Número do Registro:",
        obj?.numCrea || obj?.numOutros || ""
      );
      row = addRow(
        row,
        "ART ou Outro Tipo?",
        obj?.tipoArt === "ART" ? "ART" : obj?.nomeArtOutros || ""
      );
      row = addRow(row, "Número:", obj?.numArt || obj?.numArtOutros || "");
      return row;
    }

    row = escreverEquipe(row, "COORDENAÇÃO", formData?.page6?.coordenacao);
    row = escreverEquipe(row, "ORÇAMENTO", formData?.page6?.orcamento);
    row = escreverEquipe(row, "RESIDENTE", formData?.page6?.residente);
    row = addBlank(row, 1);

    row = addSubtitle("LIVRO DE ORDEM", row);
    row = addRow(
      row,
      "Livro de Ordem no Local?",
      formData?.page7?.livroOrdem || ""
    );
    row = addRow(row, "Observações:", formData?.page7?.observacoes || "");
    row = addBlank(row, 2);

    // -------------------------
    // PÁGINAS 8–24 (SUBCONTRATADOS)
    // -------------------------
    row = addTitle("SUBCONTRATADOS (profissionais e empresas)", row);

    for (let pageNum = 8; pageNum <= 24; pageNum++) {
      const pageKey = "page" + pageNum;

      // page8 não vem de "subcontratados", então criamos um item virtual
      const items =
        pageNum === 8
          ? [
              {
                id: "gerenciamento",
                titulo: "Gerenciamento",
                tipo: "padrao", // qualquer coisa que NÃO seja 'terceirizadas'
              },
            ]
          : subcontratados?.[pageKey];

      if (!items || items.length === 0) continue;

      // título bonito: page8 você pode dar um subtítulo específico
      if (pageNum === 8) {
        row = addSubtitle("🧭 Subcontratados - Gerenciamento", row);
      } else {
        row = addSubtitle(
          titulosSub[pageNum] || `Subcontratados - Página ${pageNum}`,
          row
        );
      }

      items.forEach((item) => {
        const renderItem = (it, pageNumLocal, indent = 0) => {
          const data = getSubData(pageNumLocal, it.id);

          const label = `${it.num ? it.num + " - " : ""}${it.titulo || it.id}`;
          row = addItem(row, label, formatOpcao(data), indent);

          if (it.tipo === "terceirizadas") {
            row = escreverTerceirizadas(row, data, indent + 1);
            row = addBlank(row, 1);
            return;
          }

          row = escreverDetalhesProf(row, data, indent + 1);

          if (Array.isArray(it.subitens) && it.subitens.length > 0) {
            it.subitens.forEach((sub) =>
              renderItem(sub, pageNumLocal, indent + 1)
            );
          }

          row = addBlank(row, 1);
        };

        renderItem(item, pageNum, 0);
      });

      row = addBlank(row, 1);
    }

    // -------------------------
    // PÁGINA 25 (FINAL)
    // -------------------------
    row = addTitle("IDENTIFICAÇÃO FINAL", row);

    row = addSubtitle("IDENTIFICAÇÃO DO DECLARANTE", row);
    row = addRow(row, "Nome", formData?.page25?.nomeDeclarante || "");
    row = addRow(
      row,
      "Qualificação e Cargo",
      formData?.page25?.qualificacaoDeclarante || ""
    );
    row = addRow(row, "Local", formData?.page25?.localDeclarante || "");
    // Usa sua função formatarData(dataISO) já existente no arquivo
    row = addRow(
      row,
      "Data",
      typeof formatarData === "function"
        ? formatarData(formData?.page25?.dataDeclarante || "")
        : formData?.page25?.dataDeclarante || ""
    );
    row = addBlank(row, 1);

    row = addSubtitle("AGENTE FISCAL", row);
    row = addRow(row, "Agente Fiscal", formData?.page25?.agenteFiscal1 || "");
    row = addRow(row, "Registro", formData?.page25?.registroAgente1 || "");
    row = addRow(
      row,
      "Data Fiscalização",
      typeof formatarData === "function"
        ? formatarData(formData?.page25?.dataFiscalizacao || "")
        : formData?.page25?.dataFiscalizacao || ""
    );
    row = addBlank(row, 1);

    row = addSubtitle("CONTATO", row);
    row = addRow(row, "Endereço", formData?.page25?.enderecoContato || "");
    row = addRow(row, "E-mail", formData?.page25?.emailContato || "");
    row = addRow(row, "Telefones", formData?.page25?.telefonesContato || "");
    row = addBlank(row, 1);

    // =========================
    // CONFIGURAÇÃO DE LARGURA / RANGE (SÓ A:B)
    // =========================
    ws["!cols"] = [
      { wch: 80 }, // A
      { wch: 45 }, // B
    ];

    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: 1 }, // até coluna B
    });

    XLSX.utils.book_append_sheet(wb, ws, "Formulário");

    const agora = new Date();
    const yyyy = agora.getFullYear();
    const mm = String(agora.getMonth() + 1).padStart(2, "0");
    const dd = String(agora.getDate()).padStart(2, "0");
    const nomeArquivo = `Formulario_Obra_${yyyy}-${mm}-${dd}.xlsx`;

    XLSX.writeFile(wb, nomeArquivo);

    showToast("✅ Excel gerado com sucesso!", "success");
  } catch (err) {
    console.error(err);
    showToast("❌ Erro ao gerar Excel", "error");
  }
}
