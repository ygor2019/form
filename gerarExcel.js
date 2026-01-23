// ============================================
// FUNÇÃO MELHORADA - Gerar Excel BONITO (2 colunas) e PADRONIZADO
// Layout: Coluna A (descrição) + Coluna B (resposta)
// Subcontratados: vertical, hierarquizado, sem “apagar” campos
// ============================================
async function gerarExcel() {
  showToast('📊 Gerando Excel formatado...', 'success');

  try {
    const wb = XLSX.utils.book_new();
    const ws = {};
    let maxRow = 0;

    // =========================
    // HELPERS BÁSICOS
    // =========================
    function safeStr(v) {
      // mantém vazio como vazio (não apaga linha/estrutura)
      if (v === null || v === undefined) return '';
      return String(v);
    }

    function setCell(cell, value, style = {}) {
      if (!ws[cell]) ws[cell] = {};
      ws[cell].v = safeStr(value);
      ws[cell].t = 's';

      const defaultStyle = {
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        },
        alignment: { vertical: 'center', wrapText: true }
      };

      ws[cell].s = { ...(ws[cell].s || {}), ...defaultStyle, ...style };

      const row = parseInt(cell.match(/\d+/)[0], 10);
      if (row > maxRow) maxRow = row;
    }

    ws['!merges'] = ws['!merges'] || [];

    // Mescla A:B na linha
    function mergeRowAB(row) {
      ws['!merges'].push({
        s: { r: row - 1, c: 0 }, // A
        e: { r: row - 1, c: 1 }  // B
      });
    }

    // =========================
    // ESTILOS
    // =========================
    const estilos = {
      titulo: {
        font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '1A365D' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
      },
      subtitulo: {
        font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '2C5282' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      },
      secao: {
        font: { bold: true, sz: 11, color: { rgb: '1A202C' } },
        fill: { fgColor: { rgb: 'E2E8F0' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      },
      label: {
        font: { bold: true, sz: 10, color: { rgb: '1A202C' } },
        fill: { fgColor: { rgb: 'EDF2F7' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      },
      item: {
        font: { bold: true, sz: 10, color: { rgb: '1A202C' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      },
      itemIndent1: {
        font: { sz: 10, color: { rgb: '1A202C' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true, indent: 1 }
      },
      itemIndent2: {
        font: { sz: 10, color: { rgb: '1A202C' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true, indent: 2 }
      },
      valor: {
        font: { sz: 10, color: { rgb: '1A202C' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      }
    };

    // =========================
    // COMPONENTES DE LINHA (2 COLUNAS)
    // =========================
    function addTitle(text, row) {
      setCell(`A${row}`, text, estilos.titulo);
      setCell(`B${row}`, '', estilos.titulo);
      mergeRowAB(row);
      return row + 2;
    }

    function addSubtitle(text, row) {
      setCell(`A${row}`, text, estilos.subtitulo);
      setCell(`B${row}`, '', estilos.subtitulo);
      mergeRowAB(row);
      return row + 1;
    }

    function addSection(text, row) {
      setCell(`A${row}`, text, estilos.secao);
      setCell(`B${row}`, '', estilos.secao);
      mergeRowAB(row);
      return row + 1;
    }

    function addRow(row, descricao, valor) {
      setCell(`A${row}`, descricao, estilos.label);
      setCell(`B${row}`, valor, estilos.valor);
      return row + 1;
    }

    function addItem(row, descricao, valor, indent = 0) {
      const styleDesc = indent === 0 ? estilos.item : (indent === 1 ? estilos.itemIndent1 : estilos.itemIndent2);
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
      nao_iniciada: 'Não iniciada',
      em_andamento: 'Em andamento',
      concluida: 'Concluída'
    };

    // Títulos dos grupos (9–24) (pra ficar bonito no Excel)
    const titulosSub = {
      9:  '📐 Subcontratados - Grupo 1 (Topografia, Sondagem, Movimento de Terra e Fundações)',
      10: '🏗️ Subcontratados - Grupo 2 (Estrutura, Concreto, Premoldados, Formas e Segurança)',
      11: '🏭 Subcontratados - Grupo 3 (Equipamentos, Geradores, Compressores e Resíduos)',
      12: '⚡ Subcontratados - Grupo 4 (Instalações Elétricas e Hidráulicas)',
      13: '🧯 Subcontratados - Grupo 5 (Segurança, AVCB, Laudos e Inspeções)',
      14: '🔥 Subcontratados - Grupo 6 (Gás, Incêndio, Energia Solar e SPDA)',
      15: '📞 Subcontratados - Grupo 7 (Telefonia, Segurança Eletrônica e Central de Concreto)',
      16: '🏊 Subcontratados - Grupo 8 (Elevadores de Carga, Piscinas e Sistemas de Água)',
      17: '🪟 Subcontratados - Grupo 9 (Fachada, Esquadrias Metálicas e Poço Artesiano)',
      18: '🧱 Subcontratados - Grupo 10 (Alvenaria, Revestimentos e Jardinagem)',
      19: '📋 Subcontratados - Grupo 11 (PGR, PGR Terceirizadas, PCMAT e LTCAT)',
      20: '🎨 Subcontratados - Grupo 12 (Drenagem, Extintores, Pintura e Instalações Gerais)',
      21: '🛣️ Subcontratados - Grupo 13 (Pavimentação e Estudos Ambientais)',
      22: '🔌 Subcontratados - Grupo 14 (Rede Elétrica, Canteiro de Obras e Viabilidade)',
      23: '🚦 Subcontratados - Grupo 15 (Sinalização, Tráfego e Controle Tecnológico)',
      24: '📝 Outras Atividades Técnicas'
    };

    // Pega dados de subcontratados de forma segura
    function getSubData(pageNum, id) {
      const pageKey = 'page' + pageNum;
      const pageObj = formData?.[pageKey] || {};
      return pageObj?.[id] || null; // cada id vira um objeto {opcao, profissional, ...}
    }

    function formatOpcao(itemData) {
      const op = itemData?.opcao || '';
      // status fica "nao_iniciada/em_andamento/concluida"
      if (statusLabel[op]) return statusLabel[op];
      return op; // Sim/Não/etc
    }

    // Escreve bloco padrão do profissional/ART/empresa (sempre embaixo, mesmo vazio)
    function escreverDetalhesProf(row, itemData, indentBase = 1) {
      // IMPORTANTE: mesmo vazio, escreve as linhas pra padronizar
      row = addItem(row, 'Profissional', itemData?.profissional || '', indentBase);
      row = addItem(row, 'CPF', itemData?.cpf || '', indentBase);
      row = addItem(row, 'Tipo de Registro', itemData?.tipoReg || '', indentBase);

      // CREA ou Outros (mantém ambos como linhas; se não usar, fica vazio)
      row = addItem(row, 'Nº CREA', itemData?.numCrea || '', indentBase);
      row = addItem(row, 'Sigla Conselho (Outros)', itemData?.sigla || '', indentBase);
      row = addItem(row, 'Nº Conselho (Outros)', itemData?.numOutros || '', indentBase);

      // ART ou Outros (mantém ambos como linhas)
      row = addItem(row, 'Tipo ART/Outros', itemData?.tipoArt || '', indentBase);
      row = addItem(row, 'Nº ART', itemData?.numArt || '', indentBase);
      row = addItem(row, 'Nome (Outros - ART)', itemData?.nomeArt || '', indentBase);
      row = addItem(row, 'Número (Outros - ART)', itemData?.valArt || '', indentBase);

      // Empresa
      row = addItem(row, 'Empresa', itemData?.empresa || '', indentBase);
      row = addItem(row, 'CNPJ', itemData?.cnpj || '', indentBase);

      // Registro empresa + placa (se existir)
      if ('regEmpresa' in (itemData || {}) || 'placa' in (itemData || {})) {
        row = addItem(row, 'Registro Empresa', itemData?.regEmpresa || '', indentBase);
        row = addItem(row, 'Placa/Equipamento', itemData?.placa || '', indentBase);
      }

      return row;
    }

    function escreverTerceirizadas(row, itemData, indentBase = 1) {
      // Sempre padronizado: escreve pelo menos 1 linha (mesmo sem nada)
      const lista = Array.isArray(itemData?.terceirizadas) ? itemData.terceirizadas : [];

      if (lista.length === 0) {
        row = addItem(row, 'Terceirizada 1', '', indentBase);
        row = addItem(row, 'ART 1', '', indentBase + 1);
        return row;
      }

      lista.forEach((t, idx) => {
        const n = idx + 1;
        row = addItem(row, `Terceirizada ${n}`, t?.terceirizada || '', indentBase);
        row = addItem(row, `ART ${n}`, t?.art || '', indentBase + 1);
      });

      return row;
    }

    // =========================
    // COMEÇA A ESCREVER O EXCEL
    // =========================
    let row = 1;

    row = addTitle('FORMULÁRIO - OBRAS DE MÉDIO E GRANDE PORTE', row);

    // -------------------------
    // PÁGINAS 1–8 (FORMULÁRIO PRINCIPAL)
    // -------------------------
    row = addSubtitle('A) LOCAL DA OBRA', row);
    row = addRow(row, 'Endereço da Obra', formData?.page1?.localObra || '');
    row = addRow(row, 'Bairro', formData?.page1?.bairro || '');
    row = addRow(row, 'Cidade', formData?.page1?.cidade || '');
    row = addRow(row, 'CEP', formData?.page1?.cep || '');
    row = addBlank(row, 1);

    row = addSubtitle('B.1) Empreendimento', row);
    row = addRow(row, 'Tipo', formData?.page2?.tipoNatureza || '');
    row = addRow(row, 'Área (m²)', formData?.page2?.area || '');
    row = addRow(row, 'Pavimentos', formData?.page2?.pavimentos || '');
    row = addRow(row, 'Nº do Processo', formData?.page2?.processoNum || '');
    row = addRow(row, 'Nº do Alvará', formData?.page2?.alvaraNum || '');
    row = addRow(row, 'Ano do Alvará', formData?.page2?.alvaraDe || '');
    row = addRow(row, 'Proprietário', formData?.page2?.proprietario || '');
    row = addRow(row, 'CPF ou CNPJ', (formData?.page2?.tipoCpfCnpj || '').toUpperCase());
    row = addRow(row, 'Número:', formData?.page2?.cpfCnpj || '');
    row = addRow(row, 'Endereço do Proprietário', formData?.page2?.enderecoProprietario || '');
    row = addBlank(row, 1);

    row = addSubtitle('3) RESPONSÁVEL TÉCNICO (RT)', row);
    row = addRow(row, 'Placa afixada', formData?.page3?.placaAfixada || '');
    row = addRow(row, 'Profissional', formData?.page3?.profissional || '');
    row = addRow(row, 'Tipo de Registro', formData?.page3?.tipoRegistro || '');
    row = addRow(row, 'Nº CREA', formData?.page3?.numeroCrea || '');
    row = addRow(row, 'CPF Profissional', formData?.page3?.cpfProfissional || '');
    row = addRow(row, 'Sigla Conselho (Outros)', formData?.page3?.siglaConselho || '');
    row = addRow(row, 'Nº Conselho (Outros)', formData?.page3?.numeroConselho || '');
    row = addRow(row, 'Título', formData?.page3?.titulo || '');
    row = addRow(row, 'Tipo ART/RRT', formData?.page3?.tipoArtRrt || '');
    row = addRow(row, 'Nº ART', formData?.page3?.artNumero || '');
    row = addRow(row, 'Nº RRT', formData?.page3?.rrtNumero || '');
    row = addRow(row, 'Empresa', formData?.page3?.empresa || '');
    row = addRow(row, 'Tipo Registro Empresa', formData?.page3?.tipoEmpresaRegistro || '');
    row = addRow(row, 'Registro Empresa', formData?.page3?.empresaRegistro || '');
    row = addBlank(row, 1);

    row = addSubtitle('4) RESPONSÁVEL TÉCNICO (RT) - COMPLEMENTO', row);
    row = addRow(row, 'Placa afixada', formData?.page4?.placaAfixadaRT || '');
    row = addRow(row, 'Profissional', formData?.page4?.profissionalRT || '');
    row = addRow(row, 'CPF', formData?.page4?.cpfRT || '');
    row = addRow(row, 'Tipo de Registro', formData?.page4?.tipoRegistroRT || '');
    row = addRow(row, 'Nº CREA', formData?.page4?.numeroCreaRT || '');
    row = addRow(row, 'Sigla Conselho (Outros)', formData?.page4?.siglaConselhoRT || '');
    row = addRow(row, 'Nº Conselho (Outros)', formData?.page4?.numeroConselhoRT || '');
    row = addRow(row, 'Tipo ART/Outros', formData?.page4?.tipoArtOutrosRT || '');
    row = addRow(row, 'Nº ART', formData?.page4?.artNumeroRT || '');
    row = addRow(row, 'Nome (Outros - ART)', formData?.page4?.nomeOutrosArtRT || '');
    row = addRow(row, 'Número (Outros - ART)', formData?.page4?.numeroOutrosArtRT || '');
    row = addRow(row, 'Empresa', formData?.page4?.empresaRT || '');
    row = addRow(row, 'Tipo Registro Empresa', formData?.page4?.tipoEmpresaRegistroRT || '');
    row = addRow(row, 'Registro Empresa', formData?.page4?.empresaRegistroRT || '');
    row = addBlank(row, 1);

    row = addSubtitle('5) DEMOLIÇÃO', row);
    row = addRow(row, 'Houve Demolição?', formData?.page5?.houveDemolicao || '');
    row = addRow(row, 'Projeto de Demolição?', formData?.page5?.projetoDemolicao || '');
    row = addSection('PROJETO / ASSESSORIA / CONSULTORIA (DEMOLIÇÃO)', row);
    row = addRow(row, 'Profissional', formData?.page5?.profissionalProjetoDemol || '');
    row = addRow(row, 'CPF', formData?.page5?.cpfProjetoDemol || '');
    row = addRow(row, 'Tipo Registro', formData?.page5?.tipoRegistroProjetoDemol || '');
    row = addRow(row, 'Nº CREA', formData?.page5?.numeroCreaProjetoDemol || '');
    row = addRow(row, 'Sigla Conselho', formData?.page5?.siglaConselhoProjetoDemol || '');
    row = addRow(row, 'Nº Conselho', formData?.page5?.numeroConselhoProjetoDemol || '');
    row = addRow(row, 'Tipo ART/Outros', formData?.page5?.tipoArtProjetoDemol || '');
    row = addRow(row, 'Nº ART', formData?.page5?.artNumeroProjetoDemol || '');
    row = addRow(row, 'Nome (Outros)', formData?.page5?.nomeOutrosArtProjetoDemol || '');
    row = addRow(row, 'Número (Outros)', formData?.page5?.numeroOutrosArtProjetoDemol || '');
    row = addSection('EXECUÇÃO (DEMOLIÇÃO)', row);
    row = addRow(row, 'Execução', formData?.page5?.execucaoDemolicao || '');
    row = addRow(row, 'Profissional', formData?.page5?.profissionalExecDemol || '');
    row = addRow(row, 'CPF', formData?.page5?.cpfExecDemol || '');
    row = addRow(row, 'Tipo Registro', formData?.page5?.tipoRegistroExecDemol || '');
    row = addRow(row, 'Nº CREA', formData?.page5?.numeroCreaExecDemol || '');
    row = addRow(row, 'Sigla Conselho', formData?.page5?.siglaConselhoExecDemol || '');
    row = addRow(row, 'Nº Conselho', formData?.page5?.numeroConselhoExecDemol || '');
    row = addRow(row, 'Tipo ART/Outros', formData?.page5?.tipoArtExecDemol || '');
    row = addRow(row, 'Nº ART', formData?.page5?.artNumeroExecDemol || '');
    row = addRow(row, 'Nome (Outros)', formData?.page5?.nomeOutrosArtExecDemol || '');
    row = addRow(row, 'Número (Outros)', formData?.page5?.numeroOutrosArtExecDemol || '');
    row = addRow(row, 'Empresa', formData?.page5?.empresaDemol || '');
    row = addRow(row, 'CNPJ', formData?.page5?.cnpjDemol || '');
    row = addBlank(row, 1);

    row = addSubtitle('6) COORDENAÇÃO / ORÇAMENTO / RESIDENTE', row);
    row = addRow(row, 'Houve Coordenação?', formData?.page6?.houveCoordenacao || '');
    row = addRow(row, 'Houve Orçamento?', formData?.page6?.houveOrcamento || '');
    row = addRow(row, 'Houve Residente?', formData?.page6?.houveResidente || '');

    function escreverEquipe(row, titulo, obj) {
      row = addSection(titulo, row);
      row = addRow(row, 'Nome', obj?.nome || '');
      row = addRow(row, 'CPF', obj?.cpf || '');
      row = addRow(row, 'Registro', obj?.reg || '');
      row = addRow(row, 'Nº CREA', obj?.numCrea || '');
      row = addRow(row, 'Sigla Conselho', obj?.sigla || '');
      row = addRow(row, 'Nº Conselho', obj?.numOutros || '');
      row = addRow(row, 'Tipo ART/Outros', obj?.tipoArt || '');
      row = addRow(row, 'Nº ART', obj?.numArt || '');
      row = addRow(row, 'Nome (Outros)', obj?.nomeArtOutros || '');
      row = addRow(row, 'Número (Outros)', obj?.numArtOutros || '');
      return row;
    }

    row = escreverEquipe(row, 'COORDENAÇÃO', formData?.page6?.coordenacao);
    row = escreverEquipe(row, 'ORÇAMENTO', formData?.page6?.orcamento);
    row = escreverEquipe(row, 'RESIDENTE', formData?.page6?.residente);
    row = addBlank(row, 1);

    row = addSubtitle('7) LIVRO DE ORDEM', row);
    row = addRow(row, 'Livro de Ordem', formData?.page7?.livroOrdem || '');
    row = addRow(row, 'Observações', formData?.page7?.observacoes || '');
    row = addBlank(row, 1);

    row = addSubtitle('8) GERENCIAMENTO', row);
    row = addRow(row, 'Houve Gerenciamento?', formData?.page8?.houveGerenciamento || '');
    row = addRow(row, 'Profissional', formData?.page8?.profissionalGerenciamento || '');
    row = addRow(row, 'CPF', formData?.page8?.cpfGerenciamento || '');
    row = addRow(row, 'Tipo Registro', formData?.page8?.tipoRegistroGerenciamento || '');
    row = addRow(row, 'Nº CREA', formData?.page8?.numeroCreaGerenciamento || '');
    row = addRow(row, 'Sigla Conselho', formData?.page8?.siglaConselhoGerenciamento || '');
    row = addRow(row, 'Nº Conselho', formData?.page8?.numeroConselhoGerenciamento || '');
    row = addRow(row, 'Tipo ART/Outros', formData?.page8?.tipoArtGerenciamento || '');
    row = addRow(row, 'Nº ART', formData?.page8?.artNumeroGerenciamento || '');
    row = addRow(row, 'Nome (Outros)', formData?.page8?.nomeOutrosArtGerenciamento || '');
    row = addRow(row, 'Número (Outros)', formData?.page8?.numeroOutrosArtGerenciamento || '');
    row = addRow(row, 'Empresa', formData?.page8?.empresaGerenciamento || '');
    row = addRow(row, 'CNPJ', formData?.page8?.cnpjGerenciamento || '');
    row = addRow(row, 'Registro Empresa (Tipo)', formData?.page8?.tipoRegistroEmpresaGer || '');
    row = addRow(row, 'Registro Empresa (Nº CREA)', formData?.page8?.numeroCreaEmpresaGer || '');
    row = addRow(row, 'Registro Empresa (Sigla)', formData?.page8?.siglaConselhoEmpresaGer || '');
    row = addRow(row, 'Registro Empresa (Nº Conselho)', formData?.page8?.numeroConselhoEmpresaGer || '');
    row = addRow(row, 'Placa / Equipamento', formData?.page8?.placaGerenciamento || '');
    row = addBlank(row, 2);

    // -------------------------
    // PÁGINAS 9–24 (SUBCONTRATADOS)
    // -------------------------
    row = addTitle('SUBCONTRATADOS (PADRÃO VERTICAL / PADRONIZADO)', row);

    for (let pageNum = 9; pageNum <= 24; pageNum++) {
      const pageKey = 'page' + pageNum;
      const items = subcontratados?.[pageKey];
      if (!items || items.length === 0) continue;

      row = addSubtitle(titulosSub[pageNum] || `Subcontratados - Página ${pageNum}`, row);

      items.forEach(item => {
        // Item pode ter subitens (statusComSubitens)
        const renderItem = (it, pageNumLocal, indent = 0) => {
          // pega dados salvos
          const data = getSubData(pageNumLocal, it.id);

          // Linha principal do item
          const label = `${it.num ? it.num + ' - ' : ''}${it.titulo || it.id}`;
          row = addItem(row, label, formatOpcao(data), indent);

          // TERCEIRIZADAS: lista dinâmica
          if (it.tipo === 'terceirizadas') {
            // Sempre imprime o bloco (mesmo se NÃO ou vazio)
            row = escreverTerceirizadas(row, data, indent + 1);
            row = addBlank(row, 1);
            return;
          }

          // Para os demais tipos: sempre imprime detalhes (mesmo se NÃO ou vazio)
          row = escreverDetalhesProf(row, data, indent + 1);

          // Se tiver subitens, renderiza abaixo (com indent)
          if (Array.isArray(it.subitens) && it.subitens.length > 0) {
            it.subitens.forEach(sub => renderItem(sub, pageNumLocal, indent + 1));
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
    row = addTitle('IDENTIFICAÇÃO FINAL', row);

    row = addSubtitle('IDENTIFICAÇÃO DO DECLARANTE', row);
    row = addRow(row, 'Nome', formData?.page25?.nomeDeclarante || '');
    row = addRow(row, 'Qualificação e Cargo', formData?.page25?.qualificacaoDeclarante || '');
    row = addRow(row, 'Local', formData?.page25?.localDeclarante || '');
    // Usa sua função formatarData(dataISO) já existente no arquivo
    row = addRow(row, 'Data', (typeof formatarData === 'function') ? formatarData(formData?.page25?.dataDeclarante || '') : (formData?.page25?.dataDeclarante || ''));
    row = addBlank(row, 1);

    row = addSubtitle('AGENTE FISCAL', row);
    row = addRow(row, 'Agente Fiscal', formData?.page25?.agenteFiscal1 || '');
    row = addRow(row, 'Registro', formData?.page25?.registroAgente1 || '');
    row = addRow(row, 'Data Fiscalização', (typeof formatarData === 'function') ? formatarData(formData?.page25?.dataFiscalizacao || '') : (formData?.page25?.dataFiscalizacao || ''));
    row = addBlank(row, 1);

    row = addSubtitle('CONTATO', row);
    row = addRow(row, 'Endereço', formData?.page25?.enderecoContato || '');
    row = addRow(row, 'E-mail', formData?.page25?.emailContato || '');
    row = addRow(row, 'Telefones', formData?.page25?.telefonesContato || '');
    row = addBlank(row, 1);

    // =========================
    // CONFIGURAÇÃO DE LARGURA / RANGE (SÓ A:B)
    // =========================
    ws['!cols'] = [
      { wch: 80 }, // A
      { wch: 45 }  // B
    ];

    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: 1 } // até coluna B
    });

    XLSX.utils.book_append_sheet(wb, ws, 'Formulário');

    const agora = new Date();
    const yyyy = agora.getFullYear();
    const mm = String(agora.getMonth() + 1).padStart(2, '0');
    const dd = String(agora.getDate()).padStart(2, '0');
    const nomeArquivo = `Formulario_Obra_${yyyy}-${mm}-${dd}.xlsx`;

    XLSX.writeFile(wb, nomeArquivo);

    showToast('✅ Excel gerado com sucesso!', 'success');

  } catch (err) {
    console.error(err);
    showToast('❌ Erro ao gerar Excel', 'error');
  }
}
