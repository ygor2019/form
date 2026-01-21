const Autocomplete = {
  lista: [],

  init: function() {
    this.carregar();
  },

  carregar: function() {
    const cached = localStorage.getItem('profissionaisCadastrados');
    if (cached) this.lista = JSON.parse(cached);
  },

  salvar: function() {
    localStorage.setItem('profissionaisCadastrados', JSON.stringify(this.lista));
  },

  add: function(nome, cpf, crea) {
    if (!nome?.trim()) return;
    const idx = this.lista.findIndex(p => p.nome.toLowerCase() === nome.toLowerCase());
    if (idx !== -1) {
      if (cpf) this.lista[idx].cpf = cpf;
      if (crea) this.lista[idx].crea = crea;
    } else {
      this.lista.push({ nome: nome.trim(), cpf: cpf || '', crea: crea || '' });
    }
    this.salvar();
  },

  setup: function(inputId, listId, callback) {
    const input = document.getElementById(inputId);
    const list = document.getElementById(listId);
    if (!input || !list) return;

    input.addEventListener('input', () => {
      const val = input.value.toLowerCase().trim();
      list.innerHTML = '';
      if (val.length < 2) { list.classList.remove('show'); return; }

      const filtrados = this.lista.filter(p => p.nome.toLowerCase().includes(val));
      if (filtrados.length === 0) { list.classList.remove('show'); return; }

      filtrados.forEach(p => {
        const item = document.createElement('div');
        item.className = 'autocomplete-item';
        item.innerHTML = `<strong>${p.nome}</strong><small>CPF: ${p.cpf || 'N/A'} ${p.crea ? '| CREA: '+p.crea : ''}</small>`;
        item.onclick = () => {
          input.value = p.nome;
          list.classList.remove('show');
          if (callback) callback(p);
        };
        list.appendChild(item);
      });
      list.classList.add('show');
    });

    document.addEventListener('click', (e) => {
      if (e.target !== input && !list.contains(e.target)) list.classList.remove('show');
    });
  }
};
