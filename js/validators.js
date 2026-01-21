// Validadores Globais
const Validator = {
  // Validar CPF
  cpf: function(cpf) {
    cpf = cpf.replace(/\D/g, '');
    if (cpf.length !== 11 || /^(\d)\1+$/.test(cpf)) return false;
    let soma = 0;
    for (let i = 0; i < 9; i++) soma += parseInt(cpf.charAt(i)) * (10 - i);
    let resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    if (resto !== parseInt(cpf.charAt(9))) return false;
    soma = 0;
    for (let i = 0; i < 10; i++) soma += parseInt(cpf.charAt(i)) * (11 - i);
    resto = (soma * 10) % 11;
    if (resto === 10 || resto === 11) resto = 0;
    return resto === parseInt(cpf.charAt(10));
  },

  // Validar CNPJ
  cnpj: function(cnpj) {
    cnpj = cnpj.replace(/\D/g, '');
    if (cnpj.length !== 14 || /^(\d)\1+$/.test(cnpj)) return false;
    let t = cnpj.length - 2, n = cnpj.substring(0, t), d = cnpj.substring(t), s = 0, p = t - 7;
    for (let i = t; i >= 1; i--) { s += n.charAt(t - i) * p--; if (p < 2) p = 9; }
    let r = s % 11 < 2 ? 0 : 11 - (s % 11);
    if (r != d.charAt(0)) return false;
    t++, n = cnpj.substring(0, t), s = 0, p = t - 7;
    for (let i = t; i >= 1; i--) { s += n.charAt(t - i) * p--; if (p < 2) p = 9; }
    r = s % 11 < 2 ? 0 : 11 - (s % 11);
    return r == d.charAt(1);
  },

  // Validar CEP de São Paulo Capital
  cepSP: function(cep, cidade) {
    const numeros = cep.replace(/\D/g, '');
    if (numeros.length !== 8) return { valid: false, msg: 'CEP deve ter 8 dígitos' };
    if (cidade.toLowerCase().includes('são paulo') || cidade.toLowerCase().includes('sao paulo')) {
      const prefixo = parseInt(numeros.substring(0, 2));
      const isSP = (prefixo >= 1 && prefixo <= 5) || (prefixo === 8);
      if (!isSP) return { valid: false, msg: 'CEP não corresponde à SP Capital (01-05 ou 08)' };
    }
    return { valid: true };
  },

  // Validar Ano
  ano: function(valor) {
    const ano = parseInt(valor);
    return valor.length === 4 && !isNaN(ano) && ano >= 1900 && ano <= 2100;
  },

  // Campo obrigatório simples
  required: function(input, errorEl) {
    if (!input.value.trim()) {
      input.classList.add('error');
      if (errorEl) errorEl.classList.add('show');
      return false;
    }
    input.classList.remove('error');
    input.classList.add('success');
    if (errorEl) errorEl.classList.remove('show');
    return true;
  }
};

// Mascaras Globais
const Mask = {
  cpf: (v) => {
    v = v.replace(/\D/g, '');
    if (v.length > 3) v = v.substring(0,3) + '.' + v.substring(3);
    if (v.length > 7) v = v.substring(0,7) + '.' + v.substring(7);
    if (v.length > 11) v = v.substring(0,11) + '-' + v.substring(11,13);
    return v;
  },
  cnpj: (v) => {
    v = v.replace(/\D/g, '');
    if (v.length > 2) v = v.substring(0,2) + '.' + v.substring(2);
    if (v.length > 6) v = v.substring(0,6) + '.' + v.substring(6);
    if (v.length > 10) v = v.substring(0,10) + '/' + v.substring(10);
    if (v.length > 15) v = v.substring(0,15) + '-' + v.substring(15,17);
    return v;
  },
  cep: (v) => {
    v = v.replace(/\D/g, '');
    if (v.length > 5) v = v.substring(0,5) + '-' + v.substring(5,8);
    return v;
  },
  number: (v) => v.replace(/\D/g, ''),
  currency: (v) => {
    v = v.replace(/\D/g, '');
    return v ? parseInt(v).toLocaleString('pt-BR') : '';
  }
};
