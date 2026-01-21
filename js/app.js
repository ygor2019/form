// Estado Global
let formData = {
  page1: { localObra: '', bairro: '', cidade: '', cep: '', saved: false },
  page2: { tipoNatureza: '', area: '', pavimentos: '', processoNum: '', alvaraNum: '', alvaraDe: '', proprietario: '', tipoCpfCnpj: 'cpf', cpfCnpj: '', enderecoProprietario: '', saved: false },
  page3: { placaAfixada: '', profissional: '', tipoRegistro: '', numeroCrea: '', cpfProfissional: '', siglaConselho: '', numeroConselho: '', titulo: '', tipoArtRrt: '', artNumero: '', rrtNumero: '', empresa: '', tipoEmpresaRegistro: '', empresaRegistro: '', saved: false },
  page4: { placaAfixadaRT: '', profissionalRT: '', cpfRT: '', tipoRegistroRT: '', numeroCreaRT: '', siglaConselhoRT: '', numeroConselhoRT: '', tipoArtOutrosRT: '', artNumeroRT: '', nomeOutrosArtRT: '', numeroOutrosArtRT: '', empresaRT: '', tipoEmpresaRegistroRT: '', empresaRegistroRT: '', saved: false },
  page5: { houveDemolicao: '', projetoDemolicao: '', profissionalProjetoDemol: '', cpfProjetoDemol: '', tipoRegistroProjetoDemol: '', numeroCreaProjetoDemol: '', siglaConselhoProjetoDemol: '', numeroConselhoProjetoDemol: '', tipoArtProjetoDemol: '', artNumeroProjetoDemol: '', nomeOutrosArtProjetoDemol: '', numeroOutrosArtProjetoDemol: '', execucaoDemolicao: '', profissionalExecDemol: '', cpfExecDemol: '', tipoRegistroExecDemol: '', numeroCreaExecDemol: '', siglaConselhoExecDemol: '', numeroConselhoExecDemol: '', tipoArtExecDemol: '', artNumeroExecDemol: '', nomeOutrosArtExecDemol: '', numeroOutrosArtExecDemol: '', empresaDemol: '', cnpjDemol: '', saved: false },
  page6: { coordenacao: {}, orcamento: {}, residente: {}, saved: false }
};

let currentPage = 1;
const totalPages = 6;
let toastTimeout = null;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
  Autocomplete.init();
  loadCache();
  updateUI();
  setupEvents();
});

function setupEvents() {
  // Mascaras Automáticas
  const masks = [
    { id: 'cep', type: 'cep' },
    { id: 'area', type: 'currency' },
    { id: 'pavimentos', type: 'number' },
    { id: 'alvaraDe', type: 'number' },
    { id: 'cpfCnpj', type: (el) => document.getElementById('btnCpf').classList.contains('active') ? 'cpf' : 'cnpj' },
    { id: 'cpfProfissional', type: 'cpf' },
    { id: 'cpfRT', type: 'cpf' },
    { id: 'cpfProjetoDemol', type: 'cpf' },
    { id: 'cpfExecDemol', type: 'cpf' },
    { id: 'cnpjDemol', type: 'cnpj' },
    { id: 'cpfCoordenacao', type: 'cpf' },
    { id: 'cpfOrcamento', type: 'cpf' },
    { id: 'cpfResidente', type: 'cpf' }
  ];

  masks.forEach(m => {
    const el = document.getElementById(m.id);
    if (!el) return;
    el.addEventListener('input', (e) => {
      const type = typeof m.type === 'function' ? m.type(el) : m.type;
      e.target.value = Mask[type](e.target.value);
      el.classList.remove('error');
      const err = document.getElementById('error' + m.id.charAt(0).toUpperCase() + m.id.slice(1));
      if (err) err.classList.remove('show');
    });
  });

  // Autocompletes
  Autocomplete.setup('profissional', 'autocompleteProfissional', p => fillProf('page3', p));
  Autocomplete.setup('profissionalRT', 'autocompleteProfissionalRT', p => fillProf('page4', p));
  Autocomplete.setup('profissionalProjetoDemol', 'autocompleteProfissionalProjetoDemol', p => fillProf('page5p', p));
  Autocomplete.setup('profissionalExecDemol', 'autocompleteProfissionalExecDemol', p => fillProf('page5e', p));
  Autocomplete.setup('profCoordenacao', 'autocompleteCoordenacao', p => fillProf('page6c', p));
  Autocomplete.setup('profOrcamento', 'autocompleteOrcamento', p => fillProf('page6o', p));
  Autocomplete.setup('profResidente', 'autocompleteResidente', p => fillProf('page6r', p));
}

// Auxiliares de UI
function showToast(msg, type = 'success', duration = 3000) {
  if (toastTimeout) clearTimeout(toastTimeout);
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `toast ${type} show`;
  toastTimeout = setTimeout(() => t.classList.remove('show'), duration);
}

function updateUI() {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById(`page${currentPage}`).classList.add('active');
  
  const fill = document.getElementById('progressFill');
  const step = document.getElementById('progressStep');
  fill.style.width = `${(currentPage / totalPages) * 100}%`;
  step.textContent = `Página ${currentPage} de ${totalPages}`;
  
  window.scrollTo(0, 0);
  if (toastTimeout) { clearTimeout(toastTimeout); document.getElementById('toast').classList.remove('show'); }
}

function loadCache() {
  const cached = localStorage.getItem('formDataObra');
  if (cached) {
    formData = JSON.parse(cached);
    // Aqui viria a lógica de repopular campos (pode ser otimizada com [name])
    showToast('Dados restaurados!');
  }
}

// Funções de Navegação
function avancar(page) {
  if (!formData[`page${page}`]?.saved) return showToast('Salve antes de avançar!', 'error');
  if (currentPage < totalPages) { currentPage++; updateUI(); }
}

function retornar(page) {
  if (currentPage > 1) { currentPage--; updateUI(); }
}
