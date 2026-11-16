// Função para trocar de abas
function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.getElementById('tab-' + tabId).classList.add('active');
}

// Lógica de Salvar Dados (LocalStorage)
const formEmprestimo = document.getElementById('formEmprestimo');

formEmprestimo.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const registro = {
        aluno: document.getElementById('nomeAluno').value,
        turma: document.getElementById('turmaSelector').value,
        livro: document.getElementById('nomeLivro').value,
        saida: document.getElementById('dataInicio').value,
        devolucao: document.getElementById('dataFim').value
    };

    // Pegar o que já existe no "banco"
    let banco = JSON.parse(localStorage.getItem('biblioteca_etegpc')) || [];
    banco.push(registro);
    
    // Salvar de volta
    localStorage.setItem('biblioteca_etegpc', JSON.stringify(banco));
    
    alert('Registrado com sucesso!');
    formEmprestimo.reset();
});