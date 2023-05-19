var primeiraColuna = document.getElementById('DDD');
var segundaColuna = document.getElementById('AT');
var terceiraColuna = document.getElementById('PP');

primeiraColuna.addEventListener('change', function() {
  var valorPrimeiraColuna = primeiraColuna.value;
  if (valorPrimeiraColuna === 'x') {
    segundaColuna.innerHTML = '<option value="">-- atividade --</option><option value="1">1</option><option value="2">2</option><option value="3">3</option>';
  } else {
    // Se o valor selecionado na primeira coluna não for "x", então as opções da segunda coluna serão diferentes
    // Aqui você pode adicionar o código para definir as opções da segunda coluna com base no valor selecionado na primeira coluna
  }
});

segundaColuna.addEventListener('change', function() {
  var valorPrimeiraColuna = primeiraColuna.value;
  var valorSegundaColuna = segundaColuna.value;
  if (valorPrimeiraColuna === 'x') {
    // Se o valor selecionado na primeira coluna for "x", então as opções da terceira coluna serão diferentes
    // Aqui você pode adicionar o código para definir as opções da terceira coluna com base nos valores selecionados na primeira e segunda colunas
  } else {
    // Se o valor selecionado na primeira coluna não for "x", então as opções da terceira coluna serão diferentes
    // Aqui você pode adicionar o código para definir as opções da terceira coluna com base nos valores selecionados na primeira e segunda colunas
  }
});

terceiraColuna.addEventListener('change', function() {
  var valorPrimeiraColuna = primeiraColuna.value;
  var valorSegundaColuna = segundaColuna.value;
  var valorTerceiraColuna = terceiraColuna.value;
  if (valorPrimeiraColuna === 'x') {
    // Se o valor selecionado na primeira coluna for "x", então você pode adicionar o código aqui para verificar se o valor selecionado na terceira coluna é válido com base nos valores selecionados nas duas primeiras colunas
  } else {
    // Se o valor selecionado na primeira coluna não for "x", então você pode adicionar o código aqui para verificar se o valor selecionado na terceira coluna é válido com base nos valores selecionados nas duas primeiras colunas
  }
});