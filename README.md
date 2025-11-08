# Cálculo da taxa de desconto média ponderada e de taxa indicativa de operação

O objetivo deste arquivo Excel é efetuar o cálculo da taxa de desconto média ponderada das operações realizadas e apresentar uma taxa indicativa da próxima operação para que a carteira tenha a taxa de desconto mínima definida pelo usuário.
Adicionamente, ele calcula a taxa esperada, taxa da operada, taxa efetiva, receita total, receita esperada, pagamento total e prazo médio do consolidado de todas as operações e também por nível de risco por cedente definido pelo usuário.

## Entradas
- Para cada operação, o usuário deve colocar denominação do cedente, nível de risco do cedente, denominação do sacado, prazo médio em dias, tarifas e a taxa de desconto da operação. A taxa de desconto da operação a ser efetuada será sugerida automaticamente, mas ela pode ser editada.

## Funcionalidades
- Salva as operações em arquivo texto .csv.
- Calcula a receita da operação, receita esperada, taxa de desconto da operação, taxa esperada de todas as operações e por nível de risco de cedente.
- Calcula a taxa indicativa da operação a ser realizada.
- Comandos para proteger células selecionadas e desbloqueio de células bloqueadas.
- Manual do usuário.
- Conexão ao API do QProf (ainda em desenvolvimento).

## Como usar
1. Abra o arquivo `.xlsm` e ler a aba "Orientações". Nela, há instruções para habilitar as macros.
2. Aba "Descrição das fórmulas" descreve como a taxa indicativa é calculada e apresenta a explicação teórica dos cálculos envolvidos nas operações de desconto de duplicatas.

## Desenvolvido por
[Akihiko Sato]  
[LinkedIn](https://www.linkedin.com/in/akihikosato/)
