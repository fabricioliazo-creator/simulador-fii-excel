# 🏢 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Ferramenta interativa desenvolvida em Microsoft Excel para simulação de acúmulo de património, projeção de rendimentos mensais (dividendos) e alocação estratégica por perfil de investidor em Fundos Imobiliários (FIIs).

---

## 🎯 Perguntas de Negócio Respondidas

A ferramenta responde de forma dinâmica às 5 perguntas essenciais de planeamento financeiro:

1. **Quanto investir por mês?** (Célula / Intervalo Nomeado: `aporte`)
2. **Por quantos anos?** (Célula / Intervalo Nomeado: `qtd_anos`)
3. **Qual a taxa de rendimento mensal?** (Célula / Intervalo Nomeado: `taxa_mensal`)
4. **Quanto de património vai acumular?** (Célula / Intervalo Nomeado: `patrimonio` – calculado via função `VF`)
5. **Quanto vai receber de dividendos por mês?** (Calculado multiplicando o património acumulado pela taxa de rendimento estimada da carteira)

---

## 🛠️ Funções e Lógicas Aplicadas

### 1. Função `VF` (Valor Futuro / `FV`)
Utilizada para projetar o crescimento do capital acumulado com base em aportes mensais constantes sob taxa de juros compostos:
```excel
=VF(taxa_mensal; qtd_anos * 12; -aporte)

taxa_mensal: Taxa de rentabilidade esperada por período mensal.

qtd_anos * 12: Total de meses de aplicação.

-aporte: Valor investido mensalmente (com sinal negativo para indicar saída de caixa e retornar um saldo final positivo).

2. Chave Composta e PROCV (VLOOKUP)
Para alocar o aporte entre os seis tipos de FIIs (Papel, Tijolo, Híbridos, FOFs, Desenvolvimentos e Hotelarias) segundo o perfil selecionado:

Aba de Apoio (Apoio_Perfis): Criação de uma chave única combinando Perfil e Tipo de FII:
=B3 & "-" & C3   --> Resultado: "Conservador-PAPEL"

Busca Dinâmica:
=PROCV($C$32 & "-" & B36; Apoio_Perfis!$A:$D; 4; FALSO)
🏷️ Intervalos Nomeados Utilizados
Para facilitar a leitura e manutenção das fórmulas, foram definidos os seguintes intervalos nomeados:

salario: Salário base informado pelo utilizador.

sugestao_investimento: Sugestão automática de investimento (30% do salário).

aporte: Valor efetivamente investido por mês.

qtd_anos: Período do investimento em anos.

taxa_mensal: Taxa mensal de rendimento esperada para a fase de acúmulo.

rendimento_carteira: Taxa de dividend yield mensal esperada para a carteira de FIIs.

patrimonio: Património total acumulado no final do período.

📸 Evidências de Funcionamento
Perfil Conservador
Perfil Agressivo
Perfil Moderado
🚀 Destaques e Evoluções em Relação ao Modelo Base
Design Estilo Aplicação: Ocultação de linhas de grelha, codificação por cores diferenciando entradas de utilizador e células calculadas protegidas visualmente.

Tabela de Cenários Automatizada: Matriz de projeção para 2, 5, 10, 20 e 30 anos utilizando intervalos nomeados e referências absolutas/mistas.

Estruturação de Dados Limpa: Separação clara entre a aba de interface (Simulador) e a aba de regras de negócio (Apoio_Perfis).
