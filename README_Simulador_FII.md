# Simulador de Investimentos em Fundos Imobiliários (FIIs)

Este projeto consiste em uma ferramenta prática e interativa desenvolvida no Microsoft Excel para simulação de investimentos em Fundos de Investimento Imobiliários (FIIs). A solução foi projetada para responder a perguntas de negócio cruciais de investidores e ajudar na tomada de decisões informadas. 

Este repositório contém a documentação completa e o passo a passo lógico para a construção e entrega do simulador.

---

## 🎯 Perguntas de Negócio Respondidas
Para garantir que a ferramenta seja altamente funcional e resolva problemas reais, ela foi estruturada em torno de cinco perguntas fundamentais:
1. **Quanto devo investir por mês?** (Aporte Mensal)
2. **Por quantos anos devo investir?** (Período de Acúmulo)
3. **Qual é a taxa de rendimento mensal esperada?**
4. **Quanto de patrimônio acumulado terei no final do período?**
5. **Qual será o meu rendimento mensal em dividendos (renda passiva)?**

---

## 🛠️ Funcionalidades da Planilha

### 1. Simulador de Patrimônio Individual
Permite ao usuário testar livremente suas projeções financeiras alterando variáveis básicas
* **Aporte Mensal (R$)**: O valor poupado e investido recorrentemente.
* **Tempo (Anos)**: Prazo da simulação.
* **Taxa Mensal (%)**: Taxa de rendimento estimada da carteira de FIIs.
* **Patrimônio Acumulado**: Calculado automaticamente usando matemática financeira.
* **Dividendos Estimados**: Projeção do fluxo de caixa mensal gerado pelo patrimônio acumulado.

### 2. Simulador de Cenários de Longo Prazo
Gera automaticamente cenários comparativos para horizontes de tempo de **2, 5, 10, 20 e 30 anos**. Isso poupa o usuário de ter que alterar manualmente o prazo para visualizar o poder dos juros compostos ao longo das décadas.

### 3. Painel de Configurações e Variáveis Globais
Um espaço dedicado para armazenar premissas que afetam múltiplos cálculos da ferramenta:
* **Salário**: Renda do usuário utilizada como base.
* **Rendimento da Carteira**: Taxa padrão de dividendos (ex: 1,00% ou 0,90%).
* **Sugestão de Investimento (Aporte ideal)**: Calculado automaticamente em **30% do salário** do usuário.

### 4. Distribuição da Carteira por Perfil de Investidor
Simulador de diversificação que sugere a alocação do aporte mensal em diferentes classes de FIIs com base no perfil do usuário:
* **Perfis**: Conservador, Moderado e Agressivo (selecionáveis via lista suspensa).
* **Tipos de FIIs Mapeados**: Tijolo, Papel (Dívidas), Híbridos, FOFs (Fundos de Fundos), Desenvolvimento e Hotelaria.

---

## 📐 Passo a Passo de Implementação Técnica

### Passo 1: Preparação do Ambiente e Interface Visual (UI/UX)
Para transformar uma simples planilha em uma aplicação com aspecto profissional e moderno:
1. **Ocultar as Linhas de Grade**: Acesse a guia **Exibir** > desmarque **Linhas de Grade**.
2. **Inserção do Banner**: Insira uma imagem de cabeçalho (banner ilustrativo) no topo da planilha.
3. **Propriedades da Imagem**: Clique com o botão direito na imagem > **Tamanho e Propriedades** > Guia **Propriedades** > marque **"Não mover ou dimensionar com células"** (isso evita distorções ao ajustar linhas/colunas).
4. **Ocultar Colunas Sobressalentes**: Selecione a primeira coluna livre à direita do simulador (ex: coluna `L`), pressione `Ctrl + Shift + Seta para a Direita`, clique com o botão direito e selecione **Ocultar**. Repita o processo para as linhas excedentes na parte inferior.

### Passo 2: Construção da Tabela de Simulação de Patrimônio
1. Crie um cabeçalho mesclando as colunas `B` e `C` na linha 12, pinte com um tom de verde moderno, configure a fonte como branca, tamanho 20, negrito e alinhada ao meio-esquerda.
2. Insira as linhas com os campos: *Aporte Mensal*, *Anos*, *Taxa Mensal*, *Patrimônio Acumulado* e *Dividendo Mensal*.
3. **Formatação de Células de Entrada vs. Calculadas**:
   * Pinte de **cinza bem claro** as células calculadas (*Patrimônio Acumulado* e *Dividendo Mensal*) para sinalizar visualmente ao usuário que são campos automáticos e bloqueados para edição.
   * Deixe os valores da coluna de dados em **negrito** para destaque visual.
4. **Formatação de Bordas Personalizadas**:
   * Aplique uma **borda externa espessa** ao redor da tabela.
   * Nas divisões internas, use **mais bordas** e escolha um tom de **cinza muito claro** com linhas finas (estilo cruz) para não carregar visualmente a tabela.
5. **Alinhamento e Recuo**: Alinhe o texto à esquerda. Para evitar que grude nas bordas da célula, aplique **Aumentar Recuo** de 1 a 3 vezes.
6. **Formato Moeda**: Formate todos os valores financeiros como **Moeda** (em vez de Contábil). O formato Moeda posiciona o símbolo `R$` colado ao valor e permite o alinhamento correto à esquerda.

### Passo 3: Implementação das Fórmulas Financeiras (Simulador)
O motor financeiro da planilha baseia-se na fórmula de juros compostos aplicada a aportes mensais recorrentes (renda fixa/variável constante).

1. **Patrimônio Acumulado (Valor Futuro)**:
   * No Excel, a fórmula utilizada é a `VF` (Valor Futuro).
   * Como a taxa de juros é mensal e o período inserido está em anos, o número de períodos (`nper`) deve ser multiplicado por 12.
   * Os aportes na fórmula do Excel são tratados como fluxo de saída (negativos). Para exibir o saldo acumulado como positivo, invertemos a polaridade multiplicando o resultado ou o pagamento por `-1`.
   * **Fórmula base**:
     ```excel
     =VF(taxa_mensal; quantidade_anos * 12; -aporte)
     ```
2. **Dividendo Mensal (Renda Passiva)**:
   * O rendimento do dividendo é calculado aplicando a taxa de rendimento sobre o patrimônio total acumulado ao final do período.
   * **Fórmula base**:
     ```excel
     =patrimônio * rendimento_carteira
     ```

### Passo 4: Nomeação de Intervalos para Legibilidade
À medida que a planilha cresce, referenciar células por coordenadas como `C14` ou `F14` torna as fórmulas confusas e difíceis de auditar. Atribuímos apelidos (intervalos nomeados) para simplificar a manutenção do projeto.

1. Selecione a célula desejada (ex: `C13`, onde está o valor de aporte).
2. No canto superior esquerdo (Caixa de Nome, ao lado da Barra de Fórmulas), digite o nome simplificado sem espaços nem acentos (use underline se necessário) e aperte `Enter`.
3. **Mapeamento de nomes recomendados**:
   * `C13` (Aporte Mensal) ➔ `aporte` 
   * `C14` (Quantidade de Anos) ➔ `quantidade_anos`
   * `C15` (Taxa Mensal) ➔ `taxa_mensal`
   * `C16` (Patrimônio Acumulado) ➔ `patrimônio`
   * `F14` (Rendimento da Carteira Global) ➔ `rendimento_carteira`
   * `F13` (Salário) ➔ `salário`
   * `F15` (Sugestão de Investimento) ➔ `sugestao_investimento`
4. **Refatoração das Fórmulas**:
   * Substitua as referências antigas nas fórmulas pelos nomes criados. Pressione `F3` dentro de uma fórmula para abrir o selecionador de nomes e colar a variável.
   * Nova fórmula do Patrimônio Acumulado:
     ```excel
     =VF(taxa_mensal; quantidade_anos * 12; -aporte)
     ```
   * Nova fórmula de Dividendos:
     ```excel
     =patrimônio * rendimento_carteira
     ```

### Passo 5: Criação do Simulador de Cenários de Longo Prazo
A tabela de cenários permite enxergar a evolução patrimonial em prazos fixos.
1. Crie uma tabela abaixo da principal com as colunas: *Cenários (Anos)*, *Patrimônio* e *Dividendo*.
2. Insira as faixas de anos de forma legível em texto (ex: `2 anos`, `5 anos`, `10 anos`, `20 anos`, `30 anos`).
3. **O Truque da Fonte Branca (Macete de UI)**: 
   * Para alimentar a fórmula de Valor Futuro de forma automatizada sem complicar a extração de texto, digite os números equivalentes em anos (`2`, `5`, `10`, `20`, `30`) em uma coluna oculta adjacente.
   * Pinte a fonte dessas células de **branco** . Assim, o valor fica invisível para o usuário, mas perfeitamente legível para as fórmulas do Excel.
4. **Cálculo da projeção dos cenários**:
   * Utilize a fórmula `VF` travando as células de aporte e taxa (usando `$`, ou chamando as variáveis nomeadas) e multiplicando o número de anos invisível por 12.
   * Arraste a fórmula para preencher todas as linhas utilizando a opção **Preencher Sem Formatação** para preservar as bordas da tabela.

### Passo 6: Criação das Variáveis Globais de Configuração
Adicione um painel lateral chamado `CONFIGURAÇÕES`.
1. **Salário**: Campo de digitação livre (ex: R$ 5.000,00).
2. **Rendimento Carteira**: Taxa padrão de mercado (ex: 1,00% ao mês).
3. **Sugestão de Investimento**: Fórmula que calcula 30% do salário, representando o percentual ideal de poupança mensal recomendado por especialistas.
   ```excel
   =salário * 30%
   ```

### Passo 7: Simulador de Perfil e Alocação Sugerida de FIIs
Para dar maior maturidade ao projeto, adicionamos uma seção que sugere onde alocar o dinheiro investido mensalmente com base no perfil de risco do usuário.

1. **Validação de Dados (Lista Suspensa)**:
   * Selecione a célula ao lado de "Perfil".
   * Vá em **Dados** > **Validação de Dados** > selecione **Lista** no campo "Permitir".
   * No campo "Fonte", digite os três perfis separados por ponto e vírgula: `Conservador;Moderado;Agressivo` e clique em OK.
2. **Distribuição Percentual de Alocação (Matriz Recomendada)**:
   * Crie uma tabela de alocação contendo as classes de FIIs: *Papel*, *Tijolo*, *Híbridos*, *FOFs*, *Desenvolvimento* e *Hotelaria*.
   * Sugestão de percentuais sugeridos com base no perfil (ex: Perfil Conservador):
     * **Tijolo**: 50% (Ativos físicos mais estáveis e seguros)
     * **Papel (Dívidas)**: 30%
     * **Híbridos**: 10%
     * **FOFs**: 10%
     * **Desenvolvimento**: 0% (Maior risco de execução de obra)
     * **Hotelaria**: 0%
3. **Cálculo da Alocação em Dinheiro**:
   * Multiplique o percentual sugerido de cada classe pelo valor de aporte mensal definido na simulação.
   * Congele a célula de aporte (com `F4` ou usando a variável nomeada `aporte`) para poder arrastar a fórmula verticalmente sem deslocar a referência.
   ```excel
   =percentual_sugerido * aporte
   ```
